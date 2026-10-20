"""
Control de monotributo - lee las facturas PDF (QR de ARCA) desde Google Drive (o desde la carpeta
"facturas" del repositorio), actualiza los datos, genera el Excel y el RESUMEN.md, y avisa por mail
si alguien pasa a ALERTA/EXCEDIDO.
"""
import base64
import configparser
import io
import json
import os
import re
import shutil
import smtplib
import sys
import tempfile
from dataclasses import dataclass
from datetime import date, datetime
from email.message import EmailMessage
from pathlib import Path
from urllib.parse import parse_qs, urlparse

import cv2
import fitz  # PyMuPDF
import numpy as np
import pandas as pd

import planilla

BASE = Path(__file__).resolve().parent
DATOS = BASE / "datos"
ENTRADA, PROCESADAS, ERRORES = BASE / "facturas", BASE / "procesadas", BASE / "con_error"
F_TRAB, F_TOPES, F_FAC = DATOS / "trabajadores.csv", DATOS / "topes.csv", DATOS / "facturas.csv"
ESTADO = DATOS / "estado.json"
F_DRIVE = DATOS / "drive_procesados.json"
F_HIST = DATOS / "historico.csv"
EXCEL = BASE / "Control_Monotributo_SERVELEC.xlsx"
RESUMEN_MD = BASE / "RESUMEN.md"
LOG = BASE / "registro.log"
CONFIG = BASE / "config.ini"

TIPOS = {
    1: "Factura A", 2: "Nota de Débito A", 3: "Nota de Crédito A",
    6: "Factura B", 7: "Nota de Débito B", 8: "Nota de Crédito B",
    11: "Factura C", 12: "Nota de Débito C", 13: "Nota de Crédito C",
    51: "Factura M", 52: "Nota de Débito M", 53: "Nota de Crédito M",
    201: "Factura de Crédito MiPyME A", 202: "Nota de Débito MiPyME A", 203: "Nota de Crédito MiPyME A",
    206: "Factura de Crédito MiPyME B", 207: "Nota de Débito MiPyME B", 208: "Nota de Crédito MiPyME B",
    211: "Factura de Crédito MiPyME C", 212: "Nota de Débito MiPyME C", 213: "Nota de Crédito MiPyME C",
}
NOTAS_DE_CREDITO = {3, 8, 13, 53, 203, 208, 213}
RANGO = {"OK": 0, "FALTA TOPE": 0, "ALERTA": 1, "EXCEDIDO": 2}
COLUMNAS_FAC = ["clave", "fecha", "cuit", "comprobante", "pto_vta", "numero",
                "importe", "moneda", "archivo", "cargado"]


# ------------------------------------------------------------------ utilidades
def log(msg):
    print(msg)
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(f"{datetime.now():%Y-%m-%d %H:%M:%S}  {msg}\n")


def pesos(v):
    return "$ " + f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def cuit_fmt(c):
    s = str(int(c)).zfill(11)
    return f"{s[:2]}-{s[2:10]}-{s[10]}"


def numero(txt):
    """Acepta 12009410.45 / 12.009.410,45 / $ 12009410,45"""
    s = str(txt).replace("$", "").replace(" ", "").strip()
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    return float(s)


def hoy_argentina():
    try:
        from zoneinfo import ZoneInfo
        return datetime.now(ZoneInfo("America/Argentina/Buenos_Aires")).date()
    except Exception:
        return date.today()


def leer_config():
    cp = configparser.ConfigParser()
    cp.read(CONFIG, encoding="utf-8")
    g = cp["general"] if cp.has_section("general") else {}
    return {
        "empresa": g.get("empresa", "SERVELEC"),
        "umbral": float(g.get("umbral", "0.80")),
        "vigencia": g.get("topes_vigencia", ""),
        "mail": cp["email"] if cp.has_section("email") else {},
    }


# ------------------------------------------------------------------ datos (CSV)
def leer_csv(ruta, sniff=True):
    if not ruta.exists():
        return pd.DataFrame()
    kw = dict(sep=None, engine="python") if sniff else dict(sep=",")
    df = pd.read_csv(ruta, dtype=str, encoding="utf-8-sig", **kw).fillna("")
    df.columns = [c.strip().lower() for c in df.columns]
    return df


def cargar_trabajadores():
    df, out = leer_csv(F_TRAB), []
    for _, r in df.iterrows():
        cuit = re.sub(r"\D", "", r.get("cuit", ""))
        if cuit:
            out.append({"cuit": int(cuit), "nombre": r.get("nombre", "").strip(),
                        "cat": r.get("categoria", "").strip().upper()})
    return out


def cargar_topes():
    df, out = leer_csv(F_TOPES), []
    for _, r in df.iterrows():
        if r.get("categoria", "").strip() and r.get("tope", "").strip():
            out.append((r["categoria"].strip().upper(), numero(r["tope"])))
    return out


def cargar_facturas():
    df = leer_csv(F_FAC, sniff=False)
    if df.empty:
        return pd.DataFrame(columns=COLUMNAS_FAC)
    df["fecha"] = pd.to_datetime(df["fecha"]).dt.date
    df["cargado"] = pd.to_datetime(df["cargado"]).dt.date
    df["cuit"] = df["cuit"].astype("int64")
    df["pto_vta"] = df["pto_vta"].astype(int)
    df["numero"] = df["numero"].astype("int64")
    df["importe"] = df["importe"].astype(float)
    return df[COLUMNAS_FAC]


def cargar_historico():
    """Totales mensuales anteriores al uso de los PDF. Columnas: cuit, mes (AAAA-MM), importe."""
    df, out = leer_csv(F_HIST), []
    for _, r in df.iterrows():
        cuit = re.sub(r"\D", "", r.get("cuit", ""))
        if not cuit or not r.get("mes", "").strip() or not r.get("importe", "").strip():
            continue
        try:
            p = pd.Period(r["mes"].strip(), "M")
            importe = numero(r["importe"])
        except ValueError:
            print(f"[AVISO] historico.csv: fila ignorada (revisar mes o importe): {dict(r)}")
            continue
        out.append({"clave": f"HIST-{cuit}-{p}", "fecha": p.end_time.date(), "cuit": int(cuit),
                    "comprobante": "Histórico (total del mes)", "pto_vta": 0, "numero": 0,
                    "importe": importe, "moneda": "PES", "archivo": "historico.csv",
                    "cargado": hoy_argentina()})
    return pd.DataFrame(out, columns=COLUMNAS_FAC)


# ------------------------------------------------------------------ lectura del QR
def _decodificar_qr(gris):
    detector = cv2.QRCodeDetector()
    for img in (gris, cv2.threshold(gris, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]):
        texto, _, _ = detector.detectAndDecode(img)
        if texto:
            return texto
    return ""


def extraer_url_qr(pdf):
    doc = fitz.open(str(pdf))
    for pagina in list(doc)[:2]:
        for dpi in (150, 220, 300, 400):
            pix = pagina.get_pixmap(dpi=dpi)
            arr = np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.width, pix.n)
            if pix.n == 1:
                gris = arr[:, :, 0].copy()
            elif pix.n == 4:
                gris = cv2.cvtColor(arr, cv2.COLOR_RGBA2GRAY)
            else:
                gris = cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY)
            texto = _decodificar_qr(gris)
            if texto:
                return texto
    raise ValueError("no se encontró ningún código QR en el PDF")


def leer_comprobante(pdf, nombre=None):
    url = extraer_url_qr(pdf)
    p = parse_qs(urlparse(url).query).get("p")
    if not p:
        raise ValueError("el QR no es de un comprobante de ARCA")
    b64 = p[0].replace(" ", "+").replace("-", "+").replace("_", "/")
    d = json.loads(base64.b64decode(b64 + "=" * (-len(b64) % 4)))
    tipo, importe = int(d["tipoCmp"]), float(d["importe"])
    if str(d.get("moneda", "PES")).upper() != "PES":
        importe *= float(d.get("ctz", 1))
    if tipo in NOTAS_DE_CREDITO:
        importe = -abs(importe)
    return {
        "clave": f'{int(d["cuit"])}-{int(d["ptoVta"]):05d}-{tipo}-{int(d["nroCmp"])}',
        "fecha": date.fromisoformat(d["fecha"]), "cuit": int(d["cuit"]),
        "comprobante": TIPOS.get(tipo, f"Tipo {tipo}"), "pto_vta": int(d["ptoVta"]),
        "numero": int(d["nroCmp"]), "importe": round(importe, 2),
        "moneda": str(d.get("moneda", "PES")), "archivo": nombre or pdf.name, "cargado": hoy_argentina(),
    }


# ------------------------------------------------------------------ cálculos
def calcular(trab, topes_lista, fac, umbral, hoy):
    topes = dict(topes_lista)
    inicio = (pd.Timestamp(hoy.year, hoy.month, 1) - pd.DateOffset(months=11)).date()
    meses = list(pd.period_range(end=pd.Period(hoy, "M"), periods=12, freq="M"))
    en_periodo = fac[(fac["fecha"] >= inicio) & (fac["fecha"] <= hoy)].copy()
    en_periodo["mes"] = [pd.Period(f, "M") for f in en_periodo["fecha"]]
    hace90 = (pd.Timestamp(hoy) - pd.Timedelta(days=90)).date()

    estados, matriz = [], {}
    for t in trab:
        mias = en_periodo[en_periodo["cuit"] == t["cuit"]]
        matriz[t["cuit"]] = [round(float(mias.loc[mias["mes"] == m, "importe"].sum()), 2) for m in meses]
        total = round(float(mias["importe"].sum()), 2)
        tope = topes.get(t["cat"])
        uso = total / tope if tope else None
        disp = tope - total if tope else None
        estado = ("FALTA TOPE" if tope is None else "EXCEDIDO" if uso >= 1
                  else "ALERTA" if uso >= umbral else "OK")
        ritmo = round(float(mias.loc[mias["fecha"] > hace90, "importe"].sum()) / 3, 2)
        if tope is None or ritmo <= 0:
            meses_tope = None
        else:
            meses_tope = 0 if disp <= 0 else disp / ritmo
        estados.append({"cuit": t["cuit"], "nombre": t["nombre"], "cat": t["cat"], "tope": tope,
                        "total": total, "disp": disp, "uso": uso, "estado": estado, "ritmo": ritmo,
                        "meses_tope": meses_tope, "ultima": mias["fecha"].max() if len(mias) else None})
    estados.sort(key=lambda x: (-RANGO[x["estado"]], -(x["uso"] or 0)))
    return estados, matriz, inicio, meses


# ------------------------------------------------------------------ salidas
def escribir_resumen_md(cfg, hoy, inicio, estados, sin_registrar, n_errores):
    ico = {"OK": "🟢", "ALERTA": "🟡", "EXCEDIDO": "🔴", "FALTA TOPE": "⚪"}
    cuenta = lambda e: sum(1 for x in estados if x["estado"] == e)
    L = [f"# Control de Monotributo — {cfg['empresa']}", "",
         f"**Actualizado:** {hoy:%d/%m/%Y} · **Período:** {inicio:%m/%Y} a {hoy:%m/%Y}", "",
         f"🟢 En regla: **{cuenta('OK')}** · 🟡 En alerta: **{cuenta('ALERTA')}** · "
         f"🔴 Excedidos: **{cuenta('EXCEDIDO')}**", "",
         "| Estado | Trabajador | Cat. | Facturado | Tope | Uso | Disponible |", "|---|---|:-:|--:|--:|--:|--:|"]
    for x in estados:
        uso = f"{x['uso']:.1%}" if x["uso"] is not None else "-"
        tope = pesos(x["tope"]) if x["tope"] else "-"
        disp = pesos(x["disp"]) if x["disp"] is not None else "-"
        L.append(f"| {ico[x['estado']]} {x['estado']} | {x['nombre']} | {x['cat']} | "
                 f"{pesos(x['total'])} | {tope} | {uso} | {disp} |")
    if sin_registrar:
        L += ["", "⚠️ CUIT con facturas que no están en `datos/trabajadores.csv`: "
              + ", ".join(cuit_fmt(c) for c in sorted(sin_registrar))]
    if n_errores:
        L += ["", f"⚠️ {n_errores} PDF no se pudieron leer en la última ejecución (quedaron en `con_error`)."]
    L += ["", "_Detalle y gráficos en `Control_Monotributo_SERVELEC.xlsx`._"]
    RESUMEN_MD.write_text("\n".join(L) + "\n", encoding="utf-8")


def mostrar(estados):
    print("\n" + "=" * 78)
    print(f"{'Trabajador':<28}{'Cat':<5}{'Facturado':>20}{'Uso':>8}  Estado")
    print("-" * 78)
    for x in estados:
        uso = f"{x['uso']:.1%}" if x["uso"] is not None else "-"
        print(f"{x['nombre'][:27]:<28}{x['cat']:<5}{pesos(x['total']):>20}{uso:>8}  {x['estado']}")
    print("=" * 78)


def enviar_avisos(cfg, estados):
    m = cfg["mail"]
    if str(m.get("activar", "no")).strip().lower() not in ("si", "sí", "yes", "true", "1"):
        return
    usuario = os.environ.get("SMTP_USUARIO") or m.get("usuario", "")
    clave = os.environ.get("SMTP_CLAVE") or m.get("clave", "")
    destinos = os.environ.get("SMTP_DESTINATARIOS") or m.get("destinatarios", "")
    if not (usuario and clave and destinos):
        log("Avisos por mail activados pero faltan usuario, clave o destinatarios.")
        return
    previo = json.loads(ESTADO.read_text(encoding="utf-8")) if ESTADO.exists() else {}
    nuevos = [x for x in estados if RANGO[x["estado"]] > RANGO.get(previo.get(str(x["cuit"]), "OK"), 0)]
    if nuevos:
        cuerpo = ["Cambios detectados en el control de monotributo:\n"]
        for x in nuevos:
            cuerpo.append(f"- {x['nombre']} ({cuit_fmt(x['cuit'])}), categoría {x['cat']}: {x['estado']} - "
                          f"{pesos(x['total'])} de {pesos(x['tope'])} ({x['uso']:.1%})")
        msg = EmailMessage()
        msg["Subject"] = f"Monotributo: {len(nuevos)} trabajador(es) en alerta o excedidos"
        msg["From"], msg["To"] = usuario, destinos
        msg.set_content("\n".join(cuerpo))
        try:
            with smtplib.SMTP(m.get("servidor", "smtp.gmail.com"), int(m.get("puerto", "587"))) as s:
                s.starttls()
                s.login(usuario, clave)
                s.send_message(msg)
            log(f"Mail de aviso enviado ({len(nuevos)} trabajador/es).")
        except Exception as e:
            log(f"No se pudo enviar el mail: {e}")
            return
    # el estado guardado baja si mejora (así vuelve a avisar si empeora de nuevo)
    ESTADO.write_text(json.dumps({str(x["cuit"]): x["estado"] for x in estados}), encoding="utf-8")


# ------------------------------------------------------------------ de dónde salen los PDF
@dataclass
class Item:
    id: str
    nombre: str
    ruta: Path


def mover(archivo, carpeta):
    destino = carpeta / archivo.name
    if destino.exists():
        destino = carpeta / f"{archivo.stem}_{datetime.now():%H%M%S}{archivo.suffix}"
    shutil.move(str(archivo), str(destino))


class OrigenLocal:
    """PDF subidos a la carpeta 'facturas' del repositorio (alternativa a Drive)."""
    nombre = "carpeta facturas del repositorio"

    def pendientes(self):
        ENTRADA.mkdir(exist_ok=True)
        return [Item(str(p), p.name, p) for p in sorted(ENTRADA.iterdir()) if p.suffix.lower() == ".pdf"]

    def terminar(self, item, resultado):
        destino = ERRORES if resultado == "error" else PROCESADAS
        destino.mkdir(exist_ok=True)
        mover(item.ruta, destino)


class OrigenDrive:
    """PDF guardados en una carpeta de Google Drive compartida con la cuenta de servicio."""
    nombre = "Google Drive"

    def __init__(self, servicio, carpeta_id, tmp):
        self.s, self.carpeta, self.tmp = servicio, carpeta_id, Path(tmp)
        self.hechos = json.loads(F_DRIVE.read_text(encoding="utf-8")) if F_DRIVE.exists() else {}
        self.sub = {n: self._buscar_subcarpeta(n) for n in ("procesadas", "con_error")}

    @classmethod
    def desde_entorno(cls, tmp):
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
        info = json.loads(os.environ["GOOGLE_CREDENTIALS"])
        creds = service_account.Credentials.from_service_account_info(
            info, scopes=["https://www.googleapis.com/auth/drive"])
        servicio = build("drive", "v3", credentials=creds, cache_discovery=False)
        return cls(servicio, os.environ["DRIVE_FOLDER_ID"].strip(), tmp)

    def _listar(self, q):
        out, token = [], None
        while True:
            r = self.s.files().list(q=q, fields="nextPageToken, files(id, name)", pageSize=100,
                                    pageToken=token, supportsAllDrives=True,
                                    includeItemsFromAllDrives=True).execute()
            out += r.get("files", [])
            token = r.get("nextPageToken")
            if not token:
                return out

    def _buscar_subcarpeta(self, nombre):
        q = (f"'{self.carpeta}' in parents and trashed=false and name='{nombre}' "
             "and mimeType='application/vnd.google-apps.folder'")
        r = self._listar(q)
        return r[0]["id"] if r else None

    def _descargar(self, file_id):
        from googleapiclient.http import MediaIoBaseDownload
        buf = io.BytesIO()
        dl = MediaIoBaseDownload(buf, self.s.files().get_media(fileId=file_id, supportsAllDrives=True))
        listo = False
        while not listo:
            _, listo = dl.next_chunk()
        return buf.getvalue()

    def _mover(self, file_id, destino):
        self.s.files().update(fileId=file_id, addParents=destino, removeParents=self.carpeta,
                              supportsAllDrives=True, fields="id").execute()

    def pendientes(self):
        q = f"'{self.carpeta}' in parents and trashed=false and mimeType='application/pdf'"
        items = []
        for f in self._listar(q):
            if f["id"] in self.hechos:
                continue
            ruta = self.tmp / f"{f['id']}.pdf"
            ruta.write_bytes(self._descargar(f["id"]))
            items.append(Item(f["id"], f["name"], ruta))
        return items

    def terminar(self, item, resultado):
        # Se anota siempre qué PDF ya se vio (así no se relee aunque no se pueda mover en Drive)
        self.hechos[item.id] = {"nombre": item.nombre, "resultado": resultado,
                                "fecha": hoy_argentina().isoformat()}
        F_DRIVE.write_text(json.dumps(self.hechos, indent=1, ensure_ascii=False), encoding="utf-8")
        destino = self.sub["con_error" if resultado == "error" else "procesadas"]
        if destino:
            try:
                self._mover(item.id, destino)
            except Exception as e:
                log(f"[AVISO] No pude mover '{item.nombre}' dentro de Drive: {e}")


# ------------------------------------------------------------------ principal
def main():
    DATOS.mkdir(exist_ok=True)
    cfg, hoy = leer_config(), hoy_argentina()
    trab, topes, fac = cargar_trabajadores(), cargar_topes(), cargar_facturas()
    hist = cargar_historico()
    hasta = hist.groupby("cuit")["fecha"].max().to_dict() if not hist.empty else {}
    if not topes:
        sys.exit("Falta datos/topes.csv (o está vacío).")

    tmp = tempfile.TemporaryDirectory()
    origenes, hubo_fallo = [OrigenLocal()], False
    if os.environ.get("GOOGLE_CREDENTIALS") and os.environ.get("DRIVE_FOLDER_ID"):
        try:
            origenes.append(OrigenDrive.desde_entorno(tmp.name))
        except Exception as e:
            log(f"[ERROR] No pude conectarme a Google Drive: {e}")
            hubo_fallo = True
    pendientes = []
    for o in origenes:
        try:
            pendientes += [(o, it) for it in o.pendientes()]
        except Exception as e:
            log(f"[ERROR] No pude leer los PDF de {o.nombre}: {e}")
            hubo_fallo = True

    claves = set(fac["clave"])
    agregadas, resultados = [], []
    for o, it in pendientes:
        try:
            c = leer_comprobante(it.ruta, it.nombre)
        except Exception as e:
            log(f"[ERROR] {it.nombre}: {e}")
            resultados.append((o, it, "error"))
            continue
        if c["fecha"] <= hasta.get(c["cuit"], date.min):
            log(f"[AVISO] {it.nombre}: es del {c['fecha']:%d/%m/%Y}, un mes que ya está en el histórico: no se suma.")
            resultados.append((o, it, "duplicado"))
        elif c["clave"] in claves:
            resultados.append((o, it, "duplicado"))
        else:
            claves.add(c["clave"])
            agregadas.append(c)
            resultados.append((o, it, "ok"))
    if agregadas:
        fac = pd.concat([fac, pd.DataFrame(agregadas)[COLUMNAS_FAC]], ignore_index=True)
        fac.sort_values(["fecha", "cuit", "numero"]).to_csv(F_FAC, index=False, encoding="utf-8")
    for o, it, r in resultados:
        o.terminar(it, r)
    if pendientes:
        cuenta = lambda r: sum(1 for x in resultados if x[2] == r)
        log(f"PDF encontrados: {len(pendientes)} | cargados: {cuenta('ok')} | "
            f"duplicados: {cuenta('duplicado')} | con error: {cuenta('error')}")
        for c in agregadas:
            log(f"  + {cuit_fmt(c['cuit'])}  {c['comprobante']} {c['pto_vta']:05d}-{c['numero']:08d}  "
                f"{pesos(c['importe'])}")
    else:
        print("No hay PDF nuevos.")
    errores_pend = sum(1 for x in resultados if x[2] == "error")

    fac = pd.concat([fac, hist], ignore_index=True) if not hist.empty else fac
    estados, matriz, inicio, meses = calcular(trab, topes, fac, cfg["umbral"], hoy)
    nombres = {t["cuit"]: t["nombre"] for t in trab}
    sin_registrar = set(fac["cuit"]) - set(nombres)
    registro = []
    for _, f in fac.sort_values(["fecha", "numero"], ascending=False).iterrows():
        registro.append({"fecha": f["fecha"], "cuit": int(f["cuit"]),
                         "trabajador": nombres.get(int(f["cuit"]), "No registrado"),
                         "tipo": f["comprobante"], "pto_vta": int(f["pto_vta"]) or None, "numero": int(f["numero"]) or None,
                         "importe": f["importe"], "moneda": f["moneda"], "archivo": f["archivo"],
                         "cargado": f["cargado"], "computa": inicio <= f["fecha"] <= hoy})
    try:
        planilla.construir(EXCEL, cfg["empresa"], cfg["umbral"], hoy, inicio, meses, estados, matriz,
                           registro, topes, cfg["vigencia"], sin_registrar, errores_pend)
    except PermissionError:
        sys.exit("El Excel está abierto. Cerralo y volvé a ejecutar.")
    escribir_resumen_md(cfg, hoy, inicio, estados, sin_registrar, errores_pend)
    for cuit in sin_registrar:
        print(f"[AVISO] El CUIT {cuit_fmt(cuit)} tiene facturas pero no está en trabajadores.csv")
    mostrar(estados)
    enviar_avisos(cfg, estados)
    if hubo_fallo:
        sys.exit("Hubo errores al leer los PDF (ver arriba): el control se generó con los datos disponibles.")


if __name__ == "__main__":
    main()
