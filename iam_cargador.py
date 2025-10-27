# -*- coding: utf-8 -*-
import re, time
import pandas as pd
from pathlib import Path
from playwright.sync_api import sync_playwright

# ====== Rutas ======
BASE = Path(r"C:\IAM_Automatizacion")
ARCHIVO_CARGA   = BASE / "carga_anticipadas.xlsx"
ARCHIVO_PUERTOS = BASE / "puertos_2025.xlsx"
URL_IAM = "https://iam.lakaut.com.ar/documentacion"

# ====== Constantes especiales ======
CODIGO_DESCONOCIDO = "ZZZZZ"
LUGAR_DESCONOCIDO  = "OTROS"
PAIS_DESCONOCIDO   = "701"       # en IAM se ve como "Desconocido"

# ====== Selectores ======
SEL_BTN_AGREGAR   = 'button[data-target="#CreateCaratulaModal"]'
SEL_ID_VIAJE      = '#IdentificadorViaje'
SEL_TITULO_MADRE  = '#IdentificadorTituloMadre'
SEL_BUQUE         = '#Buque'
SEL_PAIS_ORIGEN   = '#CodigoPaisLugarOrigen'   # autocomplete
SEL_LUGAR_PROC    = '#LugarOrigen'
SEL_COD_PUERTO    = '#CodigoPuertoEmbarque'    # autocomplete
SEL_FECHA_EMB     = '#FechaEmbarque'
SEL_BTN_CREAR     = 'input.btn.btn-primary[type="submit"][value="Crear"]'

# ====== Utilidades ======
def norm_col(s: str) -> str:
    return s.strip().lower()

def normalizar_id_viaje(v):
    if v is None: return ""
    s = str(v).strip().replace(".0", "")
    return re.sub(r"[^A-Za-z0-9-]", "", s)

def seleccionar_autocomplete_pais(page, input_sel: str, pais: str):
    """Selecciona la opción cuyo texto coincida EXACTO con el país."""
    if not pais: return
    pais = str(pais).strip()
    page.click(input_sel)
    page.fill(input_sel, "")
    page.type(input_sel, pais, delay=25)
    page.wait_for_timeout(300)
    intentos = [
        f"//div[contains(@class,'easy-autocomplete-container')]//li[normalize-space()='{pais}']",
        f"//ul[contains(@class,'ui-autocomplete')]//li[normalize-space()='{pais}']",
        f"//li[@role='option' and normalize-space()='{pais}']",
    ]
    for xp in intentos:
        loc = page.locator(xp)
        try:
            if loc.count() > 0:
                loc.first.click()
                return
        except Exception:
            pass
    # Fallback
    page.keyboard.press("ArrowDown"); page.keyboard.press("Enter")
    page.wait_for_timeout(150)

def seleccionar_autocomplete_codigo(page, input_sel: str, codigo: str):
    """Selecciona la opción cuyo texto COMIENCE con el código (o maneja ZZZZZ)."""
    if not codigo: return
    codigo = str(codigo).strip().upper()
    page.click(input_sel)
    page.fill(input_sel, "")
    page.type(input_sel, codigo, delay=25)
    page.wait_for_timeout(300)
    # 1) Prefijo por código
    intentos_prefijo = [
        f"//div[contains(@class,'easy-autocomplete-container')]//li[starts-with(normalize-space(), '{codigo}')]",
        f"//ul[contains(@class,'ui-autocomplete')]//li[starts-with(normalize-space(), '{codigo}')]",
        f"//li[@role='option' and starts-with(normalize-space(), '{codigo}')]",
    ]
    for xp in intentos_prefijo:
        loc = page.locator(xp)
        try:
            if loc.count() > 0:
                loc.first.click()
                return
        except Exception:
            pass
    # 2) Caso ZZZZZ → buscar "OTROS" visible junto al código
    if codigo == CODIGO_DESCONOCIDO:
        intentos_otro = [
            # item que contenga "OTROS"
            "//li[contains(normalize-space(), 'OTROS')]",
            "//div[contains(@class,'easy-autocomplete-container')]//li[contains(normalize-space(), 'OTROS')]",
            "//ul[contains(@class,'ui-autocomplete')]//li[contains(normalize-space(), 'OTROS')]",
        ]
        for xp in intentos_otro:
            loc = page.locator(xp)
            try:
                if loc.count() > 0:
                    loc.first.click()
                    return
            except Exception:
                pass
    # 3) Último recurso
    page.keyboard.press("ArrowDown"); page.keyboard.press("Enter")
    page.wait_for_timeout(150)

# ---- Helpers para asegurar el envío ----
def campo_lleno(page, selector):
    try:
        return page.locator(selector).input_value().strip() != ""
    except:
        return False

def esperar_boton_crear_habilitado(page, sel_btn, timeout_ms=10000):
    fin = time.time() + timeout_ms/1000
    btn = page.locator(sel_btn)
    while time.time() < fin:
        try:
            btn.scroll_into_view_if_needed()
            if btn.is_visible() and not btn.is_disabled():
                return True
        except:
            pass
        page.wait_for_timeout(150)
    return False

def enviar_form_por_js(page, sel_btn):
    """Si el click no funciona, dispara el submit del <form> contenedor."""
    btn = page.locator(sel_btn).first
    try:
        page.evaluate(
            """(el)=>{
                const form = el.closest('form');
                if (form) { form.requestSubmit ? form.requestSubmit() : form.submit(); }
            }""",
            btn.element_handle()
        )
        page.wait_for_timeout(1200)
    except:
        pass

# ====== Leer Excels (forzando ID como texto) ======
print("📂 Leyendo planillas...")
df_carga   = pd.read_excel(ARCHIVO_CARGA,   dtype={"Identificador del viaje": str})
df_puertos = pd.read_excel(ARCHIVO_PUERTOS)

# Normalizar encabezados
df_carga.columns   = [norm_col(c) for c in df_carga.columns]
df_puertos.columns = [norm_col(c) for c in df_puertos.columns]

# Columnas esperadas (en carga)
col_id_viaje = "identificador del viaje"
col_titulo   = "título madre"
col_fecha    = "fecha de embarque"
col_buque    = "buque"
col_codigo   = "codigo del puerto de embarque"

# Columnas en puertos (tolerantes a variantes)
def pick(colnames, *ops):
    for o in ops:
        o = norm_col(o)
        if o in colnames: return o
    return None

col_puerto_key = pick(df_puertos.columns, "codigo del puerto de embarque", "codigo puerto", "codigo_puerto")
col_lugar      = pick(df_puertos.columns, "lugar de procedencia", "lugar procedencia", "lugar")
col_pais       = pick(df_puertos.columns, "pais de origen", "país de origen", "pais")

if not col_puerto_key or not col_lugar or not col_pais:
    print("❌ Encabezados en puertos_2025.xlsx no encontrados (código/lugar/país).")
    input("ENTER para cerrar..."); raise SystemExit

# ====== Cruce robusto por clave normalizada ======
df_carga["_key_puerto"]   = df_carga[col_codigo].astype(str).str.strip().str.upper()
df_puertos["_key_puerto"] = df_puertos[col_puerto_key].astype(str).str.strip().str.upper()

df = pd.merge(
    df_carga,
    df_puertos[["_key_puerto", col_lugar, col_pais]],
    on="_key_puerto",
    how="left"
)

# Columnas finales: completar si faltan
if "lugar de procedencia" in df.columns:
    df["lugar de procedencia"] = df["lugar de procedencia"].fillna(df[col_lugar])
else:
    df["lugar de procedencia"] = df[col_lugar]

if "pais de origen" in df.columns:
    df["pais de origen"] = df["pais de origen"].fillna(df[col_pais])
else:
    df["pais de origen"] = df[col_pais]

print(f"✅ {len(df)} filas preparadas para cargar.")

# ====== Automatización ======
with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(user_data_dir=str(BASE/"iam_profile"), headless=False)
    page = ctx.new_page()
    page.goto(URL_IAM)

    print("\n🔐 Logueate si lo pide y dejá visible 'Agregar Carátula'.")
    input("ENTER para continuar... ")

    for i, row in df.iterrows():
        try:
            id_viaje = normalizar_id_viaje(row.get(col_id_viaje, ""))
            titulo   = str(row.get(col_titulo, "")).strip()
            buque    = str(row.get(col_buque, "")).strip()
            codigo   = str(row.get(col_codigo, "")).strip().upper()
            lugar    = str(row.get("lugar de procedencia", "")).strip()
            pais     = str(row.get("pais de origen", "")).strip()
            fecha    = str(row.get(col_fecha, "")).strip()

            # Caso especial: ZZZZZ = “OTROS / 701(Desconocido)”
            if codigo == CODIGO_DESCONOCIDO:
                lugar = LUGAR_DESCONOCIDO
                pais  = PAIS_DESCONOCIDO

            print(f"\n— Fila {i+2}: ID={id_viaje}  COD={codigo}  PAIS={pais}  LUGAR={lugar}")

            # Abrir modal
            page.locator(SEL_BTN_AGREGAR).first.click(force=True)
            page.wait_for_selector(SEL_ID_VIAJE, timeout=20000)
            time.sleep(0.4)

            # Completar campos
            page.fill(SEL_ID_VIAJE, "");      page.type(SEL_ID_VIAJE, id_viaje, delay=20)
            page.fill(SEL_TITULO_MADRE, titulo)
            page.fill(SEL_BUQUE, buque)

            # País (si es 701, IAM mostrará “Desconocido”)
            seleccionar_autocomplete_pais(page, SEL_PAIS_ORIGEN, pais)
            page.wait_for_timeout(200)

            # Lugar
            page.fill(SEL_LUGAR_PROC, lugar)

            # Código de puerto
            seleccionar_autocomplete_codigo(page, SEL_COD_PUERTO, codigo)

            # Fecha
            page.fill(SEL_FECHA_EMB, "");     page.type(SEL_FECHA_EMB, fecha)

            # ---- Enviar de forma robusta ----
            page.locator(SEL_FECHA_EMB).press("Tab")
            page.wait_for_timeout(200)

            if esperar_boton_crear_habilitado(page, SEL_BTN_CREAR, timeout_ms=12000):
                page.locator(SEL_BTN_CREAR).click(force=True)
                page.wait_for_timeout(1500)

            if page.locator(SEL_ID_VIAJE).is_visible():
                page.locator(SEL_FECHA_EMB).press("Enter")
                page.wait_for_timeout(1200)

            if page.locator(SEL_ID_VIAJE).is_visible():
                enviar_form_por_js(page, SEL_BTN_CREAR)
                page.wait_for_timeout(1200)

        except Exception as e:
            print(f"⚠️ Error en fila {i+2}: {e}")
            continue

    print("\n🎯 Listo.")
    input("ENTER para cerrar...")
    ctx.close()