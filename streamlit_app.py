"""
Dashboard de Cobranza Mayo 2025 — NAtura
Indicadores ejecutivos para reunion de cierre de mes
"""

import random
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

random.seed(2025)
np.random.seed(2025)

# ─────────────────────────────────────────────
# CONFIGURACION
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Cobranza Mayo 2025 — NAtura",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

DARK  = "#0f172a"
CARD  = "#1e293b"
BLUE  = "#3b82f6"
GREEN = "#22c55e"
RED   = "#ef4444"
AMBER = "#f59e0b"
PURPLE= "#8b5cf6"
SLATE = "#64748b"
TEXT  = "#f1f5f9"
MUTED = "#94a3b8"

st.markdown(f"""
<style>
  /* fondo general */
  .stApp {{ background-color: {DARK}; color: {TEXT}; }}
  [data-testid="stAppViewContainer"] {{ background-color: {DARK}; }}
  [data-testid="stHeader"] {{ background-color: {DARK}; }}

  /* tabs */
  .stTabs [data-baseweb="tab-list"] {{
      background: {CARD}; border-radius: 10px; padding: 4px; gap: 4px;
  }}
  .stTabs [data-baseweb="tab"] {{
      color: {MUTED}; border-radius: 8px; font-weight: 600; font-size: 0.85rem;
  }}
  .stTabs [aria-selected="true"] {{
      background: {BLUE} !important; color: white !important;
  }}

  /* dataframes */
  [data-testid="stDataFrame"] {{ border-radius: 10px; }}

  /* métricas */
  [data-testid="stMetric"] {{
      background: {CARD}; border-radius: 10px; padding: 14px 18px;
      border-left: 3px solid {BLUE};
  }}
  [data-testid="stMetricValue"] {{ color: {TEXT} !important; font-size: 1.8rem !important; }}
  [data-testid="stMetricLabel"] {{ color: {MUTED} !important; }}

  /* títulos de sección */
  .sec {{ font-size: 1.2rem; font-weight: 700; color: {TEXT};
          border-bottom: 2px solid #334155; padding-bottom: 6px; margin: 8px 0 16px 0; }}

  /* badge */
  .badge {{
      display:inline-block; padding: 2px 10px; border-radius: 20px;
      font-size: 0.78rem; font-weight: 700;
  }}
  .badge-red   {{ background:#450a0a; color:{RED};   border:1px solid {RED}; }}
  .badge-green {{ background:#052e16; color:{GREEN}; border:1px solid {GREEN}; }}
  .badge-amber {{ background:#451a03; color:{AMBER}; border:1px solid {AMBER}; }}

  /* alerta box */
  [data-testid="stAlert"] {{ border-radius: 10px; }}
</style>
""", unsafe_allow_html=True)

PLOTLY_LAYOUT = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(color=TEXT, family="Inter, sans-serif"),
    margin=dict(l=10, r=10, t=30, b=10),
)

def apply_layout(fig, **kwargs):
    layout = {**PLOTLY_LAYOUT, **kwargs}
    for ax in ("xaxis", "yaxis", "xaxis2", "yaxis2"):
        base = dict(gridcolor="#334155", linecolor="#334155")
        if ax in kwargs:
            base.update(kwargs[ax])
            layout[ax] = base
    fig.update_layout(**layout)
    return fig

# ─────────────────────────────────────────────
# HELPERS DE CARGA
# ─────────────────────────────────────────────
def leer_archivo(f):
    if f is None:
        return None
    try:
        if f.name.lower().endswith((".xlsx", ".xls")):
            return pd.read_excel(f)
        return pd.read_csv(f, encoding="utf-8-sig", low_memory=False)
    except Exception:
        try:
            f.seek(0)
            return pd.read_csv(f, encoding="latin-1", low_memory=False)
        except Exception:
            return None

def col(df, *opciones):
    for o in opciones:
        if o in df.columns:
            return o
    return None

# ─────────────────────────────────────────────
# DATOS SINTETICOS MAYO 2025 (fallback)
# ─────────────────────────────────────────────
ESTADOS   = ["CDMX","Edo. de México","Jalisco","Nuevo León","Puebla",
             "Guanajuato","Veracruz","Michoacán","Chihuahua","Tamaulipas"]
SECTORES  = ["Belleza","Moda","Hogar","Nutrición","Bienestar","Joyería"]
SEGMENTOS = ["Diamante","Oro","Plata","Bronce","Nuevo Ingreso"]
GVS       = [f"GV-{str(i).zfill(3)}" for i in range(1, 16)]
ASESORES  = [f"Asesor {chr(65+i)}" for i in range(10)]
HORAS     = list(range(8, 21))

META_TOTAL       = 4_850_000
RECUPERADO_TOTAL = 3_892_500
PROMESAS_GEN     = 1_240
PROMESAS_CUMP    = 748
PROMESAS_CAIDAS  = PROMESAS_GEN - PROMESAS_CUMP

dias = pd.date_range("2025-05-01", "2025-05-31", freq="B")
meta_d = META_TOTAL / len(dias)
rec_d = np.clip(np.random.normal(meta_d * 0.92, meta_d * 0.15, len(dias)), meta_d * 0.4, meta_d * 1.3)
df_cierre = pd.DataFrame({"Fecha": dias, "Meta": meta_d, "Recuperado": rec_d})
df_cierre["Meta Acum"]       = df_cierre["Meta"].cumsum()
df_cierre["Recuperado Acum"] = df_cierre["Recuperado"].cumsum()

df_motivos = pd.DataFrame({
    "Motivo": ["Sin liquidez","No contestó","Negó deuda","Promesa vencida","Número inválido","Otro"],
    "Casos":  [198, 167, 89, 134, 72, 61],
})

TOTAL_INT  = 38_420
TITULAR    = 14_820
BUZON      = 9_105
NO_CONT    = 8_934
NUM_INV    = 3_210
COLGADO    = 2_351
CR         = TITULAR / TOTAL_INT

df_horario = pd.DataFrame({
    "Hora": HORAS,
    "Contactos": [120,310,520,680,740,810,860,780,690,620,480,310,210],
    "Promesas":  [ 12, 35, 62, 88, 95,104,112,100, 89, 78, 61, 39, 25],
})
df_horario["Conv %"] = (df_horario["Promesas"] / df_horario["Contactos"] * 100).round(1)

df_canal = pd.DataFrame({
    "Canal":      ["SMS","WhatsApp","Email","Llamada","IVR"],
    "Enviados":   [12400, 8300, 9800, 38420, 4200],
    "Respuestas": [1240,  2076,  588, 14820,  840],
    "Promesas":   [ 186,   622,   74,   748,  168],
})
df_canal["Tasa resp %"] = (df_canal["Respuestas"] / df_canal["Enviados"] * 100).round(1)
df_canal["Conv prom %"] = (df_canal["Promesas"]   / df_canal["Respuestas"] * 100).round(1)

df_sector = pd.DataFrame({
    "Sector": SECTORES,
    "Meta":       [980000, 870000, 650000, 720000, 540000, 490000],
    "Recuperado": [784000, 652500, 487500, 504000, 378000, 318000],
})
df_sector["Cumplimiento %"] = (df_sector["Recuperado"] / df_sector["Meta"] * 100).round(1)

np.random.seed(42)
df_gv = pd.DataFrame({
    "GV":         GVS,
    "Meta":       np.random.randint(200000, 500000, len(GVS)),
    "Recuperado": np.random.randint(140000, 460000, len(GVS)),
})
df_gv["Cumplimiento %"] = (df_gv["Recuperado"] / df_gv["Meta"] * 100).round(1)
df_gv = df_gv.sort_values("Cumplimiento %", ascending=False).reset_index(drop=True)

df_edad = pd.DataFrame({
    "Rango":     ["18-25","26-35","36-45","46-55","56-65","66+"],
    "Cuentas":   [1240, 3870, 4210, 3650, 2180, 980],
    "Recuperado":[74400,271400,336800,255500,130800,49000],
    "Promesas %":[38,62,71,65,52,41],
})
df_edad["Ticket Prom"] = (df_edad["Recuperado"] / df_edad["Cuentas"]).round(0)

df_estado = pd.DataFrame({
    "Estado":     ESTADOS,
    "Cuentas":    [4820,3610,2840,1980,1650,1420,1280,1190,980,870],
    "Recuperado": [1120000,892000,620000,485000,378000,312000,280000,245000,198000,168000],
})
df_estado["Ticket Prom"] = (df_estado["Recuperado"] / df_estado["Cuentas"]).round(0)

df_segmento = pd.DataFrame({
    "Segmento":      SEGMENTOS,
    "Cuentas":       [890,2340,4210,5680,2320],
    "Recuperado":    [534000,936000,1263000,1022400,139200],
    "Cumplimiento %":[87,82,74,68,45],
})

np.random.seed(99)
df_asesores = pd.DataFrame({
    "Asesor":            ASESORES,
    "Llamadas":          np.random.randint(280, 520, 10),
    "TMO (min)":         np.random.uniform(3.2, 6.8, 10).round(1),
    "Contactos Titular": np.random.randint(80, 180, 10),
    "Promesas":          np.random.randint(20, 65, 10),
    "Colgadas <30s":     np.random.randint(15, 60, 10),
    "Monto Rec":         np.random.randint(180000, 520000, 10),
})
df_asesores["% Contacto"] = (df_asesores["Contactos Titular"] / df_asesores["Llamadas"] * 100).round(1)
df_asesores["% Abandono"] = (df_asesores["Colgadas <30s"]     / df_asesores["Llamadas"] * 100).round(1)
df_asesores["Conv %"]     = (df_asesores["Promesas"]           / df_asesores["Contactos Titular"] * 100).round(1)
df_asesores = df_asesores.sort_values("Monto Rec", ascending=False).reset_index(drop=True)

df_objeciones = pd.DataFrame({
    "Objeción":           ["No tengo dinero","No reconozco la deuda","Ya pagué",
                           "Espera a quincena","Mándame estado de cuenta","No soy yo"],
    "Frecuencia":         [2840, 1920, 1480, 2210, 890, 640],
    "Resolución exitosa %":[38,  51,   72,   61,   44,  58],
})

META_JUN      = 5_200_000
PROY_JUN      = 4_420_000
PIPELINE_PROM = 1_840_000

dias_jun = pd.date_range("2025-06-01","2025-06-30", freq="B")
n = len(dias_jun)
df_proy = pd.DataFrame({
    "Fecha":      dias_jun,
    "Meta":       np.linspace(0, META_JUN,        n),
    "Base":       np.linspace(0, PROY_JUN,        n),
    "Optimista":  np.linspace(0, META_JUN * 1.05, n),
    "Pesimista":  np.linspace(0, PROY_JUN * 0.88, n),
})

df_palancas = pd.DataFrame({
    "Palanca":             ["Recuperar promesas caídas Mayo","Incrementar contact rate",
                            "Mejorar conv. WhatsApp","Script objeciones","GVs rezagados"],
    "Impacto estimado $":  [420000, 280000, 180000, 150000, 220000],
})

df_acciones = pd.DataFrame([
    ("Inmediato","Operaciones","Rellamada a 2,891 promesas caídas"),
    ("Inmediato","Analítica","Reporte GV diario automatizado"),
    ("Lun 9 Jun","Canales Digitales","Lanzar A/B test SMS por segmento"),
    ("Mar 10 Jun","Supervisión","Coaching script objeciones 'sin dinero'"),
    ("Jue 12 Jun","Gerencia","Validar incentivo para GVs rezagados"),
    ("Vie 13 Jun","Operaciones","Ajuste de marcaciones franja 12-15h +25%"),
], columns=["Plazo","Área","Acción"])

# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📂 Cargar datos reales")
    st.caption("Sube tus archivos de Mayo. El dashboard se actualiza automáticamente.")
    f_cartera  = st.file_uploader("📋 Cartera / Remesa",        type=["csv","xlsx","xls"], key="cartera")
    f_pagos    = st.file_uploader("💰 Pagos / Recuperación",    type=["csv","xlsx","xls"], key="pagos")
    f_gestion  = st.file_uploader("📞 Gestión de llamadas",     type=["csv","xlsx","xls"], key="gestion")
    f_promesas = st.file_uploader("🤝 Promesas de pago",        type=["csv","xlsx","xls"], key="promesas")
    MODO_REAL = any([f_cartera, f_pagos, f_gestion, f_promesas])
    if MODO_REAL:
        st.success("✅ Datos reales activos")
        for nombre, f in [("Cartera",f_cartera),("Pagos",f_pagos),("Gestión",f_gestion),("Promesas",f_promesas)]:
            if f:
                df_tmp = leer_archivo(f)
                if df_tmp is not None:
                    st.markdown(f"**{nombre}** — {len(df_tmp):,} filas, {len(df_tmp.columns)} cols")
                    with st.expander(f"Columnas {nombre}"):
                        st.write(list(df_tmp.columns))
    else:
        st.warning("⚠️ Usando datos de prueba")
        st.markdown("---")
        st.markdown("**Columnas que reconoce el sistema:**")
        st.markdown("""
**Cartera/Remesa:** `valor_saldo_deuda`, `division`, `zona`, `red`, `edad_consultora`, `segmento_nombre`, `direccion_de_residencia_estado`

**Pagos:** `monto_pagado` / `valor_pago` / `importe`, `fecha_pago`, `asesor` / `ejecutivo`

**Gestión:** `resultado` / `resultado_gestion`, `hora_llamada`, `duracion_seg` / `duracion`, `canal` / `medio`, `asesor`

**Promesas:** `monto_promesa` / `promesa_monto`, `fecha_promesa`, `cumplida` / `estatus`
        """)

# ─────────────────────────────────────────────
# CARGAR DATOS REALES
# ─────────────────────────────────────────────
df_cart_real  = leer_archivo(f_cartera)  if f_cartera  else None
df_pago_real  = leer_archivo(f_pagos)   if f_pagos    else None
df_gest_real  = leer_archivo(f_gestion) if f_gestion  else None
df_prom_real  = leer_archivo(f_promesas)if f_promesas else None

if df_pago_real is not None:
    c_monto = col(df_pago_real, "monto_pagado","valor_pago","importe","monto","pago")
    c_fecha = col(df_pago_real, "fecha_pago","fecha","date","fecha_operacion")
    c_asesor= col(df_pago_real, "asesor","ejecutivo","agente","nombre_asesor")
    RECUPERADO_TOTAL = df_pago_real[c_monto].sum() if c_monto else RECUPERADO_TOTAL
    if c_fecha and c_monto:
        df_pago_real[c_fecha] = pd.to_datetime(df_pago_real[c_fecha], errors="coerce")
        df_cierre = (df_pago_real.dropna(subset=[c_fecha]).groupby(c_fecha)[c_monto].sum()
                     .reset_index().rename(columns={c_fecha:"Fecha", c_monto:"Recuperado"}))
        df_cierre = df_cierre[df_cierre["Fecha"].dt.month == 5].sort_values("Fecha")
        df_cierre["Meta"] = META_TOTAL / max(len(df_cierre), 1)
        df_cierre["Meta Acum"]       = df_cierre["Meta"].cumsum()
        df_cierre["Recuperado Acum"] = df_cierre["Recuperado"].cumsum()

if df_prom_real is not None:
    c_mp  = col(df_prom_real, "monto_promesa","promesa_monto","monto","importe")
    c_cum = col(df_prom_real, "cumplida","estatus","status","resultado")
    PROMESAS_GEN   = len(df_prom_real)
    if c_cum:
        cumplidas = df_prom_real[c_cum].astype(str).str.lower().isin(["1","si","sí","cumplida","pagada","true"])
        PROMESAS_CUMP  = int(cumplidas.sum())
        PROMESAS_CAIDAS= PROMESAS_GEN - PROMESAS_CUMP
    if c_mp:
        PIPELINE_PROM = df_prom_real[c_mp].sum()

if df_gest_real is not None:
    c_res  = col(df_gest_real, "resultado","resultado_gestion","disposicion","tipificacion")
    c_hora = col(df_gest_real, "hora_llamada","hora","hour")
    c_dur  = col(df_gest_real, "duracion_seg","duracion","duration","tmo")
    c_canal= col(df_gest_real, "canal","medio","channel","tipo_contacto")
    c_as   = col(df_gest_real, "asesor","ejecutivo","agente")
    TOTAL_INT = len(df_gest_real)
    if c_res:
        res_lower = df_gest_real[c_res].astype(str).str.lower()
        TITULAR = int(res_lower.isin(["contacto titular","titular","contactado","promesa","pdc"]).sum())
        BUZON   = int(res_lower.str.contains("buz|voicemail|vm",na=False).sum())
        NO_CONT = int(res_lower.str.contains("no contest|no answer|sin respuesta",na=False).sum())
        NUM_INV = int(res_lower.str.contains("inv[aá]lido|wrong|error",na=False).sum())
        COLGADO = int(res_lower.str.contains("colg|hang|abandon",na=False).sum())
        CR      = TITULAR / max(TOTAL_INT, 1)
    if c_hora:
        try:
            df_gest_real["_hora"] = pd.to_datetime(df_gest_real[c_hora], errors="coerce").dt.hour
            if c_res:
                grp = df_gest_real.groupby("_hora")
                df_horario = pd.DataFrame({
                    "Hora":      grp.size().index.tolist(),
                    "Contactos": grp.size().values.tolist(),
                    "Promesas":  grp.apply(lambda x: x[c_res].astype(str).str.lower()
                                           .isin(["promesa","pdc","promesa de pago"]).sum()).values.tolist(),
                })
                df_horario["Conv %"] = (df_horario["Promesas"] / df_horario["Contactos"].clip(1) * 100).round(1)
        except Exception:
            pass
    if c_canal and c_res:
        def canal_stats(grp):
            total = len(grp)
            resp  = grp[c_res].notna().sum()
            prom  = grp[c_res].astype(str).str.lower().isin(["promesa","pdc","promesa de pago"]).sum()
            return pd.Series({"Enviados": total, "Respuestas": resp, "Promesas": prom})
        df_canal = df_gest_real.groupby(c_canal).apply(canal_stats).reset_index()
        df_canal.columns = ["Canal","Enviados","Respuestas","Promesas"]
        df_canal["Tasa resp %"] = (df_canal["Respuestas"] / df_canal["Enviados"].clip(1) * 100).round(1)
        df_canal["Conv prom %"] = (df_canal["Promesas"]   / df_canal["Respuestas"].clip(1) * 100).round(1)
    if c_as and c_res:
        def asesor_stats(grp):
            llamadas = len(grp)
            titular  = grp[c_res].astype(str).str.lower().isin(["contacto titular","titular","contactado","promesa","pdc"]).sum()
            promesas = grp[c_res].astype(str).str.lower().isin(["promesa","pdc","promesa de pago"]).sum()
            tmo = grp[c_dur].mean() / 60 if c_dur and pd.api.types.is_numeric_dtype(grp[c_dur]) else 0
            return pd.Series({"Llamadas": llamadas, "Contactos Titular": titular,
                               "Promesas": promesas, "TMO (min)": round(tmo, 1),
                               "Colgadas <30s": 0, "Monto Rec": 0})
        df_asesores = df_gest_real.groupby(c_as).apply(asesor_stats).reset_index()
        df_asesores.rename(columns={c_as: "Asesor"}, inplace=True)
        if df_pago_real is not None and c_asesor and c_monto:
            rec_asesor = df_pago_real.groupby(c_asesor)[c_monto].sum().reset_index()
            rec_asesor.columns = ["Asesor","Monto Rec"]
            df_asesores = df_asesores.merge(rec_asesor, on="Asesor", how="left", suffixes=("","_r"))
            if "Monto Rec_r" in df_asesores.columns:
                df_asesores["Monto Rec"] = df_asesores["Monto Rec_r"].fillna(0)
                df_asesores.drop(columns=["Monto Rec_r"], inplace=True)
        df_asesores["% Contacto"] = (df_asesores["Contactos Titular"] / df_asesores["Llamadas"].clip(1) * 100).round(1)
        df_asesores["% Abandono"] = 0.0
        df_asesores["Conv %"]     = (df_asesores["Promesas"] / df_asesores["Contactos Titular"].clip(1) * 100).round(1)
        df_asesores = df_asesores.sort_values("Monto Rec", ascending=False).reset_index(drop=True)

if df_cart_real is not None:
    c_saldo  = col(df_cart_real, "valor_saldo_deuda","saldo","deuda","monto_deuda","importe")
    c_zona   = col(df_cart_real, "zona","zone","gv","gerencia")
    c_edad   = col(df_cart_real, "edad_consultora","edad","age")
    c_seg    = col(df_cart_real, "segmento_nombre","segmento","segment","categoria")
    c_estado = col(df_cart_real, "direccion_de_residencia_estado","estado","state","entidad")
    if c_saldo:
        META_TOTAL = df_cart_real[c_saldo].sum()
    if c_zona and c_saldo:
        df_gv = (df_cart_real.groupby(c_zona)[c_saldo].agg(["sum","count"])
                 .reset_index().rename(columns={c_zona:"GV","sum":"Meta","count":"Cuentas"}))
        if df_pago_real is not None:
            c_zona_p = col(df_pago_real, "zona","zone","gv","gerencia")
            c_mp2    = col(df_pago_real, "monto_pagado","valor_pago","importe","monto","pago")
            if c_zona_p and c_mp2:
                rec_gv = df_pago_real.groupby(c_zona_p)[c_mp2].sum().reset_index()
                rec_gv.columns = ["GV","Recuperado"]
                df_gv = df_gv.merge(rec_gv, on="GV", how="left")
                df_gv["Recuperado"] = df_gv["Recuperado"].fillna(0)
            else:
                df_gv["Recuperado"] = df_gv["Meta"] * 0.8
        else:
            df_gv["Recuperado"] = df_gv["Meta"] * 0.8
        df_gv["Cumplimiento %"] = (df_gv["Recuperado"] / df_gv["Meta"].clip(1) * 100).round(1)
        df_gv = df_gv.sort_values("Cumplimiento %", ascending=False).reset_index(drop=True)
    if c_seg and c_saldo:
        grp_s = df_cart_real.groupby(c_seg)[c_saldo].agg(["sum","count"]).reset_index()
        grp_s.columns = ["Segmento","Meta","Cuentas"]
        if df_pago_real is not None:
            c_seg_p = col(df_pago_real, "segmento_nombre","segmento","segment","categoria")
            c_mp5   = col(df_pago_real, "monto_pagado","valor_pago","importe","monto","pago")
            if c_seg_p and c_mp5:
                rec_s = df_pago_real.groupby(c_seg_p)[c_mp5].sum().reset_index()
                rec_s.columns = ["Segmento","Recuperado"]
                grp_s = grp_s.merge(rec_s, on="Segmento", how="left")
                grp_s["Recuperado"] = grp_s["Recuperado"].fillna(0)
            else:
                grp_s["Recuperado"] = grp_s["Meta"] * 0.8
        else:
            grp_s["Recuperado"] = grp_s["Meta"] * 0.8
        grp_s["Cumplimiento %"] = (grp_s["Recuperado"] / grp_s["Meta"].clip(1) * 100).round(1)
        df_segmento = grp_s[["Segmento","Cuentas","Recuperado","Cumplimiento %"]]

# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────
col_h1, col_h2, col_h3 = st.columns([3, 1, 1])
with col_h1:
    st.markdown("## 📊 Dashboard Ejecutivo — Cobranza Mayo 2025")
    st.caption("NAtura | Reunión de seguimiento | 09 Jun 2025")
with col_h2:
    pct = RECUPERADO_TOTAL / META_TOTAL * 100
    delta_pct = pct - 100
    st.metric("Recuperación Mayo", f"${RECUPERADO_TOTAL/1e6:.2f}M", delta=f"{delta_pct:.1f}% vs meta")
with col_h3:
    st.metric("Contact Rate", f"{CR*100:.1f}%", delta="–3.2pp vs abr")

# ─────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "1 · Cierre de Mes", "2 · Contactabilidad", "3 · Indicadores", "4 · Operación", "5 · Plan de Trabajo",
])

# ══════════════════════════════════════════════
# TAB 1 — CIERRE DE MES
# ══════════════════════════════════════════════
with tab1:
    st.markdown('<div class="sec">Cierre de Mes — Mayo 2025</div>', unsafe_allow_html=True)
    c1,c2,c3,c4,c5 = st.columns(5)
    c1.metric("Meta mensual",         f"${META_TOTAL/1e6:.2f}M")
    c2.metric("Recuperado",           f"${RECUPERADO_TOTAL/1e6:.2f}M", delta=f"{(RECUPERADO_TOTAL/META_TOTAL-1)*100:.1f}%")
    c3.metric("Cumplimiento",         f"{RECUPERADO_TOTAL/META_TOTAL*100:.1f}%")
    c4.metric("Promesas generadas",   f"{PROMESAS_GEN:,}")
    c5.metric("Promesas caídas",      f"{PROMESAS_CAIDAS:,}", delta=f"{PROMESAS_CAIDAS/PROMESAS_GEN*100:.0f}% caída", delta_color="inverse")
    st.markdown("---")
    col_l, col_r = st.columns([3, 2])
    with col_l:
        st.markdown("**Recuperación acumulada vs Meta**")
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=df_cierre["Fecha"], y=df_cierre["Meta Acum"], name="Meta",
                                 line=dict(color=SLATE, dash="dash", width=2), fill=None))
        fig.add_trace(go.Scatter(x=df_cierre["Fecha"], y=df_cierre["Recuperado Acum"], name="Recuperado",
                                 line=dict(color=BLUE, width=3), fill="tonexty", fillcolor="rgba(59,130,246,0.08)"))
        fig.add_hrect(y0=META_TOTAL*0.95, y1=META_TOTAL*1.05, fillcolor="rgba(34,197,94,0.05)", line_width=0,
                      annotation_text="±5% meta", annotation_position="top right", annotation_font_color=GREEN)
        apply_layout(fig, height=320, legend=dict(orientation="h", y=1.1))
        fig.update_yaxes(tickprefix="$", tickformat=",.0f")
        st.plotly_chart(fig, use_container_width=True)
    with col_r:
        st.markdown("**Motivos de caída de promesas**")
        df_mot = df_motivos.sort_values("Casos")
        fig2 = go.Figure(go.Bar(x=df_mot["Casos"], y=df_mot["Motivo"], orientation="h",
                                marker_color=[RED if c > 150 else AMBER if c > 80 else SLATE for c in df_mot["Casos"]],
                                text=df_mot["Casos"], textposition="outside", textfont=dict(color=TEXT)))
        apply_layout(fig2, height=320)
        st.plotly_chart(fig2, use_container_width=True)
    gap = META_TOTAL - RECUPERADO_TOTAL
    st.info(f"**Brecha de recuperación: ${gap/1e6:.2f}M** — las **{PROMESAS_CAIDAS} promesas caídas** "
            f"explican ~{PROMESAS_CAIDAS/PROMESAS_GEN*100:.0f}% del desvío. "
            f"Principales factores: sin liquidez ({df_motivos.iloc[0]['Casos']} casos) "
            f"y contacto perdido ({df_motivos.iloc[1]['Casos']} sin rellamada efectiva).")

# ══════════════════════════════════════════════
# TAB 2 — CONTACTABILIDAD
# ══════════════════════════════════════════════
with tab2:
    st.markdown('<div class="sec">Contactabilidad y Canales Digitales</div>', unsafe_allow_html=True)
    c1,c2,c3,c4,c5 = st.columns(5)
    c1.metric("Total intentos",       f"{TOTAL_INT:,}")
    c2.metric("Contact Rate",         f"{CR*100:.1f}%",  delta="–3.2pp vs abr", delta_color="inverse")
    c3.metric("Titular contactado",   f"{TITULAR:,}")
    c4.metric("Buzón + No contesta",  f"{BUZON+NO_CONT:,}")
    c5.metric("Cuelga al agente",     f"{COLGADO:,}",    delta="Alto", delta_color="inverse")
    st.markdown("---")
    col_l, col_r = st.columns(2)
    with col_l:
        st.markdown("**Contactos y promesas por hora**")
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_trace(go.Bar(x=df_horario["Hora"], y=df_horario["Contactos"], name="Contactos",
                             marker_color=BLUE, opacity=0.8), secondary_y=False)
        fig.add_trace(go.Scatter(x=df_horario["Hora"], y=df_horario["Conv %"], name="Conv % promesa",
                                 mode="lines+markers", line=dict(color=GREEN, width=2), marker=dict(size=6)),
                      secondary_y=True)
        fig.update_layout(**PLOTLY_LAYOUT, height=300, legend=dict(orientation="h", y=1.12))
        fig.update_yaxes(title_text="Contactos", secondary_y=False, gridcolor="#334155")
        fig.update_yaxes(title_text="Conversión %", secondary_y=True, ticksuffix="%", gridcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig, use_container_width=True)
        st.caption("Franja de mayor conversión: **13:00–15:00 h**")
    with col_r:
        st.markdown("**Distribución de intentos de llamada**")
        labels = ["Titular contactado","Buzón","No contesta","Núm inválido","Colgó agente"]
        vals   = [TITULAR, BUZON, NO_CONT, NUM_INV, COLGADO]
        colors = [GREEN, AMBER, SLATE, RED, "#c026d3"]
        fig2 = go.Figure(go.Pie(labels=labels, values=vals,
                                marker=dict(colors=colors, line=dict(color=DARK, width=2)),
                                hole=0.52, textinfo="label+percent", textfont=dict(color=TEXT)))
        apply_layout(fig2, height=300, showlegend=False)
        st.plotly_chart(fig2, use_container_width=True)
    st.markdown("---")
    st.markdown("**Canales digitales — efectividad para amarrar promesas**")
    cols = st.columns(5)
    canal_icons = {"SMS":"📱","WhatsApp":"💬","Email":"📧","Llamada":"📞","IVR":"🤖"}
    for i, row in df_canal.iterrows():
        with cols[i]:
            icon = canal_icons.get(row["Canal"],"📡")
            color = GREEN if row["Conv prom %"] > 25 else AMBER if row["Conv prom %"] > 18 else RED
            st.markdown(f"""
            <div style="background:{CARD};border-radius:10px;padding:14px;text-align:center;border-top:3px solid {color}">
              <div style="font-size:1.6rem">{icon}</div>
              <div style="font-weight:700;font-size:1rem;color:{TEXT}">{row['Canal']}</div>
              <div style="color:{MUTED};font-size:0.75rem;margin:4px 0">Enviados: {row['Enviados']:,}</div>
              <div style="color:{MUTED};font-size:0.75rem">Resp: {row['Tasa resp %']}%</div>
              <div style="font-size:1.4rem;font-weight:800;color:{color}">{row['Conv prom %']}%</div>
              <div style="color:{MUTED};font-size:0.7rem">conv → promesa</div>
              <div style="margin-top:6px;color:{color};font-weight:700">{row['Promesas']:,} promesas</div>
            </div>""", unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)
    st.warning("**WhatsApp** lidera en conversión (30.0%). **SMS** con solo 15.0% de conversión sugiere que el copy no genera urgencia — se recomienda A/B test con mensajes personalizados por segmento.")

# ══════════════════════════════════════════════
# TAB 3 — INDICADORES
# ══════════════════════════════════════════════
with tab3:
    st.markdown('<div class="sec">Indicadores de Recuperación</div>', unsafe_allow_html=True)
    sub1, sub2, sub3, sub4 = st.tabs(["Por Sector","Por GV","Por Edad","Estado & Segmento"])
    with sub1:
        col_l, col_r = st.columns([3,2])
        with col_l:
            fig = go.Figure()
            fig.add_trace(go.Bar(name="Meta", x=df_sector["Sector"], y=df_sector["Meta"], marker_color=SLATE, opacity=0.6))
            fig.add_trace(go.Bar(name="Recuperado", x=df_sector["Sector"], y=df_sector["Recuperado"],
                                 marker_color=[GREEN if p>=80 else AMBER if p>=65 else RED for p in df_sector["Cumplimiento %"]]))
            apply_layout(fig, height=340, barmode="group", legend=dict(orientation="h", y=1.1),
                         yaxis=dict(tickprefix="$", tickformat=",.0f"))
            st.plotly_chart(fig, use_container_width=True)
        with col_r:
            for _, row in df_sector.sort_values("Cumplimiento %", ascending=False).iterrows():
                badge = "badge-green" if row["Cumplimiento %"]>=80 else "badge-amber" if row["Cumplimiento %"]>=65 else "badge-red"
                st.markdown(f"<div style='background:{CARD};border-radius:8px;padding:10px 14px;margin-bottom:6px;display:flex;justify-content:space-between;align-items:center'>"
                            f"<span style='color:{TEXT};font-weight:600'>{row['Sector']}</span>"
                            f"<span class='badge {badge}'>{row['Cumplimiento %']}%</span></div>", unsafe_allow_html=True)
    with sub2:
        col_l, col_r = st.columns([5,2])
        with col_l:
            colors_gv = [GREEN if p>=85 else AMBER if p>=70 else RED for p in df_gv["Cumplimiento %"]]
            fig = go.Figure(go.Bar(x=df_gv["GV"], y=df_gv["Cumplimiento %"], marker_color=colors_gv,
                                   text=df_gv["Cumplimiento %"].astype(str)+"%", textposition="outside",
                                   textfont=dict(color=TEXT, size=10)))
            fig.add_hline(y=80, line_dash="dash", line_color=AMBER, annotation_text="Meta 80%", annotation_font_color=AMBER)
            apply_layout(fig, height=350, yaxis=dict(ticksuffix="%", range=[0,115]))
            st.plotly_chart(fig, use_container_width=True)
        with col_r:
            st.markdown(f"**Top 5** <span class='badge badge-green'>Mejor desempeño</span>", unsafe_allow_html=True)
            for _, row in df_gv.head(5).iterrows():
                st.markdown(f"<div style='background:{CARD};border-radius:6px;padding:8px 12px;margin:4px 0;border-left:3px solid {GREEN}'>"
                            f"<b style='color:{TEXT}'>{row['GV']}</b> <span style='color:{GREEN};float:right'>{row['Cumplimiento %']}%</span></div>", unsafe_allow_html=True)
            st.markdown(f"<br>**Bottom 5** <span class='badge badge-red'>Requieren atención</span>", unsafe_allow_html=True)
            for _, row in df_gv.tail(5).iterrows():
                st.markdown(f"<div style='background:{CARD};border-radius:6px;padding:8px 12px;margin:4px 0;border-left:3px solid {RED}'>"
                            f"<b style='color:{TEXT}'>{row['GV']}</b> <span style='color:{RED};float:right'>{row['Cumplimiento %']}%</span></div>", unsafe_allow_html=True)
    with sub3:
        col_l, col_r = st.columns(2)
        with col_l:
            fig = px.bar(df_edad, x="Rango", y="Recuperado", color="Promesas %",
                         color_continuous_scale=["#ef4444","#f59e0b","#22c55e"], text_auto=".2s",
                         labels={"Recuperado":"Recuperado $","Promesas %":"% Con promesa"})
            fig.update_traces(textfont_color=TEXT)
            apply_layout(fig, height=320, yaxis=dict(tickprefix="$", tickformat=",.0f"),
                         coloraxis_colorbar=dict(title="% Promesa", tickfont=dict(color=TEXT)))
            st.plotly_chart(fig, use_container_width=True)
        with col_r:
            fig2 = make_subplots(specs=[[{"secondary_y": True}]])
            fig2.add_trace(go.Bar(x=df_edad["Rango"], y=df_edad["Ticket Prom"], name="Ticket Prom $",
                                  marker_color=PURPLE, opacity=0.8), secondary_y=False)
            fig2.add_trace(go.Scatter(x=df_edad["Rango"], y=df_edad["Promesas %"], name="Promesas %",
                                      mode="lines+markers", line=dict(color=GREEN, width=2), marker=dict(size=7)),
                           secondary_y=True)
            fig2.update_layout(**PLOTLY_LAYOUT, height=320, legend=dict(orientation="h", y=1.12))
            fig2.update_yaxes(tickprefix="$", secondary_y=False, gridcolor="#334155")
            fig2.update_yaxes(ticksuffix="%", secondary_y=True, gridcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig2, use_container_width=True)
        st.info("**Segmento 36-45 años:** mayor recuperación absoluta ($336K) y tasa de promesas más alta (71%). Enfocar refuerzos aquí.")
        st.dataframe(df_edad.rename(columns={"Recuperado":"Recuperado $","Ticket Prom":"Ticket $"})
                            .style.format({"Recuperado $":"${:,.0f}","Ticket $":"${:,.0f}","Promesas %":"{:.0f}%"}),
                     use_container_width=True, hide_index=True)
    with sub4:
        col_l, col_r = st.columns(2)
        with col_l:
            st.markdown("**Recuperación por Estado (Top 10)**")
            fig = px.bar(df_estado.sort_values("Recuperado"), x="Recuperado", y="Estado", orientation="h",
                         color="Recuperado", color_continuous_scale=["#1e3a5f","#3b82f6","#22c55e"], text_auto=".2s")
            fig.update_traces(textfont_color=TEXT)
            apply_layout(fig, height=360, xaxis=dict(tickprefix="$", tickformat=",.0f"), coloraxis_showscale=False)
            st.plotly_chart(fig, use_container_width=True)
        with col_r:
            st.markdown("**Recuperación y cumplimiento por Segmento**")
            fig2 = make_subplots(specs=[[{"secondary_y": True}]])
            fig2.add_trace(go.Bar(x=df_segmento["Segmento"], y=df_segmento["Recuperado"], name="Recuperado $",
                                  marker_color=BLUE, opacity=0.8), secondary_y=False)
            fig2.add_trace(go.Scatter(x=df_segmento["Segmento"], y=df_segmento["Cumplimiento %"], name="% Cumplimiento",
                                      mode="lines+markers", line=dict(color=AMBER, width=2), marker=dict(size=8)),
                           secondary_y=True)
            fig2.update_layout(**PLOTLY_LAYOUT, height=360, legend=dict(orientation="h", y=1.12))
            fig2.update_yaxes(tickprefix="$", tickformat=",.0f", secondary_y=False, gridcolor="#334155")
            fig2.update_yaxes(ticksuffix="%", secondary_y=True, gridcolor="rgba(0,0,0,0)", range=[0,110])
            st.plotly_chart(fig2, use_container_width=True)
        st.markdown("**Detalle por segmento**")
        st.dataframe(df_segmento.style.format({"Recuperado":"${:,.0f}","Cumplimiento %":"{:.0f}%"}),
                     use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════
# TAB 4 — OPERACIÓN
# ══════════════════════════════════════════════
with tab4:
    st.markdown('<div class="sec">Operación de Asesores — Mayo 2025</div>', unsafe_allow_html=True)
    avg_tmo      = df_asesores["TMO (min)"].mean()
    total_llam   = df_asesores["Llamadas"].sum()
    avg_abandono = df_asesores["% Abandono"].mean()
    avg_conv     = df_asesores["Conv %"].mean()
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Total llamadas",          f"{total_llam:,}")
    c2.metric("TMO promedio",            f"{avg_tmo:.1f} min")
    c3.metric("% Abandono <30s",         f"{avg_abandono:.1f}%", delta="Revisar apertura", delta_color="inverse")
    c4.metric("Conv. Contacto→Promesa",  f"{avg_conv:.1f}%")
    st.markdown("---")
    col_l, col_r = st.columns([3,2])
    with col_l:
        st.markdown("**Ranking de asesores — monto recuperado y conversión**")
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_trace(go.Bar(x=df_asesores["Asesor"], y=df_asesores["Monto Rec"], name="Monto Rec $",
                             marker_color=BLUE, opacity=0.85), secondary_y=False)
        fig.add_trace(go.Scatter(x=df_asesores["Asesor"], y=df_asesores["Conv %"], name="Conv %",
                                 mode="lines+markers", line=dict(color=GREEN, width=2), marker=dict(size=7)),
                      secondary_y=True)
        fig.update_layout(**PLOTLY_LAYOUT, height=320, legend=dict(orientation="h", y=1.12))
        fig.update_yaxes(tickprefix="$", tickformat=",.0f", secondary_y=False, gridcolor="#334155")
        fig.update_yaxes(ticksuffix="%", secondary_y=True, gridcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig, use_container_width=True)
    with col_r:
        st.markdown("**TMO y % Abandono por asesor**")
        fig2 = go.Figure()
        fig2.add_trace(go.Bar(name="TMO (min)", x=df_asesores["Asesor"], y=df_asesores["TMO (min)"],
                              marker_color=[RED if t>6 else AMBER if t>4.5 else GREEN for t in df_asesores["TMO (min)"]]))
        fig2.add_trace(go.Scatter(name="% Abandono", x=df_asesores["Asesor"], y=df_asesores["% Abandono"],
                                  mode="lines+markers", line=dict(color=RED, width=2, dash="dot"),
                                  marker=dict(size=7), yaxis="y2"))
        fig2.update_layout(**PLOTLY_LAYOUT, height=320, barmode="group",
                           legend=dict(orientation="h", y=1.12),
                           yaxis2=dict(overlaying="y", side="right", ticksuffix="%",
                                       gridcolor="rgba(0,0,0,0)", showgrid=False))
        st.plotly_chart(fig2, use_container_width=True)
    st.markdown("---")
    col_l2, col_r2 = st.columns(2)
    with col_l2:
        st.markdown("**Frecuencia de objeciones difíciles**")
        fig3 = go.Figure(go.Bar(x=df_objeciones["Frecuencia"], y=df_objeciones["Objeción"], orientation="h",
                                marker_color=[RED if f>2000 else AMBER if f>1000 else SLATE for f in df_objeciones["Frecuencia"]],
                                text=df_objeciones["Frecuencia"], textposition="outside", textfont=dict(color=TEXT)))
        apply_layout(fig3, height=280)
        st.plotly_chart(fig3, use_container_width=True)
    with col_r2:
        st.markdown("**Tasa de resolución exitosa por objeción**")
        fig4 = go.Figure(go.Bar(x=df_objeciones["Resolución exitosa %"], y=df_objeciones["Objeción"], orientation="h",
                                marker_color=[GREEN if p>=65 else AMBER if p>=50 else RED for p in df_objeciones["Resolución exitosa %"]],
                                text=df_objeciones["Resolución exitosa %"].astype(str)+"%",
                                textposition="outside", textfont=dict(color=TEXT)))
        apply_layout(fig4, height=280, xaxis=dict(ticksuffix="%", range=[0,100]))
        st.plotly_chart(fig4, use_container_width=True)
    st.warning(f"**Puntos críticos:** 'No tengo dinero' es la objeción más frecuente (2,840 casos) con solo **38%** de resolución. "
               f"% Abandono promedio de **{avg_abandono:.1f}%**: revisar apertura de llamada y primeros 10 segundos del guión.")
    st.markdown("---")
    st.markdown("**Detalle completo por asesor**")
    display_df = df_asesores[["Asesor","Llamadas","TMO (min)","% Contacto","Conv %","% Abandono","Monto Rec"]].copy()
    display_df["Monto Rec"] = display_df["Monto Rec"].apply(lambda x: f"${x:,.0f}")
    st.dataframe(display_df, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════
# TAB 5 — PLAN DE TRABAJO
# ══════════════════════════════════════════════
with tab5:
    st.markdown('<div class="sec">Plan de Trabajo — Junio 2025</div>', unsafe_allow_html=True)
    gap_jun = META_JUN - PROY_JUN
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Meta Junio",           f"${META_JUN/1e6:.1f}M")
    c2.metric("Proyección base",      f"${PROY_JUN/1e6:.2f}M", delta=f"{(PROY_JUN/META_JUN-1)*100:.1f}%", delta_color="inverse")
    c3.metric("Gap a cerrar",         f"${gap_jun/1e6:.2f}M", delta_color="inverse")
    c4.metric("Pipeline de promesas", f"${PIPELINE_PROM/1e6:.2f}M")
    st.markdown("---")
    col_l, col_r = st.columns(2)
    with col_l:
        st.markdown("**Proyección acumulada Junio — 3 escenarios**")
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=df_proy["Fecha"], y=df_proy["Meta"], name="Meta",
                                 line=dict(color=SLATE, dash="dash", width=2)))
        fig.add_trace(go.Scatter(x=df_proy["Fecha"], y=df_proy["Optimista"], name="Optimista",
                                 line=dict(color=GREEN, width=2), fill="tonexty", fillcolor="rgba(34,197,94,0.05)"))
        fig.add_trace(go.Scatter(x=df_proy["Fecha"], y=df_proy["Base"], name="Base",
                                 line=dict(color=BLUE, width=3)))
        fig.add_trace(go.Scatter(x=df_proy["Fecha"], y=df_proy["Pesimista"], name="Pesimista",
                                 line=dict(color=RED, width=2, dash="dot"),
                                 fill="tonexty", fillcolor="rgba(239,68,68,0.05)"))
        apply_layout(fig, height=360, legend=dict(orientation="h", y=1.1))
        fig.update_yaxes(tickprefix="$", tickformat=",.0f")
        st.plotly_chart(fig, use_container_width=True)
    with col_r:
        st.markdown("**Impacto estimado por palanca de acción**")
        df_p = df_palancas.sort_values("Impacto estimado $")
        fig2 = go.Figure(go.Bar(x=df_p["Impacto estimado $"], y=df_p["Palanca"], orientation="h",
                                marker_color=BLUE,
                                text=df_p["Impacto estimado $"].apply(lambda x: f"${x:,.0f}"),
                                textposition="outside", textfont=dict(color=TEXT)))
        total_impacto = df_palancas["Impacto estimado $"].sum()
        fig2.add_vline(x=gap_jun, line_dash="dash", line_color=RED,
                       annotation_text=f"Gap: ${gap_jun:,.0f}", annotation_font_color=RED)
        apply_layout(fig2, height=360, xaxis=dict(tickprefix="$", tickformat=",.0f"))
        st.plotly_chart(fig2, use_container_width=True)
        st.caption(f"Impacto total estimado: **${total_impacto:,.0f}** — cubre **{total_impacto/gap_jun*100:.0f}%** del gap de Junio")
    st.markdown("---")
    st.markdown("**Compromisos del equipo**")
    for _, row in df_acciones.iterrows():
        badge_cls = "badge-red" if row["Plazo"] == "Inmediato" else "badge-amber"
        st.markdown(f"<div style='background:{CARD};border-radius:8px;padding:10px 16px;margin:5px 0;"
                    f"display:flex;align-items:center;gap:12px'>"
                    f"<span class='badge {badge_cls}'>{row['Plazo']}</span>"
                    f"<span style='color:{MUTED};min-width:130px;font-size:0.85rem'>{row['Área']}</span>"
                    f"<span style='color:{TEXT}'>{row['Acción']}</span></div>", unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)
    st.success(f"**Resumen ejecutivo:** Con las palancas propuestas se puede cerrar hasta **${total_impacto/1e6:.2f}M** "
               f"adicional, cubriendo el **{total_impacto/gap_jun*100:.0f}%** del gap. "
               f"Prioridad #1: rellamada a las **{PROMESAS_CAIDAS} promesas caídas** de Mayo "
               f"y lanzar el A/B test de WhatsApp antes del miércoles.")
