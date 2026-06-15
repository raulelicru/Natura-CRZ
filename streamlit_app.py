"""
Dashboard de Cobranza Mayo 2025 — NAtura
"""
import random
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import openpyxl

random.seed(2025)
np.random.seed(2025)

st.set_page_config(page_title="Cobranza Mayo 2025 — NAtura", page_icon="📊",
                   layout="wide", initial_sidebar_state="collapsed")

BG="#f8fafc"; CARD="#ffffff"; BORD="#e2e8f0"
BLUE="#2563eb"; GREEN="#16a34a"; RED="#dc2626"; AMBER="#d97706"
PURPLE="#7c3aed"; SLATE="#64748b"; TEXT="#0f172a"; MUTED="#64748b"
PLOT_BG="rgba(0,0,0,0)"

st.markdown(f"""
<style>
  .stApp{{background-color:{BG};color:{TEXT};}}
  [data-testid="stAppViewContainer"]{{background-color:{BG};}}
  [data-testid="stHeader"]{{background-color:{BG};border-bottom:1px solid {BORD};}}
  [data-testid="stSidebar"]{{background-color:{CARD};border-right:1px solid {BORD};}}
  .stTabs [data-baseweb="tab-list"]{{background:{BORD};border-radius:10px;padding:4px;gap:4px;}}
  .stTabs [data-baseweb="tab"]{{color:{MUTED};border-radius:8px;font-weight:600;font-size:0.85rem;}}
  .stTabs [aria-selected="true"]{{background:{BLUE}!important;color:white!important;}}
  [data-testid="stMetric"]{{background:{CARD};border-radius:10px;padding:14px 18px;
      border-left:3px solid {BLUE};border:1px solid {BORD};box-shadow:0 1px 3px rgba(0,0,0,0.06);}}
  [data-testid="stMetricValue"]{{color:{TEXT}!important;font-size:1.8rem!important;}}
  [data-testid="stMetricLabel"]{{color:{MUTED}!important;}}
  .sec{{font-size:1.2rem;font-weight:700;color:{TEXT};border-bottom:2px solid {BORD};
        padding-bottom:6px;margin:8px 0 16px 0;}}
  .badge{{display:inline-block;padding:2px 10px;border-radius:20px;font-size:0.78rem;font-weight:700;}}
  .badge-red{{background:#fee2e2;color:{RED};border:1px solid #fca5a5;}}
  .badge-green{{background:#dcfce7;color:{GREEN};border:1px solid #86efac;}}
  .badge-amber{{background:#fef3c7;color:{AMBER};border:1px solid #fcd34d;}}
  [data-testid="stDataFrame"]{{border-radius:10px;}}
  [data-testid="stAlert"]{{border-radius:10px;}}
</style>
""", unsafe_allow_html=True)

PLOTLY_LAYOUT=dict(paper_bgcolor=PLOT_BG,plot_bgcolor=PLOT_BG,
                   font=dict(color=TEXT,family="Inter, sans-serif"),margin=dict(l=10,r=10,t=30,b=10))

def apply_layout(fig,**kwargs):
    layout={**PLOTLY_LAYOUT,**kwargs}
    for ax in("xaxis","yaxis","xaxis2","yaxis2"):
        base=dict(gridcolor="#e2e8f0",linecolor="#cbd5e1")
        if ax in kwargs: base.update(kwargs[ax]); layout[ax]=base
    fig.update_layout(**layout); return fig

def leer_archivo(f):
    if f is None: return None
    try:
        if f.name.lower().endswith((".xlsx",".xls")): return pd.read_excel(f)
        return pd.read_csv(f,encoding="utf-8-sig",low_memory=False)
    except Exception:
        try: f.seek(0); return pd.read_csv(f,encoding="latin-1",low_memory=False)
        except: return None

def col(df,*ops):
    for o in ops:
        if o in df.columns: return o
    return None

# ── SIDEBAR ──
with st.sidebar:
    st.markdown("## 📂 Cargar datos reales")
    st.caption("Sube tus archivos de Mayo.")
    f_cartera =st.file_uploader("📋 Cartera / Remesa",    type=["csv","xlsx","xls"],key="cartera")
    f_pagos   =st.file_uploader("💰 Pagos / Recuperación",type=["csv","xlsx","xls"],key="pagos")
    f_gestion =st.file_uploader("📞 Gestión de llamadas", type=["csv","xlsx","xls"],key="gestion")
    f_promesas=st.file_uploader("🤝 Promesas de pago",    type=["csv","xlsx","xls"],key="promesas")
    f_comparativo=st.file_uploader("📑 Comparativo Cobranza (Pagos + Comparativo)",type=["xlsx","xls"],key="comparativo")
    MODO_REAL=any([f_cartera,f_pagos,f_gestion,f_promesas,f_comparativo])
    st.markdown("---")
    btn_ejecutar = st.button("▶ Ejecutar", type="primary", use_container_width=True,
                             help="Procesa los archivos cargados y actualiza el dashboard")
    if btn_ejecutar:
        st.session_state["ejecutado"] = True
        st.rerun()
    if st.button("↺ Limpiar", use_container_width=True):
        st.session_state["ejecutado"] = False
        st.rerun()
    st.markdown("---")
    if st.session_state.get("ejecutado"):
        if MODO_REAL:
            st.success("✅ Datos reales activos")
        else:
            st.info("ℹ️ Usando datos de prueba")
    else:
        if MODO_REAL:
            st.warning("⬆️ Archivos listos — presiona **Ejecutar**")
        else:
            st.caption("Carga archivos o presiona Ejecutar para ver datos de prueba.")

# ── PANTALLA DE ESPERA ──
if not st.session_state.get("ejecutado"):
    st.markdown("## 📊 Dashboard Ejecutivo — Cobranza Mayo 2025")
    st.caption("NAtura | Reunión de seguimiento")
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        f"<div style='text-align:center;padding:60px 20px;background:{CARD};border-radius:16px;"
        f"border:2px dashed {BORD};margin-top:20px'>"
        f"<div style='font-size:3rem'>📂</div>"
        f"<div style='font-size:1.4rem;font-weight:700;color:{TEXT};margin:12px 0'>Carga tus archivos y presiona Ejecutar</div>"
        f"<div style='color:{MUTED};font-size:0.95rem'>O bien presiona <b>Ejecutar</b> directamente para ver el dashboard con datos de ejemplo.</div>"
        f"<div style='margin-top:24px;color:{MUTED};font-size:0.85rem'>"
        f"📋 Cartera &nbsp;|&nbsp; 💰 Pagos &nbsp;|&nbsp; 📞 Gestión &nbsp;|&nbsp; 🤝 Promesas &nbsp;|&nbsp; 📑 Comparativo</div>"
        f"</div>",
        unsafe_allow_html=True
    )
    st.stop()

# ── DATOS SINTÉTICOS ──
ESTADOS=["CDMX","Edo. de México","Jalisco","Nuevo León","Puebla",
         "Guanajuato","Veracruz","Michoacán","Chihuahua","Tamaulipas"]
SECTORES=["Belleza","Moda","Hogar","Nutrición","Bienestar","Joyería"]
SEGMENTOS=["Diamante","Oro","Plata","Bronce","Nuevo Ingreso"]
GVS=[f"GV-{str(i).zfill(3)}" for i in range(1,16)]
ASESORES=[f"Asesor {chr(65+i)}" for i in range(10)]
HORAS=list(range(8,21))
DIAS_SEM=["Lunes","Martes","Miércoles","Jueves","Viernes"]

META_TOTAL=4_850_000; RECUPERADO_TOTAL=3_892_500
PROMESAS_GEN=1_240; PROMESAS_CUMP=748; PROMESAS_CAIDAS=PROMESAS_GEN-PROMESAS_CUMP

dias=pd.date_range("2025-05-01","2025-05-31",freq="B")
meta_d=META_TOTAL/len(dias)
rec_d=np.clip(np.random.normal(meta_d*0.92,meta_d*0.15,len(dias)),meta_d*0.4,meta_d*1.3)
df_cierre=pd.DataFrame({"Fecha":dias,"Meta":meta_d,"Recuperado":rec_d})
df_cierre["Meta Acum"]=df_cierre["Meta"].cumsum()
df_cierre["Recuperado Acum"]=df_cierre["Recuperado"].cumsum()

df_motivos=pd.DataFrame({
    "Motivo":["Sin liquidez","No contestó","Negó deuda","Promesa vencida","Número inválido","Otro"],
    "Casos":[198,167,89,134,72,61],
})

TOTAL_INT=38_420; TITULAR=14_820; BUZON=9_105
NO_CONT=8_934; NUM_INV=3_210; COLGADO=2_351; CR=TITULAR/TOTAL_INT

df_horario=pd.DataFrame({
    "Hora":HORAS,
    "Contactos":[120,310,520,680,740,810,860,780,690,620,480,310,210],
    "Promesas":[ 12, 35, 62, 88, 95,104,112,100, 89, 78, 61, 39, 25],
})
df_horario["Conv %"]=(df_horario["Promesas"]/df_horario["Contactos"]*100).round(1)

df_canal=pd.DataFrame({
    "Canal":["SMS","WhatsApp","Email","Llamada","IVR"],
    "Enviados":[12400,8300,9800,38420,4200],
    "Respuestas":[1240,2076,588,14820,840],
    "Promesas":[186,622,74,748,168],
})
df_canal["Tasa resp %"]=(df_canal["Respuestas"]/df_canal["Enviados"]*100).round(1)
df_canal["Conv prom %"]=(df_canal["Promesas"]/df_canal["Respuestas"]*100).round(1)

df_sector=pd.DataFrame({
    "Sector":SECTORES,
    "Meta":[980000,870000,650000,720000,540000,490000],
    "Recuperado":[784000,652500,487500,504000,378000,318000],
})
df_sector["Cumplimiento %"]=(df_sector["Recuperado"]/df_sector["Meta"]*100).round(1)

np.random.seed(42)
df_gv=pd.DataFrame({"GV":GVS,
    "Meta":np.random.randint(200000,500000,len(GVS)),
    "Recuperado":np.random.randint(140000,460000,len(GVS))})
df_gv["Cumplimiento %"]=(df_gv["Recuperado"]/df_gv["Meta"]*100).round(1)
df_gv=df_gv.sort_values("Cumplimiento %",ascending=False).reset_index(drop=True)

df_edad=pd.DataFrame({
    "Rango":["18-25","26-35","36-45","46-55","56-65","66+"],
    "Cuentas":[1240,3870,4210,3650,2180,980],
    "Recuperado":[74400,271400,336800,255500,130800,49000],
    "Promesas %":[38,62,71,65,52,41],
})
df_edad["Ticket Prom"]=(df_edad["Recuperado"]/df_edad["Cuentas"]).round(0)

df_estado=pd.DataFrame({
    "Estado":ESTADOS,
    "Cuentas":[4820,3610,2840,1980,1650,1420,1280,1190,980,870],
    "Recuperado":[1120000,892000,620000,485000,378000,312000,280000,245000,198000,168000],
})
df_estado["Ticket Prom"]=(df_estado["Recuperado"]/df_estado["Cuentas"]).round(0)

df_segmento=pd.DataFrame({
    "Segmento":SEGMENTOS,
    "Cuentas":[890,2340,4210,5680,2320],
    "Recuperado":[534000,936000,1263000,1022400,139200],
    "Cumplimiento %":[87,82,74,68,45],
})

np.random.seed(99)
df_asesores=pd.DataFrame({
    "Asesor":ASESORES,
    "Llamadas":np.random.randint(280,520,10),
    "TMO (min)":np.random.uniform(3.2,6.8,10).round(1),
    "Contactos Titular":np.random.randint(80,180,10),
    "Promesas":np.random.randint(20,65,10),
    "Colgadas <30s":np.random.randint(15,60,10),
    "Monto Rec":np.random.randint(180000,520000,10),
})
df_asesores["% Contacto"]=(df_asesores["Contactos Titular"]/df_asesores["Llamadas"]*100).round(1)
df_asesores["% Abandono"]=(df_asesores["Colgadas <30s"]/df_asesores["Llamadas"]*100).round(1)
df_asesores["Conv %"]=(df_asesores["Promesas"]/df_asesores["Contactos Titular"]*100).round(1)
df_asesores=df_asesores.sort_values("Monto Rec",ascending=False).reset_index(drop=True)

df_objeciones=pd.DataFrame({
    "Objeción":["No tengo dinero","No reconozco la deuda","Ya pagué",
               "Espera a quincena","Mándame estado de cuenta","No soy yo"],
    "Frecuencia":[2840,1920,1480,2210,890,640],
    "Resolución exitosa %":[38,51,72,61,44,58],
})

META_JUN=5_200_000; PROY_JUN=4_420_000; PIPELINE_PROM=1_840_000
dias_jun=pd.date_range("2025-06-01","2025-06-30",freq="B"); n=len(dias_jun)
df_proy=pd.DataFrame({"Fecha":dias_jun,
    "Meta":np.linspace(0,META_JUN,n),"Base":np.linspace(0,PROY_JUN,n),
    "Optimista":np.linspace(0,META_JUN*1.05,n),"Pesimista":np.linspace(0,PROY_JUN*0.88,n)})
df_palancas=pd.DataFrame({
    "Palanca":["Recuperar promesas caídas Mayo","Incrementar contact rate",
               "Mejorar conv. WhatsApp","Script objeciones","GVs rezagados"],
    "Impacto estimado $":[420000,280000,180000,150000,220000],
})
df_acciones=pd.DataFrame([
    ("Inmediato","Operaciones","Rellamada a 2,891 promesas caídas"),
    ("Inmediato","Analítica","Reporte GV diario automatizado"),
    ("Lun 9 Jun","Canales Digitales","Lanzar A/B test SMS por segmento"),
    ("Mar 10 Jun","Supervisión","Coaching script objeciones 'sin dinero'"),
    ("Jue 12 Jun","Gerencia","Validar incentivo para GVs rezagados"),
    ("Vie 13 Jun","Operaciones","Ajuste de marcaciones franja 12-15h +25%"),
],columns=["Plazo","Área","Acción"])

# ── DATOS SINTÉTICOS NUEVOS INDICADORES ──
df_temp=pd.DataFrame({
    "Semana":["Sem 1 (1-7)","Sem 2 (8-14)","Sem 3 (15-21)","Sem 4 (22-31)"],
    "Meta":[META_TOTAL*0.22,META_TOTAL*0.26,META_TOTAL*0.26,META_TOTAL*0.26],
    "Recuperado":[META_TOTAL*0.18,META_TOTAL*0.22,META_TOTAL*0.24,META_TOTAL*0.16],
})
df_temp["Cumplimiento %"]=(df_temp["Recuperado"]/df_temp["Meta"]*100).round(1)
df_quincena=pd.DataFrame({
    "Período":["1era Quincena (1-15)","2da Quincena (16-31)"],
    "Meta":[META_TOTAL*0.48,META_TOTAL*0.52],
    "Recuperado":[META_TOTAL*0.40,META_TOTAL*0.40],
})
df_quincena["Cumplimiento %"]=(df_quincena["Recuperado"]/df_quincena["Meta"]*100).round(1)

np.random.seed(77)
heatmap_data=np.random.randint(40,200,(5,13))
heatmap_data[0,:]*=1.1; heatmap_data[2,:]*=1.05
heatmap_data[:,4:7]*=1.2
heatmap_data=heatmap_data.astype(int)

df_hora_sem=pd.DataFrame(index=HORAS,columns=["Sem 1","Sem 2","Sem 3","Sem 4"])
for sem,factor in zip(["Sem 1","Sem 2","Sem 3","Sem 4"],[0.85,1.0,1.1,0.95]):
    base=[120,310,520,680,740,810,860,780,690,620,480,310,210]
    df_hora_sem[sem]=[int(v*factor+np.random.randint(-20,20)) for v in base]
df_hora_sem=df_hora_sem.reset_index().rename(columns={"index":"Hora"})

np.random.seed(55)
n_cons=300
edades_cons=np.random.randint(18,68,n_cons)
seg_cons=np.random.choice(SEGMENTOS,n_cons,p=[0.06,0.15,0.28,0.37,0.14])
montos_cons=np.where(edades_cons<30, np.random.uniform(200,1500,n_cons),
            np.where(edades_cons<45, np.random.uniform(400,3000,n_cons),
            np.where(edades_cons<55, np.random.uniform(300,2500,n_cons),
                                     np.random.uniform(150,1800,n_cons))))
df_scatter_edad=pd.DataFrame({"Edad":edades_cons,"Monto Pagado":montos_cons.round(0),"Segmento":seg_cons})

# ── LECTURA DE ARCHIVOS REALES ──
df_cart_real =leer_archivo(f_cartera)  if f_cartera  else None
df_pago_real =leer_archivo(f_pagos)   if f_pagos    else None
df_gest_real =leer_archivo(f_gestion) if f_gestion  else None
df_prom_real =leer_archivo(f_promesas)if f_promesas else None

if df_pago_real is not None:
    c_monto=col(df_pago_real,"monto_pagado","valor_pago","importe","monto","pago")
    c_fecha=col(df_pago_real,"fecha_pago","fecha","date","fecha_operacion")
    c_asesor=col(df_pago_real,"asesor","ejecutivo","agente","nombre_asesor")
    RECUPERADO_TOTAL=df_pago_real[c_monto].sum() if c_monto else RECUPERADO_TOTAL
    if c_fecha and c_monto:
        df_pago_real[c_fecha]=pd.to_datetime(df_pago_real[c_fecha],errors="coerce")
        df_p=df_pago_real.dropna(subset=[c_fecha])
        df_p["_sem"]=df_p[c_fecha].dt.isocalendar().week
        df_p["_dia_sem"]=df_p[c_fecha].dt.day_name()
        df_cierre=(df_p.groupby(c_fecha)[c_monto].sum().reset_index()
                   .rename(columns={c_fecha:"Fecha",c_monto:"Recuperado"}))
        df_cierre=df_cierre[df_cierre["Fecha"].dt.month==5].sort_values("Fecha")
        df_cierre["Meta"]=META_TOTAL/max(len(df_cierre),1)
        df_cierre["Meta Acum"]=df_cierre["Meta"].cumsum()
        df_cierre["Recuperado Acum"]=df_cierre["Recuperado"].cumsum()
        df_p2=df_p[df_p[c_fecha].dt.month==5].copy()
        df_p2["_sem_mes"]=pd.cut(df_p2[c_fecha].dt.day,bins=[0,7,14,21,31],
                                  labels=["Sem 1 (1-7)","Sem 2 (8-14)","Sem 3 (15-21)","Sem 4 (22-31)"])
        rec_sem=df_p2.groupby("_sem_mes",observed=True)[c_monto].sum().reset_index()
        rec_sem.columns=["Semana","Recuperado"]
        df_temp=df_temp.merge(rec_sem,on="Semana",how="left",suffixes=("","_r"))
        if "Recuperado_r" in df_temp.columns:
            df_temp["Recuperado"]=df_temp["Recuperado_r"].fillna(df_temp["Recuperado"])
            df_temp.drop(columns=["Recuperado_r"],inplace=True)
        df_temp["Cumplimiento %"]=(df_temp["Recuperado"]/df_temp["Meta"]*100).round(1)

if df_prom_real is not None:
    c_mp=col(df_prom_real,"monto_promesa","promesa_monto","monto","importe")
    c_cum=col(df_prom_real,"cumplida","estatus","status","resultado")
    PROMESAS_GEN=len(df_prom_real)
    if c_cum:
        cumplidas=df_prom_real[c_cum].astype(str).str.lower().isin(["1","si","sí","cumplida","pagada","true"])
        PROMESAS_CUMP=int(cumplidas.sum()); PROMESAS_CAIDAS=PROMESAS_GEN-PROMESAS_CUMP
    if c_mp: PIPELINE_PROM=df_prom_real[c_mp].sum()

if df_gest_real is not None:
    c_res=col(df_gest_real,"resultado","resultado_gestion","disposicion","tipificacion")
    c_hora=col(df_gest_real,"hora_llamada","hora","hour")
    c_fecha_g=col(df_gest_real,"fecha_llamada","fecha","date")
    TOTAL_INT=len(df_gest_real)
    if c_res:
        res_lower=df_gest_real[c_res].astype(str).str.lower()
        TITULAR=int(res_lower.isin(["contacto titular","titular","contactado","promesa","pdc"]).sum())
        BUZON=int(res_lower.str.contains("buz|voicemail|vm",na=False).sum())
        NO_CONT=int(res_lower.str.contains("no contest|no answer|sin respuesta",na=False).sum())
        NUM_INV=int(res_lower.str.contains("inv[aá]lido|wrong|error",na=False).sum())
        COLGADO=int(res_lower.str.contains("colg|hang|abandon",na=False).sum())
        CR=TITULAR/max(TOTAL_INT,1)
    if c_hora and c_fecha_g and c_res:
        try:
            df_gest_real["_hora"]=pd.to_datetime(df_gest_real[c_hora],errors="coerce").dt.hour
            df_gest_real["_fecha"]=pd.to_datetime(df_gest_real[c_fecha_g],errors="coerce")
            df_gest_real["_dia_sem"]=df_gest_real["_fecha"].dt.day_name()
            df_gest_real["_sem_mes"]=pd.cut(df_gest_real["_fecha"].dt.day,bins=[0,7,14,21,31],
                                            labels=["Sem 1","Sem 2","Sem 3","Sem 4"])
            dias_map={"Monday":"Lunes","Tuesday":"Martes","Wednesday":"Miércoles",
                      "Thursday":"Jueves","Friday":"Viernes"}
            df_gest_real["_dia_esp"]=df_gest_real["_dia_sem"].map(dias_map)
            df_hw=df_gest_real.dropna(subset=["_hora","_dia_esp"])
            is_tit=df_hw[c_res].astype(str).str.lower().isin(["contacto titular","titular","contactado","promesa","pdc"])
            df_hw2=df_hw[is_tit].groupby(["_dia_esp","_hora"]).size().reset_index(name="Contactos")
            pivot=df_hw2.pivot(index="_dia_esp",columns="_hora",values="Contactos").fillna(0)
            heatmap_data=pivot.reindex(DIAS_SEM).values.astype(int)
            df_hs=df_gest_real.dropna(subset=["_hora","_sem_mes"])
            is_tit2=df_hs[c_res].astype(str).str.lower().isin(["contacto titular","titular","contactado","promesa","pdc"])
            df_hsw=df_hs[is_tit2].groupby(["_sem_mes","_hora"],observed=True).size().reset_index(name="Contactos")
            df_hora_sem2=df_hsw.pivot(index="_hora",columns="_sem_mes",values="Contactos").fillna(0).reset_index()
            df_hora_sem2.columns.name=None; df_hora_sem2.rename(columns={"_hora":"Hora"},inplace=True)
            df_hora_sem=df_hora_sem2
        except Exception:
            pass

if df_cart_real is not None:
    c_saldo=col(df_cart_real,"valor_saldo_deuda","saldo","deuda","monto_deuda","importe")
    c_zona=col(df_cart_real,"zona","zone","gv","gerencia")
    c_seg=col(df_cart_real,"segmento_nombre","segmento","segment","categoria")
    c_estado_c=col(df_cart_real,"direccion_de_residencia_estado","estado","state","entidad")
    c_edad_c=col(df_cart_real,"edad_consultora","edad","age")
    if c_saldo: META_TOTAL=df_cart_real[c_saldo].sum()
    if c_edad_c and df_pago_real is not None:
        c_monto2=col(df_pago_real,"monto_pagado","valor_pago","importe","monto","pago")
        c_edad_p=col(df_pago_real,"edad_consultora","edad")
        c_seg_p=col(df_pago_real,"segmento_nombre","segmento","segment","categoria")
        if c_edad_p and c_monto2:
            df_scatter_edad=df_pago_real[[c_edad_p,c_monto2]].copy()
            df_scatter_edad.columns=["Edad","Monto Pagado"]
            if c_seg_p: df_scatter_edad["Segmento"]=df_pago_real[c_seg_p]
            else: df_scatter_edad["Segmento"]="Sin segmento"
            df_scatter_edad=df_scatter_edad.dropna()

# ── COMPARATIVO COBRANZA POR TRAMO (sincronización Pagos → Comparativo) ──
df_pagos_diario=None
df_comparativo=None
RESUMEN_COMP=None
DEBUG_COMP=None
if f_comparativo is not None:
    DEBUG_COMP={}
    try:
        wb=openpyxl.load_workbook(f_comparativo,data_only=True)
        DEBUG_COMP["hojas"]=wb.sheetnames
        sh_pagos=next((s for s in wb.sheetnames if "PAGO" in s.upper()),None)
        sh_comp =next((s for s in wb.sheetnames if "COMPARATIV" in s.upper()),None)
        DEBUG_COMP["hoja_pagos"]=sh_pagos
        DEBUG_COMP["hoja_comparativo"]=sh_comp
        if sh_pagos and sh_comp:
            ws_p=wb[sh_pagos]; ws_c=wb[sh_comp]

            # localizar fila de encabezados (T1..Tn) en hoja de pagos
            header_row=None
            for r in range(1,min(ws_p.max_row,10)+1):
                row_vals=[ws_p.cell(row=r,column=c).value for c in range(1,ws_p.max_column+1)]
                if any(isinstance(v,str) and v.strip().upper()=="T1" for v in row_vals):
                    header_row=r; break
            tramo_cols={}
            if header_row:
                for c in range(1,ws_p.max_column+1):
                    v=ws_p.cell(row=header_row,column=c).value
                    if isinstance(v,str):
                        vu=v.strip().upper()
                        if len(vu)>=2 and vu[0]=="T" and vu[1:].isdigit():
                            tramo_cols[vu]=c
            DEBUG_COMP["header_row_pagos"]=header_row
            DEBUG_COMP["tramo_cols"]=tramo_cols

            # localizar fila TOTAL
            total_row=None
            if header_row:
                for r in range(header_row+1,ws_p.max_row+1):
                    v1=ws_p.cell(row=r,column=1).value
                    v2=ws_p.cell(row=r,column=2).value
                    if (isinstance(v1,str) and "TOTAL" in v1.upper()) or (isinstance(v2,str) and "TOTAL" in v2.upper()):
                        total_row=r; break

            DEBUG_COMP["total_row_pagos"]=total_row

            # totales por tramo (fila TOTAL)
            totales_tramo={t:float(ws_p.cell(row=total_row,column=c).value or 0) for t,c in tramo_cols.items()} if total_row else {}
            DEBUG_COMP["totales_tramo"]=totales_tramo

            # pagos diarios (columna B = fecha, suma de tramos por día)
            registros=[]
            if header_row and total_row:
                for r in range(header_row+1,total_row):
                    fecha=ws_p.cell(row=r,column=2).value
                    if fecha is None: continue
                    total_dia=sum(float(ws_p.cell(row=r,column=c).value or 0) for c in tramo_cols.values())
                    registros.append({"Fecha":fecha,"Total Día":total_dia})
            if registros:
                df_pagos_diario=pd.DataFrame(registros)
                df_pagos_diario["Fecha"]=pd.to_datetime(df_pagos_diario["Fecha"],errors="coerce")
                df_pagos_diario=df_pagos_diario.dropna(subset=["Fecha"]).sort_values("Fecha").reset_index(drop=True)
                df_pagos_diario["Acumulado"]=df_pagos_diario["Total Día"].cumsum()

            # localizar encabezados en hoja Comparativo
            header_row_c=None
            for r in range(1,min(ws_c.max_row,10)+1):
                row_vals=[ws_c.cell(row=r,column=c).value for c in range(1,ws_c.max_column+1)]
                if any(isinstance(v,str) and "CARTERA" in v.upper() for v in row_vals):
                    header_row_c=r; break

            DEBUG_COMP["header_row_comparativo"]=header_row_c

            cols_c={}
            if header_row_c:
                for c in range(1,ws_c.max_column+1):
                    v=ws_c.cell(row=header_row_c,column=c).value
                    if not isinstance(v,str): continue
                    vu=v.strip().upper()
                    if "CARTERA" in vu: cols_c["Cartera"]=c
                    elif "OBJETIVO" in vu and "%" in vu: cols_c["Objetivo %"]=c
                    elif "OBJETIVO" in vu: cols_c["Objetivo $"]=c
                    elif "CUENTA" in vu: cols_c["Cuentas"]=c
                    elif "COBRANZA" in vu: cols_c["Cobranza"]=c
                    elif "TRAMO" in vu or "TEMPORALIDAD" in vu: cols_c["Tramo"]=c
            DEBUG_COMP["cols_comparativo"]=cols_c

            # valores de la columna Tramo detectados (para comparar con totales_tramo)
            tramo_vals_c=[]
            if header_row_c:
                col_tramo_dbg=cols_c.get("Tramo",1)
                for r in range(header_row_c+1,ws_c.max_row+1):
                    v=ws_c.cell(row=r,column=col_tramo_dbg).value
                    if v is not None:
                        tramo_vals_c.append(v)
            DEBUG_COMP["tramo_vals_comparativo"]=tramo_vals_c

            # sincronizar Cobranza $ por tramo (escribe el valor calculado en la hoja en memoria)
            comp_rows=[]
            if header_row_c and "Cobranza" in cols_c:
                col_tramo=cols_c.get("Tramo",1)
                for r in range(header_row_c+1,ws_c.max_row+1):
                    tramo_val=ws_c.cell(row=r,column=col_tramo).value
                    if not isinstance(tramo_val,str): continue
                    tramo_key=tramo_val.strip().upper()
                    if tramo_key not in totales_tramo: continue
                    cobranza_val=totales_tramo[tramo_key]
                    ws_c.cell(row=r,column=cols_c["Cobranza"]).value=cobranza_val
                    comp_rows.append({
                        "Tramo":tramo_val.strip(),
                        "Cuentas":ws_c.cell(row=r,column=cols_c["Cuentas"]).value if "Cuentas" in cols_c else None,
                        "Cartera $":float(ws_c.cell(row=r,column=cols_c["Cartera"]).value or 0) if "Cartera" in cols_c else 0.0,
                        "Objetivo $":float(ws_c.cell(row=r,column=cols_c["Objetivo $"]).value or 0) if "Objetivo $" in cols_c else 0.0,
                        "Objetivo %":ws_c.cell(row=r,column=cols_c["Objetivo %"]).value if "Objetivo %" in cols_c else None,
                        "Cobranza $":cobranza_val,
                    })

            if comp_rows:
                df_comparativo=pd.DataFrame(comp_rows)
                df_comparativo["Cumplimiento %"]=(df_comparativo["Cobranza $"]/df_comparativo["Objetivo $"].replace(0,np.nan)*100).round(1)

                total_cartera =df_comparativo["Cartera $"].sum()
                total_objetivo=df_comparativo["Objetivo $"].sum()
                total_cobranza=df_comparativo["Cobranza $"].sum()

                if df_pagos_diario is not None and len(df_pagos_diario)>0:
                    ult_fecha=df_pagos_diario["Fecha"].max()
                    dias_transcurridos=df_pagos_diario["Fecha"].nunique()
                    dias_totales_mes=pd.Period(ult_fecha,freq="M").days_in_month
                    dias_pendientes=max(dias_totales_mes-ult_fecha.day,0)
                    promedio_diario=total_cobranza/max(dias_transcurridos,1)
                    proyeccion_total=total_cobranza+promedio_diario*dias_pendientes
                else:
                    ult_fecha=None; dias_transcurridos=0; dias_totales_mes=0
                    dias_pendientes=0; promedio_diario=0.0; proyeccion_total=total_cobranza

                RESUMEN_COMP=dict(
                    total_cartera=total_cartera,total_objetivo=total_objetivo,total_cobranza=total_cobranza,
                    cumplimiento=(total_cobranza/total_objetivo*100) if total_objetivo else 0.0,
                    dias_transcurridos=dias_transcurridos,dias_pendientes=dias_pendientes,
                    dias_totales_mes=dias_totales_mes,promedio_diario=promedio_diario,
                    proyeccion_total=proyeccion_total,ultima_fecha=ult_fecha,
                )
    except Exception as e:
        st.sidebar.error(f"Error procesando archivo Comparativo: {e}")
        DEBUG_COMP["error"]=str(e)

# ajustar motivos de caída para que la suma coincida con PROMESAS_CAIDAS
df_motivos["Casos"]=(df_motivos["Casos"]/df_motivos["Casos"].sum()*PROMESAS_CAIDAS).round().astype(int)
_diff=int(PROMESAS_CAIDAS-df_motivos["Casos"].sum())
df_motivos.loc[df_motivos["Casos"].idxmax(),"Casos"]+=_diff

# ── HEADER ──
c1h,c2h,c3h=st.columns([3,1,1])
with c1h:
    st.markdown("## 📊 Dashboard Ejecutivo — Cobranza Mayo 2025")
    st.caption("NAtura | Reunión de seguimiento | 09 Jun 2025")
with c2h:
    st.metric("Recuperación Mayo",f"${RECUPERADO_TOTAL/1e6:.2f}M",
              delta=f"{(RECUPERADO_TOTAL/META_TOTAL*100-100):.1f}% vs meta")
with c3h:
    st.metric("Contact Rate",f"{CR*100:.1f}%",delta="–3.2pp vs abr")

tab1,tab2,tab3,tab4,tab5,tab6=st.tabs([
    "1 · Cierre de Mes","2 · Contactabilidad","3 · Indicadores","4 · Operación","5 · Plan de Trabajo",
    "6 · Comparativo Cobranza"])

# ══ TAB 1 ─ CIERRE DE MES ══
with tab1:
    st.markdown('<div class="sec">Cierre de Mes — Mayo 2025</div>',unsafe_allow_html=True)
    c1,c2,c3,c4,c5=st.columns(5)
    c1.metric("Meta mensual",f"${META_TOTAL/1e6:.2f}M")
    c2.metric("Recuperado",f"${RECUPERADO_TOTAL/1e6:.2f}M",delta=f"{(RECUPERADO_TOTAL/META_TOTAL-1)*100:.1f}%")
    c3.metric("Cumplimiento",f"{RECUPERADO_TOTAL/META_TOTAL*100:.1f}%")
    c4.metric("Promesas generadas",f"{PROMESAS_GEN:,}")
    c5.metric("Promesas caídas",f"{PROMESAS_CAIDAS:,}",delta=f"{PROMESAS_CAIDAS/PROMESAS_GEN*100:.0f}% caída",delta_color="inverse")
    st.markdown("---")
    cl,cr=st.columns([3,2])
    with cl:
        st.markdown("**Recuperación acumulada vs Meta**")
        fig=go.Figure()
        fig.add_trace(go.Scatter(x=df_cierre["Fecha"],y=df_cierre["Meta Acum"],name="Meta",line=dict(color=SLATE,dash="dash",width=2)))
        fig.add_trace(go.Scatter(x=df_cierre["Fecha"],y=df_cierre["Recuperado Acum"],name="Recuperado",
                                 line=dict(color=BLUE,width=3),fill="tonexty",fillcolor="rgba(37,99,235,0.07)"))
        fig.add_hrect(y0=META_TOTAL*0.95,y1=META_TOTAL*1.05,fillcolor="rgba(22,163,74,0.05)",line_width=0,
                      annotation_text="±5% meta",annotation_position="top right",annotation_font_color=GREEN)
        apply_layout(fig,height=320,legend=dict(orientation="h",y=1.1))
        fig.update_yaxes(tickprefix="$",tickformat=",.0f"); st.plotly_chart(fig,use_container_width=True)
    with cr:
        st.markdown("**Motivos de caída de promesas**")
        df_mot=df_motivos.sort_values("Casos")
        fig2=go.Figure(go.Bar(x=df_mot["Casos"],y=df_mot["Motivo"],orientation="h",
                              marker_color=[RED if c>150 else AMBER if c>80 else SLATE for c in df_mot["Casos"]],
                              text=df_mot["Casos"],textposition="outside",textfont=dict(color=TEXT)))
        apply_layout(fig2,height=320); st.plotly_chart(fig2,use_container_width=True)
    gap=META_TOTAL-RECUPERADO_TOTAL
    st.info(f"**Brecha: ${gap/1e6:.2f}M** — las **{PROMESAS_CAIDAS} promesas caídas** explican ~{PROMESAS_CAIDAS/PROMESAS_GEN*100:.0f}% del desvío.")

# ══ TAB 2 ─ CONTACTABILIDAD ══
with tab2:
    st.markdown('<div class="sec">Contactabilidad y Canales Digitales</div>',unsafe_allow_html=True)
    c1,c2,c3,c4,c5=st.columns(5)
    c1.metric("Total intentos",f"{TOTAL_INT:,}")
    c2.metric("Contact Rate",f"{CR*100:.1f}%",delta="–3.2pp vs abr",delta_color="inverse")
    c3.metric("Titular contactado",f"{TITULAR:,}")
    c4.metric("Buzón + No contesta",f"{BUZON+NO_CONT:,}")
    c5.metric("Cuelga al agente",f"{COLGADO:,}",delta="Alto",delta_color="inverse")
    st.markdown("---")
    cl,cr=st.columns(2)
    with cl:
        st.markdown("**Contactos y promesas por hora**")
        fig=make_subplots(specs=[[{"secondary_y":True}]])
        fig.add_trace(go.Bar(x=df_horario["Hora"],y=df_horario["Contactos"],name="Contactos",marker_color=BLUE,opacity=0.75),secondary_y=False)
        fig.add_trace(go.Scatter(x=df_horario["Hora"],y=df_horario["Conv %"],name="Conv %",
                                 mode="lines+markers",line=dict(color=GREEN,width=2),marker=dict(size=6)),secondary_y=True)
        fig.update_layout(**PLOTLY_LAYOUT,height=300,legend=dict(orientation="h",y=1.12))
        fig.update_yaxes(gridcolor="#e2e8f0",secondary_y=False)
        fig.update_yaxes(ticksuffix="%",gridcolor="rgba(0,0,0,0)",secondary_y=True)
        st.plotly_chart(fig,use_container_width=True); st.caption("Franja de mayor conversión: **13:00–15:00 h**")
    with cr:
        st.markdown("**Distribución de intentos**")
        labels=["Titular contactado","Buzón","No contesta","Núm inválido","Colgó agente"]
        vals=[TITULAR,BUZON,NO_CONT,NUM_INV,COLGADO]
        colors_pie=[GREEN,AMBER,SLATE,RED,"#9333ea"]
        fig2=go.Figure(go.Pie(labels=labels,values=vals,
                              marker=dict(colors=colors_pie,line=dict(color="white",width=2)),
                              hole=0.52,textinfo="label+percent",textfont=dict(color=TEXT)))
        apply_layout(fig2,height=300,showlegend=False); st.plotly_chart(fig2,use_container_width=True)
    st.markdown("---"); st.markdown("**Canales digitales — efectividad para amarrar promesas**")
    cols=st.columns(5); canal_icons={"SMS":"📱","WhatsApp":"💬","Email":"📧","Llamada":"📞","IVR":"🤖"}
    for i,row in df_canal.iterrows():
        with cols[i]:
            icon=canal_icons.get(row["Canal"],"📡")
            color=GREEN if row["Conv prom %"]>25 else AMBER if row["Conv prom %"]>18 else RED
            bg="#f0fdf4" if color==GREEN else "#fffbeb" if color==AMBER else "#fef2f2"
            st.markdown(f"""<div style="background:{bg};border-radius:10px;padding:14px;text-align:center;
                border:1px solid {BORD};border-top:3px solid {color}">
              <div style="font-size:1.6rem">{icon}</div>
              <div style="font-weight:700;font-size:1rem;color:{TEXT}">{row['Canal']}</div>
              <div style="color:{MUTED};font-size:0.75rem;margin:4px 0">Enviados: {row['Enviados']:,}</div>
              <div style="color:{MUTED};font-size:0.75rem">Resp: {row['Tasa resp %']}%</div>
              <div style="font-size:1.4rem;font-weight:800;color:{color}">{row['Conv prom %']}%</div>
              <div style="color:{MUTED};font-size:0.7rem">conv → promesa</div>
              <div style="margin-top:6px;color:{color};font-weight:700">{row['Promesas']:,} promesas</div>
            </div>""",unsafe_allow_html=True)
    st.markdown("<br>",unsafe_allow_html=True)
    st.warning("**WhatsApp** lidera en conversión (30%). **SMS** al 15% — A/B test recomendado por segmento.")

# ══ TAB 3 ─ INDICADORES ══
with tab3:
    st.markdown('<div class="sec">Indicadores de Recuperación</div>',unsafe_allow_html=True)
    sub1,sub2,sub3,sub4,sub5,sub6,sub7=st.tabs([
        "Por Sector","Por GV","Por Edad","Estado & Segmento",
        "📅 Temporalidad","🕒 Horas por Semana","👤 Edad Consultora"])

    with sub1:
        cl,cr=st.columns([3,2])
        with cl:
            fig=go.Figure()
            fig.add_trace(go.Bar(name="Meta",x=df_sector["Sector"],y=df_sector["Meta"],marker_color="#cbd5e1",opacity=0.8))
            fig.add_trace(go.Bar(name="Recuperado",x=df_sector["Sector"],y=df_sector["Recuperado"],
                                 marker_color=[GREEN if p>=80 else AMBER if p>=65 else RED for p in df_sector["Cumplimiento %"]]))
            apply_layout(fig,height=340,barmode="group",legend=dict(orientation="h",y=1.1),
                         yaxis=dict(tickprefix="$",tickformat=",.0f"))
            st.plotly_chart(fig,use_container_width=True)
        with cr:
            for _,row in df_sector.sort_values("Cumplimiento %",ascending=False).iterrows():
                badge="badge-green" if row["Cumplimiento %"]>=80 else "badge-amber" if row["Cumplimiento %"]>=65 else "badge-red"
                st.markdown(f"<div style='background:{CARD};border:1px solid {BORD};border-radius:8px;"
                            f"padding:10px 14px;margin-bottom:6px;display:flex;justify-content:space-between;align-items:center'>"
                            f"<span style='color:{TEXT};font-weight:600'>{row['Sector']}</span>"
                            f"<span class='badge {badge}'>{row['Cumplimiento %']}%</span></div>",unsafe_allow_html=True)

    with sub2:
        cl,cr=st.columns([5,2])
        with cl:
            colors_gv=[GREEN if p>=85 else AMBER if p>=70 else RED for p in df_gv["Cumplimiento %"]]
            fig=go.Figure(go.Bar(x=df_gv["GV"],y=df_gv["Cumplimiento %"],marker_color=colors_gv,
                                 text=df_gv["Cumplimiento %"].astype(str)+"%",textposition="outside",textfont=dict(color=TEXT,size=10)))
            fig.add_hline(y=80,line_dash="dash",line_color=AMBER,annotation_text="Meta 80%",annotation_font_color=AMBER)
            apply_layout(fig,height=350,yaxis=dict(ticksuffix="%",range=[0,115]))
            st.plotly_chart(fig,use_container_width=True)
        with cr:
            st.markdown(f"**Top 5** <span class='badge badge-green'>Mejor desempeño</span>",unsafe_allow_html=True)
            for _,row in df_gv.head(5).iterrows():
                st.markdown(f"<div style='background:{CARD};border:1px solid {BORD};border-radius:6px;"
                            f"padding:8px 12px;margin:4px 0;border-left:3px solid {GREEN}'>"
                            f"<b style='color:{TEXT}'>{row['GV']}</b>"
                            f"<span style='color:{GREEN};float:right'>{row['Cumplimiento %']}%</span></div>",unsafe_allow_html=True)
            st.markdown(f"<br>**Bottom 5** <span class='badge badge-red'>Requieren atención</span>",unsafe_allow_html=True)
            for _,row in df_gv.tail(5).iterrows():
                st.markdown(f"<div style='background:{CARD};border:1px solid {BORD};border-radius:6px;"
                            f"padding:8px 12px;margin:4px 0;border-left:3px solid {RED}'>"
                            f"<b style='color:{TEXT}'>{row['GV']}</b>"
                            f"<span style='color:{RED};float:right'>{row['Cumplimiento %']}%</span></div>",unsafe_allow_html=True)

    with sub3:
        cl,cr=st.columns(2)
        with cl:
            fig=px.bar(df_edad,x="Rango",y="Recuperado",color="Promesas %",
                       color_continuous_scale=["#fca5a5","#fcd34d","#86efac"],text_auto=".2s",
                       labels={"Recuperado":"Recuperado $","Promesas %":"% Con promesa"})
            fig.update_traces(textfont_color=TEXT)
            apply_layout(fig,height=320,yaxis=dict(tickprefix="$",tickformat=",.0f"),
                         coloraxis_colorbar=dict(title="% Promesa",tickfont=dict(color=TEXT)))
            st.plotly_chart(fig,use_container_width=True)
        with cr:
            fig2=make_subplots(specs=[[{"secondary_y":True}]])
            fig2.add_trace(go.Bar(x=df_edad["Rango"],y=df_edad["Ticket Prom"],name="Ticket Prom $",marker_color=PURPLE,opacity=0.8),secondary_y=False)
            fig2.add_trace(go.Scatter(x=df_edad["Rango"],y=df_edad["Promesas %"],name="Promesas %",
                                      mode="lines+markers",line=dict(color=GREEN,width=2),marker=dict(size=7)),secondary_y=True)
            fig2.update_layout(**PLOTLY_LAYOUT,height=320,legend=dict(orientation="h",y=1.12))
            fig2.update_yaxes(tickprefix="$",secondary_y=False,gridcolor="#e2e8f0")
            fig2.update_yaxes(ticksuffix="%",secondary_y=True,gridcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig2,use_container_width=True)
        st.info("**Segmento 36-45 años:** mayor recuperación absoluta ($336K) y tasa de promesas más alta (71%).")
        st.dataframe(df_edad.rename(columns={"Recuperado":"Recuperado $","Ticket Prom":"Ticket $"})
                            .style.format({"Recuperado $":"${:,.0f}","Ticket $":"${:,.0f}","Promesas %":"{:.0f}%"}),
                     use_container_width=True,hide_index=True)

    with sub4:
        cl,cr=st.columns(2)
        with cl:
            st.markdown("**Recuperación por Estado (Top 10)**")
            fig=px.bar(df_estado.sort_values("Recuperado"),x="Recuperado",y="Estado",orientation="h",
                       color="Recuperado",color_continuous_scale=["#bfdbfe","#3b82f6","#1d4ed8"],text_auto=".2s")
            fig.update_traces(textfont_color=TEXT)
            apply_layout(fig,height=360,xaxis=dict(tickprefix="$",tickformat=",.0f"),coloraxis_showscale=False)
            st.plotly_chart(fig,use_container_width=True)
        with cr:
            st.markdown("**Recuperación por Segmento**")
            fig2=make_subplots(specs=[[{"secondary_y":True}]])
            fig2.add_trace(go.Bar(x=df_segmento["Segmento"],y=df_segmento["Recuperado"],name="Recuperado $",marker_color=BLUE,opacity=0.8),secondary_y=False)
            fig2.add_trace(go.Scatter(x=df_segmento["Segmento"],y=df_segmento["Cumplimiento %"],name="% Cumplimiento",
                                      mode="lines+markers",line=dict(color=AMBER,width=2),marker=dict(size=8)),secondary_y=True)
            fig2.update_layout(**PLOTLY_LAYOUT,height=360,legend=dict(orientation="h",y=1.12))
            fig2.update_yaxes(tickprefix="$",tickformat=",.0f",secondary_y=False,gridcolor="#e2e8f0")
            fig2.update_yaxes(ticksuffix="%",secondary_y=True,gridcolor="rgba(0,0,0,0)",range=[0,110])
            st.plotly_chart(fig2,use_container_width=True)
        st.dataframe(df_segmento.style.format({"Recuperado":"${:,.0f}","Cumplimiento %":"{:.0f}%"}),
                     use_container_width=True,hide_index=True)

    with sub5:
        st.markdown("**Recuperación por Semana del Mes**")
        cl,cr=st.columns(2)
        with cl:
            fig=go.Figure()
            fig.add_trace(go.Bar(name="Meta",x=df_temp["Semana"],y=df_temp["Meta"],marker_color="#cbd5e1",opacity=0.8))
            fig.add_trace(go.Bar(name="Recuperado",x=df_temp["Semana"],y=df_temp["Recuperado"],
                                 marker_color=[GREEN if p>=90 else AMBER if p>=75 else RED for p in df_temp["Cumplimiento %"]]))
            apply_layout(fig,height=320,barmode="group",legend=dict(orientation="h",y=1.1),
                         yaxis=dict(tickprefix="$",tickformat=",.0f"))
            st.plotly_chart(fig,use_container_width=True)
        with cr:
            for _,row in df_temp.iterrows():
                badge="badge-green" if row["Cumplimiento %"]>=90 else "badge-amber" if row["Cumplimiento %"]>=75 else "badge-red"
                st.markdown(f"<div style='background:{CARD};border:1px solid {BORD};border-radius:8px;"
                            f"padding:12px 16px;margin-bottom:8px;display:flex;justify-content:space-between;align-items:center'>"
                            f"<span style='color:{TEXT};font-weight:600'>{row['Semana']}</span>"
                            f"<div style='text-align:right'>"
                            f"<div style='color:{MUTED};font-size:0.8rem'>${row['Recuperado']/1e6:.2f}M de ${row['Meta']/1e6:.2f}M</div>"
                            f"<span class='badge {badge}'>{row['Cumplimiento %']}%</span></div></div>",unsafe_allow_html=True)
        st.markdown("---")
        st.markdown("**Por Quincena**")
        cl2,cr2=st.columns(2)
        with cl2:
            fig3=go.Figure()
            fig3.add_trace(go.Bar(name="Meta",x=df_quincena["Período"],y=df_quincena["Meta"],marker_color="#cbd5e1",opacity=0.8))
            fig3.add_trace(go.Bar(name="Recuperado",x=df_quincena["Período"],y=df_quincena["Recuperado"],
                                  marker_color=[GREEN if p>=90 else AMBER if p>=75 else RED for p in df_quincena["Cumplimiento %"]]))
            apply_layout(fig3,height=280,barmode="group",legend=dict(orientation="h",y=1.1),
                         yaxis=dict(tickprefix="$",tickformat=",.0f"))
            st.plotly_chart(fig3,use_container_width=True)
        with cr2:
            st.markdown("**Recuperación diaria — Mayo**")
            fig4=go.Figure()
            fig4.add_trace(go.Bar(x=df_cierre["Fecha"],y=df_cierre["Recuperado"],name="Diario",
                                  marker_color=BLUE,opacity=0.7))
            fig4.add_hline(y=meta_d,line_dash="dash",line_color=RED,
                           annotation_text=f"Meta diaria: ${meta_d:,.0f}",annotation_font_color=RED)
            apply_layout(fig4,height=280,yaxis=dict(tickprefix="$",tickformat=",.0f"))
            st.plotly_chart(fig4,use_container_width=True)
        sem_mejor=df_temp.loc[df_temp["Recuperado"].idxmax(),"Semana"]
        sem_peor=df_temp.loc[df_temp["Recuperado"].idxmin(),"Semana"]
        st.info(f"**Mejor semana:** {sem_mejor} — **Semana con menor recuperación:** {sem_peor}. "
                f"La última semana suele caer por cierre de mes; enfocar marcaciones los días 28-31.")

    with sub6:
        st.markdown("**Heatmap de Contactos Titular — Hora × Día de Semana**")
        fig_heat=go.Figure(go.Heatmap(
            z=heatmap_data,
            x=[f"{h}:00" for h in HORAS],
            y=DIAS_SEM,
            colorscale=[[0,"#f0f9ff"],[0.4,"#7dd3fc"],[0.7,"#2563eb"],[1.0,"#1e3a8a"]],
            text=heatmap_data,texttemplate="%{text}",
            hovertemplate="<b>%{y}</b><br>%{x}<br>Contactos: %{z}<extra></extra>",
            showscale=True,colorbar=dict(title="Contactos",tickfont=dict(color=TEXT))
        ))
        apply_layout(fig_heat,height=320)
        fig_heat.update_xaxes(side="bottom")
        st.plotly_chart(fig_heat,use_container_width=True)
        st.caption("Las celdas más oscuras indican mayor volumen de contactos titulares. Concentra marcaciones en esas franjas.")
        st.markdown("---")
        st.markdown("**Contactos por Hora según Semana del Mes**")
        fig_hs=go.Figure()
        sem_cols={"Sem 1":SLATE,"Sem 2":AMBER,"Sem 3":BLUE,"Sem 4":GREEN}
        for sem,color in sem_cols.items():
            if sem in df_hora_sem.columns:
                fig_hs.add_trace(go.Scatter(
                    x=df_hora_sem["Hora"],y=df_hora_sem[sem],
                    name=sem,mode="lines+markers",
                    line=dict(color=color,width=2),marker=dict(size=6)
                ))
        apply_layout(fig_hs,height=320,legend=dict(orientation="h",y=1.12))
        fig_hs.update_xaxes(title_text="Hora del día",tickvals=HORAS,ticktext=[f"{h}:00" for h in HORAS])
        fig_hs.update_yaxes(title_text="Contactos titulares")
        st.plotly_chart(fig_hs,use_container_width=True)
        st.info("**Semana 3** tiene mayor contactabilidad en casi todas las franjas. La franja 12:00–15:00 es consistentemente alta en todas las semanas.")

    with sub7:
        st.markdown("**Dispersión: Edad vs Monto Pagado por Consultora**")
        seg_colors={"Diamante":"#7c3aed","Oro":"#d97706","Plata":"#64748b",
                    "Bronce":"#92400e","Nuevo Ingreso":"#2563eb"}
        fig_sc=px.scatter(
            df_scatter_edad,x="Edad",y="Monto Pagado",color="Segmento",
            color_discrete_map=seg_colors,
            opacity=0.65,size_max=10,
            labels={"Edad":"Edad de la consultora","Monto Pagado":"Monto pagado ($)","Segmento":"Segmento"},
            hover_data={"Edad":True,"Monto Pagado":":,.0f","Segmento":True}
        )
        fig_sc.update_traces(marker=dict(size=7))
        apply_layout(fig_sc,height=380,yaxis=dict(tickprefix="$",tickformat=",.0f"),
                     legend=dict(orientation="h",y=1.08))
        st.plotly_chart(fig_sc,use_container_width=True)
        st.markdown("---")
        st.markdown("**Resumen por rango de edad**")
        cl,cr=st.columns(2)
        with cl:
            fig_e=go.Figure()
            fig_e.add_trace(go.Bar(x=df_edad["Rango"],y=df_edad["Recuperado"],name="Recuperado $",
                                   marker_color=[GREEN if p>=65 else AMBER if p>=50 else RED for p in df_edad["Promesas %"]]))
            apply_layout(fig_e,height=300,yaxis=dict(tickprefix="$",tickformat=",.0f"))
            st.plotly_chart(fig_e,use_container_width=True)
        with cr:
            fig_t=go.Figure(go.Bar(
                x=df_edad["Rango"],y=df_edad["Ticket Prom"],name="Ticket Promedio $",
                marker_color=PURPLE,
                text=df_edad["Ticket Prom"].apply(lambda x: f"${x:,.0f}"),
                textposition="outside",textfont=dict(color=TEXT)
            ))
            apply_layout(fig_t,height=300,yaxis=dict(tickprefix="$",tickformat=",.0f"))
            st.plotly_chart(fig_t,use_container_width=True)
        st.dataframe(
            df_edad.rename(columns={"Recuperado":"Recuperado $","Ticket Prom":"Ticket $"})
                   .style.format({"Recuperado $":"${:,.0f}","Ticket $":"${:,.0f}","Promesas %":"{:.0f}%"})
                   .background_gradient(subset=["Promesas %"],cmap="RdYlGn"),
            use_container_width=True,hide_index=True
        )
        st.info("**36-45 años** tiene el mayor ticket promedio y tasa de promesas. "
                "Las consultoras **18-25** tienen el ticket más bajo: considerar estrategia diferenciada (planes de pago pequeños).")

# ══ TAB 4 ─ OPERACIÓN ══
with tab4:
    st.markdown('<div class="sec">Operación de Asesores — Mayo 2025</div>',unsafe_allow_html=True)
    avg_tmo=df_asesores["TMO (min)"].mean()
    total_llam=df_asesores["Llamadas"].sum()
    avg_abandono=df_asesores["% Abandono"].mean()
    avg_conv=df_asesores["Conv %"].mean()
    c1,c2,c3,c4=st.columns(4)
    c1.metric("Total llamadas",f"{total_llam:,}")
    c2.metric("TMO promedio",f"{avg_tmo:.1f} min")
    c3.metric("% Abandono <30s",f"{avg_abandono:.1f}%",delta="Revisar apertura",delta_color="inverse")
    c4.metric("Conv. Contacto→Promesa",f"{avg_conv:.1f}%")
    st.markdown("---")
    cl,cr=st.columns([3,2])
    with cl:
        st.markdown("**Ranking de asesores — monto recuperado y conversión**")
        fig=make_subplots(specs=[[{"secondary_y":True}]])
        fig.add_trace(go.Bar(x=df_asesores["Asesor"],y=df_asesores["Monto Rec"],name="Monto Rec $",marker_color=BLUE,opacity=0.85),secondary_y=False)
        fig.add_trace(go.Scatter(x=df_asesores["Asesor"],y=df_asesores["Conv %"],name="Conv %",
                                 mode="lines+markers",line=dict(color=GREEN,width=2),marker=dict(size=7)),secondary_y=True)
        fig.update_layout(**PLOTLY_LAYOUT,height=320,legend=dict(orientation="h",y=1.12))
        fig.update_yaxes(tickprefix="$",tickformat=",.0f",secondary_y=False,gridcolor="#e2e8f0")
        fig.update_yaxes(ticksuffix="%",secondary_y=True,gridcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig,use_container_width=True)
    with cr:
        st.markdown("**TMO y % Abandono por asesor**")
        fig2=go.Figure()
        fig2.add_trace(go.Bar(name="TMO (min)",x=df_asesores["Asesor"],y=df_asesores["TMO (min)"],
                              marker_color=[RED if t>6 else AMBER if t>4.5 else GREEN for t in df_asesores["TMO (min)"]]))
        fig2.add_trace(go.Scatter(name="% Abandono",x=df_asesores["Asesor"],y=df_asesores["% Abandono"],
                                  mode="lines+markers",line=dict(color=RED,width=2,dash="dot"),marker=dict(size=7),yaxis="y2"))
        fig2.update_layout(**PLOTLY_LAYOUT,height=320,legend=dict(orientation="h",y=1.12),
                           yaxis2=dict(overlaying="y",side="right",ticksuffix="%",gridcolor="rgba(0,0,0,0)",showgrid=False))
        st.plotly_chart(fig2,use_container_width=True)
    st.markdown("---")
    cl2,cr2=st.columns(2)
    with cl2:
        st.markdown("**Frecuencia de objeciones difíciles**")
        fig3=go.Figure(go.Bar(x=df_objeciones["Frecuencia"],y=df_objeciones["Objeción"],orientation="h",
                              marker_color=[RED if f>2000 else AMBER if f>1000 else SLATE for f in df_objeciones["Frecuencia"]],
                              text=df_objeciones["Frecuencia"],textposition="outside",textfont=dict(color=TEXT)))
        apply_layout(fig3,height=280); st.plotly_chart(fig3,use_container_width=True)
    with cr2:
        st.markdown("**Tasa de resolución exitosa por objeción**")
        fig4=go.Figure(go.Bar(x=df_objeciones["Resolución exitosa %"],y=df_objeciones["Objeción"],orientation="h",
                              marker_color=[GREEN if p>=65 else AMBER if p>=50 else RED for p in df_objeciones["Resolución exitosa %"]],
                              text=df_objeciones["Resolución exitosa %"].astype(str)+"%",
                              textposition="outside",textfont=dict(color=TEXT)))
        apply_layout(fig4,height=280,xaxis=dict(ticksuffix="%",range=[0,100]))
        st.plotly_chart(fig4,use_container_width=True)
    st.warning(f"**Puntos críticos:** 'No tengo dinero' (2,840 casos, 38% resolución). Abandono promedio **{avg_abandono:.1f}%**.")
    st.markdown("---"); st.markdown("**Detalle completo por asesor**")
    display_df=df_asesores[["Asesor","Llamadas","TMO (min)","% Contacto","Conv %","% Abandono","Monto Rec"]].copy()
    display_df["Monto Rec"]=display_df["Monto Rec"].apply(lambda x: f"${x:,.0f}")
    st.dataframe(display_df,use_container_width=True,hide_index=True)

# ══ TAB 5 ─ PLAN DE TRABAJO ══
with tab5:
    st.markdown('<div class="sec">Plan de Trabajo — Junio 2025</div>',unsafe_allow_html=True)
    gap_jun=META_JUN-PROY_JUN
    c1,c2,c3,c4=st.columns(4)
    c1.metric("Meta Junio",f"${META_JUN/1e6:.1f}M")
    c2.metric("Proyección base",f"${PROY_JUN/1e6:.2f}M",delta=f"{(PROY_JUN/META_JUN-1)*100:.1f}%",delta_color="inverse")
    c3.metric("Gap a cerrar",f"${gap_jun/1e6:.2f}M",delta_color="inverse")
    c4.metric("Pipeline de promesas",f"${PIPELINE_PROM/1e6:.2f}M")
    st.markdown("---")
    cl,cr=st.columns(2)
    with cl:
        st.markdown("**Proyección acumulada Junio — 3 escenarios**")
        fig=go.Figure()
        fig.add_trace(go.Scatter(x=df_proy["Fecha"],y=df_proy["Meta"],name="Meta",line=dict(color=SLATE,dash="dash",width=2)))
        fig.add_trace(go.Scatter(x=df_proy["Fecha"],y=df_proy["Optimista"],name="Optimista",
                                 line=dict(color=GREEN,width=2),fill="tonexty",fillcolor="rgba(22,163,74,0.06)"))
        fig.add_trace(go.Scatter(x=df_proy["Fecha"],y=df_proy["Base"],name="Base",line=dict(color=BLUE,width=3)))
        fig.add_trace(go.Scatter(x=df_proy["Fecha"],y=df_proy["Pesimista"],name="Pesimista",
                                 line=dict(color=RED,width=2,dash="dot"),fill="tonexty",fillcolor="rgba(220,38,38,0.05)"))
        apply_layout(fig,height=360,legend=dict(orientation="h",y=1.1))
        fig.update_yaxes(tickprefix="$",tickformat=",.0f"); st.plotly_chart(fig,use_container_width=True)
    with cr:
        st.markdown("**Impacto estimado por palanca**")
        df_p=df_palancas.sort_values("Impacto estimado $")
        fig2=go.Figure(go.Bar(x=df_p["Impacto estimado $"],y=df_p["Palanca"],orientation="h",marker_color=BLUE,
                              text=df_p["Impacto estimado $"].apply(lambda x: f"${x:,.0f}"),
                              textposition="outside",textfont=dict(color=TEXT)))
        total_impacto=df_palancas["Impacto estimado $"].sum()
        fig2.add_vline(x=gap_jun,line_dash="dash",line_color=RED,
                       annotation_text=f"Gap: ${gap_jun:,.0f}",annotation_font_color=RED)
        apply_layout(fig2,height=360,xaxis=dict(tickprefix="$",tickformat=",.0f"))
        st.plotly_chart(fig2,use_container_width=True)
        st.caption(f"Impacto total: **${total_impacto:,.0f}** — cubre **{total_impacto/gap_jun*100:.0f}%** del gap")
    st.markdown("---"); st.markdown("**Compromisos del equipo**")
    for _,row in df_acciones.iterrows():
        badge_cls="badge-red" if row["Plazo"]=="Inmediato" else "badge-amber"
        bg_row="#fef2f2" if row["Plazo"]=="Inmediato" else "#fffbeb"
        st.markdown(f"<div style='background:{bg_row};border:1px solid {BORD};border-radius:8px;"
                    f"padding:10px 16px;margin:5px 0;display:flex;align-items:center;gap:12px'>"
                    f"<span class='badge {badge_cls}'>{row['Plazo']}</span>"
                    f"<span style='color:{MUTED};min-width:130px;font-size:0.85rem'>{row['Área']}</span>"
                    f"<span style='color:{TEXT}'>{row['Acción']}</span></div>",unsafe_allow_html=True)
    st.markdown("<br>",unsafe_allow_html=True)
    st.success(f"**Resumen ejecutivo:** Palancas propuestas: **${total_impacto/1e6:.2f}M** adicional, "
               f"cubriendo **{total_impacto/gap_jun*100:.0f}%** del gap. "
               f"Prioridad #1: rellamada a las **{PROMESAS_CAIDAS} promesas caídas** de Mayo.")

# ══ TAB 6 ─ COMPARATIVO COBRANZA ══
with tab6:
    st.markdown('<div class="sec">Comparativo Cobranza por Tramo</div>',unsafe_allow_html=True)
    if df_comparativo is None:
        st.markdown(
            f"<div style='text-align:center;padding:50px 20px;background:{CARD};border-radius:16px;"
            f"border:2px dashed {BORD}'>"
            f"<div style='font-size:2.5rem'>📑</div>"
            f"<div style='font-size:1.2rem;font-weight:700;color:{TEXT};margin:10px 0'>Sube el archivo Comparativo</div>"
            f"<div style='color:{MUTED};font-size:0.9rem'>Carga el Excel con las hojas <b>'TABLA DE PAGOS'</b> y <b>'COMPARATIVO'</b> "
            f"en la barra lateral (📑 Comparativo Cobranza) y presiona <b>Ejecutar</b>.<br>"
            f"El sistema sincronizará automáticamente la Cobranza $ por tramo desde los totales de Pagos.</div>"
            f"</div>", unsafe_allow_html=True)
        if DEBUG_COMP is not None:
            st.markdown("---")
            with st.expander("🔍 Diagnóstico del archivo subido (clic para ver detalle)"):
                st.write("**Hojas encontradas:**", DEBUG_COMP.get("hojas"))
                st.write("**Hoja de Pagos detectada:**", DEBUG_COMP.get("hoja_pagos"))
                st.write("**Hoja de Comparativo detectada:**", DEBUG_COMP.get("hoja_comparativo"))
                st.write("**Fila de encabezados T1..Tn (Pagos):**", DEBUG_COMP.get("header_row_pagos"))
                st.write("**Columnas de tramo detectadas (Pagos):**", DEBUG_COMP.get("tramo_cols"))
                st.write("**Fila TOTAL (Pagos):**", DEBUG_COMP.get("total_row_pagos"))
                st.write("**Totales por tramo (fila TOTAL):**", DEBUG_COMP.get("totales_tramo"))
                st.write("**Fila de encabezados (Comparativo):**", DEBUG_COMP.get("header_row_comparativo"))
                st.write("**Columnas detectadas (Comparativo):**", DEBUG_COMP.get("cols_comparativo"))
                st.write("**Valores de la columna Tramo (Comparativo):**", DEBUG_COMP.get("tramo_vals_comparativo"))
                if "error" in DEBUG_COMP:
                    st.error(DEBUG_COMP["error"])
    else:
        r=RESUMEN_COMP
        c1,c2,c3,c4=st.columns(4)
        c1.metric("Cartera asignada",f"${r['total_cartera']/1e6:.2f}M")
        c2.metric("Objetivo del mes",f"${r['total_objetivo']/1e6:.2f}M")
        c3.metric("Cobranza actual",f"${r['total_cobranza']/1e6:.2f}M",
                  delta=f"{r['cumplimiento']-100:.1f}pp vs objetivo")
        c4.metric("Días pendientes",f"{r['dias_pendientes']:.0f} de {r['dias_totales_mes']}")
        st.markdown("---")

        cl,cr=st.columns(2)
        with cl:
            st.markdown("**Cartera, Objetivo y Cobranza por Tramo**")
            fig=go.Figure()
            fig.add_trace(go.Bar(name="Cartera $",x=df_comparativo["Tramo"],y=df_comparativo["Cartera $"],marker_color="#cbd5e1",opacity=0.85))
            fig.add_trace(go.Bar(name="Objetivo $",x=df_comparativo["Tramo"],y=df_comparativo["Objetivo $"],marker_color=SLATE,opacity=0.85))
            fig.add_trace(go.Bar(name="Cobranza $",x=df_comparativo["Tramo"],y=df_comparativo["Cobranza $"],marker_color=BLUE))
            apply_layout(fig,height=340,barmode="group",legend=dict(orientation="h",y=1.1),
                         yaxis=dict(tickprefix="$",tickformat=",.0f"))
            st.plotly_chart(fig,use_container_width=True)
        with cr:
            st.markdown("**% de Cumplimiento de Objetivo por Tramo**")
            df_cs=df_comparativo.sort_values("Cumplimiento %")
            fig2=go.Figure(go.Bar(x=df_cs["Cumplimiento %"],y=df_cs["Tramo"],orientation="h",
                marker_color=[GREEN if p>=100 else AMBER if p>=70 else RED for p in df_cs["Cumplimiento %"]],
                text=df_cs["Cumplimiento %"].apply(lambda x: f"{x:.1f}%" if pd.notna(x) else "—"),
                textposition="outside",textfont=dict(color=TEXT)))
            fig2.add_vline(x=100,line_dash="dash",line_color=GREEN,annotation_text="Meta 100%",annotation_font_color=GREEN)
            apply_layout(fig2,height=340,xaxis=dict(ticksuffix="%"))
            st.plotly_chart(fig2,use_container_width=True)

        if df_pagos_diario is not None and len(df_pagos_diario)>0:
            st.markdown("---")
            st.markdown("**Cobranza diaria y acumulado**")
            fig3=make_subplots(specs=[[{"secondary_y":True}]])
            fig3.add_trace(go.Bar(x=df_pagos_diario["Fecha"],y=df_pagos_diario["Total Día"],name="Cobranza diaria",
                                  marker_color=BLUE,opacity=0.75),secondary_y=False)
            fig3.add_trace(go.Scatter(x=df_pagos_diario["Fecha"],y=df_pagos_diario["Acumulado"],name="Acumulado",
                                      mode="lines+markers",line=dict(color=GREEN,width=3),marker=dict(size=5)),secondary_y=True)
            fig3.add_hline(y=r["promedio_diario"],line_dash="dash",line_color=AMBER,
                           annotation_text=f"Promedio diario: ${r['promedio_diario']:,.0f}",annotation_font_color=AMBER)
            fig3.update_layout(**PLOTLY_LAYOUT,height=340,legend=dict(orientation="h",y=1.1))
            fig3.update_yaxes(tickprefix="$",tickformat=",.0f",secondary_y=False,gridcolor="#e2e8f0")
            fig3.update_yaxes(tickprefix="$",tickformat=",.0f",secondary_y=True,gridcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig3,use_container_width=True)

        st.markdown("---")
        st.markdown("**Detalle por Tramo (sincronizado desde Pagos)**")
        fmt={"Cartera $":"${:,.0f}","Objetivo $":"${:,.0f}","Cobranza $":"${:,.0f}",
             "Cumplimiento %":"{:.1f}%"}
        if "Objetivo %" in df_comparativo.columns:
            fmt["Objetivo %"]="{:.1f}%"
        if "Cuentas" in df_comparativo.columns:
            fmt["Cuentas"]="{:,.0f}"
        st.dataframe(df_comparativo.style.format(fmt),use_container_width=True,hide_index=True)

        st.markdown("---")
        if r["proyeccion_total"]>=r["total_objetivo"]:
            st.success(f"**Proyección de cierre de mes:** ${r['proyeccion_total']/1e6:.2f}M — al ritmo actual "
                       f"(${r['promedio_diario']:,.0f}/día durante {r['dias_pendientes']:.0f} días restantes) "
                       f"se **superaría el objetivo** de ${r['total_objetivo']/1e6:.2f}M "
                       f"(cumplimiento actual {r['cumplimiento']:.1f}%).")
        else:
            falta=r["total_objetivo"]-r["proyeccion_total"]
            st.warning(f"**Proyección de cierre de mes:** ${r['proyeccion_total']/1e6:.2f}M — al ritmo actual "
                       f"(${r['promedio_diario']:,.0f}/día durante {r['dias_pendientes']:.0f} días restantes) "
                       f"**faltarían ${falta/1e6:.2f}M** para alcanzar el objetivo de ${r['total_objetivo']/1e6:.2f}M "
                       f"(cumplimiento actual {r['cumplimiento']:.1f}%).")
