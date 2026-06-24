import streamlit as strl
import firebase_admin
from firebase_admin import credentials, firestore
import pandas as pd
from io import BytesIO

# Oldal konfigurációja (szélesvásznú asztali nézet)
strl.set_page_config(
    page_title="FRD Raktár - Vezetői Dashboard",
    page_icon="📊",
    layout="wide"
)

# 1. Firebase csatlakozás inicializálása (Biztonságos felhős verzió - szigorú JSON fix)
@strl.cache_resource
def init_firebase():
    import json
    try:
        # A 'strict=False' engedélyezi a rejtett vezérlőkaraktereket (pl. \n) a szövegen belül
        key_dict = json.loads(strl.secrets["firebase_key"], strict=False)
        cred = credentials.Certificate(key_dict)
        if not firebase_admin._apps:
            firebase_admin.initialize_app(cred)
        return firestore.client()
    except Exception as e:
        strl.error(f"X BIZTONSÁGI HIBA: A Firebase kulcs nem olvasható a Secrets-ből! {e}")
        return None

db = init_firebase()

# 2. Adatok letöltése a Firestore-ból
def get_raktar_adatok():
    if db is None:
        return []
    
    docs = db.collection('materials').stream()
    adatok = []
    for doc in docs:
        d = doc.to_dict()
        # Biztonsági konverziók a tiszta megjelenítéshez
        try:
            current = float(d.get('currentStock', 0))
            minimum = float(d.get('minStock', 20))
        except:
            current, minimum = 0.0, 20.0

        adatok.append({
            "Cikkszám (SKU)": d.get('sku', doc.id),
            "Megnevezés": d.get('name', 'Névtelen alapanyag'),
            "Kategória": d.get('type', 'Egyéb'),
            "Készlet": int(current) if current % 1 == 0 else current,
            "Minimum szint": int(minimum) if minimum % 1 == 0 else minimum,
            "Egység": d.get('unit', 'Pár'),
            "Státusz": "🚨 HIÁNY" if current <= minimum else "✅ Rendben"
        })
    return adatok

# --- UI FELÉPÍTÉSE ---

strl.title("📊 FRD Alapanyag Raktár - Vezetői Műszerfal")
strl.caption("Élő, irodai betekintő felület az üzemben lévő tabletek készletéhez")
strl.write("---")

if db is not None:
    # Adatok frissítése gomb és adatok beolvasása
    nyers_adatok = get_raktar_adatok()
    df = pd.DataFrame(nyers_adatok)

    if not df.empty:
        # Készlethiányos termékek kiszűrése
        hianyzo_df = df[df["Státusz"] == "🚨 HIÁNY"]
        
        # --- VEZETŐI MUTATÓK (METRICS) ---
        col1, col2, col3 = strl.columns(3)
        with col1:
            strl.metric(label="Összes egyedi alapanyag", value=len(df))
        with col2:
            strl.metric(
                label="Készlethiányos tételek száma", 
                value=len(hianyzo_df),
                delta=f"{len(hianyzo_df)} azonnali beszerzés szükséges" if len(hianyzo_df) > 0 else "Minden rendben",
                delta_color="inverse" if len(hianyzo_df) > 0 else "normal"
            )
        with col3:
            # Excel export előkészítése a háttérben
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Aktuális Készlet')
            
            strl.write("### Riport letöltése")
            strl.download_button(
                label="📥 Teljes készlet letöltése Excelben",
                data=buffer.getvalue(),
                file_name="frd_raktarkeszlet_riport.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        strl.write("---")

        # --- FIGYELMEZTETŐ PANEL ---
        if not hianyzo_df.empty:
            strl.error("### 🚨 Az alábbi alapanyagok készlete a kritikus minimum alá süllyedt!")
            strl.dataframe(hianyzo_df, use_container_width=True, hide_index=True)
            strl.write("---")

        # --- SZŰRŐK AZ ASZTALI TÁBLÁZATHOZ ---
        strl.subheader("🔍 Keresés és szűrés a teljes raktárban")
        f_col1, f_col2 = strl.columns([1, 2])
        
        with f_col1:
            kategoriak = ["Mind"] + sorted(list(df["Kategória"].unique()))
            valasztott_kat = strl.selectbox("Szűrés kategória szerint:", kategoriak)
            
        with f_col2:
            kereses = strl.text_input("Keresés név vagy cikkszám alapján:", "").strip().lower()

        # Szűrések alkalmazása
        megjelenitendo_df = df.copy()
        if valasztott_kat != "Mind":
            megjelenitendo_df = megjelenitendo_df[megjelenitendo_df["Kategória"] == valasztott_kat]
        if kereses:
            megjelenitendo_df = megjelenitendo_df[
                megjelenitendo_df["Megnevezés"].str.lower().str.contains(kereses) | 
                megjelenitendo_df["Cikkszám (SKU)"].str.lower().str.contains(kereses)
            ]

        # --- A NAGY TÁBLÁZAT ---
        strl.dataframe(megjelenitendo_df, use_container_width=True, hide_index=True)
        
    else:
        strl.info("Az adatbázis csatlakozott, de jelenleg nem találhatók benne anyagok.")
