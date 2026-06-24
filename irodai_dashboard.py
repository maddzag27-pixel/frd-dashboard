import streamlit as strl
import firebase_admin
from firebase_admin import credentials, firestore
import pandas as pd
from io import BytesIO
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# Oldal konfigurációja (szélesvásznú asztali nézet)
strl.set_page_config(
    page_title="FRD Raktár - Vezetői Dashboard",
    page_icon="📊",
    layout="wide"
)

# 1. Firebase csatlakozás inicializálása (Letisztított, üzenetek nélküli verzió)
@strl.cache_resource
def init_firebase():
    try:
        if "p_key" not in strl.secrets:
            return None
            
        key_dict = {
            "type": "service_account",
            "project_id": "frd-alapanyag",
            "private_key_id": strl.secrets["p_id"],
            "private_key": strl.secrets["p_key"],
            "client_email": "firebase-adminsdk-fbsvc@frd-alapanyag.iam.gserviceaccount.com",
            "client_id": "118377480036110848051",
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
            "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/firebase-adminsdk-fbsvc%40frd-alapanyag.iam.gserviceaccount.com",
            "universe_domain": "googleapis.com"
        }
            
        cred = credentials.Certificate(key_dict)
        
        if not firebase_admin._apps:
            firebase_admin.initialize_app(cred)
            
        return firestore.client()
    except Exception as e:
        strl.error(f"X Rendszerhiba a kapcsolódáskor: {e}")
        return None

db = init_firebase()

# A zöld sikeres csatlakozás sávot itt teljesen töröltük a letisztult felület érdekében, 
# csak a kritikus hibaüzenet maradt meg háttér-biztosítéknak.
if db is None:
    strl.error("X HIBA: A 'db' objektum None maradt!")

# 2. Adatok letöltése a Firestore-ból (Felhőre optimalizált .get() verzió)
@strl.cache_data(ttl=10)
def get_raktar_adatok():
    if db is None:
        strl.error("X HIBA: Nem lehet adatot letölteni, mert a 'db' kliens None!")
        return []
    
    try:
        docs_snapshot = db.collection('materials').get()
        adatok = []
        
        for doc in docs_snapshot:
            d = doc.to_dict()
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
    except Exception as e:
        strl.error(f"X HIBA az adatok letöltése közben: {e}")
        return []

# 3. Formázott Excel generálása színes fejléccel és automatikus oszlopszélességgel
def generalk_formazott_excel(dataframe):
    excel_buffer = BytesIO()
    
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        dataframe.to_excel(writer, index=False, sheet_name='Aktuális Készlet')
        
        workbook = writer.book
        worksheet = writer.sheets['Aktuális Készlet']
        
        # Elegáns sötétkék háttér és félkövér fehér betűstílus a fejlécnek
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
        # Fejléc formázása és sormagasság igazítása
        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment
        worksheet.row_dimensions[1].height = 26
        
        # Automatikus oszlopszélesség kiszámítása a leghosszabb szövegek alapján
        for col in worksheet.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            
            for cell in col:
                if cell.value is not None:
                    val_str = str(cell.value)
                    # Ha van benne sortörés, a leghosszabb sort vesszük alapul
                    lines = val_str.split('\n')
                    for line in lines:
                        if len(line) > max_len:
                            max_len = len(line)
            
            # 3 karakter biztonsági ráhagyás, de minimum 12 egység széles oszlopok
            worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)
            
            # Adatcellák finom balra/középre igazítása a jobb olvashatóságért
            if col_letter in ['A', 'D', 'E', 'F', 'G']:  # Cikkszám, Készletek, Egység, Státusz mehet középre
                for cell in col[1:]:
                    cell.alignment = Alignment(horizontal="center")
            else:
                for cell in col[1:]:
                    cell.alignment = Alignment(horizontal="left")
                    
    return excel_buffer.getvalue()

# --- UI FELÉPÍTÉSE ---

strl.title("📊 FRD Alapanyag Raktár - Vezetői Műszerfal")
strl.caption("Élő, irodai betekintő felület az üzemben lévő tabletek készletéhez")
strl.write("---")

if db is not None:
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
            strl.write("### Riport letöltése")
            
            # Dinamikus, formázott Excel fájl generálása gombnyomásra
            excel_adatok = generalk_formazott_excel(df)
            
            strl.download_button(
                label="📥 Teljes készlet letöltése Excelben",
                data=excel_adatok,
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
