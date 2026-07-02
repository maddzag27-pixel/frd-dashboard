import streamlit as strl
import firebase_admin
from firebase_admin import credentials, firestore
import pandas as pd
from datetime import datetime, time
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

# 3. ÚJ: Raktári naplófájlok (logs) lekérése adott napra vonatkozóan
@strl.cache_data(ttl=5)
def get_napi_mozgasok(valasztott_datum):
    if db is None:
        return []
    
    try:
        # A választott nap kezdetének és végének beállítása (időzóna-független Firestore szűréshez)
        start_datetime = datetime.combine(valasztott_datum, time.min)
        end_datetime = datetime.combine(valasztott_datum, time.max)
        
        logs_snapshot = db.collection('logs')\
            .where('timestamp', '>=', start_datetime)\
            .where('timestamp', '<=', end_datetime)\
            .get()
            
        mozgasok = []
        tipus_leforditott = {
            "felhasznalas": "Felhasználás",
            "selejt": "Selejt",
            "atgolyositas_ki": "Átalakítás (Kiadás)",
            "atgolyositas_be": "Átalakítás (Beérkezés)"
        }
        
        for doc in logs_snapshot:
            d = doc.to_dict()
            nyers_tipus = d.get('logType', 'felhasznalas')
            mennyiseg = float(d.get('quantity', 0))
            
            mozgasok.append({
                "Időpont": d.get('timestamp').strftime('%H:%M:%S') if d.get('timestamp') else '-',
                "Cikkszám (SKU)": d.get('sku', '-'),
                "Megnevezés": d.get('name', 'Névtelen'),
                "Kategória": d.get('type', 'Egyéb'),
                "Típus": tipus_leforditott.get(nyers_tipus, nyers_tipus),
                "Mennyiség": int(mennyiseg) if mennyiseg % 1 == 0 else mennyiseg,
                "Egység": d.get('unit', 'pár')
            })
            
        # Időrendi sorrendbe rakjuk
        if mozgasok:
            mozgasok.sort(key=lambda x: x["Időpont"])
        return mozgasok
    except Exception as e:
        strl.error(f"X HIBA a naplózott adatok letöltése közben: {e}")
        return []

# 4. Kétfüles, formázott Excel generálása (Készlet + Napi Mozgások)
def generalk_formazott_excel(df_keszlet, df_mozgasok, datum_str):
    excel_buffer = BytesIO()
    
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        # 1. Fül: Aktuális Készlet
        df_keszlet.to_excel(writer, index=False, sheet_name='Aktuális Készlet')
        
        # 2. Fül: Napi Mozgások
        if df_mozgasok.empty:
            # Ha nincs mozgás, egy üres, de fejléces táblát mentünk el
            üres_mozgas = pd.DataFrame(columns=["Időpont", "Cikkszám (SKU)", "Megnevezés", "Kategória", "Típus", "Mennyiség", "Egység"])
            üres_mozgas.to_excel(writer, index=False, sheet_name=f'Mozgások {datum_str}')
        else:
            # Ha van adat, kategóriák szerint rendezve rakjuk bele az Excelbe
            df_rendezett_mozgasok = df_mozgasok.sort_values(by=["Kategória", "Időpont"])
            df_rendezett_mozgasok.to_excel(writer, index=False, sheet_name=f'Mozgások {datum_str}')
            
        workbook = writer.book
        
        # Stílusok meghatározása
        header_fill_keszlet = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid") # Sötétkék
        header_fill_mozgas = PatternFill(start_color="366092", end_color="366092", fill_type="solid")  # Világosabb acélkék
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
        # --- Formázás: 1. Fül ---
        ws1 = workbook['Aktuális Készlet']
        ws1.row_dimensions[1].height = 26
        for cell in ws1[1]:
            cell.fill = header_fill_keszlet
            cell.font = header_font
            cell.alignment = header_alignment
            
        for col in ws1.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws1.column_dimensions[col_letter].width = max(max_len + 3, 12)
            if col_letter in ['A', 'D', 'E', 'F', 'G']:
                for cell in col[1:]: cell.alignment = Alignment(horizontal="center")
            else:
                for cell in col[1:]: cell.alignment = Alignment(horizontal="left")

        # --- Formázás: 2. Fül ---
        ws2 = workbook[f'Mozgások {datum_str}']
        ws2.row_dimensions[1].height = 26
        for cell in ws2[1]:
            cell.fill = header_fill_mozgas
            cell.font = header_font
            cell.alignment = header_alignment
            
        for col in ws2.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws2.column_dimensions[col_letter].width = max(max_len + 3, 12)
            if col_letter in ['A', 'B', 'E', 'F', 'G']:
                for cell in col[1:]: cell.alignment = Alignment(horizontal="center")
            else:
                for cell in col[1:]: cell.alignment = Alignment(horizontal="left")
                
    return excel_buffer.getvalue()

# --- UI FELÉPÍTÉSE ---

strl.title("📊 FRD Alapanyag Raktár - Vezetői Műszerfal")
strl.caption("Élő, irodai betekintő felület az üzemben lévő tabletek készletéhez és napi naplózásához")
strl.write("---")

if db is not None:
    nyers_adatok = get_raktar_adatok()
    df = pd.DataFrame(nyers_adatok)

    if not df.empty:
        # Készlethiányos termékek kiszűrése
        hianyzo_df = df[df["Státusz"] == "🚨 HIÁNY"]
        
        # --- DÁTUMVÁLASZTÓ ÉS EXCEL GENERÁLÁS ---
        col1, col2, col3 = strl.columns([1, 1, 1])
        
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
            # Dátumválasztó beillesztése (alapértelmezetten a mai nap)
            ma = datetime.now().date()
            valasztott_datum = strl.date_input("Válaszd ki a riport napját:", ma)
            datum_str = valasztott_datum.strftime('%Y-%m-%d')

        strl.write("---")
        
        # Lekérjük a naplózott mozgásokat a kiválasztott napra
        nyers_mozgasok = get_napi_mozgasok(valasztott_datum)
        df_mozgasok = pd.DataFrame(nyers_mozgasok)

        # --- NAPI JELENTÉS / KOMBINÁLT RIORT PANEL ---
        strl.subheader(f"📈 Raktári Mozgások és Riport Letöltés: {datum_str}")
        
        rep_col1, rep_col2 = strl.columns([2, 1])
        
        with rep_col1:
            if not df_mozgasok.empty:
                # Gyors összesítés a menedzsmentnek kategória és típus alapján
                szumma = df_mozgasok.groupby(["Kategória", "Típus"])["Mennyiség"].count().reset_index(name="Események száma")
                strl.write(f"**Napi aktivitás összesítve:** {len(df_mozgasok)} könyvelt tranzakció történt a mai napon.")
            else:
                strl.info(f"A választott napon ({datum_str}) még nem történt anyagkiadás vagy selejtezés az üzemben.")
        
        with rep_col2:
            # Kombinált Excel letöltése
            excel_adatok = generalk_formazott_excel(df, df_mozgasok, datum_str)
            strl.download_button(
                label=f"📥 Összetett Excel Riport Letöltése ({datum_str})",
                data=excel_adatok,
                file_name=f"frd_raktar_riport_{datum_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        # Ha vannak mozgások, vizuálisan is kirakjuk kategória szerint rendezhetően
        if not df_mozgasok.empty:
            with strl.expander(f"👀 Részletes napi mozgáslista megtekintése ({datum_str})", expanded=True):
                # Kiválasztható szűrő kifejezetten a mozgástípusra (Felhasználás vs Selejt)
                mozgas_szuro = strl.multiselect("Szűrés mozgástípus szerint:", list(df_mozgasok["Típus"].unique()), default=list(df_mozgasok["Típus"].unique()))
                megjelenitendo_mozgas_df = df_mozgasok[df_mozgasok["Típus"].isin(mozgas_szuro)]
                strl.dataframe(megjelenitendo_mozgas_df, use_container_width=True, hide_index=True)

        strl.write("---")

        # --- FIGYELMEZTETŐ PANEL ---
        if not hianyzo_df.empty:
            strl.error("### 🚨 Az alábbi alapanyagok készlete a kritikus minimum alá süllyedt!")
            strl.dataframe(hianyzo_df, use_container_width=True, hide_index=True)
            strl.write("---")

        # --- SZŰRŐK AZ ASZTALI TÁBLÁZATHOZ ---
        strl.subheader("🔍 Keresés és szűrés a teljes aktuális raktárban")
        f_col1, f_col2 = strl.columns([1, 2])
        
        with f_col1:
            kategoriak = ["Mind"] + sorted(list(df["Kategória"].unique()))
            valasztott_kat = strl.selectbox("Szűrés kategória szerint:", kategoriak)
            
        with f_col2:
            kereses = strl.text_input("Keresés név vagy cikkszám alapján:", "").strip().lower()

        # Szűrések alkalmazása a törzsadatra
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
