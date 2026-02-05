import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Excel Namenslisten-Vergleich", layout="wide")

st.title("📊 Excel Namenslisten-Vergleich")
st.markdown("Vergleiche zwei Excel-Dateien und finde Unterschiede zwischen Soll- und Ist-Tabelle")

# Sidebar für Datei-Uploads
with st.sidebar:
    st.header("📁 Dateien hochladen")
    
    st.subheader("Soll-Tabelle (Tabelle 1)")
    datei1 = st.file_uploader("Excel-Datei auswählen (Soll)", type=['xlsx', 'xls'], key="datei1")
    
    st.subheader("Ist-Tabelle (Tabelle 2)")
    datei2 = st.file_uploader("Excel-Datei auswählen (Ist)", type=['xlsx', 'xls'], key="datei2")

# Hauptbereich
if datei1 and datei2:
    try:
        # Dateien einlesen (Zeile 3 als Header)
        df1 = pd.read_excel(datei1, header=2)
        df2 = pd.read_excel(datei2, header=2)
        
        # Automatische Erkennung von Vorname/Nachname-Spalten
        def finde_spalte(df, suchbegriff):
            """Findet eine Spalte, die den Suchbegriff enthält (case-insensitive)"""
            for col in df.columns:
                if suchbegriff.lower() in str(col).lower():
                    return col
            return None
        
        # Für Tabelle 1
        vorname1_auto = finde_spalte(df1, "vorname")
        nachname1_auto = finde_spalte(df1, "nachname")
        
        # Für Tabelle 2
        vorname2_auto = finde_spalte(df2, "vorname")
        nachname2_auto = finde_spalte(df2, "nachname")
        
        # Standard-Index setzen
        vorname1_idx = list(df1.columns).index(vorname1_auto) if vorname1_auto else 0
        nachname1_idx = list(df1.columns).index(nachname1_auto) if nachname1_auto else 0
        vorname2_idx = list(df2.columns).index(vorname2_auto) if vorname2_auto else 0
        nachname2_idx = list(df2.columns).index(nachname2_auto) if nachname2_auto else 0
        
        # Spaltenauswahl
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Soll-Tabelle")
            if vorname1_auto and nachname1_auto:
                st.success(f"✅ Automatisch erkannt: '{vorname1_auto}' und '{nachname1_auto}'")
            else:
                st.warning("⚠️ 'Vorname' oder 'Nachname' nicht automatisch gefunden")
            
            spalte1_vorname = st.selectbox("Vorname-Spalte:", df1.columns, index=vorname1_idx, key="spalte1_vorname")
            spalte1_nachname = st.selectbox("Nachname-Spalte:", df1.columns, index=nachname1_idx, key="spalte1_nachname")
        
        with col2:
            st.subheader("Ist-Tabelle")
            if vorname2_auto and nachname2_auto:
                st.success(f"✅ Automatisch erkannt: '{vorname2_auto}' und '{nachname2_auto}'")
            else:
                st.warning("⚠️ 'Vorname' oder 'Nachname' nicht automatisch gefunden")
            
            spalte2_vorname = st.selectbox("Vorname-Spalte:", df2.columns, index=vorname2_idx, key="spalte2_vorname")
            spalte2_nachname = st.selectbox("Nachname-Spalte:", df2.columns, index=nachname2_idx, key="spalte2_nachname")
        
        # Vergleichen-Button
        if st.button("🔍 Listen vergleichen", type="primary", use_container_width=True):
            # Namen kombinieren: Vorname + Nachname
            df1['vollname'] = df1[spalte1_vorname].astype(str).str.strip() + ' ' + df1[spalte1_nachname].astype(str).str.strip()
            df2['vollname'] = df2[spalte2_vorname].astype(str).str.strip() + ' ' + df2[spalte2_nachname].astype(str).str.strip()
            
            # Daten extrahieren und bereinigen
            namen1 = set(df1['vollname'].dropna().str.strip())
            namen2 = set(df2['vollname'].dropna().str.strip())
            
            # Leere Strings und "nan nan" entfernen
            namen1 = {name for name in namen1 if name and name.strip() and name != 'nan nan'}
            namen2 = {name for name in namen2 if name and name.strip() and name != 'nan nan'}
            
            # Vergleich durchführen
            fehlen_in_ist = sorted(namen1 - namen2)
            ueberfluessig_in_ist = sorted(namen2 - namen1)
            
            # Ergebnisse anzeigen
            st.markdown("---")
            st.header("📋 Ergebnisse")
            
            col_result1, col_result2 = st.columns(2)
            
            with col_result1:
                st.subheader(f"🔴 Im Soll, aber nicht im Ist")
                st.caption(f"{len(fehlen_in_ist)} Einträge fehlen")
                
                if fehlen_in_ist:
                    for name in fehlen_in_ist:
                        st.write(f"• {name}")
                else:
                    st.success("✅ Keine fehlenden Einträge")
            
            with col_result2:
                st.subheader(f"🟠 Im Ist, aber nicht im Soll")
                st.caption(f"{len(ueberfluessig_in_ist)} überflüssige Einträge")
                
                if ueberfluessig_in_ist:
                    for name in ueberfluessig_in_ist:
                        st.write(f"• {name}")
                else:
                    st.success("✅ Keine überflüssigen Einträge")
            
            # Export-Option
            st.markdown("---")
            st.subheader("💾 Ergebnis exportieren")
            
            # DataFrame für Export erstellen
            max_len = max(len(fehlen_in_ist), len(ueberfluessig_in_ist)) if fehlen_in_ist or ueberfluessig_in_ist else 0
            fehlen_padded = fehlen_in_ist + [''] * (max_len - len(fehlen_in_ist))
            ueberfluessig_padded = ueberfluessig_in_ist + [''] * (max_len - len(ueberfluessig_in_ist))
            
            df_export = pd.DataFrame({
                'Fehlt im Ist (aus Soll-Tabelle)': fehlen_padded,
                'Überflüssig im Ist (nicht in Soll)': ueberfluessig_padded
            })
            
            # Excel-Export vorbereiten mit korrekter Engine
            output = io.BytesIO()
            df_export.to_excel(output, index=False, engine='openpyxl')
            excel_data = output.getvalue()
            
            st.download_button(
                label="📥 Als Excel herunterladen",
                data=excel_data,
                file_name="vergleichsergebnis.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
    except Exception as e:
        st.error(f"❌ Fehler beim Verarbeiten der Dateien: {e}")

elif datei1 or datei2:
    st.info("ℹ️ Bitte beide Excel-Dateien hochladen, um den Vergleich zu starten.")
else:
    st.info("👈 Bitte lade beide Excel-Dateien in der Sidebar hoch, um zu beginnen.")
    
    # Anleitung
    with st.expander("📖 Anleitung"):
        st.markdown("""
        ### So funktioniert's:
        
        1. **Soll-Tabelle hochladen** - Die Referenztabelle mit den erwarteten Einträgen
        2. **Ist-Tabelle hochladen** - Die zu prüfende Tabelle
        3. **Spalten auswählen** - Wähle für jede Tabelle die zu vergleichende Spalte
        4. **Vergleichen** - Klicke auf "Listen vergleichen"
        5. **Exportieren** - Optional: Lade das Ergebnis als Excel herunter
        
        ### Format der Excel-Dateien:
        - Zeile 3 muss die Spaltenüberschriften enthalten
        - Daten ab Zeile 4
        - Unterstützte Formate: .xlsx, .xls
        """)

