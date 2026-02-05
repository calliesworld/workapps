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
        
        # Spaltenauswahl
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Soll-Tabelle")
            spalte1 = st.selectbox("Spalte auswählen:", df1.columns, key="spalte1")
        
        with col2:
            st.subheader("Ist-Tabelle")
            spalte2 = st.selectbox("Spalte auswählen:", df2.columns, key="spalte2")
        
        # Vergleichen-Button
        if st.button("🔍 Listen vergleichen", type="primary", use_container_width=True):
            # Daten extrahieren und bereinigen
            namen1 = set(df1[spalte1].dropna().astype(str).str.strip())
            namen2 = set(df2[spalte2].dropna().astype(str).str.strip())
            
            # Leere Strings entfernen
            namen1 = {name for name in namen1 if name}
            namen2 = {name for name in namen2 if name}
            
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
