import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Excel Namenslisten-Vergleich", layout="wide")

st.title("📊 Excel Namenslisten-Vergleich")
st.markdown("Vergleiche zwei Excel-Dateien und finde Unterschiede zwischen Soll- und Ist-Tabelle")

# Sidebar für Datei-Uploads
with st.sidebar:
    st.header("📁 Dateien hochladen")
    
    st.subheader("Soll-Tabelle (Referenz)")
    datei_soll = st.file_uploader("Excel-Datei auswählen (Soll)", type=['xlsx', 'xls'], key="datei_soll")
    
    st.subheader("Ist-Tabellen (Studiengänge)")
    st.caption("Du kannst mehrere Dateien hochladen")
    dateien_ist = st.file_uploader(
        "Excel-Dateien auswählen (Ist)", 
        type=['xlsx', 'xls'], 
        accept_multiple_files=True,
        key="dateien_ist"
    )

# Hauptbereich
if datei_soll and dateien_ist:
    try:
        # Soll-Datei einlesen (Zeile 1 als Header)
        df_soll = pd.read_excel(datei_soll, header=0)
        
        # Automatische Erkennung von Vorname/Nachname-Spalten
        def finde_vorname_spalte(df):
            """Findet die Vorname-Spalte"""
            for col in df.columns:
                col_str = str(col).strip().lower()
                if col_str == "vorname" or col_str.startswith("vorname ") or col_str.startswith("vorname:"):
                    return col
            return None
        
        def finde_name_spalte(df):
            """Findet die Name/Nachname-Spalte (aber NICHT Vorname!)"""
            for col in df.columns:
                col_str = str(col).strip().lower()
                # Exakte Übereinstimmung für "name" oder "nachname"
                if col_str == "name" or col_str == "nachname":
                    return col
                # Mit Suffix wie "Name:" oder "Nachname "
                if col_str.startswith("name ") or col_str.startswith("name:"):
                    return col
                if col_str.startswith("nachname ") or col_str.startswith("nachname:"):
                    return col
            return None
        
        # Soll-Tabelle prüfen
        vorname_soll = finde_vorname_spalte(df_soll)
        name_soll = finde_name_spalte(df_soll)
        
        st.subheader("Soll-Tabelle")
        if vorname_soll and name_soll:
            st.success(f"✅ Erkannt: '{vorname_soll}' und '{name_soll}'")
        else:
            st.error(f"❌ Spalten 'Vorname' und 'Name' in Soll-Tabelle nicht gefunden!")
            st.info(f"Gefunden: Vorname={vorname_soll}, Name={name_soll}")
            st.info(f"Verfügbare Spalten: {', '.join(df_soll.columns)}")
            st.stop()
        
        # Alle Ist-Tabellen einlesen und kombinieren
        st.subheader(f"Ist-Tabellen ({len(dateien_ist)} Studiengänge)")
        df_ist_kombiniert = pd.DataFrame()
        
        for i, datei_ist in enumerate(dateien_ist):
            df_ist_temp = pd.read_excel(datei_ist, header=0)
            
            vorname_ist = finde_vorname_spalte(df_ist_temp)
            name_ist = finde_name_spalte(df_ist_temp)
            
            if vorname_ist and name_ist:
                st.success(f"✅ {datei_ist.name}: '{vorname_ist}' und '{name_ist}'")
                df_ist_kombiniert = pd.concat([df_ist_kombiniert, df_ist_temp], ignore_index=True)
            else:
                st.error(f"❌ {datei_ist.name}: Spalten nicht gefunden!")
                st.info(f"Verfügbare Spalten: {', '.join(df_ist_temp.columns)}")
                st.stop()
        
        # Vergleichen-Button
        if st.button("🔍 Listen vergleichen", type="primary", use_container_width=True):
            # Namen kombinieren: Vorname + Nachname
            df_soll['vollname'] = df_soll[vorname_soll].astype(str).str.strip() + ' ' + df_soll[name_soll].astype(str).str.strip()
            df_ist_kombiniert['vollname'] = df_ist_kombiniert[vorname_soll].astype(str).str.strip() + ' ' + df_ist_kombiniert[name_soll].astype(str).str.strip()
            
            # Daten extrahieren und bereinigen
            namen_soll = set(df_soll['vollname'].dropna().str.strip())
            namen_ist = set(df_ist_kombiniert['vollname'].dropna().str.strip())
            
            # Leere Strings und "nan nan" entfernen
            namen_soll = {name for name in namen_soll if name and name.strip() and name != 'nan nan'}
            namen_ist = {name for name in namen_ist if name and name.strip() and name != 'nan nan'}
            
            # Vergleich durchführen
            fehlen_in_ist = sorted(namen_soll - namen_ist)
            ueberfluessig_in_ist = sorted(namen_ist - namen_soll)
            
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

elif datei_soll or dateien_ist:
    st.info("ℹ️ Bitte lade die Soll-Tabelle UND mindestens eine Ist-Tabelle hoch.")
else:
    st.info("👈 Bitte lade die Soll-Tabelle und Ist-Tabellen in der Sidebar hoch, um zu beginnen.")
    
    # Anleitung
    with st.expander("📖 Anleitung"):
        st.markdown("""
        ### So funktioniert's:
        
        1. **Soll-Tabelle hochladen** - Die Referenztabelle mit allen erwarteten Namen
        2. **Ist-Tabellen hochladen** - Eine oder mehrere Excel-Dateien (z.B. verschiedene Studiengänge)
        3. **Automatische Erkennung** - Spalten "Vorname" und "Name" werden automatisch gefunden
        4. **Vergleichen** - Alle Ist-Tabellen werden kombiniert und mit dem Soll verglichen
        5. **Exportieren** - Optional: Lade das Ergebnis als Excel herunter
        
        ### Format der Excel-Dateien:
        - Zeile 1 muss die Spaltenüberschriften "Vorname" und "Name" enthalten
        - Daten ab Zeile 2
        - Unterstützte Formate: .xlsx, .xls
        - Du kannst mehrere Ist-Tabellen gleichzeitig hochladen
        """)
