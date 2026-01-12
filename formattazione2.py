import streamlit as st
import os
import json
from google import genai
from google.genai import types
from docxtpl import DocxTemplate
from pptx import Presentation
import io

# --- CONFIGURAZIONE PAGINA ---
st.set_page_config(page_title="Generatore Documenti AI", page_icon="🪄", layout="wide")

st.title("🪄 Generatore Report GEOCOM® (Word & PPT)")
st.markdown("Trasforma i tuoi appunti grezzi in documenti ufficiali & ben formattati con un click!")

# --- CONFIGURAZIONE API KEY ---
# Se vuoi fissare la chiave nel codice, inseriscila tra le virgolette qui sotto.
API_KEY_FISSA = "" 

if API_KEY_FISSA:
    api_key = API_KEY_FISSA
else:
    api_key = st.sidebar.text_input("Google AI Studio Key", type="password")
    st.sidebar.info("Richiedi la chiave su aistudio.google.com")

# --- SELETTORE FORMATO ---
col1, col2 = st.columns([1, 2])
with col1:
    formato_scelto = st.radio(
        "Scegli il formato di output:",
        ["Documento Word (.docx)", "Presentazione PowerPoint (.pptx)"],
        index=0
    )

# --- INPUT UTENTE ---
testo_grezzo = st.text_area("Incolla qui il testo del report o gli appunti:", height=350)

# --- FUNZIONI ---

def get_gemini_data(text, key, formato):
    """
    Chiama Gemini per strutturare i dati.
    Cambia il prompt in base se stiamo facendo un Word (dettagliato) o PPT (sintetico).
    """
    client = genai.Client(api_key=key)
    
    # Prompt Base
    istruzioni_base = """
    Sei un assistente editoriale esperto. Estrai i dati dal testo fornito e restituisci un JSON.
    Struttura JSON richiesta:
    {
        "titolo_report": "Titolo breve e incisivo",
        "sottotitolo_report": "Sottotitolo o contesto",
        "data_odierna": "DD MMMM YYYY",
        "lista_sezioni": [
            { "titolo": "Titolo Sezione", "testo": "Contenuto..." }
        ]
    }
    """

    # Personalizzazione Prompt in base al formato
    if formato == "ppt":
        istruzioni_extra = """
        ATTENZIONE: Stiamo creando una PRESENTAZIONE POWERPOINT.
        1. Il campo 'testo' DEVE essere un elenco puntato sintetico (usa simboli come • o -).
        2. Sii molto conciso. Niente muri di testo.
        3. Se un argomento è lungo, dividilo in più sezioni logiche.
        """
    else:
        istruzioni_extra = """
        ATTENZIONE: Stiamo creando un REPORT WORD DETTAGLIATO.
        1. Il campo 'testo' deve essere discorsivo, professionale e completo.
        2. Usa un linguaggio formale aziendale.
        3. Usa \\n per andare a capo nei paragrafi.
        """

    prompt_finale = f"{istruzioni_base}\n{istruzioni_extra}\n\nTESTO DA ANALIZZARE:\n{text}"

    # Chiamata al modello (Usa l'alias 'latest' per evitare errori di quota)
    response = client.models.generate_content(
        model='gemini-flash-latest', 
        contents=prompt_finale,
        config=types.GenerateContentConfig(
            response_mime_type='application/json'
        )
    )
    
    return json.loads(response.text)

def generate_doc(data):
    """Genera il file Word usando docxtpl"""
    if not os.path.exists("template_aziendale.docx"):
        st.error("ERRORE: Manca il file 'template_aziendale.docx' nella cartella!")
        return None

    doc = DocxTemplate("template_aziendale.docx")
    doc.render(data)
    
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

def generate_ppt(data):
    """Genera il file PowerPoint usando python-pptx"""
    if not os.path.exists("template_aziendale.pptx"):
        st.error("ERRORE: Manca il file 'template_aziendale.pptx' nella cartella!")
        return None

    prs = Presentation("template_aziendale.pptx")

    # --- SLIDE 1: COPERTINA (Layout 0) ---
    try:
        layout_copertina = prs.slide_layouts[0] 
        slide = prs.slides.add_slide(layout_copertina)
        
        # Titolo
        if slide.shapes.title:
            slide.shapes.title.text = data.get('titolo_report', 'Report')
        
        # Sottotitolo (Placeholder 1)
        if len(slide.placeholders) > 1:
            slide.placeholders[1].text = f"{data.get('sottotitolo_report', '')}\n{data.get('data_odierna', '')}"
    except Exception as e:
        st.warning(f"Attenzione nella slide copertina: {e}")

    # --- SLIDE SUCCESSIVE: CONTENUTI (Layout 1) ---
    try:
        layout_contenuto = prs.slide_layouts[1] # Solitamente Titolo + Contenuto
    except:
        layout_contenuto = prs.slide_layouts[0] # Fallback se manca il layout 1

    for sezione in data.get('lista_sezioni', []):
        slide = prs.slides.add_slide(layout_contenuto)
        
        # Titolo Slide
        if slide.shapes.title:
            slide.shapes.title.text = sezione.get('titolo', '')
        
        # Corpo Slide (Placeholder 1 è il box testo principale)
        # Nota: controlliamo che esista un secondo placeholder (il corpo)
        if len(slide.placeholders) > 1:
            body = slide.placeholders[1]
            body.text = sezione.get('testo', '')

    bio = io.BytesIO()
    prs.save(bio)
    bio.seek(0)
    return bio

# --- LOGICA APPLICAZIONE ---
if st.button("🚀 Genera Documento", type="primary"):
    if not api_key:
        st.error("⚠️ Manca la API Key di Google Gemini!")
    elif not testo_grezzo:
        st.warning("⚠️ Inserisci del testo prima di generare.")
    else:
        
        # Determina tipo formato per passare il parametro corretto alla funzione
        tipo_formato = "ppt" if "PowerPoint" in formato_scelto else "word"
        
        with st.spinner(f"L'AI sta analizzando il testo per creare un {tipo_formato.upper()}..."):
            try:
                # 1. Estrazione Dati con AI
                dati_strutturati = get_gemini_data(testo_grezzo, api_key, tipo_formato)
                st.success("✅ Dati analizzati con successo!")
                
                # Mostra anteprima JSON (utile per debug)
                with st.expander("🔍 Vedi dati estratti (JSON)"):
                    st.json(dati_strutturati)

                # 2. Generazione File
                file_output = None
                nome_file = ""
                mime_type = ""

                if tipo_formato == "word":
                    file_output = generate_doc(dati_strutturati)
                    nome_file = "Report_Finale.docx"
                    mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                else:
                    file_output = generate_ppt(dati_strutturati)
                    nome_file = "Presentazione_Finale.pptx"
                    mime_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

                # 3. Download Button
                if file_output:
                    st.download_button(
                        label=f"📥 Scarica {nome_file}",
                        data=file_output,
                        file_name=nome_file,
                        mime=mime_type
                    )
                
            except Exception as e:
                st.error(f"❌ Si è verificato un errore: {e}")
                
                
                ## per runnarlo ---> python -m streamlit run formattazione3.py