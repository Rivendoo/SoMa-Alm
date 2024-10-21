import pandas as pd
import streamlit as st
from nltk.stem import PorterStemmer
from nltk.tokenize import word_tokenize
from rapidfuzz import fuzz
import nltk
import io

# Ladda NLTK-resurser
nltk.download('punkt')

# Initiera stemmer
stemmer = PorterStemmer()

def stem_text(text):
    tokens = word_tokenize(text.lower())
    stemmed = [stemmer.stem(token) for token in tokens]
    return ' '.join(stemmed)

# Streamlit UI
st.title('SoMa-Alm-24')

# Ladda upp Excel-fil istället för att sätta arbetskatalog
uploaded_file = st.file_uploader("Välj en Excel-fil", type=["xlsx"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Ett fel inträffade vid inläsning av filen: {e}")
        st.stop()
    
    # Initiera session_state för att lagra filtrerad_df
    if 'filtrerad_df' not in st.session_state:
        st.session_state['filtrerad_df'] = None

    # Inmatningsfält för nyckelord
    nyckelord_input = st.text_input('Ange nyckelord (separera med kommatecken):')

    # Filtrera data när knappen klickas
    if st.button('Visa Antal Matchningar'):
        nyckelord = [word.strip() for word in nyckelord_input.split(',')]
        stemmed_keywords = [stemmer.stem(word.lower()) for word in nyckelord]
        
        try:
            # Stemma alla relevanta kolumner i DataFrame
            df_stemmed = df.applymap(lambda x: stem_text(str(x)) if isinstance(x, str) else x)
            
            # Filtrera med fuzzy matching
            def fuzzy_match(row):
                for item in row:
                    if isinstance(item, str):
                        for keyword in stemmed_keywords:
                            ratio = fuzz.partial_ratio(keyword, item)
                            if ratio > 80:  # Justera tröskelvärdet vid behov
                                return True
                return False
            
            mask = df_stemmed.apply(fuzzy_match, axis=1)
            filtrerad_df = df[mask]
            
            # Uppdatera session_state med det filtrerade DataFrame
            st.session_state['filtrerad_df'] = filtrerad_df
            
            # Uppdatera antal matchande rader och procentandel
            match_count = filtrerad_df.shape[0]
            procentandel = (match_count / df.shape[0]) * 100
            
            # Visa resultatet
            st.write(f"Antal matchande rader: {match_count}")
            st.write(f"Procentandel av totalen: {procentandel:.2f}%")
            
            if match_count > 0:
                st.success("Filtreringen lyckades! Klicka nedan för att spara filen.")
            else:
                st.info("Inga rader matchar nyckelorden. Inget att spara.")
            
        except Exception as e:
            st.error(f"Ett fel inträffade vid filtreringen: {e}")

    # Spara den filtrerade filen som en nedladdning med anpassad knapp
    if st.session_state.get('filtrerad_df') is not None and not st.session_state['filtrerad_df'].empty:
        try:
            output = io.BytesIO()
            st.session_state['filtrerad_df'].to_excel(output, index=False, engine='openpyxl')
            output.seek(0)
            st.download_button(
                label="Spara fil förberedd för Cassidy",
                data=output,
                file_name='Program2024-sherlock.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        except Exception as e:
            st.error(f"Ett fel inträffade vid generering av nedladdningslänk: {e}")
