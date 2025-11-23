import streamlit as st
import pandas as pd

st.title("BME PYQ Analysis Tool")

syllabus_file = st.file_uploader("Upload Syllabus File (.xlsx)", type="xlsx")
pyq_file = st.file_uploader("Upload PYQ File (.xlsx)", type="xlsx")

if syllabus_file and pyq_file:
    syllabus = pd.read_excel(syllabus_file)
    pyq = pd.read_excel(pyq_file)
    results = []
    for _, row in pyq.iterrows():
        question = str(row['Question']).lower()
        year = row['Year']
        best_topic = "Unknown"
        best_unit = "?"
        best_score = 0
        for _, syl_row in syllabus.iterrows():
            concept = str(syl_row['Concept']).lower()
            keywords = str(syl_row['Keywords']).lower()
            score = sum(10 for k in keywords.split(',') if k.strip() in question)
            if concept in question:
                score += 20
            if score > best_score:
                best_score = score
                best_topic = syl_row['Concept']
                best_unit = syl_row['Unit']
        results.append({'Unit': best_unit, 'Topic': best_topic, 'Question': row['Question'], 'Year': year})
    df_results = pd.DataFrame(results)
    topic_counts = df_results.groupby(['Unit', 'Topic']).size().reset_index(name='Frequency')
    topic_counts = topic_counts.sort_values('Frequency', ascending=False)
    st.write("### Top 15 Most Asked Topics")
    st.dataframe(topic_counts.head(15))
    st.write("### Full Question Data")
    st.dataframe(df_results)
    csv = df_results.to_csv(index=False)
    st.download_button("Download Results CSV", csv, "BME_PYQ_Results.csv")
else:
    st.info("Upload both files to start.")

