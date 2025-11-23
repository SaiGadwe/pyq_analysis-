import pandas as pd
import os

# ==================== STEP 1: CHECK FILES ====================
# print("Checking files in current folder...")
# print("Current folder:", os.getcwd())
# print("\nFiles found:")
# for file in os.listdir():
#     print(f"  - {file}")
# print()

# File names - MUST match exactly!
SYLLABUS_FILE = 'RGPV_Basic_Mechanical_Engineering_Syllabus.xlsx'
PYQ_FILE = 'BME_PYQ.xlsx'

# ==================== STEP 2: LOAD FILES ====================
print("="*60)
print("BASIC MECHANICAL ENGINEERING - PYQ ANALYSIS")
print("="*60)
print(f"\nLoading files...")

try:
    # Read syllabus (must have: Unit, Concept, Keywords columns)
    syllabus = pd.read_excel(SYLLABUS_FILE)
    print(f"✓ Syllabus loaded: {len(syllabus)} topics")

    # Read PYQ (must have: Question, Year columns)
    pyq = pd.read_excel(PYQ_FILE)
    print(f"✓ PYQ loaded: {len(pyq)} questions")

except FileNotFoundError as e:
    print(f"\n❌ ERROR: File not found!")
    print(f"Missing file: {e}")
    print(f"\nMake sure these files are in the same folder as this script:")
    print(f"  1. {SYLLABUS_FILE}")
    print(f"  2. {PYQ_FILE}")
    print(f"\n💡 TIP: Check the file list above to see what files are actually in your folder!")
    exit()

except Exception as e:
    print(f"\n❌ ERROR: {e}")
    print(f"\n💡 Make sure your Excel files have the correct column names:")
    print(f"   Syllabus file: Unit, Concept, Keywords")
    print(f"   PYQ file: QUESTION, YEAR")
    exit()

# ==================== STEP 3: MATCH QUESTIONS TO TOPICS ====================
print(f"\nAnalyzing questions...")

results = []

for index, row in pyq.iterrows():
    question_text = str(row['QUESTION']).lower()
    year = row['YEAR']

    # Tokenize question words for more robust matching
    question_words = set(question_text.split())

    # Find best matching topic from syllabus
    best_topic = "Unknown"
    best_unit = "?"
    best_score = 0

    for _, syl_row in syllabus.iterrows():
        concept_raw = str(syl_row['Concept'])
        keywords_raw = str(syl_row['Keywords'])

        # Preprocess syllabus concept and keywords into sets of words
        syl_concept_words = set(concept_raw.lower().split())
        
        # Split keywords by comma, then further split each by space to get individual words
        syl_keywords_split = [word.strip().lower() for k in keywords_raw.split(',') for word in k.split() if word.strip()]
        syl_keywords_words = set(syl_keywords_split)

        score = 0

        # Score for direct concept phrase match (higher weight)
        if concept_raw.lower() in question_text:
            score += 50
        else:
            # Score for individual concept words matching
            overlapping_concept_words = syl_concept_words.intersection(question_words)
            score += len(overlapping_concept_words) * 10 # 10 points per matching unique concept word

        # Score for individual keyword words matching
        overlapping_keyword_words = syl_keywords_words.intersection(question_words)
        score += len(overlapping_keyword_words) * 5 # 5 points per matching unique keyword word

        # Save if this is the best match so far
        if score > best_score:
            best_score = score
            best_topic = syl_row['Concept']
            best_unit = syl_row['Unit']

    # Store result
    results.append({
        'Unit': best_unit,
        'Topic': best_topic,
        'Question': row['QUESTION'],
        'Year': year,
        'Match_Score': best_score
    })

# Convert to DataFrame
df_results = pd.DataFrame(results)

# ==================== STEP 4: COUNT FREQUENCY ====================
# Count how many times each topic appears
topic_counts = df_results.groupby(['Unit', 'Topic']).size().reset_index(name='Frequency')
topic_counts = topic_counts.sort_values('Frequency', ascending=False)

# ==================== STEP 5: SHOW RESULTS ====================
print("\n" + "="*80)
print("RESULTS - MOST FREQUENTLY ASKED TOPICS IN BME")
print("="*80)

# Show top 20 topics
for idx, row in topic_counts.head(20).iterrows():
    print(f"\n📌 Unit {row['Unit']} | {row['Topic']}")
    print(f"   ⭐ Asked {row['Frequency']} times")

    # Show all questions for this topic
    topic_questions = df_results[df_results['Topic'] == row['Topic']]
    print(f"   Questions:")
    for i, q_row in topic_questions.iterrows():
        q_text = q_row['Question']
        # Shorten if too long
        if len(q_text) > 75:
            q_text = q_text[:75] + "..."
        year_val = q_row['Year']
        # Extract year if it's a Timestamp, else use as string/int
        if hasattr(year_val, 'year'):
            year_str = str(year_val.year)
        else:
            year_str = str(year_val)
        print(f"      [{year_str}] {q_text}")

print("\n" + "="*80)

# ==================== STEP 6: SAVE TO EXCEL ====================
output_file = "BME_Analysis_Results.xlsx"
df_results.to_excel(output_file, index=False)
print(f"\n💾 Results saved to: {output_file}")
print(f"📊 Total questions analyzed: {len(df_results)}")
print(f"📚 Total unique topics found: {len(topic_counts)}")

# Show unit-wise summary
print(f"\n📋 UNIT-WISE SUMMARY:")
# Convert 'Unit' column to string before sorting to handle mixed types (numbers and '?')
unit_summary = df_results.groupby('Unit').size().sort_index(key=lambda x: x.astype(str))
for unit, count in unit_summary.items():
    if unit != '?':
        print(f"   Unit {unit}: {count} questions")

print("\n✅ ANALYSIS COMPLETE!")
print("="*80)

# ==================== INVESTIGATE UNKNOWN QUESTIONS ====================
print("\n" + "="*80)
print("INVESTIGATING UNKNOWN QUESTIONS")
print("="*80)

unknown_questions_df = df_results[df_results['Topic'] == 'Unknown']

if not unknown_questions_df.empty:
    for index, q_row in unknown_questions_df.iterrows():
        question_text_lower = str(q_row['Question']).lower()
        print(f"\nQuestion: {q_row['Question']}")
        print(f"  Year: {q_row['Year']}")
        print(f"  Match Score: {q_row['Match_Score']}")
        print(f"  Attempting to find partial matches with syllabus keywords/concepts:")

        found_partial_match = False
        for _, syl_row in syllabus.iterrows():
            concept_raw = str(syl_row['Concept'])
            keywords_raw = str(syl_row['Keywords'])
            syl_concept_words = set(concept_raw.lower().split())
            syl_keywords_split = [word.strip().lower() for k in keywords_raw.split(',') for word in k.split() if word.strip()]
            syl_keywords_words = set(syl_keywords_split)

            # Check for any overlapping words (even with low scores)
            overlap_concept = syl_concept_words.intersection(question_text_lower.split())
            overlap_keyword = syl_keywords_words.intersection(question_text_lower.split())

            if overlap_concept or overlap_keyword:
                print(f"    Potential Partial Match in Unit {syl_row['Unit']} | Concept: {syl_row['Concept']}")
                if overlap_concept:
                    print(f"      Overlapping Concept Words: {', '.join(overlap_concept)}")
                if overlap_keyword:
                    print(f"      Overlapping Keyword Words: {', '.join(overlap_keyword)}")
                found_partial_match = True
        
        if not found_partial_match:
            print(f"    No significant keyword overlap found even for partial matching.")
else:
    print("No questions found categorized as 'Unit ? | Unknown'.")

print("="*80)
