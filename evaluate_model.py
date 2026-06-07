import warnings
warnings.filterwarnings("ignore")

from datasets import load_dataset
from rouge_score import rouge_scorer
import numpy as np

# We can import your existing algorithm directly from your web app!
import app 

print("Downloading/Loading CNN/DailyMail dataset (this might take a minute)...")
# We grab just the first 50 articles from the 'test' split so it runs quickly
dataset = load_dataset("cnn_dailymail", "3.0.0", split="test[:50]")

# Initialize the ROUGE grader
scorer = rouge_scorer.RougeScorer(['rouge1', 'rouge2', 'rougeL'], use_stemmer=True)
rouge_1_scores = []
rouge_2_scores = []
rouge_L_scores = []

print(f"\nStarting evaluation on {len(dataset)} articles...\n")

for i, example in enumerate(dataset):
    article_text = example['article']
    human_reference = example['highlights'] # These are the true, human-written bullet points

    # 1. Feed the article into YOUR algorithm
    raw_sents, clean_sents, full_doc = app.preprocess_text(article_text)
    generated_bullets, _ = app.extract_top_sentences(raw_sents, clean_sents, full_doc, top_n=3)

    # Skip if the text was too short to summarize
    if not generated_bullets:
        continue

    # Convert the list of bullets back into a single string so the grader can read it
    generated_summary = " ".join(generated_bullets)

    # 2. Grade your algorithm against the human reference
    scores = scorer.score(human_reference, generated_summary)
    
    # Save the F1-Measures (the harmonic mean of precision and recall)
    rouge_1_scores.append(scores['rouge1'].fmeasure)
    rouge_2_scores.append(scores['rouge2'].fmeasure)
    rouge_L_scores.append(scores['rougeL'].fmeasure)

    # Print progress so you know it hasn't frozen
    if (i + 1) % 10 == 0:
        print(f"Graded {i + 1}/50 articles...")

# 3. Calculate and Print the Final Report
print("\n" + "="*40)
print("🏆 FINAL EXTRACTIVE MODEL GRADES 🏆")
print("="*40)
print(f"ROUGE-1 (Single Word Match):   {np.mean(rouge_1_scores) * 100:.2f}%")
print(f"ROUGE-2 (Two-Word Phrase Match): {np.mean(rouge_2_scores) * 100:.2f}%")
print(f"ROUGE-L (Sentence Flow Match):   {np.mean(rouge_L_scores) * 100:.2f}%")
print("="*40)
print("Note: In unsupervised extractive NLP, ROUGE-1 scores above 35% are considered highly successful!")

import json

# 4. Bundle the final averages into a dictionary and convert NumPy floats to standard Python floats
final_results = {
    "ROUGE-1_Percentage": round(float(np.mean(rouge_1_scores)) * 100, 2),
    "ROUGE-2_Percentage": round(float(np.mean(rouge_2_scores)) * 100, 2),
    "ROUGE-L_Percentage": round(float(np.mean(rouge_L_scores)) * 100, 2)
}

# Save to JSON file
with open("rouge_evaluation_scores.json", "w") as outfile:
    json.dump(final_results, outfile, indent=4)
    
print("✅ Successfully saved ROUGE scores to rouge_evaluation_scores.json!")