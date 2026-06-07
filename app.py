import warnings
warnings.filterwarnings("ignore", category=UserWarning)

import json # <-- NEW: Needed to read your benchmark scores
from flask import Flask, render_template, request
import spacy
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import numpy as np
import nltk
from nltk.corpus import wordnet as wn

nltk.download('wordnet', quiet=True)
nltk.download('omw-1.4', quiet=True)

print("Loading spaCy model...")
nlp = spacy.load("en_core_web_sm")

app = Flask(__name__)

def preprocess_text(text):
    doc = nlp(text)
    original_sentences = [sent.text.strip() for sent in doc.sents if sent.text.strip()]
    cleaned_sentences = []
    
    for sent in original_sentences:
        sent_doc = nlp(sent)
        clean_tokens = [
            token.lemma_.lower() for token in sent_doc 
            if not token.is_stop and not token.is_punct and token.text.strip()
        ]
        cleaned_sentences.append(" ".join(clean_tokens))
        
    return original_sentences, cleaned_sentences

def get_wup_similarity(sent_doc, top_doc_words):
    total_sim = 0
    valid_words = 0
    for token in sent_doc:
        if token.pos_ in ["NOUN", "VERB"]:
            synsets1 = wn.synsets(token.lemma_)
            if not synsets1: continue
            s1 = synsets1[0]
            
            max_sim = 0
            for word in top_doc_words:
                synsets2 = wn.synsets(word)
                if not synsets2: continue
                s2 = synsets2[0]
                
                sim = s1.wup_similarity(s2)
                if sim and sim > max_sim:
                    max_sim = sim
            
            total_sim += max_sim
            valid_words += 1
    return total_sim / valid_words if valid_words > 0 else 0

def extract_top_sentences(original_sentences, cleaned_sentences, top_n=3):
    alpha, beta, gamma, delta = 0.4, 0.2, 0.2, 0.2

    target_n = min(top_n, len(original_sentences))

    if target_n == 0:
        return [], [], None

    vectorizer = TfidfVectorizer()
    tfidf_matrix = vectorizer.fit_transform(cleaned_sentences)
    
    tfidf_scores = np.array(tfidf_matrix.sum(axis=1)).flatten()
    if tfidf_scores.max() > 0:
        tfidf_scores = tfidf_scores / tfidf_scores.max()

    feature_names = vectorizer.get_feature_names_out()
    doc_word_scores = tfidf_matrix.sum(axis=0).A1
    top_doc_words = [feature_names[i] for i in doc_word_scores.argsort()[-5:]]

    final_scores = []
    scorecard = []
    
    for idx, sent_text in enumerate(original_sentences):
        sent_doc = nlp(sent_text)
        
        important_words = sum(1 for token in sent_doc if token.pos_ in ["NOUN", "PROPN", "VERB"])
        pos_score = important_words / len(sent_doc) if len(sent_doc) > 0 else 0
        semantic_score = get_wup_similarity(sent_doc, top_doc_words)
        structural_score = 1.0 / (idx + 1)
        
        w_tfidf, w_pos = alpha * tfidf_scores[idx], beta * pos_score
        w_semantic, w_struct = gamma * semantic_score, delta * structural_score
        total_score = w_tfidf + w_pos + w_semantic + w_struct
        
        final_scores.append(total_score)
        
        scorecard.append({
            "index": idx + 1, "sentence": sent_text,
            "tfidf": round(w_tfidf, 3), "pos": round(w_pos, 3),
            "semantic": round(w_semantic, 3), "struct": round(w_struct, 3),
            "total": round(total_score, 3), "is_selected": False
        })

    final_scores = np.array(final_scores)
    similarity_matrix = cosine_similarity(tfidf_matrix)
    lambda_param = 0.7  
    
    selected_indices = [int(np.argmax(final_scores))]
    candidates = list(set(range(len(original_sentences))) - set(selected_indices))

    while len(selected_indices) < target_n and candidates:
        mmr_scores = []
        for candidate in candidates:
            relevance = final_scores[candidate]
            max_similarity = max([similarity_matrix[candidate][s] for s in selected_indices])
            mmr_score = (lambda_param * relevance) - ((1 - lambda_param) * max_similarity)
            mmr_scores.append((candidate, mmr_score))

        best_candidate = max(mmr_scores, key=lambda x: x[1])[0]
        selected_indices.append(best_candidate)
        candidates.remove(best_candidate)

    top_indices = sorted(selected_indices)

    avg_tfidf, avg_pos, avg_sem, avg_str = 0, 0, 0, 0
    for i in top_indices:
        scorecard[i]["is_selected"] = True
        avg_tfidf += scorecard[i]["tfidf"]
        avg_pos += scorecard[i]["pos"]
        avg_sem += scorecard[i]["semantic"]
        avg_str += scorecard[i]["struct"]
    
    chart_data = {
        "tfidf": round(avg_tfidf / target_n, 3), 
        "pos": round(avg_pos / target_n, 3),
        "semantic": round(avg_sem / target_n, 3), 
        "struct": round(avg_str / target_n, 3)
    }
    
    summary = [original_sentences[i] for i in top_indices]
    return summary, scorecard, chart_data

@app.route("/", methods=["GET", "POST"])
def home():
    summary_bullets, analysis_data, chart_data = None, None, None
    user_input = ""
    num_bullets = 3

    # ==========================================
    # NEW: Load the Benchmark Scores from the JSON
    # ==========================================
    try:
        with open('rouge_evaluation_scores.json', 'r') as file:
            benchmark_scores = json.load(file)
    except FileNotFoundError:
        # Fallback if the file gets deleted or moved
        benchmark_scores = {"ROUGE-1_Percentage": "N/A", "ROUGE-2_Percentage": "N/A", "ROUGE-L_Percentage": "N/A"}

    if request.method == "POST":
        user_input = request.form.get("raw_text", "")
        num_bullets = int(request.form.get("num_bullets", 3))

        if user_input.strip():
            raw_sents, clean_sents = preprocess_text(user_input)
            summary_bullets, analysis_data, chart_data = extract_top_sentences(raw_sents, clean_sents, top_n=num_bullets)

    return render_template("index.html", 
                           summary=summary_bullets, analysis=analysis_data,
                           chart_data=chart_data, original_text=user_input, 
                           num_bullets=num_bullets,
                           scores=benchmark_scores) # <-- NEW: Pass scores to HTML

if __name__ == "__main__":
    app.run(debug=True)