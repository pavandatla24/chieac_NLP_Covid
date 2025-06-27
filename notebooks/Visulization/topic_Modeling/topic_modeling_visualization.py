import pandas as pd
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.decomposition import LatentDirichletAllocation
import pyLDAvis
import pyLDAvis.lda_model as pyldavis_sklearn

import os

# Load cleaned + topic-ready data
input_path = r"C:/Users/91970/source/repos/Chieac/chieac_NLP_Covid/data/COVID_Student_Survey_TOPIC_READY.xlsx"
df = pd.read_excel(input_path)

# List of topic columns and friendly names for file naming
columns = {
    "What have you appreciated most about your institution’s response to COVID-19?_topic": "appreciation",
    "What are your biggest worries or concerns as you think about what’s coming up in the next few months?_topic": "worries",
    "Is there anything else you’d like to tell us about the way your institution responded to COVID-19 and your experience this term?_topic": "other_feedback"
}

output_dir = r"C:/Users/91970/source/repos/Chieac/chieac_NLP_Covid/output"
os.makedirs(output_dir, exist_ok=True)

# Loop through each column
for col, label in columns.items():
    print(f"Processing column: {col}")
    
    texts = df[col].dropna().astype(str).tolist()
    
    vectorizer = CountVectorizer(max_df=0.95, min_df=2, stop_words='english')
    dtm = vectorizer.fit_transform(texts)
    
    lda_model = LatentDirichletAllocation(n_components=5, random_state=42)
    lda_model.fit(dtm)
    
    vis = pyldavis_sklearn.prepare(lda_model, dtm, vectorizer)
    html_path = os.path.join(output_dir, f"topic_vis_{label}.html")
    pyLDAvis.save_html(vis, html_path)
    
    print(f"✅ Saved visualization to {html_path}")

print("\n🎉 All topic visualizations completed.")
