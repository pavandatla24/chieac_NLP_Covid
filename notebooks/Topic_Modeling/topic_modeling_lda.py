import pandas as pd
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.decomposition import LatentDirichletAllocation

# Load the cleaned data
input_path = r"C:/Users/91970/source/repos/Chieac/chieac_NLP_Covid/data/COVID_Student_Survey_TOPIC_READY.xlsx"
df = pd.read_excel(input_path)

# Columns that were cleaned for topic modeling
topic_columns = [
    "What have you appreciated most about your institution’s response to COVID-19?_topic",
    "What are your biggest worries or concerns as you think about what’s coming up in the next few months?_topic",
    "Is there anything else you’d like to tell us about the way your institution responded to COVID-19 and your experience this term?_topic"
]

# Function to perform LDA topic modeling
def run_lda(text_series, n_topics=5, top_words=10):
    vectorizer = CountVectorizer(max_df=0.95, min_df=2, stop_words='english')
    dtm = vectorizer.fit_transform(text_series.dropna())
    
    lda = LatentDirichletAllocation(n_components=n_topics, random_state=42)
    lda.fit(dtm)

    words = vectorizer.get_feature_names_out()
    
    print("\n💡 Top keywords per topic:")
    for i, topic in enumerate(lda.components_):
        print(f"\nTopic {i + 1}:")
        print(", ".join([words[i] for i in topic.argsort()[-top_words:][::-1]]))

# Run LDA on each open-ended question
for col in topic_columns:
    print(f"\n==============================")
    print(f"📝 LDA Topic Modeling for Column: {col}")
    print("==============================")
    run_lda(df[col], n_topics=5, top_words=10)
