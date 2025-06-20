import nltk
# nltk.download('vader_lexicon')  # Run only once
from nltk.sentiment import SentimentIntensityAnalyzer
import pandas as pd

df = pd.read_excel(r"C:/Users/91970/source/repos/Chieac/chieac_NLP_Covid/data/COVID_Student_Survey_CLEANED.xlsx")

sia = SentimentIntensityAnalyzer()

text_columns = [
    "What have you appreciated most about your institution’s response to COVID-19?",
    "What are your biggest worries or concerns as you think about what’s coming up in the next few months?",
    "Is there anything else you’d like to tell us about the way your institution responded to COVID-19 and your experience this term?"
]

for col in text_columns:
    df[f'{col}_compound'] = df[col].astype(str).apply(lambda x: sia.polarity_scores(x)['compound'])
    df[f'{col}_label'] = df[f'{col}_compound'].apply(
        lambda score: 'Positive' if score > 0.05 else ('Negative' if score < -0.05 else 'Neutral')
    )

for col in text_columns:
    print(f"\nSentiment summary for '{col}':")
    print(df[f'{col}_label'].value_counts())

df.to_excel(r"C:/Users/91970/source/repos/Chieac/chieac_NLP_Covid/data/COVID_Student_Survey_SENTIMENT.xlsx", index=False)
