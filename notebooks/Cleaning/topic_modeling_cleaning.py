import pandas as pd
import re
import nltk
from nltk.corpus import stopwords

# Download stopwords if not already done
nltk.download('stopwords')
stop_words = set(stopwords.words('english'))

# Load raw Excel file
input_path = r"C:/Users/91970/source/repos/Chieac/chieac_NLP_Covid/data/COVID_Student_Survey.xlsx"
df = pd.read_excel(input_path)

# Define relevant text columns
text_columns = [
    "What have you appreciated most about your institution’s response to COVID-19?",
    "What are your biggest worries or concerns as you think about what’s coming up in the next few months?",
    "Is there anything else you’d like to tell us about the way your institution responded to COVID-19 and your experience this term?"
]

# Clean text for topic modeling
def clean_for_topic_modeling(text):
    text = str(text)
    text = re.sub(r'<.*?>', '', text)  # remove HTML tags
    text = re.sub(r'http\S+', '', text)  # remove URLs
    text = re.sub(r'[^A-Za-z0-9\s]', '', text)  # remove special characters
    text = re.sub(r'\s+', ' ', text).strip()  # remove extra spaces
    text = text.lower()

    tokens = text.split()
    tokens = [word for word in tokens if word not in stop_words and len(word) >= 3]
    return " ".join(tokens)

# Apply cleaning
for col in text_columns:
    cleaned_col = f"{col}_topic"
    df[cleaned_col] = df[col].apply(clean_for_topic_modeling)

# Save to cleaned file for topic modeling
output_path = r"C:/Users/91970/source/repos/Chieac/chieac_NLP_Covid/data/COVID_Student_Survey_TOPIC_READY.xlsx"
df.to_excel(output_path, index=False)
print(f"\n✅ Topic-modeling ready data saved to: {output_path}")
