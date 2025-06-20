import pandas as pd
import re

# Step 1: Read the Excel file
input_path =  r"C:/Users/91970/source/repos/Chieac/chieac_NLP_Covid/data/COVID_Student_Survey.xlsx"
df = pd.read_excel(input_path)
print(df.shape)
print(df.columns)

# Show top 5 rows
print("\n🔹 First few rows:")
print(df.head())

# Data types and nulls
print("\n🔹 Data types and null values:")
print(df.info())

# Summary statistics for numeric/text data
print("\n🔹 Summary stats:")
print(df.describe(include='all'))

# Count of missing values
print("\n🔹 Missing values per column:")
print(df.isnull().sum())

# Column names
print("\n🔹 All column names:")
print(df.columns.tolist())



# Step 2: Add response length columns
text_columns = [
    "What have you appreciated most about your institution’s response to COVID-19?",
    "What are your biggest worries or concerns as you think about what’s coming up in the next few months?",
    "Is there anything else you’d like to tell us about the way your institution responded to COVID-19 and your experience this term?"
]

# Add a new column for each text column's length
for col in text_columns:
    df[f'{col}_length'] = df[col].astype(str).apply(len)

# Quick summary of response lengths
for col in text_columns:
    print(f"\n--- Length stats for '{col}' ---")
    print(df[f'{col}_length'].describe())



# Step 3: Clean text responses
def clean_text(text):
    text = str(text)
    text = re.sub(r'<.*?>', '', text)  # remove HTML tags
    text = re.sub(r'http\S+', '', text)  # remove URLs
    text = re.sub(r'[^A-Za-z0-9\s]', '', text)  # remove special characters
    text = re.sub(r'\s+', ' ', text).strip()  # remove extra spaces
    return text.lower()

# Apply cleaning to each text column
for col in text_columns:
    cleaned_col = f"{col}_cleaned"
    df[cleaned_col] = df[col].apply(clean_text)

print("\n✅ Text cleaning complete. Sample cleaned values:")
for col in text_columns:
    print(f"\n--- {col} ---")
    print(df[f"{col}_cleaned"].head(3))

# Save cleaned data to new Excel file
output_path = r"C:/Users/91970/source/repos/Chieac/chieac_NLP_Covid/data/COVID_Student_Survey_CLEANED.xlsx"
df.to_excel(output_path, index=False)

print(f"\n✅ Cleaned data saved to: {output_path}")
