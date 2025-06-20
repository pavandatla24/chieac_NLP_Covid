import pandas as pd
import matplotlib.pyplot as plt

# Load the dataset with sentiment results
df = pd.read_excel(r"C:/Users/91970/source/repos/Chieac/chieac_NLP_Covid/data/COVID_Student_Survey_SENTIMENT.xlsx")

# List of text columns we processed
text_columns = [
    "What have you appreciated most about your institution’s response to COVID-19?",
    "What are your biggest worries or concerns as you think about what’s coming up in the next few months?",
    "Is there anything else you’d like to tell us about the way your institution responded to COVID-19 and your experience this term?"
]

# Set consistent style
plt.style.use("seaborn-v0_8")

# Create visualizations for each question
for col in text_columns:
    label_col = f"{col}_label"

    # Bar Chart
    plt.figure(figsize=(6, 4))
    df[label_col].value_counts().plot(kind='bar', color=['green', 'gray', 'red'])
    plt.title(f"Sentiment Distribution -\n{col}", fontsize=10)
    plt.ylabel("Number of Responses")
    plt.xlabel("Sentiment")
    plt.xticks(rotation=0)
    plt.tight_layout()
    plt.show()

    # Pie Chart
    plt.figure(figsize=(5, 5))
    df[label_col].value_counts().plot(
        kind='pie',
        autopct='%1.1f%%',
        startangle=90,
        colors=['green', 'gray', 'red'],
        labels=['Positive', 'Neutral', 'Negative']
    )
    plt.title(f"Sentiment % Breakdown -\n{col}", fontsize=10)
    plt.ylabel("")
    plt.tight_layout()
    plt.show()
