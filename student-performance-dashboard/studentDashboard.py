import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load the Excel file
file_path = r"data/MGS E05.xlsx"
df = pd.read_excel(file_path, skiprows=10)

# Drop completely empty rows and columns
df.dropna(how='all', inplace=True)
df.dropna(axis=1, how='all', inplace=True)

# Optional: Rename columns if needed
df.rename(columns={
    'Name': 'Name',
    'Total   100%': 'Total_Score'
}, inplace=True)

# Ensure Total Score is numeric
df['Total_Score'] = pd.to_numeric(df['Total_Score'], errors='coerce')
df.dropna(subset=['Total_Score'], inplace=True)

# Descriptive statistics
print("Descriptive Statistics:")
print(df['Total_Score'].describe())

# Plot distribution of scores
plt.figure(figsize=(8, 6))
sns.histplot(df['Total_Score'], bins=10, kde=True, color='skyblue')
plt.title('Distribution of Student Total Scores')
plt.xlabel('Total Score')
plt.ylabel('Number of Students')
plt.tight_layout()
plt.savefig('visuals/score_distribution.png')
plt.show()

# Highlight low performers
low_scores = df[df['Total_Score'] < 50]
print("\nStudents Scoring Below 50%:")
print(low_scores[['Name', 'Total_Score']])

# Save cleaned and analyzed data
df.to_csv('data/cleaned_student_scores.csv', index=False)
print("\nCleaned data saved as 'data/cleaned_student_scores.csv'")
