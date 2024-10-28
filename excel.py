import pandas as pd

# Define your data (replace this with your actual data)
data = {
    'Rentang Nilai': ['50-54', '55-59', '60-64', '65-69', '70-74', '75-79', '80-84', '85-89', '90-94', '95-99'],
    'Frekuensi': [1, 2, 11, 10, 12, 21, 6, 9, 4, 4]
}

# Create a DataFrame
df = pd.DataFrame(data)

# Calculate the midpoint for each class interval
df['Midpoint'] = [52, 57, 62, 67, 72, 77, 82, 87, 92, 97]

# Calculate cumulative frequency
df['Cumulative Frequency'] = df['Frekuensi'].cumsum()

# Total frequency
n = df['Frekuensi'].sum()

# Mean calculation
df['fx'] = df['Frekuensi'] * df['Midpoint']  # f * x
mean = df['fx'].sum() / n

# Median calculation
# Locate the median class: n/2
median_class_index = df[df['Cumulative Frequency'] >= n / 2].index[0]
L_median = 50 + (median_class_index * 5)  # Lower boundary of the median class
F_median = df['Cumulative Frequency'][median_class_index - 1] if median_class_index > 0 else 0
f_median = df['Frekuensi'][median_class_index]
h = 5  # Class interval width
median = L_median + ((n/2 - F_median) / f_median) * h

# Mode calculation
modal_class_index = df['Frekuensi'].idxmax()  # Find the modal class index
L_mode = 50 + (modal_class_index * 5)  # Lower boundary of the modal class
f1 = df['Frekuensi'][modal_class_index - 1] if modal_class_index > 0 else 0
f2 = df['Frekuensi'][modal_class_index + 1] if modal_class_index < len(df) - 1 else 0
f_mode = df['Frekuensi'][modal_class_index]
mode = L_mode + ((f_mode - f1) / ((f_mode - f1) + (f_mode - f2))) * h

# Q1 (First Quartile) Calculation (25th percentile)
q1_class_index = df[df['Cumulative Frequency'] >= n / 4].index[0]
L_q1 = 50 + (q1_class_index * 5)
F_q1 = df['Cumulative Frequency'][q1_class_index - 1] if q1_class_index > 0 else 0
f_q1 = df['Frekuensi'][q1_class_index]
q1 = L_q1 + ((n/4 - F_q1) / f_q1) * h

# Q3 (Third Quartile) Calculation (75th percentile)
q3_class_index = df[df['Cumulative Frequency'] >= 3 * n / 4].index[0]
L_q3 = 50 + (q3_class_index * 5)
F_q3 = df['Cumulative Frequency'][q3_class_index - 1] if q3_class_index > 0 else 0
f_q3 = df['Frekuensi'][q3_class_index]
q3 = L_q3 + ((3 * n / 4 - F_q3) / f_q3) * h

# D5 (50th percentile) is the same as the median
d5 = median

# P10 (10th percentile)
p10_class_index = df[df['Cumulative Frequency'] >= n * 10 / 100].index[0]
L_p10 = 50 + (p10_class_index * 5)
F_p10 = df['Cumulative Frequency'][p10_class_index - 1] if p10_class_index > 0 else 0
f_p10 = df['Frekuensi'][p10_class_index]
p10 = L_p10 + ((n * 10 / 100 - F_p10) / f_p10) * h

# P25 (25th percentile) is the same as Q1
p25 = q1

# Creating a summary table of the results
summary_data = {
    'Statistic': ['Mean', 'Median', 'Mode', 'Q1', 'Q3', 'D5', 'P10', 'P25'],
    'Value': [mean, median, mode, q1, q3, d5, p10, p25]
}

summary_df = pd.DataFrame(summary_data)

# Saving the summary table and detailed data to a new Excel file
with pd.ExcelWriter('statistics_summary.xlsx') as writer:
    df.to_excel(writer, sheet_name='Data', index=False)
    summary_df.to_excel(writer, sheet_name='Summary', index=False)

'statistics_summary.xlsx'
