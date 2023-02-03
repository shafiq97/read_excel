import pandas as pd
import numpy as np
import scipy.stats as stats
import matplotlib.pyplot as plt

# Task 1
# Read the xlsx file into a DataFrame
df = pd.read_excel('final_grades.xlsx')
column_names = df.columns
print(column_names)
# Convert the DataFrame to a numpy array
data_array = np.array(df)

# Check the data types of each column
print(df.dtypes)

# Convert a column to a specific data type, if necessary
# df['Student Id'] = df['Student Id'].astype(int)
# df['Session 2'] = df['Session 2'].astype(int)
# df['Session 3'] = df['Session 3'].astype(int)
# df['Session 4'] = df['Session 4'].astype(int)
# df['Session 5'] = df['Session 5'].astype(int)


# Calculate summary statistics for each column in the array
statistics = np.zeros((data_array.shape[1], 6))
for i in range(data_array.shape[1]):
    statistics[i, 0] = np.min(data_array[:, i])
    statistics[i, 1] = np.max(data_array[:, i])
    statistics[i, 2] = np.mean(data_array[:, i])
    statistics[i, 3] = np.median(data_array[:, i])
    statistics[i, 4] = np.std(data_array[:, i])
    statistics[i, 5] = stats.skew(data_array[:, i])

# Print the summary statistics
# print(statistics)

# Detect outliers using the Z-score method
for i in range(data_array.shape[1]):
    z = np.abs(stats.zscore(data_array[:, i]))
    outliers = np.where(z > 3)
    print(f"Outliers in column {i}: {outliers}")

# Print the summary statistics in a table format
print('Column\t\tMinimum\tMaximum\tMean\tMedian\tStandard Deviation')
for i in range(1, data_array.shape[1]):
    print(
        f"{column_names[i]}\t{statistics[i, 0]:.2f}\t{statistics[i, 1]:.2f}\t{statistics[i, 2]:.2f}\t{statistics[i, 3]:.2f}\t{statistics[i, 4]:.2f}")

# for i in range(data_array.shape[1]):
#     plt.hist(data_array[:, i], edgecolor='black')
#     plt.title(f"Column {i}")
#     plt.xlabel("Value")
#     plt.ylabel("Count")
#     plt.show()

marks = data_array[1:, 17]

# Plot a histogram of the marks
plt.hist(marks, edgecolor='black')
plt.title("Marks")
plt.xlabel("Mark")
plt.ylabel("Number of Students")
plt.show()

# Define the grade boundaries
grade_boundaries = [0, 40, 50, 60, 70, 80]
grade_labels = ['Fail', 'E', 'D', 'C', 'B', 'A']

# Determine the grade for each mark
grades = np.zeros_like(marks, dtype=int)
for i, boundary in enumerate(grade_boundaries):
    grades[marks >= boundary] = i


# Plot a pie chart of the grades
unique_grades, counts = np.unique(grades, return_counts=True)
plt.pie(counts, labels=[f"{grade_labels[g]} ({counts[i]} students, {grade_boundaries[g]}-{grade_boundaries[g+1]})" for i, g in enumerate(unique_grades[:-1])] + [f"{grade_labels[unique_grades[-1]]} ({counts[-1]} students, {grade_boundaries[-1]}-100)"])
plt.title("Grades")
plt.show()