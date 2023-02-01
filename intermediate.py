import pandas as pd
import numpy as np
import scipy.stats as stats

# Task 1
# Read the xlsx file into a DataFrame
df = pd.read_excel('intermediate_grades.xlsx')

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
print('Column\tMinimum\tMaximum\tMean\tMedian\tStandard Deviation')
for i in range(data_array.shape[1]):
    print(
        f"{i}\t{statistics[i, 0]:.2f}\t{statistics[i, 1]:.2f}\t{statistics[i, 2]:.2f}\t{statistics[i, 3]:.2f}\t{statistics[i, 4]:.2f}")
