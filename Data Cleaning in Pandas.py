import pandas as pd
import numpy as np
import random

np.random.seed(10)

names = [f"Employee_{i}" for i in range(1,121)]   # 120 rows
departments = ["IT", "HR", "Finance", "Marketing", "Sales", "Support"]

data = {
    "Name": names,
    "Age": np.random.randint(22,60,120),
    "Department": [random.choice(departments) for _ in range(120)],
    "Salary": np.random.randint(30000,120000,120),
    "Experience_Years": np.random.randint(0,20,120)
}

df = pd.DataFrame(data)

print("Original Dataset")
print(df.head())

for col in ["Age","Department","Salary","Experience_Years"]:
    idx = np.random.choice(df.index, size=10)
    df.loc[idx, col] = np.nan

print("\nMissing Values Count")
print(df.isnull().sum())


df["Age"] = df["Age"].fillna(df["Age"].mean())

df["Department"] = df["Department"].fillna("Unknown")

df["Salary"] = df["Salary"].fillna(df["Salary"].median())

df["Experience_Years"] = df["Experience_Years"].fillna(0)

print("\nDataset After Handling Missing Values")
print(df.head())


sorted_df = df.sort_values(by="Salary", ascending=False)

print("\nTop 5 Highest Salaries")
print(sorted_df.head())


filtered_df = df[(df["Age"] > 30) & (df["Salary"] > 50000)]

print("\nFiltered Data (Age>30 & Salary>50000)")
print(filtered_df.head())


grouped = df.groupby("Department")["Salary"].mean()

print("\nAverage Salary by Department")
print(grouped)


with pd.ExcelWriter("employee_dataset.xlsx") as writer:
    df.to_excel(writer, sheet_name="Cleaned_Data", index=False)
    sorted_df.to_excel(writer, sheet_name="Sorted_Data", index=False)
    filtered_df.to_excel(writer, sheet_name="Filtered_Data", index=False)

print("\nExcel file 'employee_dataset.xlsx' created successfully")