import pandas as pd
import numpy as np
from datetime import datetime

# Read the Excel file
df = pd.read_excel('input_data.xlsx')

# Remove duplicate rows
df = df.drop_duplicates()

# Infer column types
numeric_cols = df.select_dtypes(include=['int64', 'float64']).columns
categorical_cols = df.select_dtypes(include=['object']).columns
date_cols = []

# Detect potential date columns by checking for date-like strings
for col in categorical_cols:
    try:
        # Sample a few values to check if they can be converted to datetime
        sample = df[col].dropna().sample(min(5, len(df[col].dropna())))
        if all(pd.to_datetime(sample, errors='coerce').notna()):
            df[col] = pd.to_datetime(df[col], errors='coerce')
            date_cols.append(col)
            categorical_cols = categorical_cols.drop(col)
    except:
        continue

# Handle missing values
# Numeric: Fill with median to reduce outlier impact
for col in numeric_cols:
    df[col] = df[col].fillna(df[col].median())

# Categorical: Fill with mode
for col in categorical_cols:
    df[col] = df[col].fillna(df[col].mode()[0] if not df[col].mode().empty else 'Unknown')

# Dates: Fill with most recent date or leave as NaT
for col in date_cols:
    df[col] = df[col].fillna(df[col].max())

# Clean strings: Trim whitespace and standardize case
if len(categorical_cols) > 0:
    df[categorical_cols] = df[categorical_cols].apply(lambda x: x.str.strip().str.lower())

# Remove outliers in numeric columns using IQR method
for col in numeric_cols:
    Q1 = df[col].quantile(0.25)
    Q3 = df[col].quantile(0.75)
    IQR = Q3 - Q1
    lower_bound = Q1 - 1.5 * IQR
    upper_bound = Q3 + 1.5 * IQR
    df[col] = df[col].clip(lower=lower_bound, upper=upper_bound)

# Save cleaned data with timestamp to avoid overwriting
timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
output_file = f'cleaned_data_{timestamp}.xlsx'
df.to_excel(output_file, index=False)

print(f"Cleaned data saved to {output_file}")