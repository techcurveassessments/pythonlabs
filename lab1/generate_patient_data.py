import pandas as pd
import numpy as np

# Create sample patient data
data = {
    'Patient_ID': [1001, 1002, 1002, 1003, 1004, 1005, 1006, 1007, 1008],
    'Patient_Name': ['Emma Watson ', 'michael chen', 'michael chen', np.nan, 'Sophia Lee', 'JAMES BROWN', '  Olivia Kim', 'Emma Watson', 'Liam Noah'],
    'Admission_Date': ['2024-07-01', '07/15/2024', '07/15/2024', '2024-08-10', np.nan, '8/20/24', '2024-09-05', '2024-09-05', '2024-10-01'],
    'Systolic_BP': [120, 180, 180, np.nan, 140, 200, 130, 110, 999],  # 999 is an outlier
    'Diagnosis': ['Hypertension', 'diabetes', 'Diabetes', 'Flu', np.nan, 'HYPERTENSION', 'flu', 'Diabetes', 'Asthma']
}

df = pd.DataFrame(data)

# Save to Excel
df.to_excel('patient_data.xlsx', index=False)
print("Patient data Excel file generated: patient_data.xlsx")