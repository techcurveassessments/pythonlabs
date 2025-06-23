# Step-by-Step Guide to Demonstrate Dynamic Excel Cleaning Script

This guide demonstrates the capability of the `dynamic_excel_cleaner.py` script using two diverse Excel files: `sales_data.xlsx` (sales domain) and `patient_data.xlsx` (healthcare domain). The script cleans data by handling duplicates, missing values, inconsistent formats, and outliers, producing Power BI-ready output.

## Prerequisites

- Python 3.x installed with pandas, numpy, and openpyxl libraries (`pip install pandas numpy openpyxl`).
- A working directory to store the scripts and Excel files.

## Step 1: Set Up the Scripts

1. **Save the Cleaning Script**:

   - Copy the `dynamic_excel_cleaner.py` script from the previous conversation (artifact_id: 1e52b09b-3210-49c5-a652-4c37536f4d46) into your working directory.
   - This script reads `input_data.xlsx`, cleans it, and outputs a file like `cleaned_data_20250623_122025.xlsx`.
2. **Save the Data Generation Scripts**:

   - **Sales Data**: Copy the `generate_sales_data.py` script (artifact_id: d184cb1d-cc17-423f-a1d8-6d21f05d1d00) to generate `sales_data.xlsx`.
   - **Patient Data**: Copy the `generate_patient_data.py` script (artifact_id: 51feefaf-38f1-4b80-8a83-98ccb47042c3) to generate `patient_data.xlsx`.
   - Save both in the same directory as `dynamic_excel_cleaner.py`.

## Step 2: Generate the Sales Data Excel File

1. **Run the Sales Data Script**:

   - Execute `python generate_sales_data.py` in your terminal or IDE.
   - Output: `sales_data.xlsx` is created in the working directory.
   - File Characteristics:
     - 10 rows, 5 columns: Order_ID, Customer_Name, Order_Date, Sales_Amount, Region.
     - Issues: Duplicates (rows 2 and 3), missing values (e.g., Customer_Name in row 6), mixed case strings (e.g., "alice smith", "SOUTH"), varied date formats (e.g., "2023-01-15", "3/15/23"), outlier in Sales_Amount (9999.99).
2. **Inspect the File**:

   - Open `sales_data.xlsx` in Excel to confirm the data.
   - Note the inconsistencies, such as "alice smith" vs. "Alice Smith" and the extreme Sales_Amount value.
3. **Prepare for Cleaning**:

   - Rename `sales_data.xlsx` to `input_data.xlsx` (the cleaning script expects this filename).

## Step 3: Clean the Sales Data

1. **Run the Cleaning Script**:

   - Execute `python dynamic_excel_cleaner.py`.
   - Output: A new file (e.g., `cleaned_data_20250623_122025.xlsx`) is created.
   - The script performs:
     - **Duplicates**: Removes row 3 (identical to row 2), reducing to 9 rows.
     - **Missing Values**:
       - Sales_Amount: Fills missing value with median (e.g., median of [100.50, 200.75, 1500.00, 400.00, 250.00, 275.50]).
       - Customer_Name, Region: Fills with mode (e.g., "alice smith", "south").
       - Order_Date: Fills with latest date (2023-05-01).
     - **Dates**: Standardizes Order_Date to datetime (e.g., "2023-01-15").
     - **Strings**: Converts Customer_Name, Region to lowercase, trims whitespace.
     - **Outliers**: Clips Sales_Amount 9999.99 to IQR upper bound (e.g., ~1500.00).
2. **Verify the Output**:

   - Open the cleaned Excel file in Excel or Power BI.
   - Check:
     - 9 rows (duplicate removed).
     - No missing values (e.g., Sales_Amount in row 6 filled with median).
     - Consistent dates (e.g., "2023-01-15" format).
     - Lowercase strings (e.g., "alice smith", "south").
     - Sales_Amount 9999.99 clipped to a reasonable value.

## Step 4: Generate the Patient Data Excel File

1. **Run the Patient Data Script**:

   - Execute `python generate_patient_data.py`.
   - Output: `patient_data.xlsx` is created.
   - File Characteristics:
     - 9 rows, 5 columns: Patient_ID, Patient_Name, Admission_Date, Systolic_BP, Diagnosis.
     - Issues: Duplicates (rows 2 and 3), missing values (e.g., Patient_Name in row 4), mixed case strings (e.g., "michael chen", "HYPERTENSION"), varied date formats (e.g., "2024-07-01", "8/20/24"), outlier in Systolic_BP (999).
2. **Inspect the File**:

   - Open `patient_data.xlsx` in Excel to confirm the data.
   - Note issues like "Diabetes" vs. "diabetes" and the unrealistic Systolic_BP value (999).
3. **Prepare for Cleaning**:

   - Rename `patient_data.xlsx` to `input_data.xlsx`.

## Step 5: Clean the Patient Data

1. **Run the Cleaning Script**:

   - Execute `python dynamic_excel_cleaner.py`.
   - Output: A new file (e.g., `cleaned_data_20250623_122026.xlsx`) is created.
   - The script performs:
     - **Duplicates**: Removes row 3 (identical to row 2), reducing to 8 rows.
     - **Missing Values**:
       - Systolic_BP: Fills missing value with median (e.g., median of [120, 180, 140, 200, 130, 110, 999]).
       - Patient_Name, Diagnosis: Fills with mode (e.g., "emma watson", "diabetes").
       - Admission_Date: Fills with latest date (2024-10-01).
     - **Dates**: Standardizes Admission_Date to datetime (e.g., "2024-07-01").
     - **Strings**: Converts Patient_Name, Diagnosis to lowercase, trims whitespace.
     - **Outliers**: Clips Systolic_BP 999 to IQR upper bound (e.g., ~270).
2. **Verify the Output**:

   - Open the cleaned Excel file in Excel or Power BI.
   - Check:
     - 8 rows (duplicate removed).
     - No missing values (e.g., Systolic_BP in row 4 filled with median).
     - Consistent dates (e.g., "2024-07-01" format).
     - Lowercase strings (e.g., "emma watson", "diabetes").
     - Systolic_BP 999 clipped to a reasonable value.

## Step 6: Compare Results Across Domains

- **Sales Data (Step 3)**:
  - Handles financial data with large numeric outliers (Sales_Amount) and varied date formats.
  - Ensures consistency in customer and region names for reporting in Power BI.
- **Patient Data (Step 5)**:
  - Manages medical data with unrealistic measurements (Systolic_BP) and diagnosis inconsistencies.
  - Standardizes dates and strings for clinical analysis in Power BI.
- **Script Capability**:
  - Infers column types (numeric, categorical, date) dynamically.
  - Applies context-appropriate cleaning (median for numerics, mode for categoricals, IQR for outliers).
  - Produces clean, Power BI-ready files regardless of domain-specific issues.

## Step 7: Import into Power BI (Optional)

1. Open Power BI Desktop.
2. Select "Get Data" > "Excel Workbook" and load the cleaned files (e.g., `cleaned_data_20250623_122025.xlsx`).
3. Verify that:
   - Data loads without errors.
   - Dates are recognized as datetime.
   - No missing values or duplicates.
   - Strings are consistent (lowercase, no extra spaces).
4. Create sample visualizations (e.g., bar charts for Sales_Amount by Region or Systolic_BP by Diagnosis) to confirm data quality.

## Conclusion

The `dynamic_excel_cleaner.py` script effectively cleans diverse datasets from sales and healthcare domains. By handling duplicates, missing values, inconsistent formats, and outliers, it ensures the output is reliable for analysis in Power BI. This guide demonstrates its robustness and adaptability, making it a valuable tool for data preparation across industries.
