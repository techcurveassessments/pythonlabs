# Step-by-Step Guide to Test Visualization Script with New Excel Files

This guide demonstrates the `excel_data_visualizer.py` script using two new Excel files: `inventory_data.xlsx` (inventory management) and `student_data.xlsx` (student performance). Optionally, use `dynamic_excel_cleaner.py` to clean the files first for optimal visualization results.

## Prerequisites

- Python 3.x with pandas, numpy, openpyxl, matplotlib, seaborn (`pip install pandas numpy openpyxl matplotlib seaborn`).
- Working directory with scripts and Excel files.

## Step 1: Set Up the Scripts

1. **Save the Visualization Script**:
   - Copy `excel_data_visualizer.py` (artifact_id: 2f90e0d1-d2d3-4dc5-b915-69e6b6f288e2) to your directory.
2. **Save the Cleaning Script (Optional)**:
   - Copy `dynamic_excel_cleaner.py` (artifact_id: 1e52b09b-3210-49c5-a652-4c37536f4d46) for cleaning data.
3. **Save the Data Generation Scripts**:
   - Copy `generate_inventory_data.py` (artifact_id: f8c7d9d3-5e8a-4c2e-8f8b-3c8e7e6f2e3) to generate `inventory_data.xlsx`.
   - Copy `generate_student_data.py` (artifact_id: c3b2a7e4-6f9b-4d1f-9a8c-4d9f8e7f3e4) to generate `student_data.xlsx`.

## Step 2: Generate and Clean Inventory Data

1. **Generate the File**:
   - Run `python generate_inventory_data.py`.
   - Output: `inventory_data.xlsx` created.
   - Inspect in Excel to note duplicates, missing values, and inconsistencies.
2. **Clean the File (Recommended)**:
   - Rename `inventory_data.xlsx` to `input_data.xlsx`.
   - Run `python dynamic_excel_cleaner.py`.
   - Output: Cleaned file (e.g., `cleaned_data_20250623_123625.xlsx`).
   - Verify: 9 rows (duplicate removed), no missing values, standardized dates, lowercase strings, clipped outliers.
3. **Prepare for Visualization**:
   - Rename the cleaned file (or raw `inventory_data.xlsx` if skipping cleaning) to `input_data.xlsx`.

## Step 3: Visualize Inventory Data

1. **Run the Visualization Script**:
   - Run `python excel_data_visualizer.py`.
   - Output:
     - Console shows data summary and chart recommendations.
     - PNGs saved in `visualizations` folder.
   - Expected Visualizations:
     - Histogram: Stock_Level distribution (e.g., most products 50–300 units).
     - Boxplot: Stock_Level range (outliers clipped, e.g., 1000 → ~300).
     - Bar: Category frequency (e.g., "electronics" dominant).
     - Line: Stock_Level over Restock_Date (stock trends).
     - Box/Bar: Stock_Level by Category (e.g., higher stock in "electronics").
2. **Verify Output**:
   - Check `visualizations` folder for PNGs (e.g., `Bar_Category_20250623_123625.png`).
   - Open images to confirm charts reflect cleaned data (e.g., no extreme Stock_Level values).

## Step 4: Generate and Clean Student Data

1. **Generate the File**:
   - Run `python generate_student_data.py`.
   - Output: `student_data.xlsx` created.
   - Inspect in Excel to note duplicates, missing values, and inconsistencies.
2. **Clean the File (Recommended)**:
   - Rename `student_data.xlsx` to `input_data.xlsx`.
   - Run `python dynamic_excel_cleaner.py`.
   - Output: Cleaned file (e.g., `cleaned_data_20250623_123626.xlsx`).
   - Verify: 7 rows (duplicate removed), no missing values, standardized dates, lowercase strings, clipped outliers.
3. **Prepare for Visualization**:
   - Rename the cleaned file (or raw `student_data.xlsx` if skipping cleaning) to `input_data.xlsx`.

## Step 5: Visualize Student Data

1. **Run the Visualization Script**:
   - Run `python excel_data_visualizer.py`.
   - Output:
     - Console shows data summary and chart recommendations.
     - PNGs saved in `visualizations` folder.
   - Expected Visualizations:
     - Histogram: Exam_Score distribution (e.g., most scores 65–95).
     - Boxplot: Exam_Score range (outliers clipped, e.g., 200 → ~100).
     - Bar: Subject frequency (e.g., "science" dominant).
     - Line: Exam_Score over Exam_Date (score trends).
     - Box/Bar: Exam_Score by Subject (e.g., higher scores in "math").
2. **Verify Output**:
   - Check `visualizations` folder for PNGs (e.g., `Line_Exam_Score_over_Exam_Date_20250623_123626.png`).
   - Open images to confirm charts reflect cleaned data (e.g., no invalid Exam_Score values).

## Step 6: Compare Results Across Domains

- **Inventory Data**: Visualizes stock levels (numeric) by category (categorical) and restock dates (time), highlighting inventory trends and distributions.
- **Student Data**: Visualizes exam scores (numeric) by subject (categorical) and exam dates (time), showing performance patterns.
- **Script Capability**:
  - Infers column types (numeric, categorical, date) for tailored chart recommendations.
  - Generates domain-relevant visualizations (e.g., stock trends for inventory, score distributions for students).
  - Benefits from cleaned data to produce accurate, Power BI-ready visuals.

## Step 7: Import into Power BI (Optional)

1. Open Power BI Desktop.
2. Import cleaned Excel files (`input_data.xlsx`) via "Get Data" > "Excel Workbook".
3. Import PNGs from `visualizations` folder as images for static reports.
4. Create dynamic charts in Power BI to replicate or extend visualizations (e.g., bar charts for Category, line charts for Exam_Score trends).

## Conclusion

The `excel_data_visualizer.py` script, combined with `dynamic_excel_cleaner.py`, effectively handles diverse datasets from inventory and education domains. The new Excel files test its ability to analyze and visualize complex data, ensuring meaningful insights for analysis in Power BI or other tools.
