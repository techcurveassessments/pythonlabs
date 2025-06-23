import pandas as pd
import numpy as np

# Create sample inventory data
data = {
    'Product_ID': [101, 102, 102, 103, 104, 105, 106, 107, 108, 109],
    'Product_Name': ['Laptop  ', 'smartphone', 'smartphone', 'Tablet', np.nan, 'HEADPHONES', 'Mouse  ', 'Laptop', 'Keyboard', 'Monitor'],
    'Restock_Date': ['2025-01-10', '01/25/2025', '01/25/2025', np.nan, '2025-02-15', '2/28/25', '2025-03-01', '2025-03-01', '2025-04-10', np.nan],
    'Stock_Level': [50, 200, 200, 75, np.nan, 150, 300, 25, 1000, 80],  # 1000 is an outlier
    'Category': ['Electronics', 'ELECTRONICS', 'electronics', 'Gadgets', 'Audio', np.nan, 'Accessories', 'Electronics', 'accessories', 'Displays']
}

df = pd.DataFrame(data)

# Save to Excel
df.to_excel('inventory_data.xlsx', index=False)
print("Inventory data Excel file generated: inventory_data.xlsx")