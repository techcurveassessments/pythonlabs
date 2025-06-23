import pandas as pd
import numpy as np
from datetime import datetime

# Create sample sales data
data = {
    'Order_ID': [1, 2, 2, 3, 4, 5, 6, 7, 8, 9],
    'Customer_Name': ['John Doe', 'alice smith', 'alice smith', 'Bob Jones', 'CHARLIE BROWN', np.nan, 'Eve Wilson', 'JOHN DOE', 'Alice Smith', 'Frank Lee'],
    'Order_Date': ['2023-01-15', '01/20/2023', '01/20/2023', '2023-02-01', np.nan, '2023-03-10', '3/15/23', '2023-04-01', '2023-04-01', '2023-05-01'],
    'Sales_Amount': [100.50, 200.75, 200.75, 1500.00, 300.25, np.nan, 400.00, 9999.99, 250.00, 275.50],
    'Region': ['North', 'SOUTH', 'south', 'West', 'East', 'North', np.nan, 'WEST', 'South', 'East']
}

df = pd.DataFrame(data)

# Save to Excel
df.to_excel('sales_data.xlsx', index=False)
print("Sales data Excel file generated: sales_data.xlsx")