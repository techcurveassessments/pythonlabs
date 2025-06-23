import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import os

# Set up plotting style
plt.style.use('seaborn')
sns.set_palette("deep")

# Read the Excel file
def read_excel_file(file_path='input_data.xlsx'):
    try:
        df = pd.read_excel(file_path)
        print(f"Successfully read {file_path}")
        return df
    except Exception as e:
        print(f"Error reading file: {e}")
        return None

# Analyze data and recommend chart types
def analyze_data(df):
    analysis = {
        'numeric_cols': [],
        'categorical_cols': [],
        'date_cols': [],
        'recommendations': []
    }
    
    # Identify column types
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            analysis['numeric_cols'].append(col)
        elif pd.api.types.is_datetime64_any_dtype(df[col]):
            analysis['date_cols'].append(col)
        else:
            analysis['categorical_cols'].append(col)
    
    # Recommend visualizations based on column types
    # 1. Numeric columns: Histogram, Boxplot
    for col in analysis['numeric_cols']:
        analysis['recommendations'].append({
            'column': col,
            'chart_types': ['Histogram', 'Boxplot'],
            'reason': f"Visualize distribution and outliers of numeric {col}"
        })
    
    # 2. Categorical columns: Bar plot
    for col in analysis['categorical_cols']:
        if df[col].nunique() <= 20:  # Limit to avoid cluttered plots
            analysis['recommendations'].append({
                'column': col,
                'chart_types': ['Bar'],
                'reason': f"Show frequency of categories in {col}"
            })
    
    # 3. Date + Numeric: Line plot
    if analysis['date_cols'] and analysis['numeric_cols']:
        for date_col in analysis['date_cols']:
            for num_col in analysis['numeric_cols']:
                analysis['recommendations'].append({
                    'column': f"{num_col} over {date_col}",
                    'chart_types': ['Line'],
                    'reason': f"Track {num_col} trends over time"
                })
    
    # 4. Categorical + Numeric: Boxplot or Bar
    if analysis['categorical_cols'] and analysis['numeric_cols']:
        for cat_col in analysis['categorical_cols']:
            if df[cat_col].nunique() <= 20:
                for num_col in analysis['numeric_cols']:
                    analysis['recommendations'].append({
                        'column': f"{num_col} by {cat_col}",
                        'chart_types': ['Box', 'Bar'],
                        'reason': f"Compare {num_col} across {cat_col} categories"
                    })
    
    return analysis

# Generate visualizations
def generate_visualizations(df, analysis, output_dir='visualizations'):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    
    for rec in analysis['recommendations']:
        col = rec['column']
        chart_types = rec['chart_types']
        
        for chart_type in chart_types:
            plt.figure(figsize=(10, 6))
            
            if chart_type == 'Histogram' and col in df.columns:
                sns.histplot(df[col].dropna(), bins=20)
                plt.title(f'Histogram of {col}')
                plt.xlabel(col)
                plt.ylabel('Frequency')
            
            elif chart_type == 'Boxplot' and col in df.columns:
                sns.boxplot(y=df[col].dropna())
                plt.title(f'Boxplot of {col}')
                plt.ylabel(col)
            
            elif chart_type == 'Bar' and col in df.columns:
                value_counts = df[col].value_counts().head(10)  # Top 10 categories
                sns.barplot(x=value_counts.values, y=value_counts.index)
                plt.title(f'Bar Plot of {col}')
                plt.xlabel('Count')
                plt.ylabel(col)
            
            elif chart_type == 'Line' and ' over ' in col:
                num_col, date_col = col.split(' over ')
                df_sorted = df.sort_values(date_col)
                plt.plot(df_sorted[date_col], df_sorted[num_col])
                plt.title(f'{num_col} over {date_col}')
                plt.xlabel(date_col)
                plt.ylabel(num_col)
                plt.xticks(rotation=45)
            
            elif chart_type == 'Box' and ' by ' in col:
                num_col, cat_col = col.split(' by ')
                sns.boxplot(x=cat_col, y=num_col, data=df)
                plt.title(f'{num_col} by {cat_col}')
                plt.xticks(rotation=45)
            
            elif chart_type == 'Bar' and ' by ' in col:
                num_col, cat_col = col.split(' by ')
                means = df.groupby(cat_col)[num_col].mean().head(10)
                sns.barplot(x=means.index, y=means.values)
                plt.title(f'Mean {num_col} by {cat_col}')
                plt.xlabel(cat_col)
                plt.ylabel(f'Mean {num_col}')
                plt.xticks(rotation=45)
            
            plt.tight_layout()
            output_file = os.path.join(output_dir, f"{chart_type}_{col.replace(' ', '_')}_{timestamp}.png")
            plt.savefig(output_file)
            plt.close()
            print(f"Saved {chart_type} for {col} to {output_file}")

# Main function
def main():
    df = read_excel_file()
    if df is None:
        return
    
    print("\nData Summary:")
    print(df.info())
    print("\nFirst few rows:")
    print(df.head())
    
    analysis = analyze_data(df)
    print("\nVisualization Recommendations:")
    for rec in analysis['recommendations']:
        print(f"- {rec['column']}: {', '.join(rec['chart_types'])} ({rec['reason']})")
    
    generate_visualizations(df, analysis)
    print("\nVisualizations generated in 'visualizations' folder.")

if __name__ == "__main__":
    main()