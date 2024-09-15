import os
import time
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# Unique Customers Function with Insights
def process_branch_data(file_first_month, file_second_month):
    # Load the Excel files
    df_first_month = pd.read_excel(file_first_month)
    df_second_month = pd.read_excel(file_second_month)

    # Create a new column 'رقم العميل' by merging 'العميل/رقم الهاتف' and 'العميل/الهاتف المحمول'
    df_first_month['رقم العميل'] = df_first_month['العميل/رقم الهاتف'].combine_first(df_first_month['العميل/الهاتف المحمول'])
    df_second_month['رقم العميل'] = df_second_month['العميل/رقم الهاتف'].combine_first(df_second_month['العميل/الهاتف المحمول'])

    # Combine data for both months
    combined_data = pd.concat([df_first_month, df_second_month], ignore_index=True)

    # Group by 'العميل' and 'رقم العميل' to get visit counts
    visit_counts = combined_data.groupby(['العميل', 'رقم العميل']).size().reset_index(name='عدد مرات الزيارة')

    # Find clients who visited in both months
    first_month_customers = set(df_first_month['رقم العميل'].dropna())
    second_month_customers = set(df_second_month['رقم العميل'].dropna())
    repeated_customers = list(first_month_customers & second_month_customers)

    # Add a column to indicate whether the customer visited in both months or only in the first
    visit_counts['زيارة في الشهرين ام الشهر الاول فقط'] = visit_counts['رقم العميل'].apply(
        lambda x: 'زيارة في الشهرين' if x in repeated_customers else 'زيارة في الشهر الاول فقط'
    )

    # Calculate distinct clients for each month
    distinct_clients_first_month = len(first_month_customers)
    distinct_clients_second_month = len(second_month_customers)

    # Calculate the percentage of repeated visits
    repeated_percentage = (len(repeated_customers) / distinct_clients_first_month) * 100 if distinct_clients_first_month > 0 else 0

    # Save the result to an Excel file in the Downloads folder with a unique name
    output_path = os.path.expanduser(f"~/Downloads/branch_visits_summary_{int(time.time())}.xlsx")
    visit_counts.to_excel(output_path, index=False)

    # Return insights and file path
    insights = {
        'distinct_clients_first_month': distinct_clients_first_month,
        'distinct_clients_second_month': distinct_clients_second_month,
        'repeated_customers_count': len(repeated_customers),
        'repeated_percentage': repeated_percentage,  # This was missing before
        'result_file': output_path
    }
    
    return insights

# Growth Rate Function
def calculate_growth_rate():
    first_file_path = filedialog.askopenfilename(title="Select the First Month File", filetypes=[("Excel files", "*.xlsx")])
    second_file_path = filedialog.askopenfilename(title="Select the Second Month File", filetypes=[("Excel files", "*.xlsx")])

    first_month = pd.read_excel(first_file_path)
    second_month = pd.read_excel(second_file_path)

    first_month_clean = first_month.iloc[2:].reset_index(drop=True)
    second_month_clean = second_month.iloc[2:].reset_index(drop=True)
    
    first_month_clean.columns = ["Product", "Orders", "Quantity", "Total_Price"]
    second_month_clean.columns = ["Product", "Orders", "Quantity", "Total_Price"]

    first_month_clean["Orders"] = pd.to_numeric(first_month_clean["Orders"], errors='coerce')
    second_month_clean["Orders"] = pd.to_numeric(second_month_clean["Orders"], errors='coerce')

    merged_data = pd.merge(first_month_clean, second_month_clean, on="Product", suffixes=('_First', '_Second'))
    
    merged_data["Orders_Growth_Rate"] = ((merged_data["Orders_Second"] - merged_data["Orders_First"]) / merged_data["Orders_First"]) * 100
    
    output_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    merged_data.to_excel(output_file_path, index=False)

# Target Prediction Function
def calculate_next_month_target(file_path):
    excel_data = pd.read_excel(file_path)
    excel_data.columns = excel_data.columns.str.strip()
    cleaned_data = excel_data.dropna()

    sales_columns = cleaned_data.columns[-4:-1]
    growth_rate_column = cleaned_data.columns[-1]

    cleaned_data['Average Sales Last 3 Months'] = cleaned_data[sales_columns].mean(axis=1)
    cleaned_data['Next Month Target'] = cleaned_data['Average Sales Last 3 Months'] * (1 + cleaned_data[growth_rate_column])

    output_file = file_path.replace('.xlsx', '_Predicted_Targets.xlsx')
    cleaned_data[['Average Sales Last 3 Months', growth_rate_column, 'Next Month Target']].to_excel(output_file, index=False)

# Main Program
def main():
    print("Choose an option:")
    print("1. Calculate Unique Customers and Repeated Visits with Insights")
    print("2. Calculate Growth Rate")
    print("3. Predict Next Month's Target")

    choice = input("Enter the number of your choice: ")

    if choice == '1':
        print("Select the files for the first and second months")
        file_first_month = filedialog.askopenfilename(title="Select First Month File", filetypes=[("Excel files", "*.xlsx")])
        file_second_month = filedialog.askopenfilename(title="Select Second Month File", filetypes=[("Excel files", "*.xlsx")])
        
        insights = process_branch_data(file_first_month, file_second_month)
        print(f"Distinct clients in the first month: {insights['distinct_clients_first_month']}")
        print(f"Distinct clients in the second month: {insights['distinct_clients_second_month']}")
        print(f"Number of repeated customers: {insights['repeated_customers_count']}")
        print(f"Percentage of repeated visits: {insights['repeated_percentage']:.2f}%")
        print(f"Results saved to: {insights['result_file']}")
    
    elif choice == '2':
        print("Calculating growth rate...")
        calculate_growth_rate()
    
    elif choice == '3':
        print("Select the file to predict next month's target")
        file_path = filedialog.askopenfilename(title="Select File", filetypes=[("Excel files", "*.xlsx")])
        calculate_next_month_target(file_path)
        print(f"Prediction saved.")

# Running the main program with a simple interface
if __name__ == '__main__':
    root = tk.Tk()
    root.withdraw()  # Hide the Tkinter main window
    main()
