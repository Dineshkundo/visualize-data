import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
import glob

# Function to load data from CSV
def load_csv(file_path):
    try:
        data = pd.read_csv(file_path)
        print(f"CSV data loaded: {file_path}")
        return data
    except Exception as e:
        print(f"Error loading CSV: {file_path}. Error: {e}")
        return None

# Function to load data from Excel
def load_excel(file_path):
    try:
        data = pd.read_excel(file_path)
        print(f"Excel data loaded: {file_path}")
        return data
    except Exception as e:
        print(f"Error loading Excel: {file_path}. Error: {e}")
        return None

# Function to load data from JSON
def load_json(file_path):
    try:
        data = pd.read_json(file_path)
        print(f"JSON data loaded: {file_path}")
        return data
    except ValueError as e:
        try:
            # Attempt to load JSON as a dictionary
            data = pd.DataFrame(pd.read_json(file_path, typ='series')).reset_index()
            print(f"JSON loaded as series and converted: {file_path}")
            return data
        except Exception as inner_e:
            print(f"Error loading JSON: {file_path}. Error: {inner_e}")
            return None
    except Exception as e:
        print(f"Error loading JSON: {file_path}. Error: {e}")
        return None

# Function to load all files from a folder
def load_files_from_folder(folder_path):
    supported_files = glob.glob(os.path.join(folder_path, "*"))
    all_data = []
    for file_path in supported_files:
        if file_path.endswith(".csv"):
            data = load_csv(file_path)
        elif file_path.endswith(".xlsx") or file_path.endswith(".xls"):
            data = load_excel(file_path)
        elif file_path.endswith(".json"):
            data = load_json(file_path)
        else:
            print(f"Unsupported file format skipped: {file_path}")
            continue
        
        if data is not None:
            all_data.append(data)
    if all_data:
        combined_data = pd.concat(all_data, ignore_index=True)
        print("All data successfully loaded and combined.")
        return combined_data
    else:
        print("No valid files found in the folder.")
        return None

# Function to detect columns for grouping and aggregation
def detect_columns(data):
    numeric_columns = data.select_dtypes(include="number").columns
    categorical_columns = data.select_dtypes(include="object").columns

    if not numeric_columns.any() or not categorical_columns.any():
        print("No valid columns for aggregation or grouping.")
        return None, None

    # Prioritize meaningful columns based on their names
    group_column = next((col for col in categorical_columns if "name" in col.lower() or "brand" in col.lower()), categorical_columns[0])
    value_column = next((col for col in numeric_columns if "value" in col.lower() or "amount" in col.lower() or "sales" in col.lower()), numeric_columns[0])
    
    return group_column, value_column

# Function to create visualizations dynamically
def create_visualizations(data, output_folder, top_n=20):
    group_column, value_column = detect_columns(data)
    if not group_column or not value_column:
        print("Visualization skipped due to missing required columns.")
        return

    print(f"Using '{group_column}' for grouping and '{value_column}' for aggregation.")

    # Aggregate data
    aggregated_data = data.groupby(group_column)[value_column].sum().reset_index()
    aggregated_data = aggregated_data.nlargest(top_n, value_column)

    # Create bar chart
    sns.set(style="whitegrid")
    plt.figure(figsize=(12, 6))
    sns.barplot(data=aggregated_data, x=group_column, y=value_column)
    plt.title(f"Top {top_n} {value_column} by {group_column}")
    plt.xticks(rotation=45)
    bar_chart_path = os.path.join(output_folder, "bar_chart.png")
    plt.savefig(bar_chart_path)
    print(f"Bar chart saved at {bar_chart_path}.")
    plt.close()

    # Create pie chart
    plt.figure(figsize=(8, 8))
    plt.pie(
        aggregated_data[value_column],
        labels=aggregated_data[group_column],
        autopct='%1.1f%%',
        startangle=140,
        colors=sns.color_palette("pastel"),
    )
    plt.title(f"Distribution of {value_column} by {group_column}")
    pie_chart_path = os.path.join(output_folder, "pie_chart.png")
    plt.savefig(pie_chart_path)
    print(f"Pie chart saved at {pie_chart_path}.")
    plt.close()

# Main function to load, process, and visualize data
def main():
    folder_path = r"C:\Users\dkundo\Desktop\visu\archive"  # Update this path
    output_folder = r"C:\Users\dkundo\Desktop\visu\output"  # Update this path

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    data = load_files_from_folder(folder_path)
    if data is not None and not data.empty:
        create_visualizations(data, output_folder)
    else:
        print("No data available for visualization.")

if __name__ == "__main__":
    main()
