from google.cloud import storage
import pandas as pd
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

# Function to export data to a CSV file and upload to GCS
def export_to_gcs(data, output_file, bucket_name, gcs_path):
    try:
        # Save the data to a CSV file
        output_file = output_file.replace(".xlsx", ".csv")
        data.to_csv(output_file, index=False)
        print(f"Data saved to CSV file: {output_file}")

        # Upload the CSV file to GCS
        client = storage.Client()
        bucket = client.bucket(bucket_name)
        blob = bucket.blob(gcs_path.replace(".xlsx", ".csv"))
        blob.upload_from_filename(output_file)
        print(f"CSV file uploaded to GCS: gs://{bucket_name}/{gcs_path.replace('.xlsx', '.csv')}")
    except Exception as e:
        print(f"Error exporting to GCS: {e}")

# Main function to load, process, visualize data, and upload results
def main():
    folder_path = r"C:\Users\dkundo\Desktop\visu\archive"  # Update this path
    output_folder = r"C:\Users\dkundo\Desktop\visu\output"  # Update this path
    output_file = os.path.join(output_folder, "combined_data.xlsx")

    # Input for GCS bucket and path
    bucket_name = input("Enter the Google Cloud Storage bucket name: ").strip()
    gcs_path = input("Enter the GCS path (e.g., output/combined_data.xlsx): ").strip()

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    data = load_files_from_folder(folder_path)
    if data is not None and not data.empty:
        export_to_gcs(data, output_file, bucket_name, gcs_path)
    else:
        print("No data available for export.")

if __name__ == "__main__":
    main()
