
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import os
from datetime import datetime
import threading

# Banner
BANNER = "Bidhan Software's: A step towards modern age"

def generate_summary(df):
    return df.describe(include='all')

def generate_missing_data_report(df):
    return df.isnull().sum()

def generate_correlation_matrix(df):
    return df.corr(numeric_only=True)

def generate_histograms(df, output_dir, progress_callback):
    numeric_cols = df.select_dtypes(include='number').columns
    total = len(numeric_cols)
    for i, col in enumerate(numeric_cols):
        plt.figure()
        df[col].hist(bins=20)
        plt.title(f'Histogram of {col}')
        plt.xlabel(col)
        plt.ylabel('Frequency')
        plt.savefig(os.path.join(output_dir, f'hist_{col}.png'))
        plt.close()
        progress_callback(40 + int(20 * (i + 1) / total))  # Update progress

def generate_boxplots(df, output_dir, progress_callback):
    numeric_cols = df.select_dtypes(include='number').columns
    total = len(numeric_cols)
    for i, col in enumerate(numeric_cols):
        plt.figure()
        sns.boxplot(x=df[col])
        plt.title(f'Boxplot of {col}')
        plt.savefig(os.path.join(output_dir, f'box_{col}.png'))
        plt.close()
        progress_callback(60 + int(20 * (i + 1) / total))  # Update progress

def generate_heatmap(df, output_dir):
    corr = df.corr(numeric_only=True)
    plt.figure(figsize=(10, 8))
    sns.heatmap(corr, annot=True, cmap='coolwarm')
    plt.title('Correlation Heatmap')
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'correlation_heatmap.png'))
    plt.close()

def run_eda(file_path, progress_callback):
    df = pd.read_excel(file_path, engine='openpyxl') if file_path.endswith('.xlsx') else pd.read_csv(file_path)
    base_dir = os.path.dirname(file_path)
    output_dir = os.path.join(base_dir, 'Data_Analysis_Results')
    os.makedirs(output_dir, exist_ok=True)

    progress_callback(10)
    summary = generate_summary(df)
    summary.to_csv(os.path.join(output_dir, 'summary_statistics.csv'))

    progress_callback(20)
    missing = generate_missing_data_report(df)
    missing.to_csv(os.path.join(output_dir, 'missing_data_report.csv'))

    progress_callback(30)
    corr = generate_correlation_matrix(df)
    corr.to_csv(os.path.join(output_dir, 'correlation_matrix.csv'))

    generate_histograms(df, output_dir, progress_callback)
    generate_boxplots(df, output_dir, progress_callback)

    progress_callback(90)
    generate_heatmap(df, output_dir)

    with open(os.path.join(output_dir, 'eda_log.txt'), 'w') as log:
        log.write(f"{BANNER}\n")
        log.write(f"Data Analysis completed on {datetime.now()}\n")
        log.write(f"Input file: {file_path}\n")
        log.write(f"Output directory: {output_dir}\n")

    progress_callback(100)
    return output_dir

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel and CSV files", "*.xlsx *.xls *.csv")])
    if file_path:
        thread = threading.Thread(target=run_analysis_thread, args=(file_path,))
        thread.start()

def run_analysis_thread(file_path):
    try:
        def update_progress(value):
            progress["value"] = value
            root.update_idletasks()

        output_dir = run_eda(file_path, update_progress)
        messagebox.showinfo("Success", f"Data Analysis completed.\nResults saved in:\n{output_dir}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
    finally:
        progress["value"] = 0

root = tk.Tk()
root.title("Bidhan Software's: A step towards modern age")
root.geometry("400x250")

label = tk.Label(root, text=BANNER, font=("Arial", 12, "bold"))
label.pack(pady=20)

button = tk.Button(root, text="Select Excel or CSV File for Data Analysis", command=select_file, font=("Arial", 10))
button.pack(pady=10)

progress = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
progress.pack(pady=10)

root.mainloop()
