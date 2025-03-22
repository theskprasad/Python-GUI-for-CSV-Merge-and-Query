import os
import pandas as pd
import gzip
import tkinter as tk
from tkinter import filedialog, messagebox
from threading import Thread
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.tooltip import ToolTip
from tkinter.ttk import Progressbar
from datetime import datetime

# Create main application window with Light Professional Theme
root = ttk.Window(themename="cosmo")
root.title("CSV Merge & Query Tool")
root.geometry("1200x800")
root.minsize(1000, 700)  # Set minimum window size

# Global variables
selected_files = []
merged_file_path = None  # Store the merged file path
merged_df = None  # Store the merged DataFrame
comparison_df = None  # Store the comparison DataFrame
comparison_result_df = None  # Store the result of the VLOOKUP-like comparison

# Functions
def select_files():
    """Opens a file dialog to select CSV (GZipped) files."""
    global selected_files
    files = filedialog.askopenfilenames(filetypes=[("GZipped CSV files", "*.gz")])
    if files:
        selected_files = list(files)
        file_listbox.delete(0, tk.END)
        for file in selected_files:
            file_listbox.insert(tk.END, os.path.basename(file))
        clear_button["state"] = "normal"
        update_status(f"Selected {len(selected_files)} file(s).")

def clear_selection():
    """Clears the selected files."""
    global selected_files
    selected_files = []
    file_listbox.delete(0, tk.END)  # Fixed indentation
    clear_button["state"] = "disabled"
    update_status("Selection cleared.")

def merge_files():
    """Merges selected CSV files and saves as a GZipped CSV with an automatic name."""
    global merged_file_path, merged_df
    if not selected_files:
        messagebox.showerror("Error", "No files selected!")
        return

    def process_files():
        global merged_file_path, merged_df
        try:
            df_list = []
            total_files = len(selected_files)
            progress_bar["maximum"] = total_files
            progress_bar["value"] = 0
            update_status("Processing files...")

            for i, file in enumerate(selected_files):
                with gzip.open(file, "rt", encoding="utf-8") as f:
                    df = pd.read_csv(f)
                    df_list.append(df)
                progress_bar["value"] = i + 1
                root.update_idletasks()

            merged_df = pd.concat(df_list, ignore_index=True)

            # Generate automatic file name based on current date and time
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            default_filename = f"merged_{timestamp}.gz"
            save_path = filedialog.asksaveasfilename(
                defaultextension=".gz",
                filetypes=[("GZipped CSV", "*.gz")],
                title="Save Merged File As",
                initialfile=default_filename  # Suggest the automatic filename
            )
            if save_path:
                with gzip.open(save_path, "wt", encoding="utf-8") as f:
                    merged_df.to_csv(f, index=False)

                merged_file_path = save_path
                update_status(f"Merged file saved: {save_path}")
                messagebox.showinfo("Success", f"Merged file saved:\n{save_path}")
                notebook.tab(1, state="normal")  # Enable the Comparison tab
                notebook.select(1)  # Switch to the Comparison tab
                setup_comparison_tab()  # Initialize the comparison tab

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            update_status("Error during processing.")
        finally:
            progress_bar["value"] = 0

    Thread(target=process_files, daemon=True).start()

def setup_comparison_tab():
    """Sets up the comparison tab after merging files."""
    global merged_df
    if merged_df is not None:
        # Clear existing Treeview
        for item in tree_comparison.get_children():
            tree_comparison.delete(item)

        # Set up columns in the Treeview
        columns_comparison = list(merged_df.columns)
        tree_comparison["columns"] = columns_comparison
        for col in columns_comparison:
            tree_comparison.heading(col, text=col)
            tree_comparison.column(col, width=150)

def compare_files():
    """Compares the merged data with another XLSX/CSV file using VLOOKUP-like functionality."""
    global comparison_df, comparison_result_df
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])
    if file_path:
        try:
            # Show progress bar
            progress_bar_comparison.start(10)
            update_status("Loading comparison file...")
            root.update_idletasks()

            # Load the comparison file
            if file_path.endswith('.xlsx'):
                comparison_df = pd.read_excel(file_path)
            else:
                comparison_df = pd.read_csv(file_path)

            # Ensure the key columns are of the same type (convert both to string)
            key_column = merged_df.columns[0]  # Assuming the first column is the key for comparison
            compare_key_column = comparison_df.columns[0]  # Assuming the first column is the key in the comparison file

            merged_df[key_column] = merged_df[key_column].astype(str)
            comparison_df[compare_key_column] = comparison_df[compare_key_column].astype(str)

            # Perform a merge (VLOOKUP-like operation)
            comparison_result_df = pd.merge(merged_df, comparison_df, how="left", left_on=key_column, right_on=compare_key_column)

            # Display the merged result in the Treeview
            for item in tree_comparison.get_children():
                tree_comparison.delete(item)

            for i, row in comparison_result_df.iterrows():
                tag = "evenrow" if i % 2 == 0 else "oddrow"
                tree_comparison.insert("", "end", values=row.tolist(), tags=(tag,))

            messagebox.showinfo("Success", "Comparison completed and results displayed.")
            notebook.tab(2, state="normal")  # Enable the Query Builder tab
            notebook.select(2)  # Switch to the Query Builder tab
            setup_query_tab()  # Initialize the query tab

        except Exception as e:
            messagebox.showerror("Error", f"Failed to compare files: {str(e)}")
        finally:
            progress_bar_comparison.stop()
            update_status("Comparison completed.")

def setup_query_tab():
    """Sets up the query tab after comparison."""
    global comparison_result_df
    if comparison_result_df is not None:
        # Clear existing Treeview
        for item in tree_query.get_children():
            tree_query.delete(item)

        # Set up columns in the Treeview
        columns_query = list(comparison_result_df.columns)
        tree_query["columns"] = columns_query
        for col in columns_query:
            tree_query.heading(col, text=col)
            tree_query.column(col, width=150)

        # Update the column dropdown
        column_dropdown["menu"].delete(0, "end")
        for col in columns_query:
            column_dropdown["menu"].add_command(label=col, command=tk._setit(column_var, col))
        column_var.set(columns_query[0])  # Set the default column

def update_status(message):
    """Updates the status bar with the given message."""
    status_var.set(message)
    root.update_idletasks()

# Main Application UI
notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True)

# Merge Tab
merge_tab = ttk.Frame(notebook)
notebook.add(merge_tab, text="Merge CSV Files")

file_select_frame = ttk.LabelFrame(merge_tab, text="Select Files", padding=10)
file_select_frame.pack(fill="x", pady=5)

select_button = ttk.Button(file_select_frame, text="Select Files", command=select_files, bootstyle="primary")
select_button.pack(side="left", padx=5)
ToolTip(select_button, "Select one or more GZipped CSV files to merge.")

clear_button = ttk.Button(file_select_frame, text="Clear Selection", command=clear_selection, bootstyle="secondary", state="disabled")
clear_button.pack(side="left", padx=5)
ToolTip(clear_button, "Clear the selected files.")

file_listbox = tk.Listbox(file_select_frame, width=70, height=5)
file_listbox.pack(fill="x", pady=5)

progress_frame = ttk.Frame(merge_tab)
progress_frame.pack(fill="x", pady=5)

progress_bar = Progressbar(progress_frame, mode="determinate")
progress_bar.pack(fill="x", pady=5)

merge_button = ttk.Button(merge_tab, text="Merge and Save", command=merge_files, bootstyle="success")
merge_button.pack(pady=10)

# Comparison Tab
comparison_tab = ttk.Frame(notebook)
notebook.add(comparison_tab, text="VLOOKUP-like Comparison", state="disabled")

ttk.Label(comparison_tab, text="VLOOKUP-like Comparison", font=("Segoe UI", 12)).pack(pady=10)

compare_button = ttk.Button(comparison_tab, text="Compare with File", command=compare_files, bootstyle="info")
compare_button.pack(pady=10)

progress_bar_comparison = Progressbar(comparison_tab, mode="indeterminate")
progress_bar_comparison.pack(fill="x", padx=10, pady=5)

table_frame_comparison = ttk.Frame(comparison_tab)
table_frame_comparison.pack(pady=5, fill="both", expand=True)

tree_comparison = ttk.Treeview(table_frame_comparison, show="headings")

style = ttk.Style()
style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), background="#0078D7", foreground="white")
style.configure("Treeview", font=("Segoe UI", 10), rowheight=25)
style.map("Treeview", background=[("selected", "#005499")])

tree_comparison.tag_configure("evenrow", background="#FFFFFF")  # White
tree_comparison.tag_configure("oddrow", background="#F0F0F0")  # Light gray

tree_scroll_y_comparison = ttk.Scrollbar(table_frame_comparison, orient="vertical", command=tree_comparison.yview)
tree_scroll_y_comparison.pack(side="right", fill="y")
tree_comparison.configure(yscrollcommand=tree_scroll_y_comparison.set)

tree_scroll_x_comparison = ttk.Scrollbar(table_frame_comparison, orient="horizontal", command=tree_comparison.xview)
tree_scroll_x_comparison.pack(side="bottom", fill="x")
tree_comparison.configure(xscrollcommand=tree_scroll_x_comparison.set)

tree_comparison.pack(fill="both", expand=True)

# Query Builder Tab
query_tab = ttk.Frame(notebook)
notebook.add(query_tab, text="Query Builder", state="disabled")

query_conditions = []  # Store multiple query conditions

def update_conditions_dropdown(*args):
    """Update condition options based on the selected column."""
    selected_column = column_var.get()
    
    if selected_column:
        # Determine if column contains numbers or text
        if pd.to_numeric(comparison_result_df[selected_column], errors="coerce").notna().all():
            conditions = ["==", "!=", ">", "<", ">=", "<=", "SUM", "AVG", "COUNT"]
        else:
            conditions = ["==", "!=", "contains", "not contains", "startswith", "endswith"]

        # Update the dropdown
        condition_dropdown["menu"].delete(0, "end")
        for condition in conditions:
            condition_dropdown["menu"].add_command(label=condition, command=tk._setit(condition_var, condition))
        
        # Set default value
        condition_var.set(conditions[0])

def add_condition():
    """Adds a query condition dynamically."""
    column = column_var.get()
    condition = condition_var.get()
    value = value_entry.get().strip()
    operator = operator_var.get()

    if not column or not condition or not value:
        messagebox.showerror("Error", "Please enter all query details!")
        return

    # Escape double quotes in the value
    value = value.replace('"', '\\"')

    # Enclose text values in double quotes
    if not value.replace('.', '', 1).isdigit():
        value = f'"{value}"'

    # Convert condition for string-based filters
    if condition == "contains":
        condition = f"{column}.str.contains({value})"
    elif condition == "not contains":
        condition = f"{column}.str.contains({value}) == False"
    elif condition == "startswith":
        condition = f"{column}.str.startswith({value})"
    elif condition == "endswith":
        condition = f"{column}.str.endswith({value})"
    else:
        condition = f"{column} {condition} {value}"  # Normal condition

    # Combine multiple conditions correctly
    if query_conditions:
        query_conditions.append(f"{operator} {condition}")
    else:
        query_conditions.append(condition)

    # Update the query display
    query_text.insert(tk.END, f"{query_conditions[-1]}\n")

def execute_query():
    """Executes multiple conditions on the comparison data."""
    global comparison_result_df
    if not query_conditions:
        messagebox.showerror("Error", "No queries added!")
        return

    # Convert `AND` → `&` and `OR` → `|` before execution
    query_string = " ".join(query_conditions).replace("AND", "&").replace("OR", "|")

    try:
        status_label.config(text="Processing Query...")
        progress_bar_query.start(10)
        root.update_idletasks()

        result_df = comparison_result_df.query(query_string, engine="python")  # Use "python" engine to handle str operations
        result_count.set(f"Total Matching Rows: {len(result_df)}")

        # Clear existing table
        for item in tree_query.get_children():
            tree_query.delete(item)

        # Insert new results with alternating row colors
        for i, row in result_df.iterrows():
            tag = "evenrow" if i % 2 == 0 else "oddrow"
            tree_query.insert("", "end", values=row.tolist(), tags=(tag,))

        if result_df.empty:
            messagebox.showwarning("Warning", "No matching data found for this query!")

        save_filtered_button["state"] = "normal" if not result_df.empty else "disabled"

    except Exception as e:
        messagebox.showerror("Query Error", f"Invalid query: {str(e)}")
    finally:
        progress_bar_query.stop()
        status_label.config(text="Query Execution Complete!")

def clear_query():
    """Clears the query conditions and resets the table."""
    query_conditions.clear()
    query_text.delete("1.0", tk.END)
    result_count.set("Total Matching Rows: 0")
    for item in tree_query.get_children():
        tree_query.delete(item)
    save_filtered_button["state"] = "disabled"

def save_results():
    """Saves the filtered results to a CSV file."""
    file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
    if file_path:
        query_string = " ".join(query_conditions).replace("AND", "&").replace("OR", "|")
        result_df = comparison_result_df.query(query_string, engine="python")
        result_df.to_csv(file_path, index=False)
        messagebox.showinfo("Success", f"Results saved to {file_path}")

# Query Builder Tab UI Elements
ttk.Label(query_tab, text="Advanced Query Builder:", font=("Segoe UI", 12)).pack(pady=5)

query_frame = ttk.Frame(query_tab)
query_frame.pack(pady=5)

column_var = tk.StringVar()
column_dropdown = ttk.OptionMenu(query_frame, column_var, "")
column_dropdown.pack(side="left", padx=5)
ToolTip(column_dropdown, "Select a column to filter on.")

condition_var = tk.StringVar()
condition_dropdown = ttk.OptionMenu(query_frame, condition_var, "==")  # Default
condition_dropdown.pack(side="left", padx=5)
ToolTip(condition_dropdown, "Select a condition for the filter.")

value_entry = ttk.Entry(query_frame, width=15)
value_entry.pack(side="left", padx=5)
ToolTip(value_entry, "Enter the value to filter by.")

operator_var = tk.StringVar(value="AND")
operator_dropdown = ttk.OptionMenu(query_frame, operator_var, "AND", "OR")
operator_dropdown.pack(side="left", padx=5)
ToolTip(operator_dropdown, "Combine multiple conditions with AND/OR.")

add_condition_button = ttk.Button(query_frame, text="Add Condition", command=add_condition, bootstyle="primary-outline")
add_condition_button.pack(side="left", padx=5)

clear_query_button = ttk.Button(query_frame, text="Clear Query", command=clear_query, bootstyle="danger-outline")
clear_query_button.pack(side="left", padx=5)

query_text = tk.Text(query_tab, width=80, height=4)
query_text.pack(pady=5)

query_button = ttk.Button(query_tab, text="Execute Query", command=execute_query, bootstyle="success")
query_button.pack(pady=5)

progress_bar_query = Progressbar(query_tab, mode="indeterminate")
progress_bar_query.pack(fill="x", padx=10, pady=5)

result_count = tk.StringVar()
result_count_label = ttk.Label(query_tab, textvariable=result_count, font=("Segoe UI", 10, "bold"))
result_count_label.pack(pady=5)

status_label = ttk.Label(query_tab, text="", font=("Segoe UI", 10, "italic"))
status_label.pack(pady=5)

table_frame_query = ttk.Frame(query_tab)
table_frame_query.pack(pady=5, fill="both", expand=True)

tree_query = ttk.Treeview(table_frame_query, show="headings")

style = ttk.Style()
style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), background="#0078D7", foreground="white")
style.configure("Treeview", font=("Segoe UI", 10), rowheight=25)
style.map("Treeview", background=[("selected", "#005499")])

tree_query.tag_configure("evenrow", background="#FFFFFF")  # White
tree_query.tag_configure("oddrow", background="#F0F0F0")  # Light gray

tree_scroll_y_query = ttk.Scrollbar(table_frame_query, orient="vertical", command=tree_query.yview)
tree_scroll_y_query.pack(side="right", fill="y")
tree_query.configure(yscrollcommand=tree_scroll_y_query.set)

tree_scroll_x_query = ttk.Scrollbar(table_frame_query, orient="horizontal", command=tree_query.xview)
tree_scroll_x_query.pack(side="bottom", fill="x")
tree_query.configure(xscrollcommand=tree_scroll_x_query.set)

tree_query.pack(fill="both", expand=True)

save_filtered_button = ttk.Button(query_tab, text="Save Results", bootstyle="secondary", state="disabled", command=save_results)
save_filtered_button.pack(pady=5)

# Status Bar
status_var = tk.StringVar()
status_var.set("Ready")
status_bar = ttk.Label(root, textvariable=status_var, relief="sunken", anchor="w")
status_bar.pack(side="bottom", fill="x", padx=10, pady=5)

# Start the main loop
root.mainloop()
