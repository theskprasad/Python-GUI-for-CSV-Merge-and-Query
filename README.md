# Python-GUI-for-CSV-Merge-and-Query
Python GUI for CSV Merge and Query
The provided code is a Python script for a GUI application that allows users to merge multiple GZipped CSV files, perform VLOOKUP-like comparisons with other CSV or Excel files, and build and execute queries on the merged data. The application is built using the tkinter library with the ttkbootstrap theme for a modern look.

Here’s a breakdown of the key components and functionality:

1. Main Application Window
The application window is created using ttk.Window with the "cosmo" theme.

The window is titled "CSV Merge & Query Tool" and has a minimum size of 1000x700 pixels.

2. Global Variables
selected_files: Stores the list of selected GZipped CSV files.

merged_file_path: Stores the path of the merged file.

merged_df: Stores the merged DataFrame.

comparison_df: Stores the DataFrame for comparison.

comparison_result_df: Stores the result of the VLOOKUP-like comparison.

3. Functions
select_files(): Opens a file dialog to select GZipped CSV files and updates the listbox with the selected files.

clear_selection(): Clears the selected files and resets the listbox.

merge_files(): Merges the selected CSV files into a single DataFrame and saves it as a GZipped CSV file. The merged file is saved with an automatic name based on the current date and time.

setup_comparison_tab(): Sets up the comparison tab by populating the Treeview with the merged DataFrame columns.

compare_files(): Compares the merged DataFrame with another CSV or Excel file using a VLOOKUP-like operation. The result is displayed in the Treeview.

setup_query_tab(): Sets up the query tab by populating the Treeview with the comparison result DataFrame columns.

update_status(): Updates the status bar with a given message.

4. UI Components
Merge Tab: Allows users to select GZipped CSV files, merge them, and save the merged file.

Comparison Tab: Allows users to compare the merged data with another CSV or Excel file using a VLOOKUP-like operation.

Query Builder Tab: Allows users to build and execute queries on the comparison result DataFrame. Users can add conditions, execute queries, and save the filtered results.

5. Query Builder
Users can select a column, choose a condition, and enter a value to build a query.

Multiple conditions can be combined using AND or OR.

The query is executed on the comparison result DataFrame, and the results are displayed in the Treeview.

Users can save the filtered results to a CSV file.

6. Progress Bars
Progress bars are used to indicate the progress of file merging and query execution.

7. Status Bar
A status bar at the bottom of the window displays messages about the current operation.

8. Threading
File merging and query execution are performed in separate threads to keep the UI responsive.

9. Error Handling
The code includes error handling to display error messages in case of exceptions.

10. Styling
The Treeview and other UI elements are styled using ttkbootstrap for a modern and professional look.

Potential Improvements:
Error Handling: The error handling could be more specific to provide more detailed feedback to the user.

Progress Bar: The progress bar could be more granular, especially for large files.

Column Selection: The VLOOKUP-like comparison assumes the first column as the key. This could be made configurable by the user.

Query Builder: The query builder could be enhanced with more complex conditions and support for more data types.

Performance: For very large datasets, the application might become slow. Consider optimizing the DataFrame operations or adding support for chunked processing.

Example Usage:
Merge CSV Files:

Select multiple GZipped CSV files using the "Select Files" button.

Click "Merge and Save" to merge the files and save the result.

VLOOKUP-like Comparison:

After merging, switch to the "Comparison" tab.

Click "Compare with File" and select a CSV or Excel file to compare with the merged data.

Query Builder:

After comparison, switch to the "Query Builder" tab.

Build a query by selecting a column, condition, and value.

Execute the query and view the results in the Treeview.

Save the filtered results to a CSV file.
