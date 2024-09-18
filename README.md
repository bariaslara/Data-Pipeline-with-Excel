# Data-Pipeline-with-Excel
An automated data pipeline for processing and filtering product data from CSV files, integrated with Excel for easy execution and results logging.

## About the Project

This project is a **data pipeline** built to process and analyze product data using Python. The pipeline automates the following tasks:

- **Data Extraction**: Reads product-related CSV files from a designated directory.
- **Data Validation**: Ensures that product transaction dates are processed correctly and only includes transactions on or before a given date.
- **Results Logging**: Appends the filtered product data (transaction ID and date) to a results file in CSV format.
- **Automated Execution**: The script integrates with an Excel sheet through `xlwings`, allowing users to execute it via a button in the spreadsheet.
  
### Key Features:
- **Date Filtering**: The script reads CSV data and filters out product transactions based on a given cutoff date, ensuring only relevant records are processed.
- **Automated Excel Integration**: The project is designed to be run from an Excel workbook, making it easy for users to execute the script and view results in a familiar interface.
- **Customizable for Different Use Cases**: The script is adaptable, allowing for easy modifications to fit various business scenarios and data structures.

### Technologies Used:
- **Python**: The core language for scripting and data manipulation.
- **Pandas**: For handling data structures and date conversions.
- **CSV**: For input/output data in CSV format.
- **xlwings**: To integrate Python scripts with Excel for user-friendly execution.
- **Datetime**: For working with date comparisons and timestamps.

---

This project is a simple yet powerful tool for managing product data pipelines efficiently. By leveraging Python’s data processing libraries and integrating with Excel, it enables seamless automation for daily workflows.

Feel free to fork, contribute, or suggest improvements!
