This project is a Python application with a graphical user interface (GUI) designed to process various data log files from DC EV chargers, including charger serial numbers, MAC addresses, and SIM card details. It consolidates this data to create structured, professional reports.

Key Features
Data Consolidation: Merges and processes information from multiple .csv files.
Data Extraction: Automatically extracts crucial details like charger serial numbers, firmware versions, MAC IDs, and SIM numbers from raw log data.
Report Generation: Generates comprehensive reports in .docx and .pdf formats.
GUI Interface: Provides a user-friendly interface for selecting files, initiating the data processing, and viewing the results.

Project File Structure
main.py: The core application file that handles the GUI and the user interaction flow.
process_data.py: Contains the main logic for reading, parsing, and processing the data from the various CSV files to generate a report.

paths.txt: A file used to store the local file path for the data.

The numerous 
.csv files are the raw data logs used by the application.

Dependencies
This project requires the following Python libraries. You can install them using pip:
pip install pandas
pip install docxtpl
pip install docx2pdf
pip install python-comtypes # for docx2pdf on Windows
pip install pyqt5
