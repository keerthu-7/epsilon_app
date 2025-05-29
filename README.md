🔋 Cycle Analysis Web App
=========================

This is a web-based tool to extract voltage-capacity graphs and combine multiple battery cycle Excel files into one comparison Excel file with charts.

Built with:
- Python
- Flask
- openpyxl

--------------------
🚀 Features
--------------------

✅ Extract Cycle Graphs
- Upload a raw Excel file generated from battery testing
- Specify material weight and cycle range (e.g., 1 to 10)
- Download a new Excel file with voltage vs capacity charts (charging + discharging)

✅ Combine Excel Files
- Upload 2 or more Excel files generated using the extract feature
- Select common cycles available in all uploaded files
- Download a combined Excel file with a comparison chart

--------------------
🛠️ Setup Instructions
--------------------

1. Install Python dependencies (Python 3.8+ recommended):

    pip install flask openpyxl

2. Run the app:

    python cycle_web_app.py

3. Open your browser and go to:

    http://localhost:5000

--------------------
📂 Folder Structure
--------------------


| File / Folder            | Description                                        |
|--------------------------|----------------------------------------------------|
| cycle_web_app.py         | Flask app entry point                              |
| combine_excel_logic.py   | Logic for combining multiple Excel files           |
| cycle_graph_logic.py     | Logic for extracting voltage-capacity from a file  |
| templates/index.html     | Frontend web interface (HTML form)                 |
| static/header.png        | Optional header image (UI banner)                  |
| README.txt               | This file                                          |

--------------------
📌 Notes
--------------------

- Only Excel files exported from this app (via Extract Graph tab) should be used in Combine mode.
- The app automatically detects available cycles written in row 1 (e.g., "Cycle 18", "Cycle 19").
- The final chart will show both Charging and Discharging curves for each selected cycle/file.
- You can compare as many files and cycles as you want — as long as they have common cycle numbers.

--------------------
🧪 Example Use Case
--------------------

1. Run the app.
2. In the "Extract Graph" tab:
   - Upload a raw battery test Excel file.
   - Enter weight and cycle range (e.g., 5 to 10).
   - Download the generated graph Excel.
3. Repeat step 2 for other test files.
4. In the "Combine Excels" tab:
   - Upload all the extracted Excel files.
   - Select 1 or more common cycles shown in the dropdown.
   - Download the final Excel with a combined chart.




Developed with ❤️ using Flask and openpyxl.
