# ergospirography
**Ergospyrography**

Simple File Manager for Excel Files

**Project Description**

This project is a simple file manager developed in Python using the Tkinter library to create a graphical user interface. The main functionality of the program includes downloading data from Excel files, performing data operations and saving the results back to an Excel file. The program accepts an Excel file exported from the SpiroLab M software for ergospyrogaf. It determines the zones before testing, the anaerobic threshold, the maximum oxygen consumption and the end of the workout. It calculates the segments 30 seconds before these zones and calculates the arithmetic mean for them, and then saves this data to a new Excel file.
**Requirements**

To use the program, the following dependencies need to be installed:

- Python (version 3.x)
- Tkinter library (usually included with Python installation)
- pandas library (`pip install pandas`)
- numpy library (`pip install numpy`)
- openpyxl library (`pip install openpyxl`)

**Installation and Usage**

1. Clone the repository using the command:
   ```
   git clone https://github.com/your-username/file-manager.git
   ```

2. Navigate to the project directory:
   ```
   cd file-manager
   ```

3. Install the required dependencies:
   ```
   pip install -r requirements.txt
   ```

4. Run the program:
   ```
   python file_manager.py
   ```

**Usage**

1. Upon launching the program, the file manager window will open.
2. Click the "Choose File" button and select the Excel file you want to work with.
3. The program will load the data from the first sheet of the Excel file and display it in the window.
4. Perform the desired operations on the data using the available functions and buttons.
5. After completing the operations, you can save the results back to the Excel file by clicking the "Save File" button.
6. Exit the program by clicking the "Exit" button.

**Contributing**

If you have suggestions for improving the project or have found any issues, please create a new pull request or submit an issue on GitHub.

**License**

This project is licensed under the [MIT License](https://opensource.org/licenses/MIT).
