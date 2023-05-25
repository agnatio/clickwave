## Abstract Form

This is a Python script that creates a graphical user interface (GUI) for processing data from an Excel file. The script utilizes the `tkinter` library for creating the GUI elements and interacts with the Windows operating system using the `win32gui` module. The processed data can be saved to a new Excel file.

### Prerequisites

The following libraries are required to run the script:

* `tkinter`
* `openpyxl`
* `win32gui`
* `os`
* `pyautogui`
* `datetime`

Install the necessary libraries using `pip`:

<pre><div class="bg-black rounded-md mb-4"><div class="flex items-center relative text-gray-200 bg-gray-800 px-4 py-2 text-xs font-sans justify-between rounded-t-md"><span>shell</span><button class="flex ml-auto gap-2"><svg stroke="currentColor" fill="none" stroke-width="2" viewBox="0 0 24 24" stroke-linecap="round" stroke-linejoin="round" class="h-4 w-4" height="1em" width="1em" xmlns="http://www.w3.org/2000/svg"><path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"></path><rect x="8" y="2" width="8" height="4" rx="1" ry="1"></rect></svg>Copy code</button></div><div class="p-4 overflow-y-auto"><code class="!whitespace-pre hljs language-shell">pip install tkinter openpyxl pywin32 pyautogui
</code></div></div></pre>

### Usage

1. Run the script using the command `python script_name.py` in the terminal.
2. The form window will appear on the screen with an "Open File" button.
3. Click on the "Open File" button to select an Excel file to process. Only files with the `.xlsx` extension are supported.
4. After selecting the file, the form will be populated with the data from the first row of the Excel file.
5. Use the navigation buttons (`<` and `>`) to move between rows and view the data.
6. Modify the data in the form as needed.
7. Click the "Submit" button to save the changes. If the "Keep log" checkbox is selected, a new Excel file with the suffix "_updated.xlsx" will be created, containing the updated data.
8. Click the "Exit" button to close the application. If the "Keep log" checkbox is selected, the script will attempt to save the updated data to a new Excel file before exiting. Ensure that the original Excel file is closed to avoid permission errors.

### Notes

* The script utilizes the `win32gui` module to interact with windows on the Windows operating system. It can retrieve the title of the currently active window and bring a window to the foreground based on its title.
* The form window is initially set to always be on top of other windows.
* The script supports processing Excel files with up to 20 columns. If the number of columns exceeds this limit, an error message will be displayed.
* The `pyautogui` library is used for screen capture and automation purposes.
* Additional functionality can be implemented by extending the `BaseForm` class and overriding the corresponding methods.
* The script provides a basic GUI for data processing and can be customized and extended to suit specific requirements.

Feel free to modify and adapt the script according to your needs.
