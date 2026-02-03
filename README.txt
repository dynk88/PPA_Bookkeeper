Here is a comprehensive guide you can share with new users or keep for your reference.
📘 Department Bookkeeper - User Manual
1. Getting Started

This application works as a "front-end" for an Excel file. It allows you to record financial transactions and allocations without opening Excel, but it stores everything in a file named data.xlsx.
The Data File

    Location: The app always looks for data.xlsx in the same folder where the .exe is running.

    First Run: If data.xlsx does not exist, the app will create a brand new one automatically.

    Portability: To move your data to another computer, simply copy the .exe AND the data.xlsx file together.

2. Setting Up Departments & Opening Balance

When you first start (or when the app creates a new file), the system is empty. You need to define your departments and their starting balance.

How to set it up:

    Close the App.

    Open data.xlsx in Microsoft Excel.

    Go to the sheet named Limits.

    Column A (Department): Type the names of your departments (e.g., "PWD EZ", "RWD WZ").

    Column B (Previous_balance): Enter the Opening Balance for the financial year.

        Note: This amount acts as the "Bank Balance" carried forward from previous years.

    Save and Close Excel.

    Open the App. Your departments will now appear in the dropdown list.

3. Allocations (Adding Money)

When the government or parent department releases funds to a specific department during the year.

In the App:

    Toggle the switch at the top left to "Allocation Mode: ON".

    The form changes to purple.

    Select the Department and enter the Amount and Date.

    Click "Add Allocation".

    Click "Validate & Save".

What happens in Excel (Limits Sheet):

    The app finds the department in Row X.

    It adds the new allocation amount to Column B (Previous_balance) to keep a running total of the "Grand Total Limit".

    It also records the specific allocation details in new columns to the right (e.g., 1st Allocation / Date_1, 2nd Allocation / Date_2).

    Why? This allows the PDF report to show exactly when money was received (Q1, Q2, Q3, or Q4).

4. Transactions (Spending Money)

When you issue a PPA (Print Payment Advice) or spend money.

In the App:

    Ensure "Allocation Mode" is OFF (Grey toggle).

    Select Department.

    Enter PPA Number (13 alphanumeric characters).

    Enter Amount and Date.

    Click "Submit PPA" -> "Validate & Save".

What happens in Excel (Transactions_YYYY_YY Sheet):

    The app creates a sheet named based on the Financial Year (e.g., Transactions_2025_26).

    Every Department gets 3 columns in this sheet:

        PPA_Number

        Date

        Amount

    The app finds the next empty row for that department and saves the PPA details there.

5. Correcting Mistakes (Manually Modifying Data)

Since the database is just an Excel file, you can fix mistakes easily.

Scenario: You entered a PPA amount as 500 instead of 5000.

    Close the App.

    Open data.xlsx.

    Go to the Transactions_2025_26 sheet.

    Find the PPA number or Department column.

    Manually change the cell value from 500 to 5000.

    Save and Close Excel.

    The App will now reflect the corrected amount in Dashboards and PDFs.

⚠️ CAUTION:

    Do not change Column Headers (Row 1 and 2). The app relies on these names to find data.

    Do not delete Department names from the Limits sheet unless you want them gone forever.

6. Financial Year Transition (April 2026)

The app is designed to handle the accounting year rollover automatically.

What happens on April 1st, 2026?

    When you enter a transaction with a date like 15-04-2026:

        The app automatically detects the new Financial Year.

        It creates a new sheet in Excel named Transactions_2026_27.

        It creates new columns for your departments in this new sheet.

    The PDF Report:

        It will calculate your Opening Balance by taking your Total Allocations (from history) minus your Total Expenditures (from the old Transactions_2025_26 sheet).

        New transactions will appear in the "Quarter 1" column of the PDF.

Do I need to do anything manually?
No. Just ensure that when you enter data for the new year, you select the correct date in the calendar.
7. Generating Reports

    Dashboard View: Shows a quick summary of Total Limit vs Total Spent vs Balance.

    PDF Export: Creates a detailed "Running Balance" report. It separates "Addl Allocations" and "Expenditure" into Quarters (Q1: Apr-Jun, Q2: Jul-Sep, etc.) so you can track cash flow throughout the year.