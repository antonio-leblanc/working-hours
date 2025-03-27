# Outlook Calendar Work Hours Analyzer

This Python script analyzes your Outlook calendar to calculate your working hours. It provides summaries by category, day, week, and month, and can export the results to an Excel file.

## Features

*   Connects to your Outlook calendar using `win32com.client`.
*   Filters appointments based on specified date ranges:
    *   This week (Monday to Sunday)
    *   This month
    *   Custom date range
    *   Specific week number (ISO week)
*   Calculates total working hours, excluding events categorized as "Pessoal" (configurable).
*   Provides summaries by category, day, week, and day of the week.
*   Calculates daily and working day average hours.
*   Exports results to an Excel file with multiple sheets.
*   Command-line arguments for controlling the script's behavior.

## Prerequisites

*   Python 3.6 or higher.
*   `pywin32` library:  `pip install pywin32`
*   `pandas` library: `pip install pandas`
*   For the week number calculation functions you need to install the `datetime` Library
*   Access to a Microsoft Outlook installation on Windows.

## Installation

1.  Clone or download this repository.
2.  Install the required Python libraries using `pip`:

    ```bash
    pip install pywin32 pandas
    ```

## Usage

Run the script from the command line:

```bash
python main.py [options]
```