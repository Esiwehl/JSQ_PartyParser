# Jumpsquare Party Parser

The Jumpsquare Party Parser is a Python script designed to transform booking information from a CSV file into a more readable and structured Excel format, which makes managing and organizing parties easier. This transformation includes extracting and formatting details such as party type, duration, special notes (e.g., slush, cake, diploma, candy bags), and calculating arrival times.

## Features

- Parse and transform booking CSV files for Jumpsquare parties.
- Extract party type and duration from booking items.
- Identify specific notes based on keywords.
- Calculate adjusted arrival times.
- Output the transformed data into a formatted Excel file.

## Installation

Before running the script, ensure you have Python installed on your system. This script requires Python 3.11.6

Next, you'll need to install the required Python libraries. You can install these dependencies by running:

```bash
pip install pandas openpyxl
```

## Usage

To use the script, you should have a CSV file containing the booking information. The script takes this file as input and generates an Excel file with the formatted and transformed data.

Run the script with the following command:

```bash
python party_parser.py <path_to_csv_file>
```