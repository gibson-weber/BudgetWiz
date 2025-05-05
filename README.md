# BudgetWiz

## Overview
**BudgetWiz** is a Python-based application designed to help users manage their monthly spending by categorizing transactions from CSV files and generating detailed Excel reports with visualizations. The application streamlines the budgeting process, allowing users to easily analyze their expenses and track their financial health.

## Features
- Load transaction data from CSV files.
- Categorize transactions automatically or manually.
- Generate an Excel report with categorized expenses and pivot tables.
- Create pie charts for visual representation of spending by category.
- Autofit Excel columns for better readability.
- Clean up and maintain a categories CSV file for future use.

## Technologies Used
- **Python**: The primary programming language.
- **Pandas**: For data manipulation and analysis.
- **Openpyxl**: For working with Excel files.
- **Pywin32**: To interact with Excel via Windows COM interface.
- **OS & Subprocess**: For handling file paths and running system commands.

## Installation

### Prerequisites
Make sure you have Python installed on your machine. You can download it from [python.org](https://www.python.org/).

### Clone the Repository
```bash
git clone https://github.com/yourusername/BudgetWiz.git
cd BudgetWiz
```

### Install Dependencies
Run the following command to install the required libraries:
```bash
pip install -r requirements.txt
```

## Usage
1. Place your transaction CSV file in the `Data` folder.
2. Run the main script:
```bash
python BudgetWiz.py
```
3. Follow the prompts to input the CSV file name and the sheet name for the Excel report.

## File Structure
```
BudgetWiz/
│
├── Data/
│   └── (Your CSV files go here)
│
├── Categories.csv
├── requirements.txt
├── BudgetWiz.py
└── CategoriesCleanup.py
```

## Contributing
Contributions are welcome! If you have suggestions for improvements or features, please feel free to fork the repository and submit a pull request.

## License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgements
- Thank you to the developers of the libraries used in this project for their excellent work.
