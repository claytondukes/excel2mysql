# Excel to SQL Converter

## Overview

I am by no means a Python Guru, but the code did work for me in converting a 50k line excel file to MySQL.

The script converts data from an Excel file into MySQL-compatible `INSERT` statements. Optionally, it can also generate a `CREATE TABLE` statement based on the data's structure.

The script supports renaming columns to ensure they comply with SQL naming conventions (no spaces or special characters). It also handles escaping problematic characters, such as single quotes and backslashes, to prevent SQL errors during execution.

## Features

- Converts Excel file data into SQL `INSERT` statements.
- Optionally generates a `CREATE TABLE` statement based on the Excel file's column structure.
- Escapes special characters like single quotes and backslashes to ensure valid SQL queries.
- Renames columns by removing or replacing spaces and special characters.
- Supports various data types, including `TEXT`, `INT`, `FLOAT`, `BOOLEAN`, and `DATETIME`.

## Requirements

- Python 3.x
- `pandas` for reading Excel files
- `openpyxl` for parsing Excel files

Install the required dependencies using:

```bash
pip install pandas openpyxl
```

## Usage

### Command Line Arguments

```bash
python script.py <input_file> <output_file> <table_name> [--create_table <table_sql_file>]
```

### Arguments

- **input_file**: Path to the Excel file to convert.
- **output_file**: Path to the SQL file where the `INSERT` statements will be written.
- **table_name**: Name of the MySQL table where the data will be inserted.
- **--create_table** (optional): Path to the SQL file where the `CREATE TABLE` statement will be written.

### Example Commands

1. **Generate SQL `INSERT` statements only**:
   ```
   python script.py data.xlsx data.sql my_table
   ```

   This will convert the `data.xlsx` file into SQL `INSERT` statements and save them in `data.sql` for the table named `my_table`.

2. **Generate both `CREATE TABLE` and `INSERT` statements**:
   ```bash
   python script.py data.xlsx data.sql my_table --create_table create_table.sql
   ```

   This command will create both `CREATE TABLE` and `INSERT` SQL statements. The `CREATE TABLE` statement will be saved in `create_table.sql`, while the `INSERT` statements will be in `data.sql`.

## How It Works

1. **Read Excel File**: 
   The script uses `pandas.read_excel()` to load the data from the Excel file into a DataFrame.

2. **Column Renaming**: 
   Column names are cleaned by removing special characters and replacing spaces with underscores to conform to SQL naming conventions.

3. **Generate `INSERT` Statements**: 
   The script generates a SQL `INSERT` statement for each row of data. It escapes special characters like single quotes (`'`) and backslashes (`\`) to ensure the SQL statements are valid.

4. **Optionally Generate `CREATE TABLE`**: 
   If the `--create_table` option is provided, the script analyzes the data types in the Excel file and generates a `CREATE TABLE` SQL statement. The data types are mapped to MySQL types like `TEXT`, `INT`, `FLOAT`, and `DATETIME`.

### Example SQL Outputs

#### `CREATE TABLE` Statement Example:

```sql
CREATE TABLE IF NOT EXISTS `my_table` (
    `Company_Name` TEXT,
    `Acquisition_Period` TEXT,
    `Number_of_Columns` INT,
    `Category` TEXT,
    `Year_Founded` INT,
    `Revenue_Millions` FLOAT
);
```

#### `INSERT` Statement Example:

```sql
INSERT INTO `my_table` (`Company_Name`, `Acquisition_Period`, `Number_of_Columns`, `Category`, `Year_Founded`, `Revenue_Millions`)
VALUES ('Example Corp', '2024 Q1', 12, 'Technology', 2005, 45.25);
```

## Handling Special Characters

The script automatically escapes single quotes and backslashes in text fields to prevent SQL syntax errors. For example:

- Data like `"It's a good deal"` becomes `"It''s a good deal"` in the SQL output to escape the single quote.

## Customization

If your data contains specific types or needs special handling, you can adjust the default SQL types in the `sql_types` dictionary. For example:

```python
sql_types = {
    'object': 'TEXT',
    'int64': 'INT',
    'float64': 'FLOAT',
    'bool': 'BOOLEAN',
    'datetime64[ns]': 'DATETIME'
}
```

This dictionary maps Python data types to MySQL data types. You can customize this mapping based on your database schema.

## Limitations

- The script is designed to work with MySQL and may need adjustments for other databases (e.g., PostgreSQL or SQLite).
- The default SQL types can be expanded based on your specific database schema needs.

## License

This project is licensed under the GNU License.

