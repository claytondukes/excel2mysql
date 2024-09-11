import pandas as pd
import argparse
import re

def excel_to_sql(input_excel, output_sql, table_name, table_sql=None):
    # Read the Excel file
    df = pd.read_excel(input_excel)

    # Generate CREATE TABLE statement if requested
    if table_sql:
        renamed_columns = create_table_sql(df, table_sql, table_name)
    else:
        # If not creating a table, just rename the columns
        renamed_columns = rename_columns(df.columns)

    # Add `id` column to renamed columns
    renamed_columns_with_id = ['id'] + renamed_columns

    # Open the output SQL file to append insert statements
    with open(output_sql, 'w') as sql_file:
        for index, row in df.iterrows():
            # Prepare the SQL insert statement using the renamed columns list
            columns = ', '.join([f'`{col}`' for col in renamed_columns_with_id])
            values = ', '.join([f"'{str(val).replace('\\', '\\\\').replace('\'', '\'\'')}'" if pd.notna(val) else 'NULL' for val in row])
            insert_statement = f"INSERT INTO `{table_name}` ({columns}) VALUES (NULL, {values});\n"  # NULL for auto-increment id
            
            # Write the SQL statement to the file
            sql_file.write(insert_statement)

    print(f"SQL file '{output_sql}' created successfully.")

def create_table_sql(df, table_sql, table_name):
    # Define basic MySQL types; you can adjust based on your preferences.
    sql_types = {
        'object': 'TEXT',
        'int64': 'INT',
        'float64': 'FLOAT',
        'bool': 'BOOLEAN',
        'datetime64[ns]': 'DATETIME'
    }

    # Rename columns
    renamed_columns = rename_columns(df.columns)

    # Map the column types from the dataframe to MySQL types
    column_definitions = ['`id` INT NOT NULL AUTO_INCREMENT PRIMARY KEY']  # Add the `id` column definition
    for original_col, renamed_col in zip(df.columns, renamed_columns):
        mysql_type = sql_types.get(str(df[original_col].dtype), 'TEXT')  # Default to TEXT for unknown types
        column_definitions.append(f'`{renamed_col}` {mysql_type}')
    
    create_table_statement = f"CREATE TABLE IF NOT EXISTS `{table_name}` (\n" + \
                             ",\n".join(column_definitions) + "\n) ENGINE=InnoDB AUTO_INCREMENT=1 DEFAULT CHARSET=utf8mb4;\n"

    # Write the CREATE TABLE statement to the table SQL file
    with open(table_sql, 'w') as table_file:
        table_file.write(create_table_statement)

    print(f"CREATE TABLE statement written to '{table_sql}'.")
    return renamed_columns

def rename_columns(columns, max_length=64):
    seen_column_names = {}
    renamed_columns = []
    for col in columns:
        renamed_col = shorten_column_name(col, seen_column_names, max_length)
        renamed_columns.append(renamed_col)
    return renamed_columns

def shorten_column_name(column_name, seen_column_names, max_length=64):
    # Remove leading and trailing spaces and replace multiple spaces or special characters with underscores
    clean_name = re.sub(r'\s+', '_', column_name.strip())
    clean_name = re.sub(r'[^a-zA-Z0-9_]', '', clean_name)  # Remove any non-alphanumeric characters except underscores

    # Truncate if the name is still too long
    if len(clean_name) > max_length:
        clean_name = clean_name[:max_length]

    # Handle duplicate column names (case-insensitive)
    lower_clean_name = clean_name.lower()
    if lower_clean_name in seen_column_names:
        count = seen_column_names[lower_clean_name] + 1
        clean_name = f"{clean_name}_{count}"  # Only append suffix for actual duplicates
        seen_column_names[lower_clean_name] = count
    else:
        seen_column_names[lower_clean_name] = 1

    return clean_name

if __name__ == '__main__':
    # Set up argument parsing
    parser = argparse.ArgumentParser(description='Convert Excel file to SQL insert statements.')
    parser.add_argument('input_file', help='Path to the input Excel file')
    parser.add_argument('output_file', help='Path to the output SQL file')
    parser.add_argument('table_name', help='Name of the SQL table')
    parser.add_argument('--create_table', help='Path to the table SQL file to generate CREATE TABLE statement')

    args = parser.parse_args()

    # If --create_table is provided, we pass the path to create_table_sql
    excel_to_sql(args.input_file, args.output_file, args.table_name, table_sql=args.create_table)
