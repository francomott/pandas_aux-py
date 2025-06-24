# pandas_aux-py

Auxiliary methods for working with pandas.

## Requirements

- Python 3.13.0
- pandas
- openpyxl

## Installation

1. Clone the repository:
   ```bash
   git clone <repository-url>
   ```

2. Navigate to the project directory:
   ```bash
   cd pandas_aux-py
   ```

3. Create a virtual environment:
   ```bash
   python3.13 -m venv pandasvenv
   ```

4. Activate the virtual environment:
   - On macOS and Linux:
     ```bash
     source pandasvenv/bin/activate
     ```
   - On Windows:
     ```bash
     .\pandasvenv\Scripts\activate
     ```

5. Install the required packages:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

Below are examples of how to use the functions available in `pandas_aux.py`.

### excel_to_dataframe

```python
from pandas_aux import excel_to_dataframe

df = excel_to_dataframe('path_to_file.xlsx', sheet_name='Sheet1')
print(df)
```

### dataframe_to_excel

```python
from pandas_aux import dataframe_to_excel

file_path = dataframe_to_excel(df, 'output_directory', 'output_file.xlsx')
print(f"DataFrame exported to {file_path}")
```

### remove_duplicates

```python
from pandas_aux import remove_duplicates

df_no_duplicates = remove_duplicates(df, 'column_name')
print(df_no_duplicates)
```

### strip_dataframe

```python
from pandas_aux import strip_dataframe

df_stripped = strip_dataframe(df)
print(df_stripped)
```

### filter_dataframe

```python
from pandas_aux import filter_dataframe

filtered_df = filter_dataframe(df, 'column_name', 'parameter')
print(filtered_df)
```

### clean_special_characters

```python
from pandas_aux import clean_special_characters

cleaned_df = clean_special_characters(df, 'column_name')
print(cleaned_df)
```

### lowercase_dataframe

```python
from pandas_aux import lowercase_dataframe

lowercased_df = lowercase_dataframe(df, 'column_name')
print(lowercased_df)
```

### format_document

```python
from pandas_aux import format_document

formatted_df = format_document(df, 'column_name', 11)
print(formatted_df)
```

### drop_column

```python
from pandas_aux import drop_column

df_dropped = drop_column(df, 'column_name')
print(df_dropped)
```

### filter_by_different_char_count

```python
from pandas_aux import filter_by_different_char_count

filtered_df = filter_by_different_char_count(df, 'column_name', 10)
print(filtered_df)
```

### limit_column_size

```python
from pandas_aux import limit_column_size

limited_df = limit_column_size(df, 'column_name', 10)
print(limited_df)
```

### format_phone_number

```python
from pandas_aux import format_phone_number

formatted_phone_df = format_phone_number(df, 'column_name')
print(formatted_phone_df)
```

### rename_column

```python
from pandas_aux import rename_column

renamed_df = rename_column(df, 'old_name', 'new_name')
print(renamed_df)
```

## Contribution

Feel free to contribute to this project by opening issues or submitting pull requests.

## License

This project is licensed under the MIT License.
