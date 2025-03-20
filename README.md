# vbaFunctionsComparer
Do you need to compares VBA functions across two Excel workbooks? Here i am!

This repository contains a Python script that compares VBA functions across two Excel workbooks. It extracts VBA code from each workbook, identifies differences in function implementations, and generates separate files for each function that has changed.

## **Features**

✅ Extracts and normalizes VBA functions from two Excel workbooks\
✅ Identifies differences in function definitions and code\
✅ Generates separate `.bas` files for each differing function\
✅ Cleans up code by removing invisible differences (e.g., spaces, empty lines, indentation)

## **Usage Example**

```python
if __name__ == "__main__":

    new_file = "yournewfilepath"
    old_file = "youroldfilepath"
    
    comparer = CompareWorkbooks(new_file, old_file)
    comparer.showDifferences()
    comparer.close_workbooks()
```

This tool is useful for developers and analysts who need to track VBA function changes between different versions of an Excel file, ensuring better version control.

## **Installation**

Ensure you have Python installed along with the required dependencies. You can install them using:

```sh
pip install -r requirements.txt
```

## **Contributing**

Feel free to open issues or submit pull requests for improvements.

## **License**

This project is licensed under the MIT License.
