# Project Title
Python Excel Data Extractor


## Project Description
The Python Excel Data Extractor is a script that reads data from a source Excel sheet and maps specific fields to different locations in a destination Excel sheet. It automates the process of transferring data between sheets based on a predefined mapping.


## Run Instructions

1. Go to https://colab.research.google.com/.
2. Click on the "GitHub" tab.
3. Enter the URL of the Git repository:
    https://github.com/humaidan/pysheets1.git
4. Select main.ipynb file from the repository
5. The notebook will open in Google Colab, and you'll be able to execute the code cells and interact with the notebook's content.
6. Input your excel in main folder and the output file will be in out folder


## Configuration
You can modify the following variables in the `config.py` file:

```python
sheetname_source = "SourceSheet"
sheetname_dest = "DestinationSheet"

field_mapping = {
 "Field1": {"row": 1, "col": 1},
 "Field2": {"row": 2, "col": 2},
 "Field3": {"row": 3, "col": 3},
 # Add more field mappings as needed
 "Field20": {"row": 20, "col": 20}
}
```


## Contributing
Contributions are welcome! If you find any issues or have suggestions for enhancements, please submit a pull request or open an issue in the repository.


## License
This project is licensed under the MIT License.


## Authors
    Ali Humaidan (@humaidan)

