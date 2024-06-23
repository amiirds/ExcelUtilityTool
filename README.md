
# Project Title: File Management and Processing Utilities

## Description
This project provides a collection of utilities for managing and processing files. It includes functionalities for reading and formatting text files, creating Excel files with hyperlinks, finding file paths based on filenames, and saving these paths to an Excel file. These utilities are designed to streamline file handling tasks and enhance productivity.

## Installation
To run this project, you need to have Python installed along with the `openpyxl` library. You can install the required dependencies using pip:

```bash
pip install openpyxl
```

## Usage
This project includes multiple functionalities that can be accessed via a menu-driven interface. The main features are as follows:

1. **Read and format a list of names:**
    - Reads a text file containing names and formats them into a single line separated by commas.
    - Usage:
      ```python
      read_and_format(input_filename, output_filename)
      ```
    - Command-line example:
      ```bash
      python app.py
      # Follow the menu prompts to input file paths
      ```

2. **Create an Excel file with hyperlinks:**
    - Reads an Excel file with names and paths, and creates a new Excel file with hyperlinks to the specified paths.
    - Usage:
      ```python
      create_hyperlink_excel(input_file, output_file)
      ```
    - Command-line example:
      ```bash
      python app.py
      # Follow the menu prompts to input file paths
      ```

3. **Find file paths and save to Excel:**
    - Searches a specified directory for files matching provided filenames and saves the paths to an Excel file.
    - Usage:
      ```python
      find_file_paths(directory, file_names)
      save_paths_to_excel(file_paths_dict, output_directory)
      ```
    - Command-line example:
      ```bash
      python app.py
      # Follow the menu prompts to input file paths and directory
      ```

4. **Menu:**
    - Provides a menu-driven interface to access the above functionalities.
    - Command-line example:
      ```bash
      python app.py
      ```

## Contributing
Contributions are welcome! Please follow these guidelines when contributing to the project:
1. Fork the repository.
2. Create a new branch for your feature or bugfix.
3. Commit your changes with descriptive commit messages.
4. Push your branch and create a pull request.
5. Ensure your code adheres to the project's coding standards and includes appropriate tests.

