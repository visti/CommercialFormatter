# Odin Data Manager

This project is a reference implementation for managing and processing specific files using the [Odin programming language](https://odin-lang.org/). 

## Features

- **Station-Specific Processing:**  
  Processes files based on predefined stations such as Bauer, Jyskfynske, and Globus, each with their own custom requirements.

- **Dynamic File Handling:**  
  Demonstrates how to read eligible files from the current directory based on file extensions and join file contents for further processing.

- **Special File Processing for Globus:**  
  When the "Globus" station is selected, each line of the input files is prepended with the file's base name (i.e., filename without its extension).

- **Stopword Filtering:**  
  Implements a stopword filtering mechanism to exclude lines with unwanted content during processing.

- **Customizable Data Extraction:**  
  Extracts specific segments of file content based on configurable positions, tailored for each station.

## Project Structure

- **main.odin:**  
  Contains the main entry point handling command-line arguments, station selection, and file processing logic.

- **Station Definitions and Processing Logic:**  
  Includes station configurations, file I/O functions, stopword checking, and filename manipulation helpers to demonstrate advanced file handling techniques in Odin.

## Note

This project is intended primarily as a reference for file handling and data processing using Odin. It may not be immediately applicable for any particular use case but can serve as a guide and educational resource for developing similar functionality in your own projects.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
