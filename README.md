# Rolladex Contact Management System

A Python-based contact management system that organizes client data and generates Excel spreadsheets with hyperlinked references to client document folders.

## Features

- **Data Organization**: Stores client information in structured text files
- **Excel Generation**: Creates formatted spreadsheets with client data
- **Hyperlink Integration**: Links directly to client document folders
- **Automated Processing**: Batch processes multiple client records

## Project Structure

```
rolladex/
├── main.py              # Main processing logic
├── startup.py           # Excel workbook management
├── dataFolder/          # Client data directory
│   └── [ClientID]/      # Individual client folders
│       ├── data.txt     # Client information file
│       └── Cdata/       # Client documents folder
└── outfile.xlsx         # Generated Excel output (excluded from git)
```

## Data Format

Each client folder contains a `data.txt` file with the following format:

```
Name: Client Company Name
Location: Street Address, City, State
Amount: 000.00
```

## Usage

### Running the Program

1. **Set up client data**: Create folders in `dataFolder/` with client information
2. **Run the script**: Execute the main program
   ```bash
   python main.py
   ```
3. **Output**: Check `outfile.xlsx` for the generated spreadsheet

### Adding New Clients

1. Create a new folder in `dataFolder/` (use a short ID as folder name)
2. Add a `data.txt` file with client information
3. Create a `Cdata/` subfolder for client documents
4. Run `main.py` to update the Excel file

## Components

### main.py
- **`getList()`**: Reads client data from text files
- **`placeHlink()`**: Creates hyperlink formulas for client folders
- **`writeRow()`**: Writes client data to Excel rows
- **`combDataFolder()`**: Processes all client folders

### startup.py
- **`Book` class**: Manages Excel workbook operations
- **File handling**: Creates, loads, and saves Excel files
- **Header management**: Sets up spreadsheet columns

## Requirements

- Python 3.x
- openpyxl library
- pathlib (built-in)

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/Sapphire-Enby/rolladex.git
   cd rolladex
   ```

2. Install dependencies:
   ```bash
   pip install openpyxl
   ```

3. Run the program:
   ```bash
   python main.py
   ```

## Example Output

The generated Excel file contains:
- **Name**: Client company name
- **Address**: Client location
- **Amount**: Associated monetary value
- **FolderPath**: Hyperlink to client document folder

## Contributing

Feel free to submit issues, fork the repository, and create pull requests for any improvements.

## License

This project is open source and available under the [MIT License](LICENSE).

