# Journal Entry ID Creator

This script reads journal lines from an Excel file and automatically creates balanced journal entries where each entry has equal debits and credits, and all lines within an entry share the same posted date.

## Features

- **Automatic Journal Entry Creation**: Groups journal lines into balanced entries (debits = credits)
- **Date-based Grouping**: Ensures all lines in a journal entry have the same posted date
- **Smart Field Combinations**: Tests different combinations of optional fields to create meaningful groupings
- **Comprehensive Reporting**: Provides detailed summary of created entries and unassigned lines
- **Excel Integration**: Reads from and writes to Excel files with proper formatting

## Requirements

- Python 3.7+
- pandas
- openpyxl

Install requirements:
```bash
pip install -r requirements.txt
```

## Excel File Format

The input Excel file should have the following structure:

**Row 1 (Template/Headers):**
- Column A: "Posted Date" 
- Column B: "Account ID"
- Column C: "Debit Amount" 
- Column D: "Credit Amount"
- Columns E-M: Optional fields (Description, Reference, Department, Project, etc.)

**Row 2+: Journal Line Data**
Each row represents a single journal line with:
- Posted Date: Date of the transaction
- Account ID: Account identifier
- Debit Amount: Debit amount (0 if credit transaction)
- Credit Amount: Credit amount (0 if debit transaction)
- Optional fields: Additional identifying information

## Usage

### GUI Application (Recommended)
The easiest way to use the Journal Entry ID Creator is through the graphical user interface:

**Launch the GUI:**
```bash
python3 journal_entry_gui.py
```

**GUI Features:**
- 📁 Browse and select input Excel file
- 📂 Choose output file location (auto-suggested)
- ▶️ One-click processing with progress indicator
- 📊 Real-time results display
- ✅ Success/error notifications

### Command Line Usage
For advanced users or automation:

**Basic Usage:**
```bash
python journal_entry_creator.py input_file.xlsx
```

**With Custom Output File:**
```bash
python journal_entry_creator.py input_file.xlsx -o output_file.xlsx
```

**Limit Optional Fields for Grouping:**
```bash
python journal_entry_creator.py input_file.xlsx --max-fields 2
```

## Algorithm

The script uses a sophisticated algorithm to create journal entries:

1. **Load Data**: Reads journal lines from Excel file
2. **Field Analysis**: Identifies optional fields that contain meaningful data
3. **Combination Testing**: Tests different combinations of:
   - Posted Date only
   - Posted Date + 1 optional field
   - Posted Date + 2 optional fields  
   - Posted Date + 3 optional fields (default max)
4. **Balance Checking**: For each grouping, verifies that:
   - Total debits = total credits
   - Group has at least 2 lines (minimum for valid journal entry)
5. **Individual Entry Assignment**: Remaining unmatched lines are assigned individual journal IDs
6. **ID Assignment**: Assigns unique Journal Entry IDs (JE0001, JE0002, etc.) to ALL lines
7. **Output Generation**: Creates new Excel file with Journal ID column added

## Output

The script generates:
- **Excel File**: Original file with added "Journal ID" column (inserted before "Posted Date")
- **Console Report**: Summary showing:
  - Number of journal entries created
  - Lines per entry with amounts
  - Grouping criteria used
  - Any unassigned lines that couldn't be balanced

## Example

Input data:
```
Posted Date | Account ID | Debit Amount | Credit Amount | Description
2024-01-15  | 1000      | 500.00       | 0             | Cash Sale
2024-01-15  | 4000      | 0            | 500.00        | Cash Sale
2024-01-16  | 6000      | 250.00       | 0             | Office Supplies  
2024-01-16  | 1000      | 0            | 250.00        | Office Supplies
```

Output:
```
Journal ID | Posted Date | Account ID | Debit Amount | Credit Amount | Description
JE0001     | 2024-01-15  | 1000      | 500.00       | 0             | Cash Sale
JE0001     | 2024-01-15  | 4000      | 0            | 500.00        | Cash Sale
JE0002     | 2024-01-16  | 6000      | 250.00       | 0             | Office Supplies
JE0002     | 2024-01-16  | 1000      | 0            | 250.00        | Office Supplies
```

## Error Handling

The script handles various scenarios:
- **Unbalanced Lines**: Lines that can't form balanced entries are reported as unassigned
- **Missing Data**: Validates required columns exist
- **Date Formats**: Automatically parses common date formats
- **Empty Fields**: Ignores empty optional fields when grouping

## GUI Screenshots & Features

The GUI provides an intuitive interface with the following features:

### Main Interface
- **Input File Selection**: Browse for your Excel file with journal line data
- **Output Location**: Choose where to save the processed file (auto-suggested)
- **Progress Tracking**: Real-time progress bar and status updates
- **Results Display**: Comprehensive summary of created journal entries

### User Experience
- **Auto-completion**: Output filename automatically suggested based on input
- **Validation**: Input validation with helpful error messages
- **Threading**: Non-blocking processing - GUI remains responsive
- **Cross-platform**: Works on Windows, macOS, and Linux

## Limitations

- Maximum 5 optional fields used for grouping by default (configurable up to 8)
- Requires exact balance (debits = credits within 0.01 tolerance)
- Requires minimum 2 lines per journal entry (accounting standard)
- All lines in a journal entry must have identical posted dates
- Does not handle multi-currency transactions
