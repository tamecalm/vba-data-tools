# VBA Data Tools

A comprehensive collection of Excel VBA macros and functions designed for data processing, validation, and automation tasks. VBA Data Tools provides specialized scripts for address normalization, data categorization, and text analysis operations.

## Table of Contents

- [Project Overview](#project-overview)
- [Installation Instructions](#installation-instructions)
- [Usage Guide](#usage-guide)
- [Scripts Overview](#scripts-overview)
  - [NIS (Address Processing)](#nis-address-processing)
  - [SCRUTINIZE (Data Analysis)](#scrutinize-data-analysis)
- [Project Structure](#project-structure)
- [Configuration](#configuration)
- [Examples](#examples)
- [Contributing Guidelines](#contributing-guidelines)
- [License Information](#license-information)

## Project Overview

VBA Data Tools provides automated solutions for common Excel data processing tasks, with a focus on:

- **Address Data Normalization**: Standardizing directional suffixes and compound directions
- **Data Validation**: Identifying and highlighting inconsistent data patterns
- **Text Analysis**: Extracting and processing numeric values from mixed text
- **Data Categorization**: Automatically classifying entries based on predefined criteria
- **Regional Classification**: Assigning geopolitical regions based on state information

All scripts are designed to work with Excel's built-in VBA environment and require no external dependencies.

## Installation Instructions

### Prerequisites
- Microsoft Excel (2010 or later recommended)
- VBA enabled in Excel
- Macro security settings configured to allow VBA execution

### Setup Process

1. **Enable Developer Tab** (if not already visible):
   - Go to `File` → `Options` → `Customize Ribbon`
   - Check "Developer" in the right panel

2. **Configure Macro Security**:
   - Go to `Developer` → `Macro Security`
   - Select "Enable all macros" or "Disable all macros with notification"

3. **Install Scripts**:
   - Open Excel and press `ALT + F11` to open VBA Editor
   - For each script you want to use, follow the specific setup instructions in the corresponding `setup.md` file
   - Save your workbook as `.xlsm` (Excel Macro-Enabled Workbook)

For detailed setup instructions for individual scripts, refer to the `setup.md` files in each script's directory.

## Usage Guide

### Running Macros

1. **Via Macro Dialog**:
   - Press `ALT + F8`
   - Select the desired macro from the list
   - Click "Run"

2. **Via Developer Tab**:
   - Go to `Developer` → `Macros`
   - Select and run the desired macro

3. **Via VBA Editor**:
   - Press `ALT + F11`
   - Navigate to the macro in the code window
   - Press `F5` or click the "Run" button

### Using Custom Functions

Custom functions (UDFs) can be used directly in Excel formulas:

```excel
=SumNumbersInCell(A1)
=SumNumbersInText(B2)
```

### General Workflow

1. Prepare your data in the specified columns (typically Column A, C, F, or L depending on the script)
2. Ensure your worksheet is active and contains the target data
3. Run the appropriate macro
4. Review the results and any generated reports

## Scripts Overview

### NIS (Address Processing)

Scripts designed for normalizing and validating address data:

#### CheckUnreplacedAST
- **Purpose**: Validates address suffixes by highlighting cells ending in "AST" but not "EAST"
- **Target Column**: F
- **Output**: Yellow highlighting of problematic cells
- **Files**: `script.vba`, `setup.md`

#### ReplaceASTWithEast
- **Purpose**: Converts address suffixes from "AST" to "EAST" format
- **Target Column**: F
- **Pattern**: `SOUTHWEST AST` → `SOUTHWEST EAST`
- **Files**: `script.vba`, `setup.md`

#### FixMalformedEastWest
- **Purpose**: Corrects malformed directional patterns
- **Target Column**: F
- **Examples**: `E EAST` → `EAST`, `W EAST` → `WEST`
- **Files**: `script.vba`, `setup.md`

#### FixCompoundDirectionsAdvanced
- **Purpose**: Fixes complex compound direction errors
- **Target Column**: F
- **Examples**: `NORTH-E EAST` → `NORTH-EAST`, `NORTH-EEAST` → `NORTH-EAST`
- **Files**: `script.vba`, `setup.md`

### SCRUTINIZE (Data Analysis)

Scripts for data analysis, categorization, and validation:

#### SumNumbersInCell
- **Purpose**: Custom Excel function to sum all numbers within a cell's text
- **Type**: User Defined Function (UDF)
- **Usage**: `=SumNumbersInCell(A1)`
- **Files**: `script.vba`, `setup.md`

#### sumNumbersIntoText
- **Purpose**: Alternative implementation for extracting and summing numbers from text
- **Type**: Function with demonstration subroutine
- **Usage**: Can be called programmatically or via formula
- **Files**: `script.vba`, `setup.md`

#### textOnlyAddresses
- **Purpose**: Identifies and catalogs addresses containing no numeric characters
- **Target Columns**: C (addresses), L (officers)
- **Output**: New worksheet "TextOnlyAddresses"
- **Files**: `scipt.vba`, `setup.md`

#### IncompleteAddresses
- **Purpose**: Identifies incomplete addresses (text-only or numbers-only) and generates a report with a Pivot Table summary
- **Target Columns**: User-prompted (default Column C for addresses, Column L for officers)
- **Output**: New worksheet "IncompleteAddresses" with Pivot Table summary
- **Files**: `script.bas`, `setup.md`

#### ShallowAddresses
- **Purpose**: Flags shallow Nigerian addresses lacking location depth (strictly 3 words or fewer: Single-word, Two-word, Bare Street like "18 SARI STREET", and short 3-word entries)
- **Target Columns**: User-prompted (default Column C for addresses, Column L for officers)
- **Output**: "ShallowAddresses" (full audit + Pivot Table) and "FinalRecord" (50% Round Up summary table & 50% raw address rows linked)
- **Files**: `script.bas`, `setup.md`

#### CatgorizedRidersAndAdhoc
- **Purpose**: Categorizes personnel into HUB OFFICER, RIDER, ADHOC, or "Not In List"
- **Target Column**: A (names)
- **Output Column**: K (categories)
- **Files**: `script.vba`, `setup.md`

#### InsertRegionBasedOnState
- **Purpose**: Assigns geopolitical regions based on Nigerian state names
- **Target Column**: A (states)
- **Output Column**: J (regions)
- **Regions**: NORTH-CENTRAL, NORTH-EAST, NORTH-WEST
- **Files**: `script.vba`, `setup.md`

## Project Structure

```
VBA SCRIPTS/
├── README.md
├── NIS/                              # Address processing scripts
│   ├── CheckUnreplacedAST/
│   │   ├── script.vba
│   │   └── setup.md
│   ├── ReplaceASTWithEast/
│   │   ├── script.vba
│   │   └── setup.md
│   ├── FixMalformedEastWest/
│   │   ├── script.vba
│   │   └── setup.md
│   └── FixCompoundDirectionsAdvanced/
│       ├── script.vba
│       └── setup.md
└── SCRUTINIZE/                       # Data analysis scripts
    ├── IncompleteAddresses/
    │   ├── script.bas
    │   └── setup.md
    ├── ShallowAddresses/
    │   ├── script.bas
    │   └── setup.md
    ├── SumNumbersInCell/
    │   ├── script.vba
    │   └── setup.md
    ├── sumNumbersIntoText/
    │   ├── script.vba
    │   └── setup.md
    ├── textOnlyAddresses/
    │   ├── scipt.vba
    │   └── setup.md
    ├── CatgorizedRidersAndAdhoc/
    │   ├── script.vba
    │   └── setup.md
    └── InsertRegionBasedOnState/
        ├── script.vba
        └── setup.md
```

## Configuration

### Worksheet Requirements

Most scripts expect specific worksheet structures:

- **NIS Scripts**: Generally work on the active worksheet with data in Column F
- **SCRUTINIZE Scripts**: May require specific worksheet names (e.g., "Pivot (Old)")
- **Column Dependencies**: Scripts target specific columns (A, C, F, J, K, L)

### Customization

Many scripts include arrays that can be customized:

```vba
' Example from CatgorizedRidersAndAdhoc
adhocList = Array("Name1", "Name2", "Name3")
hubOfficerList = Array("Officer1", "Officer2")
ridersList = Array("Rider1", "Rider2")
```

### Data Backup

**Important**: Always backup your Excel files before running bulk processing macros, as changes may be irreversible.

## Examples

### Address Normalization Workflow

```excel
' Original data in Column F:
SOUTHWEST AST
NORTH-E EAST
E EAST

' After running ReplaceASTWithEast:
SOUTHWEST EAST
NORTH-E EAST
E EAST

' After running FixCompoundDirectionsAdvanced:
SOUTHWEST EAST
NORTH-EAST
E EAST

' After running FixMalformedEastWest:
SOUTHWEST EAST
NORTH-EAST
EAST
```

### Numeric Extraction Example

```excel
' Cell A1 contains: "Order123Item45Total"
=SumNumbersInCell(A1)
' Result: 168 (123 + 45)
```

### Regional Classification Example

```excel
' Column A contains: "Kaduna"
' After running InsertRegionBasedOnState
' Column J will contain: "NORTH-CENTRAL"
```

## Contributing Guidelines

### Code Standards

1. **Documentation**: All functions and subroutines must include header comments with:
   - Purpose description
   - Author information
   - Creation date
   - Parameter descriptions
   - Return value descriptions

2. **Naming Conventions**:
   - Use descriptive names for variables and functions
   - Follow camelCase for variables, PascalCase for functions
   - Include meaningful prefixes (e.g., `ws` for worksheets, `regex` for RegExp objects)

3. **Error Handling**:
   - Include appropriate error handling for file operations
   - Use `On Error Resume Next` judiciously
   - Validate input parameters

### Submission Process

1. Create a new directory under the appropriate category (NIS or SCRUTINIZE)
2. Include both `script.vba` and `setup.md` files
3. Follow the existing documentation format
4. Test thoroughly with sample data
5. Update this README.md with script information

### Setup Documentation

Each script must include a `setup.md` file containing:
- Overview and purpose
- Step-by-step installation instructions
- Usage examples
- Expected input/output formats
- Author information and creation date

## License Information

This project is created by **Ilesanmi Kehinde John (aka Calm)** and is available for personal, educational, and commercial use. Individual scripts may have specific licensing terms as noted in their respective documentation.

### Author Information

- **Primary Author**: Ilesanmi Kehinde John
- **Alias**: Calm
- **Contact**: Available in individual script documentation
- **Creation Period**: 2024-2025

### Usage Rights

- ✅ Personal use
- ✅ Educational use  
- ✅ Commercial use
- ✅ Modification and adaptation
- ✅ Distribution with attribution

### Disclaimer

These scripts are provided "as-is" without warranty. Users are responsible for testing scripts with their specific data and use cases. Always backup your data before running any macro that modifies worksheet content.

---

**Last Updated**: January 2025  
**Version**: 1.0  
**Compatibility**: Excel 2010+