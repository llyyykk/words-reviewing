# Intelligent Vocabulary Review System

## Overview 
A Python-based vocabulary learning assistant that enables systematic word review through Excel integration, featuring adaptive learning strategies and data persistence.

## Key Features
- **Multi-file Integration**: Merge multiple Excel worksheets dynamically
- **Error Tracking**: Automatic mistake counting (`times` column)
- **Priority Management**: Importance-based filtering (`importance` column)
- **Flexible Review Modes**:
  - Sequential/Random order selection
  - Custom session length configuration
- **Data Safety**:
  - Auto-backup/restore system
  - Version-controlled file handling
- **Interactive Learning**:
  - Spacebar-controlled progression
  - Real-time progress tracking

## Installation
```bash
pip install pandas pathlib keyboard sys random shutil os
```

## Configuration
Modify `CONFIG` in the source code:

```python
CONFIG = {
    "input_dir": r"your/input/path",       # Primary storage for word files
    "backup_dir": r"your/backup/path",     # Versioned backups
    "file_prefix": "words_day",            # File naming pattern
    "file_suffix": ".xlsx",                # File extension
    "required_columns": [                  # Mandatory data schema
        "words", 
        "definition",
        "times", 
        "importance"
    ],
    "enable_backup": True                  # Backup toggle
}
```

## Workflow

### 1. File Initialization
```python
numbers, m = get_user_input()  # Get file numbers to load
loader = MultiExcelLoader(CONFIG, numbers)
combined_df = loader.combine_dataframes()
```

### 2. Review Session Configuration
- **Order Control**: 
  - Sequential (1) or Random (2)
  - Error-priority sorting option
- **Session Length**:
  - Full dataset or partial selection

### 3. Interactive Review
```python
# Sample word presentation flow
for j in shuffled_indices:
    print(combined_df.at[j, "words"])
    keyboard.wait('space')
    print(combined_df.at[j, "definition"])
    mistake_count(j, combined_df)
```

### 4. Data Persistence
- Automatic backup during file loading
- Session results saved to original files

## Usage Example
1. Prepare Excel files following `words_day[X].xlsx` format

2. Follow prompts:
```
Enter number of files to load: 2
Enter 1/2 number: 1
Enter 1/2 number: 2
Choose review order (1=Sequential/2=Random): 1
```

## Key Components
| Component | Functionality |
|-----------|---------------|
| `MultiExcelLoader` | Validates & loads Excel files |
| `create_backup()` | Creates timestamped backups |
| `data_back()` | Saves modified data to source files |
| `weigh_judge()` | Sorts by mistake frequency |
| `present_value()` | Controls word presentation flow |

## Error Handling
- Column validation for input files
- File existence checks
- Type checking for user inputs
- Graceful backup restoration

---

**Version**: 1.0 Beta  
**License**: MIT  
**Note**: Requires Excel files to be closed during operation for proper file handling.
