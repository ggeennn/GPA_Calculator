# GPA Calculator

An automated tool for retrieving Seneca course grades and calculating GPA.

## Features

- Automated login to Seneca Learning Platform
- Automatic grade retrieval for all courses
- Excel export functionality
- Detailed logging system
- Support for multiple semesters (Semester 1 currently available)

## Supported Courses

### Semester 1
- IPC (Introduction to Programming Using C)
- OPS (Introduction to Operating Systems)
- CPR (Computer Principles for Programmers)
- APS (Applied Problem Solving)
- COM111 (Communication Course)

## Requirements

- Python 3.x
- Chrome Browser
- Python Packages:
  ```bash
  pip install selenium webdriver-manager openpyxl requests pandas numpy
  ```

## Installation

1. Clone the repository:
```bash
git clone [repository-url]
cd gpa-calculator
```

2. Set up environment variables:

### Linux/macOS: 
```bash
export SENECA_EMAIL="your_email"
export SENECA_PASSWORD="your_password"
```

### Windows: 
```bash
set SENECA_EMAIL=your_email
set SENECA_PASSWORD=your_password
```

## Project Structure

- `s1.py`: Main program for Semester 1
- `course_config.json`: Course configuration settings
- `semester1_mapping.json`: Grade mapping configuration
- `GPA_formula_guide.txt`: GPA calculation rules
- `SEMESTER 1 MARKINGS.xlsx`: Grade template file
- `gpa_calculator.log`: Program execution logs

## Usage

1. Ensure environment variables are set correctly
2. Verify the presence of `SEMESTER 1 MARKINGS.xlsx` template
3. Run the program:
```bash
python s1.py  # For Semester 1
```

## Important Notes

1. Ensure stable internet connection
2. Chrome WebDriver will be downloaded automatically on first run
3. Do not close the browser window during execution
4. All operations are logged in `gpa_calculator.log`

## Troubleshooting

If you encounter issues:
1. Verify environment variables are set correctly
2. Check internet connectivity
3. Review `gpa_calculator.log` for error messages
4. Ensure all configuration files are properly formatted
5. Update Chrome browser to the latest version

## Future Updates

- Support for Semester 2, 3, and 4
- Enhanced error handling
- Additional course support
- Grade trend analysis
- GPA prediction features

## Changelog

- Added merged cell handling support
- Implemented detailed logging system
- Optimized grade calculation logic
- Enhanced error handling mechanisms
- Structured for multi-semester support

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.
