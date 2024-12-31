# GPA Calculator

An automated tool for retrieving Seneca College course grades and calculating GPA.

[![GitHub repo](https://img.shields.io/badge/github-GPA__Calculator-blue)](https://github.com/ggeennn/GPA_Calculator.git)

## Features

- 🔐 Automated Seneca login
- 📊 Automatic grade retrieval
- 📈 GPA calculation
- 📑 Excel export
- 🔄 Multi-semester support

## Prerequisites

- Python 3.11+
- Chrome Browser
- Required packages:
  ```bash
  pip install selenium webdriver-manager openpyxl pandas
  ```

## Quick Start

1. Clone and setup:
```bash
git clone https://github.com/ggeennn/GPA_Calculator.git
cd GPA_Calculator
```

2. Set environment variables:
```bash
# Windows
set SENECA_EMAIL=your_email
set SENECA_PASSWORD=your_password

# Linux/macOS
export SENECA_EMAIL="your_email"
export SENECA_PASSWORD="your_password"
```

3. Run:
```bash
python s1.py
```

## Project Structure

```
├── s1.py                    # Main program
├── course_config.json       # Course settings
├── semester1_mapping.json   # Grade mapping rules
├── GPA_formula_guide.txt    # GPA calculation guide
└── README.md
```

## Supported Courses

Currently supports Semester 1 courses:
- IPC (Introduction to Programming Using C)
- OPS (Introduction to Operating Systems)
- CPR (Computer Principles for Programmers)
- APS (Applied Problem Solving)
- COM111 (Communication Course)

## Troubleshooting

- Verify environment variables
- Check internet connection
- Review `gpa_calculator.log`
- Update Chrome browser

## License

MIT License - See LICENSE file for details

## Contributing

Contributions welcome! Feel free to submit a Pull Request.
