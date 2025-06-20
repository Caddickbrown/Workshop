# Workshop - Development & Automation Toolkit

A comprehensive workspace for personal development projects, work automation tools, and experimental applications. This repository serves as both a testing ground for new ideas and a storage location for production-ready utilities.

## 🚀 Quick Start

```bash
# Clone the repository
git clone <repository-url>
cd Workshop

# For Python projects
cd Python && python -r requirements.txt  # (if applicable)

# For web applications
cd Pomo && open index.html  # Or serve with local server

# For work tools (BVI)
# Import VBA macros into Excel as needed
# Run SQL scripts in your database management tool
```

## 📁 Project Structure

```
Workshop/
├── Python/              # Python utilities and applications
│   ├── CSV Editing/     # CSV file manipulation tools
│   ├── Film Time Check/ # Movie duration utilities
│   ├── Norwegian Fund Tracker/ # Financial tracking application
│   └── RecFile/         # File processing utilities
├── BVI/                 # Work-related automation tools
│   ├── Demand Plan/     # Production demand forecasting
│   ├── Daily Plan/      # Daily planning templates
│   ├── Issues Log/      # Problem tracking systems
│   └── BacklogPower/    # Backlog management tools
├── Pomo/                # Pomodoro timer web application
├── Hemingway/           # Text editor applications
├── Languahtml/          # Language learning web tool
├── Startup Standard/    # Computer setup automation guide
├── Document Maker/      # Document generation utilities
├── KenJournal/          # Raspberry Pi setup scripts
└── General/             # Miscellaneous VBA utilities
```

## 🛠️ Projects Overview

### Python Tools
- **Pomodoro Timer** (`Pomo.py`) - Command-line productivity timer
- **Norwegian Fund Tracker** - Financial portfolio monitoring
- **Random Number Generator** - Utility for generating random numbers
- **CSV Editing Tools** - Batch processing of CSV files
- **Film Time Check** - Movie duration verification utilities

### Web Applications
- **Pomodoro Web App** (`/Pomo/`) - Browser-based productivity timer with modern UI
- **Hemingway Editor** (`/Hemingway/`) - Text editing application with writing analytics
- **Language Learning Tool** (`/Languahtml/`) - Interactive language practice application
- **Fade Effects** - CSS transition and animation demonstrations

### Work Automation (BVI)
- **VBA Macro Collection** - 10+ Excel automation scripts for:
  - Issues logging and tracking
  - Backorder trending analysis
  - Sterilization list generation
  - Demand planning automation
- **SQL Scripts** - Database utilities for:
  - Assembly demand forecasting
  - Manufacturing planning
  - Packaging demand analysis
- **Production Tools** - Daily planning and control systems

### Development Setup
- **Startup Standard** - Comprehensive guide for setting up new development environments
- **Computer Configuration** - Automated installation scripts and setup procedures

## 🔧 Setup & Installation

### Prerequisites
- Python 3.7+ (for Python projects)
- Modern web browser (for web applications)
- Microsoft Excel (for VBA macros)
- SQL Server Management Studio (for database scripts)

### Python Dependencies
```bash
# Navigate to Python project directories and install requirements as needed
pip install pandas numpy requests  # Common dependencies
```

### Web Applications
Most web applications can be run by opening the HTML files in a browser or serving with a local web server:
```bash
# For development server
python -m http.server 8000
# Then visit http://localhost:8000
```

## 📖 Usage Examples

### Running Python Tools
```bash
cd Python
python Pomo.py                    # Start Pomodoro timer
python RandomNumberGenerator.py   # Generate random numbers
python Compare\ to\ 50.py        # Run comparison utility
```

### Using Web Applications
1. **Pomodoro Timer**: Open `Pomo/index.html` in browser
2. **Hemingway Editor**: Open `Hemingway/hemingway.html` 
3. **Language Tool**: Open `Languahtml/languahtml.html`

### Work Tools (BVI)
1. Import VBA files into Excel using Developer tab
2. Run SQL scripts in SQL Server Management Studio
3. Follow specific tool documentation in respective directories

## 🏗️ Development

### Adding New Projects
1. Create appropriate directory structure
2. Follow existing naming conventions
3. Add documentation to relevant README sections
4. Include usage examples

### Code Style
- Python: Follow PEP 8 guidelines
- JavaScript: Use modern ES6+ syntax
- VBA: Include error handling and comments
- SQL: Use clear, readable formatting

## 🔄 Automation Features

### Current Automations
- **Computer Setup** - Standardized development environment configuration
- **Excel Macros** - Automated data processing and reporting
- **SQL Utilities** - Database maintenance and analysis scripts

### Planned Automations
- Initialization script for complete environment setup
- Automated testing for Python utilities
- CI/CD pipeline integration

## 📝 Contributing

This is primarily a personal workspace, but contributions and suggestions are welcome:

1. Fork the repository
2. Create a feature branch
3. Make your changes with proper documentation
4. Submit a pull request with detailed description

## 🐛 Troubleshooting

### Common Issues
- **Python script errors**: Check Python version and installed packages
- **VBA macro security**: Ensure Excel macro security settings allow execution
- **Web app not loading**: Try serving with local web server instead of file:// protocol
- **SQL connection issues**: Verify database connection strings and permissions

### Getting Help
- Check individual project directories for specific documentation
- Review error logs and console output
- Ensure all prerequisites are properly installed

## 📊 Project Statistics
- **Languages**: Python, JavaScript, VBA, SQL, HTML/CSS
- **Total Projects**: 25+ utilities and applications
- **Active Development**: Ongoing maintenance and feature additions
- **Use Cases**: Personal productivity, work automation, learning experiments

## 📄 License

This project is for personal and educational use. Individual tools may have specific licensing requirements.

---

*Last Updated: December 2024*  
*Maintained by: Daniel Caddick-Brown*
