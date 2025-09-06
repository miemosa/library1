# Bank Fees Accrual Processing System

A professional web application for processing bank transaction data and generating comprehensive Excel reports with Panda Bank Transaction ID matching.

## Features

### 🏦 Professional Financial Processing
- **Multi-Account Support**: Process transactions for 5 key bank accounts (Chase & PNC)
- **Automated Matching**: Intelligent transaction matching between NetSuite and bank data
- **Panda Bank Transaction IDs**: Automatic assignment for non-batch activity accounts
- **Professional Excel Reports**: Financial institution-grade formatting and structure

### 📊 Comprehensive Reporting
- **Individual Account Tabs**: Separate worksheets for each account
- **Transaction Summary**: Detailed breakdown with amounts, dates, and memos
- **Professional Formatting**: Calibri fonts, borders, and alternating row colors
- **Account Totals**: Automatic calculation of account-level totals

### 🎯 Supported Accounts
- **10068** - Chase Reverse Wire Payrolls
- **10069** - Chase Recovery Ops  
- **10071** - Chase Wire In
- **10504** - Chase 3rd Party Processors
- **10513** - PNC Customer Wire Ins

### 💼 Business Rules
- Filters by Settlement State (Released/Settled)
- Transaction Status filtering (APPROVED)
- Transaction Type processing (Remittance_Debit_External/Remittance_Reversal)
- Comprehensive fee schedule for 25+ currencies
- Formula-driven calculations (Count × Fee Rate)

## Installation

### Prerequisites
- Python 3.8 or higher
- pip (Python package installer)

### Setup Instructions

1. **Clone the repository**
   ```bash
   git clone https://github.com/miemosa/library1.git
   cd library1
   ```

2. **Create virtual environment**
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install dependencies**
   ```bash
   pip3 install -r requirements.txt
   ```

4. **Run the application**
   ```bash
   python3 app.py
   ```

5. **Open browser**
   Navigate to `http://localhost:5000`

## Usage

### 1. File Upload
- **Bank Transaction Files**: Upload CSV files containing bank transaction data
- **NetSuite File**: Upload Excel file with NetSuite transaction details (header starts at row 7)

### 2. Processing
The system will automatically:
- Read and parse all uploaded files
- Filter NetSuite data by Customer Funds Obligation criteria
- Match transactions between NetSuite and bank data by amount
- Generate Panda Bank Transaction IDs for non-batch accounts

### 3. Report Generation
- Click "Generate & Download Report" to create Excel workbook
- Each account gets its own tab with formatted data
- Professional financial institution styling applied
- Panda Bank Transaction IDs included where applicable

## File Formats

### Bank Transaction CSV Files
Expected columns:
- `Panda Bank Transaction Id`
- `Signed Amount` 
- `Bank Transaction Date`

### NetSuite Excel File
- Header row starts at row 7
- Must contain columns: Date, Document Number, Memo, Amount, Account, Split
- Filters applied for "Customer Funds Obligation" transactions

## Technical Architecture

### Backend
- **Flask**: Web framework
- **pandas**: Data processing and manipulation
- **openpyxl**: Excel file generation with professional formatting
- **Werkzeug**: File upload handling

### Frontend  
- **Bootstrap 5**: Responsive UI framework
- **Font Awesome**: Professional icons
- **Custom CSS**: Financial institution styling
- **JavaScript**: Real-time status updates and file handling

### Data Processing
- Intelligent CSV parsing with multiple encoding support
- Amount-based transaction matching with tolerance
- Professional Excel formatting with Calibri fonts
- Account-specific processing rules

## Security Features
- File upload validation and sanitization
- Secure filename handling
- File size limits (100MB max)
- Input data validation

## Development

### Project Structure
```
├── app.py                 # Main Flask application
├── requirements.txt       # Python dependencies
├── templates/
│   └── index.html        # Main UI template
├── static/
│   ├── css/
│   │   └── style.css     # Professional styling
│   └── js/
│       └── app.js        # Frontend functionality
├── uploads/              # File upload directory
└── outputs/              # Generated reports
```

### Key Classes
- `BankFeesProcessor`: Core data processing logic
- `BankFeesApp` (JS): Frontend application controller

## API Endpoints
- `GET /` - Main application interface
- `POST /upload` - File upload and processing
- `GET /generate_report` - Excel report generation
- `GET /status` - Processing status information

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly with sample data
5. Submit a pull request

## License

This project is proprietary software for financial institution use.

## Support

For technical support or questions about the Bank Fees Accrual Processing System, please contact the development team.

---

**Note**: This system processes sensitive financial data. Ensure proper security measures are in place when deploying to production environments.
