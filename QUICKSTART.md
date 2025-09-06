# Quick Start Guide - Bank Fees Accrual Processing System

## 🚀 Run the Application

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

### 2. Start the Application
```bash
python app.py
```

### 3. Open Browser
Navigate to: `http://localhost:5000`

## 📋 How to Use

### Step 1: Upload Files
- **Bank Transaction CSV Files**: Upload your CSV files (accounts 10068, 10069, 10071, 10504, 10513)
- **NetSuite Excel File**: Upload your NetSuite transaction details Excel file

### Step 2: Process
- Click "Process Files" button
- Wait for processing to complete
- View processing status and results

### Step 3: Download Report
- Click "Generate & Download Report" 
- Excel file will be automatically downloaded
- Each account gets its own professionally formatted tab

## 🔧 Features Included

✅ **Complete Flask Web Application**  
✅ **Professional Financial UI** (Calibri fonts, borders, alternating colors)  
✅ **Multi-Account Processing** (Chase & PNC accounts)  
✅ **Intelligent Transaction Matching**  
✅ **Panda Bank Transaction ID Assignment**  
✅ **Excel Report Generation** with professional formatting  
✅ **Real-time Processing Status**  
✅ **File Upload Validation**  
✅ **Responsive Design**  

## 📁 Data Files

The following files should be uploaded via the web interface:
- Bank Transaction CSV files (various accounts)
- NetSuite Transaction Details Excel file
- Any other data files needed for processing

These files are **excluded** from git tracking for security and size reasons.

## 🎯 Account Support

The system processes these specific bank accounts:
- **10068** - Chase Reverse Wire Payrolls
- **10069** - Chase Recovery Ops  
- **10071** - Chase Wire In
- **10504** - Chase 3rd Party Processors
- **10513** - PNC Customer Wire Ins

## 💡 Next Steps

1. Upload your data files through the web interface
2. Test the processing with your actual data
3. Customize styling or business rules as needed
4. Deploy to production environment when ready

---

**Your Bank Fees Accrual Processing System is now ready to use! 🏦✨**
