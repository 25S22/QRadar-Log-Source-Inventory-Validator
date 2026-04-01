# QRadar Log Source Inventory Validator 🔍

[![Python 3.6+](https://img.shields.io/badge/python-3.6+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![QRadar API](https://img.shields.io/badge/QRadar-API%20v14.0-red.svg)](https://www.ibm.com/docs/en/qradar-common?topic=overview-qradar-api)

> **Automated QRadar log source inventory validation and health monitoring tool**

A comprehensive Python script that validates log source inventories against live QRadar deployments, identifying inactive sources, configuration issues, and generating actionable reports for security operations teams.

## 📋 Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [Configuration](#configuration)
- [Usage](#usage)
- [API Endpoints](#api-endpoints)
- [Output Files](#output-files)
- [Customization](#customization)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [License](#license)

## 🎯 Overview

### What This Tool Does

The **QRadar Log Source Inventory Validator** automates the critical process of validating log source inventories against live QRadar environments. It cross-references documented log sources from Excel spreadsheets with real-time QRadar API data to provide comprehensive operational status reports.

### Why Use This Tool?

- **🔍 Inventory Auditing**: Automatically verify that documented log sources are actually configured and operational
- **📊 Health Monitoring**: Identify inactive or problematic log sources before they impact security monitoring
- **📈 Compliance Reporting**: Generate audit-ready reports showing log source operational status
- **⚡ Operational Efficiency**: Reduce manual verification time from hours to minutes
- **📧 Automated Alerting**: Generate email reports for stakeholders when issues are detected

### Business Impact

- **Security Operations**: Ensure complete visibility into your environment
- **Compliance**: Demonstrate active monitoring of security infrastructure
- **Cost Optimization**: Identify and remediate underutilized log sources
- **Risk Management**: Proactively address monitoring gaps

## ✨ Features

### Core Functionality
- ✅ **Dual Lookup Strategy**: Name-based and IP-based log source identification
- ✅ **Activity Analysis**: Configurable thresholds for determining log source health
- ✅ **Comprehensive Reporting**: Multiple output formats for different stakeholders
- ✅ **Error Handling**: Robust error detection and reporting
- ✅ **Batch Processing**: Handle multiple Excel sheets and large inventories
- ✅ **Email Integration**: Automated Outlook draft generation

### Advanced Features
- 🔐 **Secure Authentication**: Support for various QRadar authentication methods
- 📅 **Timestamp Validation**: Intelligent handling of QRadar timestamp formats
- 🔄 **Retry Logic**: Automatic retry mechanism for API failures
- 📊 **Progress Tracking**: Real-time progress indicators and statistics
- 🎨 **Customizable Outputs**: Flexible report generation and formatting

## 🛠 Prerequisites

### System Requirements
- **Python**: 3.6 or higher
- **Operating System**: Windows (for Outlook integration), Linux/macOS (core functionality)
- **Memory**: Minimum 512MB RAM (2GB+ recommended for large inventories)
- **Network**: Direct access to QRadar management interface

### QRadar Requirements
- **QRadar Version**: 7.3.0 or higher
- **API Access**: Valid user account with API permissions
- **Network Access**: HTTPS connectivity to QRadar console
- **Permissions**: Read access to log source management APIs

### Microsoft Outlook (Optional)
- **Version**: Outlook 2016 or higher
- **Purpose**: Automated email draft generation
- **Note**: Script functions without Outlook, but email features will be disabled

## 📦 Installation

### 1. Clone the Repository
```bash
git clone https://github.com/yourusername/qradar-inventory-validator.git
cd qradar-inventory-validator
```

### 2. Install Dependencies
```bash
pip install -r requirements.txt
```

**Or install manually:**
```bash
pip install pandas requests urllib3 pywin32 numpy openpyxl
```

### 3. Verify Installation
```bash
python qradar_log_source_checker.py --help
```

## 🚀 Quick Start

### 1. Prepare Your Excel File
Create an Excel file with the following structure:

| log source name | IP | Additional Columns |
|----------------|----|--------------------|
| firewall-01 | 192.168.1.100 | ... |
| server-web-01 | 10.0.1.50 | ... |
| switch-core | 172.16.1.1 | ... |

### 2. Configure the Script
Edit the configuration section in `qradar_log_source_checker.py`:

```python
# ─── CONFIGURATION ─────────────────────────────────────────────────────────────
INPUT_EXCEL_PATH = r'C:\path\to\your\inventory.xlsx'
SHEETS_TO_PROCESS = ['Sheet1']  # or ['all'] for all sheets
LOGSOURCE_COLUMN = 'log source name'
IP_COLUMN = 'IP'
QRADAR_HOST = 'https://your-qradar-host'
QRADAR_USERNAME = 'your-username'
QRADAR_PASSWORD = 'your-password'
ACTIVITY_THRESHOLD_DAYS = 7  # Consider inactive if no events in X days
```

### 3. Run the Script
```bash
python qradar_log_source_checker.py
```

### 4. Run the EPS Burn Rate Monitor
```bash
python eps_burn_rate_monitor.py
```

### 5. Run the Rule Effectiveness Auditor
```bash
python rule_effectiveness_auditor.py
```

### 6. Review Results
- **Console Output**: Real-time progress and statistics
- **Updated Excel**: Original file with new status columns
- **Filtered Report**: `inactive_and_errors.xlsx` with issues only
- **Email Draft**: Automatically generated Outlook draft (if available)

## ⚙️ Configuration

### Basic Configuration

#### File and Path Settings
```python
# Primary input file
INPUT_EXCEL_PATH = r'C:\path\to\your\input.xlsx'

# Sheets to process
SHEETS_TO_PROCESS = ['Sheet1', 'Sheet2']  # Specific sheets
SHEETS_TO_PROCESS = ['all']              # All sheets

# Column mapping
LOGSOURCE_COLUMN = 'log source name'  # Excel column with log source names
IP_COLUMN = 'IP'                      # Excel column with IP addresses
```

#### QRadar Connection
```python
# QRadar server details
QRADAR_HOST = 'https://your-qradar-host'
QRADAR_USERNAME = 'your-username'
QRADAR_PASSWORD = 'your-password'
VERIFY_SSL = False  # Set to True for production environments
```

### Advanced Configuration

#### Activity Monitoring
```python
# Activity threshold configuration
ACTIVITY_THRESHOLD_DAYS = 7  # Days to consider log source inactive

# Recommended values:
# - Critical systems: 1-2 days
# - Standard monitoring: 7 days  
# - Legacy systems: 14-30 days
```

#### Performance Tuning
```python
# Request timeout and retry settings
REQUEST_TIMEOUT = 30          # API request timeout (seconds)
MAX_SEARCH_RETRIES = 20       # Maximum retry attempts
SEARCH_RETRY_DELAY = 3        # Delay between retries (seconds)

# Timestamp validation range
MIN_TIMESTAMP = 0             # Unix epoch start
MAX_TIMESTAMP = 2147483647    # 32-bit system limit
```

## 📖 Usage

### Command Line Execution
```bash
# Basic execution
python qradar_log_source_checker.py

# With custom configuration (modify script)
python qradar_log_source_checker.py
```

### Execution Flow

1. **🔗 Connection Test**: Validates QRadar connectivity and authentication
2. **📖 Excel Processing**: Reads and validates input Excel file
3. **🔍 Log Source Lookup**: Performs API calls for each log source
4. **📊 Status Analysis**: Determines activity status and health
5. **💾 Report Generation**: Creates updated Excel and filtered reports
6. **📧 Email Creation**: Generates Outlook draft with results

### Monitoring Progress

The script provides detailed console output:

```
🚀 Starting QRadar Log Source Checker (Enhanced Version)...
🔗 Testing QRadar connection...
✅ QRadar connection successful!
📖 Reading Excel file: C:\inventory.xlsx
📋 Processing sheet: Sheet1
[1/100] Processing row 1
   🔍 Lookup by name: 'firewall-01'
   📊 Last Event: 2023-12-01 15:30:25 | Status: Active | Enabled: Yes
✅ Sheet Sheet1 completed: 95/100 log sources found
```

## 🔌 API Endpoints

### QRadar REST API Integration

#### 1. Connection Testing
```http
GET /api/help/versions
```
- **Purpose**: Validate connectivity and authentication
- **Headers**: `Accept: application/json`, `Version: 14.0`
- **Response**: QRadar version information

#### 2. Log Source Management
```http
GET /api/config/event_sources/log_source_management/log_sources
```
- **Purpose**: Retrieve log source configuration and status
- **Parameters**: 
  - `filter=name="log_source_name"` (name-based lookup)
  - `filter=ip_address="192.168.1.100"` (IP-based lookup)
- **Response**: Complete log source details including:
  - `id`: QRadar log source ID
  - `name`: Log source name
  - `enabled`: Configuration status
  - `last_event_time`: Last event timestamp (milliseconds)
  - `ip_address`: Configured IP address

#### 3. Ariel Search Submission
```http
POST /api/ariel/searches
```
- **Purpose**: Submit AQL queries for EPS analytics and rule fire counts
- **Usage in this repo**:
  - `eps_burn_rate_monitor.py`: 7-day EPS by log source and trend calculations
  - `rule_effectiveness_auditor.py`: 30-day rule fire counts for enabled rules

#### 4. Ariel Search Polling and Results
```http
GET /api/ariel/searches/{search_id}
GET /api/ariel/searches/{search_id}/results
```
- **Purpose**: Poll async AQL search status and retrieve results

#### 5. Analytics Rules
```http
GET /api/analytics/rules
```
- **Purpose**: Retrieve all enabled rules for effectiveness auditing
- **Filter**: `enabled=true`

### API Response Handling

#### Successful Response
```json
{
  "id": 123,
  "name": "firewall-01",
  "enabled": true,
  "last_event_time": 1701436825000,
  "ip_address": "192.168.1.100"
}
```

#### Error Handling
- **401 Unauthorized**: Invalid credentials
- **404 Not Found**: Log source doesn't exist
- **500 Internal Server Error**: QRadar system error
- **Timeout**: Network connectivity issues

## 📄 Output Files

### 1. Updated Excel File
**Location**: Same as input file  
**Content**: Original data plus new columns:

| Original Columns | status | qradar_id | enabled | last_seen | activity_status | days_since_last_event |
|------------------|--------|-----------|---------|-----------|-----------------|----------------------|
| ... | Found | 123 | Yes | 2023-12-01 15:30:25 | Active | 2 |

### 2. Filtered Report (`inactive_and_errors.xlsx`)
**Location**: Same directory as input file  
**Content**: Only problematic log sources:
- Inactive sources (no events within threshold)
- API errors (configuration issues)
- Not found sources (missing from QRadar)

### 3. Email Draft
**Platform**: Microsoft Outlook  
**Content**: Executive summary with statistics:
```
Subject: QRadar Log Source Status Report - 15 Issues Found

Hello,

Attached is the QRadar log source status report.

Summary:
- Total flagged log sources: 15
- Inactive/No Activity: 8
- API Errors: 2
- Not Found: 5

Please review the attached Excel file for detailed information.
```

### 4. EPS Burn Rate Report (`eps_burn_rate_report.xlsx`)
**Location**: Script working directory  
**Sheets**:
- `Summary`: Current/previous weekly EPS totals, license cap, days-until-cap projection
- `Source_Ranking`: Ranked log sources with EPS trend arrows and license-share %

### 5. Rule Effectiveness Audit (`rule_effectiveness_audit.xlsx`)
**Location**: Script working directory  
**Sheets**:
- `Summary`: Enabled rule counts with dead/noise totals
- `All_Enabled_Rules`: Full enabled-rule list with 30-day fire counts and classification
- `Dead_Rules`: Rules with 0 fires in lookback window
- `Noise_Generators`: Rules over configured fire threshold (default 500)

## 🎛 Customization

### Custom Column Names
If your Excel file uses different column names:

```python
# Modify these variables to match your Excel structure
LOGSOURCE_COLUMN = 'Device Name'      # Instead of 'log source name'
IP_COLUMN = 'Management IP'           # Instead of 'IP'
```

### Custom Activity Thresholds
```python
# Different thresholds for different environments
ACTIVITY_THRESHOLD_DAYS = 1   # Critical systems
ACTIVITY_THRESHOLD_DAYS = 7   # Standard systems
ACTIVITY_THRESHOLD_DAYS = 30  # Legacy systems
```

### Custom Report Filtering
Modify the `filter_and_email` function to customize what gets flagged:

```python
def filter_and_email(df_dict, draft_path):
    # Add custom filtering logic here
    # Example: Only flag sources with 0 events
    mask_no_events = df['activity_status'] == 'No Activity'
    
    # Example: Custom threshold per log source type
    for idx, row in df.iterrows():
        if 'critical' in row['log source name'].lower():
            custom_threshold = 1  # 1 day for critical systems
        else:
            custom_threshold = ACTIVITY_THRESHOLD_DAYS
```

### Environment-Specific Configuration
```python
# Development environment
if 'dev' in QRADAR_HOST.lower():
    ACTIVITY_THRESHOLD_DAYS = 30
    VERIFY_SSL = False
    
# Production environment
if 'prod' in QRADAR_HOST.lower():
    ACTIVITY_THRESHOLD_DAYS = 7
    VERIFY_SSL = True
```

## 🔧 Troubleshooting

### Common Issues and Solutions

#### Connection Problems
```bash
❌ Connection error! Check QRadar host URL.
```
**Solutions**:
- Verify `QRADAR_HOST` URL is correct
- Check network connectivity: `ping your-qradar-host`
- Ensure firewall rules allow HTTPS traffic
- Confirm QRadar management interface is accessible

#### Authentication Issues
```bash
❌ Authentication failed! Check username/password.
```
**Solutions**:
- Verify credentials are correct
- Check if account has API access permissions
- Ensure account is not locked or expired
- Test login via QRadar web interface

#### SSL Certificate Problems
```bash
❌ SSL Certificate verification failed
```
**Solutions**:
- Set `VERIFY_SSL = False` for self-signed certificates
- Install proper SSL certificates on QRadar
- Use IP address instead of hostname if DNS issues exist

#### Excel File Issues
```bash
❌ Missing columns in Sheet1: ['log source name']
```
**Solutions**:
- Verify column names match exactly (case-sensitive)
- Check for extra spaces in column headers
- Ensure Excel file is not corrupted
- Verify sheet names are correct

#### Performance Issues
```bash
# Slow processing or timeouts
```
**Solutions**:
- Increase `REQUEST_TIMEOUT` for slow networks
- Reduce `MAX_SEARCH_RETRIES` for faster failure detection
- Process smaller batches of log sources
- Check QRadar system performance

### Debug Mode
Enable verbose logging by modifying the script:

```python
import logging
logging.basicConfig(level=logging.DEBUG)

# Add debug prints in functions
def get_log_source_details(qradar_host, username, password, identifier, is_ip=False):
    print(f"DEBUG: Looking up {identifier} (IP: {is_ip})")
    # ... rest of function
```

### Log File Analysis
Monitor QRadar logs for API access:
```bash
# On QRadar system
tail -f /var/log/qradar.log | grep -i api
```

## 🤝 Contributing

### Development Setup
```bash
# Clone repository
git clone https://github.com/yourusername/qradar-inventory-validator.git
cd qradar-inventory-validator

# Create virtual environment
python -m venv venv
source venv/bin/activate  # Linux/macOS
# or
venv\Scripts\activate     # Windows

# Install development dependencies
pip install -r requirements-dev.txt
```

### Code Style
- Follow PEP 8 guidelines
- Use type hints where appropriate
- Add docstrings for all functions
- Include error handling for all API calls

### Testing
```bash
# Run unit tests
python -m pytest tests/

# Run integration tests (requires QRadar access)
python -m pytest tests/integration/
```

### Pull Request Process
1. Fork the repository
2. Create a feature branch: `git checkout -b feature/new-feature`
3. Make your changes and add tests
4. Ensure all tests pass
5. Submit a pull request with detailed description

## 📊 Performance Metrics

### Typical Performance
- **Small inventories** (< 100 sources): 2-5 minutes
- **Medium inventories** (100-500 sources): 10-30 minutes
- **Large inventories** (> 500 sources): 30+ minutes

### Optimization Tips
- Process during off-peak hours
- Use dedicated service account with minimal permissions
- Implement connection pooling for large inventories
- Consider parallel processing for very large datasets

## 📚 Additional Resources

### QRadar API Documentation
- [IBM QRadar API Guide](https://www.ibm.com/docs/en/qradar-common?topic=overview-qradar-api)
- [QRadar REST API Reference](https://www.ibm.com/docs/en/qradar-common?topic=api-rest-api-overview)

### Python Libraries
- [Pandas Documentation](https://pandas.pydata.org/docs/)
- [Requests Library](https://requests.readthedocs.io/)
- [OpenPyXL Documentation](https://openpyxl.readthedocs.io/)

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- IBM QRadar development team for comprehensive API documentation
- Python community for excellent libraries
- Security operations teams who provided requirements and feedback

---

**⭐ Star this repository if it helps you maintain your QRadar environment!**

**🐛 Found a bug? Please [open an issue](https://github.com/yourusername/qradar-inventory-validator/issues)**

**💡 Have a feature request? [Start a discussion](https://github.com/yourusername/qradar-inventory-validator/discussions)**
