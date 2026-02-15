📦 Enterprise Inventory Management System
A production-ready Excel-based inventory tracking system enhanced with Python automation, designed for small to medium operational environments requiring structured inventory governance and proactive restocking.
 
 
 
________________________________________
🎯 Problem Statement
Manual inventory tracking in operational environments led to:
.Inconsistent stock visibility
.Reactive ordering (frequent stockouts)
.No cost transparency
.No centralized reporting
.Time-consuming reconciliation processes
 The absence of structured digital tracking created operational inefficiencies and accountability gaps.
________________________________________
🏗 System Architecture

Excel User Interface
        ↓
Excel Validation & Formula Layer
        ↓
Python Automation Engine (OOP-based)
        ↓
Data Processing (pandas + openpyxl)
        ↓
Reporting & CSV Export
Architectural Layers
.Presentation Layer: Excel sheets for user-friendly interaction
.Business Logic Layer: Python InventoryManager class
.Data Layer: Structured Excel workbook
.Processing Layer: pandas-based bulk operations
.Reporting Layer: Automated text and CSV exports
________________________________________
⚙️ Technical Stack
Layer	Technology	Purpose
Backend Logic	Python 3.8+	Core automation engine
Excel Integration	openpyxl	Structured read/write
Data Processing	pandas	Data transformation
UI	Microsoft Excel	Operational interface
Environment	Virtualenv	Dependency isolation
________________________________________
🚀 Core Features
.Automated total cost calculations
.Dynamic reorder threshold alerts
.Multi-department support
.Cross-sheet references
.Dashboard with live summary metrics
.CSV backup export
.Batch item processing via Python
.Structured error handling
________________________________________
🔐 Data Integrity & Validation
.Excel-level data validation rules
.Duplicate item detection safeguards
.Formula error handling (#REF prevention)
.Input constraint enforcement
.Exception handling for file corruption
.Cross-platform file handling (Windows + WSL)
________________________________________
🧠 Key Engineering Decisions
.Excel chosen as UI layer to maximize accessibility in non-technical environments.
.Object-Oriented Python design to ensure extensibility.
.Dual-layer reorder logic (Excel + Python redundancy).
.Modular architecture to allow future migration to REST API backend.
.CSV export included for portability and backup resilience.
________________________________________
📈 Scalability Considerations
Designed for future extension into:
.REST API (FastAPI)
.PostgreSQL backend
.Web-based dashboard
.Role-based access control
.Cloud deployment (AWS S3 + RDS)
.Docker containerization
________________________________________
🧪 Testing
Functional testing performed across:
.Add/update item workflows
.Reorder threshold triggers
.Dashboard updates
.Report generation
.Cross-sheet references
.Zero and edge-case quantities
Validated for both Windows and Linux (WSL) environments.
________________________________________
📊 Operational Impact
.Significant reduction in manual reconciliation time
.Improved proactive restocking visibility
.Full cost transparency across departments
.Structured inventory governance implemented
31+ inventory items tracked across multiple departments.
________________________________________
🛠 Installation
git clone https://github.com/OCHHQ/inventory-management-system.git
cd inventory-management-system

python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
________________________________________
📖 Usage
Excel Mode
1.	Open inventory_tracker_template.xlsx
2.	Enter items and quantities
3.	Set reorder thresholds
4.	Monitor dashboard updates
Python Mode
from scripts.inventory_automation import InventoryManager

manager = InventoryManager('inventory_tracker.xlsx')
manager.add_item('Cleaning Items', 'Dish Soap', 12, 'bottles', 350, 5)
manager.update_quantity('Cleaning Items', 'Dish Soap', 8)
manager.generate_report('inventory_report.txt')
manager.save()
________________________________________
👤 Author
Enoseje Collins
Backend Engineer | Systems Automation | Information Management
GitHub: https://github.com/OCHHQ

