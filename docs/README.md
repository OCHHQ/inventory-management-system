# 📊 INVENTORY TRACKING SYSTEM

Complete Excel-based inventory management system with Python automation for Cashier and Cleaning supplies.

---

## 📁 PROJECT STRUCTURE

```
inventory-system/
├── inventory_tracker.xlsx        # Main Excel file with all data
├── inventory_automation.py       # Python automation module
├── inventory_tracker.py          # Excel file generator
└── README.md                     # This file
```

---

## 📋 PART 1: EXCEL STRUCTURE

### **Sheet 1: Cashier Items**
Tracks office supplies for cashier operations.

**Columns:**
- **A: Item Name** - Name of the item
- **B: Quantity** - Current stock level
- **C: Unit** - Unit of measurement (pic, PAC, etc.)
- **D: Unit Price (₦)** - Price per unit (BLUE = user input)
- **E: Total Cost (₦)** - Auto-calculated (Quantity × Unit Price)
- **F: Reorder Level** - Minimum quantity threshold (BLUE = user input)
- **G: Status** - Auto-calculated ("REORDER" or "OK")

**Key Formulas:**
```excel
# Total Cost (Cell E6)
=B6*D6

# Status (Cell G6)
=IF(B6<F6,"REORDER","OK")

# Grand Total (Cell E20)
=SUM(E6:E18)

# Items to Reorder Count (Cell E21)
=COUNTIF(G6:G18,"REORDER")
```

### **Sheet 2: Cleaning Items**
Tracks cleaning supplies and maintenance items.

**Same structure as Cashier Items** with green color scheme.

### **Sheet 3: Dashboard**
Overview of both inventories with summary metrics.

**Key Formulas:**
```excel
# Cashier Total Cost (Cell B6)
='Cashier Items'!E20

# Cleaning Total Cost (Cell B10)
='Cleaning Items'!E26

# Grand Total (Cell B13)
=B6+B10
```

---

## 🎨 COLOR CODING GUIDE

| Color | Meaning | Usage |
|-------|---------|-------|
| **Blue Text** | User Input | Unit prices, reorder levels |
| **Black Text** | Formula | All calculations |
| **Green Headers** | Cleaning Items | Sheet-specific color |
| **Blue Headers** | Cashier Items | Sheet-specific color |

---

## 📊 PART 2: KEY FORMULAS EXPLAINED

### 1. **Total Cost Calculation**
```excel
=B6*D6
```
Multiplies Quantity (B6) by Unit Price (D6) to get total cost.

### 2. **Status Check**
```excel
=IF(B6<F6,"REORDER","OK")
```
Compares current quantity with reorder level:
- If quantity < reorder level → Shows "REORDER"
- Otherwise → Shows "OK"

### 3. **Sum Total**
```excel
=SUM(E6:E18)
```
Adds up all total costs in the range.

### 4. **Count Items Needing Reorder**
```excel
=COUNTIF(G6:G18,"REORDER")
```
Counts how many cells in status column show "REORDER".

### 5. **Cross-Sheet References**
```excel
='Cashier Items'!E20
```
Pulls data from another sheet. Format: `'SheetName'!CellAddress`

---

## 🐍 PART 3: PYTHON AUTOMATION

### **Setup**

1. **Install Dependencies:**
```bash
pip install openpyxl pandas --break-system-packages
```

2. **Import the Module:**
```python
from inventory_automation import InventoryManager

# Initialize
manager = InventoryManager('inventory_tracker.xlsx')
```

### **Basic Operations**

#### **Add New Item**
```python
manager.add_item(
    sheet_name='Cashier Items',
    item_name='Highlighter',
    quantity=12,
    unit='pic',
    unit_price=150,
    reorder_level=5
)
```

#### **Update Quantity**
```python
manager.update_quantity('Cashier Items', 'Marker', 25)
```

#### **Update Price**
```python
manager.update_price('Cleaning Items', 'Vim', 450)
```

#### **Get Items to Reorder**
```python
reorder_df = manager.get_items_to_reorder()
print(reorder_df)
```

#### **Generate Report**
```python
report = manager.generate_report('my_report.txt')
print(report)
```

#### **Export to CSV**
```python
manager.export_to_csv(output_dir='my_exports')
```

#### **Save Changes**
```python
manager.save()
```

---

## 🚀 ADVANCED AUTOMATION PROJECT

### **Project 1: Automated Reorder System**

```python
import smtplib
from email.message import EmailMessage
from inventory_automation import InventoryManager

def check_and_notify():
    manager = InventoryManager('inventory_tracker.xlsx')
    reorder_items = manager.get_items_to_reorder()
    
    if not reorder_items.empty:
        # Generate alert email
        msg = EmailMessage()
        msg['Subject'] = '⚠️ Inventory Reorder Alert'
        msg['From'] = 'inventory@company.com'
        msg['To'] = 'purchasing@company.com'
        
        body = "The following items need to be reordered:\n\n"
        body += reorder_items.to_string(index=False)
        msg.set_content(body)
        
        # Send email (configure your SMTP server)
        # with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
        #     smtp.send_message(msg)
        
        print("📧 Reorder notification sent!")
    else:
        print("✅ All items adequately stocked")

# Run daily check
if __name__ == "__main__":
    check_and_notify()
```

### **Project 2: Barcode Scanner Integration**

```python
import cv2
from pyzbar.pyzbar import decode
from inventory_automation import InventoryManager

def scan_and_update():
    manager = InventoryManager('inventory_tracker.xlsx')
    
    # Initialize camera
    cap = cv2.VideoCapture(0)
    
    while True:
        ret, frame = cap.read()
        
        # Decode barcodes
        for barcode in decode(frame):
            barcode_data = barcode.data.decode('utf-8')
            
            # Map barcode to item and update quantity
            item_map = {
                '1234567': ('Cashier Items', 'Marker', -1),
                '7654321': ('Cleaning Items', 'Vim', -1)
            }
            
            if barcode_data in item_map:
                sheet, item, qty_change = item_map[barcode_data]
                # Implement quantity update logic
                print(f"Scanned: {item}")
        
        cv2.imshow('Barcode Scanner', frame)
        
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break
    
    cap.release()
    cv2.destroyAllWindows()
    manager.save()
```

### **Project 3: Web Dashboard with Flask**

```python
from flask import Flask, render_template, jsonify
from inventory_automation import InventoryManager
import pandas as pd

app = Flask(__name__)

@app.route('/')
def dashboard():
    return render_template('dashboard.html')

@app.route('/api/inventory')
def get_inventory():
    manager = InventoryManager('inventory_tracker.xlsx')
    
    cashier_df = pd.read_excel('inventory_tracker.xlsx', sheet_name='Cashier Items', skiprows=4)
    cleaning_df = pd.read_excel('inventory_tracker.xlsx', sheet_name='Cleaning Items', skiprows=4)
    
    return jsonify({
        'cashier': cashier_df.to_dict('records'),
        'cleaning': cleaning_df.to_dict('records')
    })

@app.route('/api/reorder')
def get_reorder_items():
    manager = InventoryManager('inventory_tracker.xlsx')
    reorder_df = manager.get_items_to_reorder()
    return jsonify(reorder_df.to_dict('records'))

if __name__ == '__main__':
    app.run(debug=True)
```

### **Project 4: Scheduled Automation with Cron**

Create `inventory_check.py`:
```python
from inventory_automation import InventoryManager
from datetime import datetime

manager = InventoryManager('inventory_tracker.xlsx')

# Generate daily report
timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
manager.generate_report(f'reports/daily_report_{timestamp}.txt')

# Export to CSV for backup
manager.export_to_csv(f'backups/{timestamp}')

print(f"✅ Daily automation completed: {timestamp}")
```

Add to crontab (runs daily at 6 PM):
```bash
0 18 * * * cd /path/to/project && python inventory_check.py
```

---

## 📚 LEARNING RESOURCES

### **Excel Formulas:**
- IF function: https://support.microsoft.com/excel/if-function
- SUMIF/COUNTIF: https://support.microsoft.com/excel/sum-countif
- Cell references: https://support.microsoft.com/excel/references

### **Python Libraries:**
- openpyxl: https://openpyxl.readthedocs.io/
- pandas: https://pandas.pydata.org/docs/
- Flask: https://flask.palletsprojects.com/

---

## 🔧 TROUBLESHOOTING

### **Excel Formula Errors**

| Error | Cause | Solution |
|-------|-------|----------|
| #REF! | Invalid cell reference | Check formula cell references |
| #DIV/0! | Division by zero | Check if price is 0 |
| #VALUE! | Wrong data type | Ensure numeric values in qty/price |

### **Python Issues**

**Module not found:**
```bash
pip install openpyxl pandas --break-system-packages
```

**File not found:**
```python
# Use absolute path
manager = InventoryManager('/full/path/to/inventory_tracker.xlsx')
```

---

## 📈 NEXT STEPS

1. **Customize for your needs** - Add more fields (supplier, location, etc.)
2. **Add conditional formatting** - Highlight low stock items in red
3. **Create charts** - Visualize inventory levels and costs
4. **Implement database** - For multi-user access (SQLite, PostgreSQL)
5. **Build mobile app** - Use frameworks like Kivy or React Native
6. **Add authentication** - Secure access to inventory system

---

## 📞 SUPPORT

For questions or improvements, refer to:
- Excel documentation: https://support.microsoft.com/excel
- Python openpyxl docs: https://openpyxl.readthedocs.io/
- pandas documentation: https://pandas.pydata.org/

---

**Created by:** Fatmintih Peace Amammun  
**Date:** 20th July, 2024  
**Last Updated:** February 15, 2026
