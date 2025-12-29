# TNB LKS Automation

Automation toolkit for processing TNB (Tenaga Nasional Berhad) LKS (Laporan Kerja Siap) reports.

## Features

- **Automated Report Generation** - Process raw data and generate formatted LKS reports
- **Image Handling** - Keyboard automation for copying meter images between Excel sheets
- **Data Validation** - Quality control checks for missing images and data integrity
- **Date Processing** - Automatic date calculations and formatting

## Quick Start

### Prerequisites
- Python 3.8+
- Microsoft Excel
- Windows OS

### Installation

```bash
# Clone the repository
git clone https://github.com/nua1m/TNB-LKS-Automation.git
cd TNB-LKS-Automation

# Install dependencies
pip install -r requirements.txt
pip install pyautogui  # For keyboard automation
```

### Usage

**Main LKS Processing:**
```bash
python main.py "path/to/data.xlsx"
```

**Keyboard Image Automation:**
```bash
# Copy ticket images (card) to Column E
python keyboard_copy_images.py

# Copy new meter images to Column F
python keyboard_copy_new_meter.py
```

## Project Structure

```
v1.4/
├── main.py                    # Main LKS automation script
├── keyboard_copy_images.py    # Ticket image automation
├── keyboard_copy_new_meter.py # New meter image automation
├── config.py                  # Configuration settings
├── requirements.txt           # Python dependencies
├── core/
│   ├── excel_handler.py       # Excel file operations
│   ├── so_utils.py            # SO number utilities
│   └── services/
│       ├── claim_service.py   # Claim data processing
│       ├── date_engine.py     # Date calculations
│       ├── image_injector.py  # Image formula injection
│       ├── preprocessor.py    # Raw data preprocessing
│       └── quality_control.py # QC checks
└── ui/                        # UI components
```

## Keyboard Automation Setup

1. Open Excel with ATTACHMENT sheet active
2. Position cursor on the first SO cell (Column B)
3. Ensure DATA sheet is the next sheet (Ctrl+PageDown)
4. Run the script and follow prompts
5. **Important:** Focus Excel immediately after pressing Enter

**Safety:** Move mouse to top-left corner of screen to abort automation.

## License

Private - For internal use only.

## Author

Developed for TNB meter replacement workflow automation.
