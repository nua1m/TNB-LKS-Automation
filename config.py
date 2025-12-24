# ================================================================
#  CONFIG — CENTRALIZED CONSTANTS FOR RAW DATA, CLAIM, ATTACHMENT
# ================================================================

# RAW DATA sheet
DATA_SHEET_NAME = None          # Use first sheet if None
HEADER_ROW      = 1             # Row index for column names

# ================================================================
# RAW DATA COLUMN HEADERS (Must match EXACT Excel column names)
# ================================================================

COL_3MS_SO       = "3MS SO No."
COL_CONTRACT     = "Contract Account"
COL_SO_STATUS    = "SO Status"
COL_USER_STATUS  = "User Status"
COL_ADDRESS      = "Address"
COL_VOLTAGE      = "Voltage"
COL_SO_TYPE      = "SO Type"
COL_SO_DESC      = "SO Description"
COL_TECHNICIAN   = "Technician"
COL_STATUS_DATE  = "Status Date"
COL_SITE_ID      = "Site ID"

COL_OLD_METER    = "Old Meter no"
COL_OLD_COMM     = "Old Comm Module"
COL_NEW_METER    = "New Meter no"
COL_NEW_COMM     = "New Comm Module"

# ------------------------------------------------
# IMAGE URL COLUMN (from RAW)
# ------------------------------------------------
COL_ATTACH_URL   = "Attachments URL"   # Your column R


# ================================================================
# TEMPLATE SHEET NAMES
# ================================================================
CLAIM_SHEET_NAME      = "CLAIM"
ATTACH_SHEET_NAME     = "ATTACHMENT"


# ================================================================
# TEMPLATE ROW + COLUMN CONFIG
# ================================================================
DATA_START_ROW        = 3    # First row of data (row 3)
SERVICE_ORDER_COL_IDX = 2    # Column B in CLAIM/ATTACHMENT


# ================================================================
# OTHER CONSTANTS
# ================================================================
# Add more here if other modules need configuration
DEFAULT_TEMPLATE_PATH = r"C:\Users\syahm\Desktop\TNB_LKS_Dev\v1.4\Data\LKS Template (M).xlsm"
