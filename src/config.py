# Excel VBA to Python Converter Configuration

# MCP Server Configuration
MCP_SERVER_NAME = "excel-vba-converter"
MCP_SERVER_VERSION = "0.1.0"
MCP_TRANSPORT = "stdio"

# Logging Configuration
LOG_LEVEL = "INFO"
LOG_FORMAT = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
LOG_FILE = "vba_converter.log"

# Conversion Settings
DEFAULT_OUTPUT_DIR = "converted_python"
BATCH_MAX_WORKERS = 4
SUPPORTED_EXCEL_EXTENSIONS = [".xlsx", ".xlsm", ".xls"]

# VBA Parsing Configuration
VBA_COMPLEXITY_THRESHOLDS = {
    "easy": 10,
    "medium": 30,
    "hard": 60
}

# Python Code Generation Settings
PYTHON_INDENT = "    "  # 4 spaces
INCLUDE_TYPE_HINTS = True
ADD_DOCSTRINGS = True
FORMAT_WITH_BLACK = False  # Set to True if black is available

# Excel Processing Settings
READ_ONLY_MODE = True
EXTRACT_FORMULAS = True
INCLUDE_WORKSHEETS_INFO = True

# Error Handling
CONTINUE_ON_ERROR = True
GENERATE_ERROR_REPORT = True

# Output Formatting
ADD_CONVERSION_COMMENTS = True
INCLUDE_ORIGINAL_VBA_AS_COMMENTS = False
GENERATE_USAGE_EXAMPLES = True