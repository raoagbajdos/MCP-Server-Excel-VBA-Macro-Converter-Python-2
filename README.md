# Excel VBA to Python Converter using FastMCP

A Model Context Protocol (MCP) server built with FastMCP that converts Excel files containing VBA macros into equivalent Python code.

## Features

- 🔄 Convert VBA macros to Python code
- 📊 Extract Excel data and structure
- 🐍 Generate clean, readable Python code
- 🛠️ Support for common VBA constructs
- 📁 Batch processing capabilities
- 🔍 Code analysis and optimization suggestions

## Installation

1. Clone the repository:
```bash
git clone <repository-url>
cd python-mcp-vba-conversion
```

2. Create a virtual environment:
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### As MCP Server

Start the FastMCP server:
```bash
python src/mcp_server.py
```

### Standalone Usage

Convert a single Excel file:
```bash
python src/converter.py path/to/your/file.xlsm
```

Convert multiple files:
```bash
python src/batch_converter.py path/to/excel/files/
```

## Project Structure

```
python-mcp-vba-conversion/
├── src/
│   ├── mcp_server.py          # FastMCP server implementation
│   ├── converter.py           # Main VBA to Python converter
│   ├── vba_parser.py          # VBA code parsing utilities
│   ├── python_generator.py    # Python code generation
│   ├── excel_extractor.py     # Excel file processing
│   └── batch_converter.py     # Batch processing utilities
├── tests/
│   ├── test_converter.py
│   ├── test_vba_parser.py
│   └── sample_files/
├── examples/
│   ├── sample_macro.xlsm
│   └── converted_output.py
├── requirements.txt
├── setup.py
└── README.md
```

## MCP Tools Available

- `convert_vba_file`: Convert a single Excel file with VBA
- `extract_vba_code`: Extract VBA code from Excel file
- `analyze_vba_complexity`: Analyze VBA code complexity
- `batch_convert_files`: Convert multiple Excel files
- `generate_python_equivalent`: Generate Python code from VBA

## Examples

See the `examples/` directory for sample Excel files and their Python conversions.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests
5. Submit a pull request

## License

MIT License