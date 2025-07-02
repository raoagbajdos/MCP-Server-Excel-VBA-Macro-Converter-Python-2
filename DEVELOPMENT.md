# Development and Deployment Guide

## Development Setup

### Prerequisites
- Python 3.8 or higher
- pip package manager
- Git

### Local Development

1. **Clone and setup:**
```bash
git clone <repository-url>
cd python-mcp-vba-conversion
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate
pip install -r requirements.txt
pip install -e .  # Install in development mode
```

2. **Run tests:**
```bash
pytest tests/ -v
pytest tests/test_converter.py -v  # Run specific test file
```

3. **Code formatting:**
```bash
black src/ tests/
flake8 src/ tests/
```

## Usage Examples

### Command Line Interface

```bash
# Convert single file
python src/cli.py convert examples/sample_macro.xlsm

# Batch convert directory
python src/cli.py batch /path/to/excel/files -o /path/to/output

# Analyze VBA complexity
python src/cli.py analyze examples/sample_macro.xlsm --detailed

# Start MCP server
python src/cli.py server
```

### Programmatic Usage

```python
from src.converter import VBAConverter
from pathlib import Path

# Initialize converter
converter = VBAConverter()

# Convert a file
result = converter.convert_file(
    Path("sample.xlsm"), 
    Path("output.py")
)

if result["success"]:
    print(f"Converted {result['modules_converted']} modules")
    print(f"Complexity: {result['complexity_analysis']['difficulty_level']}")
else:
    print(f"Conversion failed: {result['error']}")
```

### MCP Client Integration

```python
import asyncio
from fastmcp import FastMCP

async def use_mcp_client():
    # Connect to the MCP server
    client = FastMCP("client")
    
    # Use the conversion tools
    result = await client.call_tool(
        "convert_vba_file",
        file_path="sample.xlsm",
        output_dir="converted/"
    )
    
    print(result)

asyncio.run(use_mcp_client())
```

## Configuration

Edit `src/config.py` to customize:

- **Logging levels**: `LOG_LEVEL = "DEBUG"`
- **Output formatting**: `INCLUDE_TYPE_HINTS = True`
- **Batch processing**: `BATCH_MAX_WORKERS = 8`
- **Complexity thresholds**: Adjust difficulty levels

## Architecture Overview

```
src/
├── mcp_server.py          # FastMCP server with tools
├── converter.py           # Main conversion orchestrator  
├── vba_parser.py         # VBA code parsing and analysis
├── python_generator.py   # Python code generation
├── excel_extractor.py    # Excel file and VBA extraction
├── batch_converter.py    # Batch processing utilities
├── cli.py               # Command-line interface
└── config.py            # Configuration settings
```

## Extending the Converter

### Adding New VBA Constructs

1. **Update VBA Parser** (`vba_parser.py`):
```python
def _parse_new_construct(self, line: str) -> Dict:
    # Add parsing logic for new VBA construct
    pass
```

2. **Update Python Generator** (`python_generator.py`):
```python
def _convert_new_construct(self, construct_info: Dict) -> List[str]:
    # Add Python code generation for the construct
    pass
```

### Adding New MCP Tools

```python
@mcp.tool()
async def new_conversion_tool(param: str) -> Dict[str, Any]:
    """New tool for specific conversion needs."""
    # Implementation
    return {"result": "success"}
```

## Testing

### Running Tests
```bash
# All tests
pytest

# Specific test categories
pytest tests/test_converter.py -v
pytest tests/test_vba_parser.py -v

# With coverage
pytest --cov=src tests/
```

### Adding Tests
```python
def test_new_feature():
    """Test new conversion feature."""
    # Setup
    converter = VBAConverter()
    
    # Test
    result = converter.new_method()
    
    # Assert
    assert result["success"] == True
```

## Deployment

### Docker Deployment

Create `Dockerfile`:
```dockerfile
FROM python:3.11-slim

WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt

COPY src/ ./src/
COPY setup.py .
RUN pip install -e .

EXPOSE 8000
CMD ["python", "src/mcp_server.py"]
```

### Production Considerations

1. **Error Handling**: Robust error handling for malformed VBA
2. **Performance**: Optimize for large Excel files
3. **Security**: Validate file inputs, sanitize VBA code
4. **Monitoring**: Add logging and metrics collection
5. **Scaling**: Consider async processing for concurrent requests

## Contributing

1. Fork the repository
2. Create a feature branch: `git checkout -b feature-name`
3. Make changes and add tests
4. Run tests: `pytest`
5. Format code: `black src/ tests/`
6. Commit changes: `git commit -m "Add feature"`
7. Push branch: `git push origin feature-name`
8. Create pull request

## Troubleshooting

### Common Issues

**VBA Extraction Fails:**
- Ensure file is macro-enabled (.xlsm)
- Check if VBA project is password protected
- Verify file is not corrupted

**Conversion Errors:**
- Complex VBA constructs may need manual review
- Check logs for specific parsing errors
- Validate generated Python syntax

**Performance Issues:**
- Reduce `BATCH_MAX_WORKERS` for memory constraints
- Process large files individually
- Consider streaming for very large datasets

### Debug Mode

Enable detailed logging:
```python
import logging
logging.basicConfig(level=logging.DEBUG)
```