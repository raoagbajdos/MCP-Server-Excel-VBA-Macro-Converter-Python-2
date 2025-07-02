#!/usr/bin/env python3
"""
FastMCP Server for Excel VBA to Python Conversion

This module implements a Model Context Protocol (MCP) server using FastMCP
to provide tools for converting Excel files with VBA macros to Python code.
"""

import asyncio
import json
import logging
import os
from pathlib import Path
from typing import Dict, List, Optional, Any

from fastmcp import FastMCP
from fastmcp.tools import Tool

from converter import VBAConverter
from excel_extractor import ExcelExtractor
from batch_converter import BatchConverter

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Initialize FastMCP server
mcp = FastMCP("Excel VBA to Python Converter")

# Initialize converters
vba_converter = VBAConverter()
excel_extractor = ExcelExtractor()
batch_converter = BatchConverter()


@mcp.tool()
async def convert_vba_file(file_path: str, output_dir: Optional[str] = None) -> Dict[str, Any]:
    """
    Convert an Excel file containing VBA macros to Python code.
    
    Args:
        file_path: Path to the Excel file (.xls, .xlsx, .xlsm)
        output_dir: Optional output directory for converted files
        
    Returns:
        Dict containing conversion results and file paths
    """
    try:
        file_path = Path(file_path)
        if not file_path.exists():
            return {"error": f"File not found: {file_path}"}
        
        if not file_path.suffix.lower() in ['.xls', '.xlsx', '.xlsm']:
            return {"error": "File must be an Excel file (.xls, .xlsx, .xlsm)"}
        
        # Extract VBA code
        vba_code = excel_extractor.extract_vba_code(file_path)
        if not vba_code:
            return {"error": "No VBA code found in the file"}
        
        # Convert to Python
        python_code = vba_converter.convert_to_python(vba_code)
        
        # Save output
        if output_dir:
            output_path = Path(output_dir) / f"{file_path.stem}_converted.py"
        else:
            output_path = file_path.parent / f"{file_path.stem}_converted.py"
        
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(python_code, encoding='utf-8')
        
        return {
            "success": True,
            "input_file": str(file_path),
            "output_file": str(output_path),
            "vba_modules": len(vba_code),
            "python_lines": len(python_code.split('\n'))
        }
        
    except Exception as e:
        logger.error(f"Error converting VBA file: {e}")
        return {"error": str(e)}


@mcp.tool()
async def extract_vba_code(file_path: str) -> Dict[str, Any]:
    """
    Extract VBA code from an Excel file without conversion.
    
    Args:
        file_path: Path to the Excel file
        
    Returns:
        Dict containing extracted VBA code and metadata
    """
    try:
        file_path = Path(file_path)
        if not file_path.exists():
            return {"error": f"File not found: {file_path}"}
        
        vba_code = excel_extractor.extract_vba_code(file_path)
        vba_modules = excel_extractor.get_vba_modules(file_path)
        
        return {
            "success": True,
            "file_path": str(file_path),
            "vba_code": vba_code,
            "modules": vba_modules,
            "module_count": len(vba_modules)
        }
        
    except Exception as e:
        logger.error(f"Error extracting VBA code: {e}")
        return {"error": str(e)}


@mcp.tool()
async def analyze_vba_complexity(file_path: str) -> Dict[str, Any]:
    """
    Analyze the complexity of VBA code in an Excel file.
    
    Args:
        file_path: Path to the Excel file
        
    Returns:
        Dict containing complexity analysis results
    """
    try:
        file_path = Path(file_path)
        vba_code = excel_extractor.extract_vba_code(file_path)
        
        if not vba_code:
            return {"error": "No VBA code found"}
        
        analysis = vba_converter.analyze_complexity(vba_code)
        
        return {
            "success": True,
            "file_path": str(file_path),
            "complexity_score": analysis["complexity_score"],
            "total_lines": analysis["total_lines"],
            "function_count": analysis["function_count"],
            "variable_count": analysis["variable_count"],
            "conversion_difficulty": analysis["difficulty_level"],
            "recommendations": analysis["recommendations"]
        }
        
    except Exception as e:
        logger.error(f"Error analyzing VBA complexity: {e}")
        return {"error": str(e)}


@mcp.tool()
async def batch_convert_files(directory_path: str, output_dir: Optional[str] = None) -> Dict[str, Any]:
    """
    Convert multiple Excel files with VBA in a directory.
    
    Args:
        directory_path: Path to directory containing Excel files
        output_dir: Optional output directory for converted files
        
    Returns:
        Dict containing batch conversion results
    """
    try:
        directory_path = Path(directory_path)
        if not directory_path.exists():
            return {"error": f"Directory not found: {directory_path}"}
        
        results = await batch_converter.convert_directory(directory_path, output_dir)
        
        return {
            "success": True,
            "directory": str(directory_path),
            "files_processed": results["processed"],
            "files_converted": results["converted"],
            "files_failed": results["failed"],
            "output_directory": results["output_dir"],
            "conversion_summary": results["summary"]
        }
        
    except Exception as e:
        logger.error(f"Error in batch conversion: {e}")
        return {"error": str(e)}


@mcp.tool()
async def generate_python_equivalent(vba_code: str) -> Dict[str, Any]:
    """
    Generate Python code equivalent for provided VBA code.
    
    Args:
        vba_code: VBA code string to convert
        
    Returns:
        Dict containing converted Python code
    """
    try:
        if not vba_code.strip():
            return {"error": "VBA code cannot be empty"}
        
        python_code = vba_converter.convert_to_python([vba_code])
        analysis = vba_converter.analyze_complexity([vba_code])
        
        return {
            "success": True,
            "input_vba": vba_code,
            "output_python": python_code,
            "conversion_notes": analysis["recommendations"],
            "complexity_score": analysis["complexity_score"]
        }
        
    except Exception as e:
        logger.error(f"Error generating Python equivalent: {e}")
        return {"error": str(e)}


async def main():
    """Run the FastMCP server."""
    try:
        logger.info("Starting Excel VBA to Python Converter MCP Server...")
        await mcp.run(transport="stdio")
    except Exception as e:
        logger.error(f"Server error: {e}")
        raise


if __name__ == "__main__":
    asyncio.run(main())