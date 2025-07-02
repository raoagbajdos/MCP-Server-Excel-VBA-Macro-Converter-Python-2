#!/usr/bin/env python3
"""
Command Line Interface for VBA to Python Converter

This script provides a command-line interface for converting Excel VBA files to Python.
"""

import argparse
import asyncio
import logging
import sys
from pathlib import Path

from converter import VBAConverter
from batch_converter import BatchConverter
from config import *

# Configure logging
logging.basicConfig(
    level=getattr(logging, LOG_LEVEL),
    format=LOG_FORMAT,
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(LOG_FILE)
    ]
)

logger = logging.getLogger(__name__)


def setup_argparser():
    """Set up command line argument parser."""
    parser = argparse.ArgumentParser(
        description="Convert Excel VBA files to Python code using FastMCP",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Convert single file
  python cli.py convert sample.xlsm
  
  # Convert single file with custom output
  python cli.py convert sample.xlsm -o converted.py
  
  # Batch convert directory
  python cli.py batch /path/to/excel/files -o /path/to/output
  
  # Analyze VBA complexity
  python cli.py analyze sample.xlsm
  
  # Start MCP server
  python cli.py server
        """
    )
    
    subparsers = parser.add_subparsers(dest='command', help='Available commands')
    
    # Convert single file command
    convert_parser = subparsers.add_parser('convert', help='Convert a single Excel file')
    convert_parser.add_argument('file', help='Excel file to convert (.xlsx, .xlsm, .xls)')
    convert_parser.add_argument('-o', '--output', help='Output Python file path')
    convert_parser.add_argument('--include-formulas', action='store_true', 
                               help='Include Excel formulas in conversion')
    
    # Batch convert command
    batch_parser = subparsers.add_parser('batch', help='Convert multiple Excel files')
    batch_parser.add_argument('directory', help='Directory containing Excel files')
    batch_parser.add_argument('-o', '--output', help='Output directory for converted files')
    batch_parser.add_argument('-w', '--workers', type=int, default=BATCH_MAX_WORKERS,
                             help='Number of parallel workers')
    batch_parser.add_argument('--report', action='store_true',
                             help='Generate detailed conversion report')
    
    # Analyze command
    analyze_parser = subparsers.add_parser('analyze', help='Analyze VBA complexity')
    analyze_parser.add_argument('file', help='Excel file to analyze')
    analyze_parser.add_argument('--detailed', action='store_true',
                               help='Show detailed analysis')
    
    # Server command
    server_parser = subparsers.add_parser('server', help='Start MCP server')
    server_parser.add_argument('--host', default='localhost', help='Server host')
    server_parser.add_argument('--port', type=int, default=8000, help='Server port')
    
    return parser


async def convert_single_file(args):
    """Convert a single Excel file."""
    file_path = Path(args.file)
    
    if not file_path.exists():
        logger.error(f"File not found: {file_path}")
        return 1
    
    if file_path.suffix.lower() not in SUPPORTED_EXCEL_EXTENSIONS:
        logger.error(f"Unsupported file type: {file_path.suffix}")
        return 1
    
    # Set up output path
    if args.output:
        output_path = Path(args.output)
    else:
        output_path = file_path.parent / f"{file_path.stem}_converted.py"
    
    try:
        converter = VBAConverter()
        logger.info(f"Converting {file_path} to {output_path}")
        
        result = converter.convert_file(file_path, output_path)
        
        if result["success"]:
            logger.info("Conversion completed successfully!")
            logger.info(f"Output file: {result['output_file']}")
            logger.info(f"Modules converted: {result['modules_converted']}")
            
            # Show complexity analysis
            analysis = result["complexity_analysis"]
            logger.info(f"Complexity level: {analysis['difficulty_level']}")
            logger.info(f"Functions: {analysis['function_count']}")
            logger.info(f"Total lines: {analysis['total_lines']}")
            
            return 0
        else:
            logger.error(f"Conversion failed: {result['error']}")
            return 1
    
    except Exception as e:
        logger.error(f"Error during conversion: {e}")
        return 1


async def batch_convert_files(args):
    """Convert multiple Excel files in a directory."""
    directory_path = Path(args.directory)
    
    if not directory_path.exists():
        logger.error(f"Directory not found: {directory_path}")
        return 1
    
    output_dir = Path(args.output) if args.output else directory_path / DEFAULT_OUTPUT_DIR
    
    try:
        batch_converter = BatchConverter(max_workers=args.workers)
        logger.info(f"Starting batch conversion of {directory_path}")
        
        results = await batch_converter.convert_directory(directory_path, output_dir)
        
        logger.info("Batch conversion completed!")
        logger.info(f"Files processed: {results['processed']}")
        logger.info(f"Successfully converted: {results['converted']}")
        logger.info(f"Failed: {results['failed']}")
        logger.info(f"Output directory: {results['output_dir']}")
        
        if args.report:
            batch_converter.generate_batch_report(results, Path(results['output_dir']))
            logger.info("Detailed report generated")
        
        return 0 if results['failed'] == 0 else 1
    
    except Exception as e:
        logger.error(f"Error during batch conversion: {e}")
        return 1


async def analyze_vba_file(args):
    """Analyze VBA complexity in an Excel file."""
    file_path = Path(args.file)
    
    if not file_path.exists():
        logger.error(f"File not found: {file_path}")
        return 1
    
    try:
        converter = VBAConverter()
        
        # Extract VBA code
        vba_modules = converter.extractor.extract_vba_code(file_path)
        
        if not vba_modules:
            logger.info("No VBA code found in the file")
            return 0
        
        # Analyze complexity
        analysis = converter.analyze_complexity(vba_modules)
        
        print(f"\nVBA Complexity Analysis for {file_path.name}")
        print("=" * 50)
        print(f"Complexity Score: {analysis['complexity_score']}")
        print(f"Difficulty Level: {analysis['difficulty_level']}")
        print(f"Total Lines: {analysis['total_lines']}")
        print(f"Functions: {analysis['function_count']}")
        print(f"Variables: {analysis['variable_count']}")
        print(f"Dependencies: {len(analysis['dependencies'])}")
        
        if args.detailed:
            print(f"\nDetailed Analysis:")
            print(f"Dependencies: {', '.join(analysis['dependencies'])}")
            print(f"\nRecommendations:")
            for rec in analysis['recommendations']:
                print(f"  • {rec}")
        
        return 0
    
    except Exception as e:
        logger.error(f"Error during analysis: {e}")
        return 1


async def start_mcp_server(args):
    """Start the MCP server."""
    try:
        from mcp_server import main as server_main
        logger.info(f"Starting MCP server on {args.host}:{args.port}")
        await server_main()
        return 0
    except Exception as e:
        logger.error(f"Error starting server: {e}")
        return 1


async def main():
    """Main CLI function."""
    parser = setup_argparser()
    args = parser.parse_args()
    
    if not args.command:
        parser.print_help()
        return 1
    
    # Execute the appropriate command
    if args.command == 'convert':
        return await convert_single_file(args)
    elif args.command == 'batch':
        return await batch_convert_files(args)
    elif args.command == 'analyze':
        return await analyze_vba_file(args)
    elif args.command == 'server':
        return await start_mcp_server(args)
    else:
        parser.print_help()
        return 1


if __name__ == "__main__":
    try:
        exit_code = asyncio.run(main())
        sys.exit(exit_code)
    except KeyboardInterrupt:
        logger.info("Operation cancelled by user")
        sys.exit(1)
    except Exception as e:
        logger.error(f"Unexpected error: {e}")
        sys.exit(1)