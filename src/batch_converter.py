"""
Batch Converter

This module provides functionality to convert multiple Excel files in batch.
"""

import asyncio
import logging
from pathlib import Path
from typing import Dict, List, Any, Optional
from concurrent.futures import ThreadPoolExecutor, as_completed

from converter import VBAConverter

logger = logging.getLogger(__name__)


class BatchConverter:
    """Convert multiple Excel files with VBA to Python."""
    
    def __init__(self, max_workers: int = 4):
        self.converter = VBAConverter()
        self.max_workers = max_workers
    
    async def convert_directory(self, directory_path: Path, output_dir: Optional[Path] = None) -> Dict[str, Any]:
        """
        Convert all Excel files in a directory.
        
        Args:
            directory_path: Directory containing Excel files
            output_dir: Optional output directory
            
        Returns:
            Batch conversion results
        """
        if not directory_path.exists():
            return {"error": f"Directory not found: {directory_path}"}
        
        # Find Excel files
        excel_files = []
        for pattern in ['*.xlsx', '*.xlsm', '*.xls']:
            excel_files.extend(directory_path.glob(pattern))
        
        if not excel_files:
            return {
                "processed": 0,
                "converted": 0,
                "failed": 0,
                "output_dir": str(output_dir) if output_dir else str(directory_path),
                "summary": "No Excel files found in directory"
            }
        
        # Set up output directory
        if not output_dir:
            output_dir = directory_path / "converted_python"
        output_dir.mkdir(exist_ok=True)
        
        # Convert files in parallel
        results = await self._convert_files_parallel(excel_files, output_dir)
        
        # Compile summary
        processed = len(excel_files)
        converted = sum(1 for r in results if r["success"])
        failed = processed - converted
        
        summary_lines = [
            f"Processed {processed} files",
            f"Successfully converted: {converted}",
            f"Failed: {failed}"
        ]
        
        if failed > 0:
            failed_files = [r["file"] for r in results if not r["success"]]
            summary_lines.append(f"Failed files: {', '.join(failed_files)}")
        
        return {
            "processed": processed,
            "converted": converted,
            "failed": failed,
            "output_dir": str(output_dir),
            "summary": "\n".join(summary_lines),
            "detailed_results": results
        }
    
    async def _convert_files_parallel(self, files: List[Path], output_dir: Path) -> List[Dict[str, Any]]:
        """Convert files in parallel using ThreadPoolExecutor."""
        loop = asyncio.get_event_loop()
        results = []
        
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            # Submit all conversion tasks
            future_to_file = {}
            for file_path in files:
                output_path = output_dir / f"{file_path.stem}_converted.py"
                future = loop.run_in_executor(
                    executor,
                    self._convert_single_file,
                    file_path,
                    output_path
                )
                future_to_file[future] = file_path
            
            # Collect results as they complete
            for future in as_completed(future_to_file):
                file_path = future_to_file[future]
                try:
                    result = await future
                    result["file"] = file_path.name
                    results.append(result)
                except Exception as e:
                    logger.error(f"Error converting {file_path}: {e}")
                    results.append({
                        "file": file_path.name,
                        "success": False,
                        "error": str(e)
                    })
        
        return results
    
    def _convert_single_file(self, file_path: Path, output_path: Path) -> Dict[str, Any]:
        """Convert a single file (for use in thread executor)."""
        try:
            result = self.converter.convert_file(file_path, output_path)
            return result
        except Exception as e:
            logger.error(f"Error in single file conversion {file_path}: {e}")
            return {
                "success": False,
                "error": str(e)
            }
    
    def generate_batch_report(self, results: Dict[str, Any], output_dir: Path) -> None:
        """Generate a batch conversion report."""
        report_path = output_dir / "conversion_report.txt"
        
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write("VBA to Python Batch Conversion Report\n")
            f.write("=" * 50 + "\n\n")
            f.write(f"Processed: {results['processed']} files\n")
            f.write(f"Converted: {results['converted']} files\n")
            f.write(f"Failed: {results['failed']} files\n\n")
            
            if "detailed_results" in results:
                f.write("Detailed Results:\n")
                f.write("-" * 20 + "\n")
                
                for result in results["detailed_results"]:
                    f.write(f"\nFile: {result['file']}\n")
                    if result["success"]:
                        f.write("Status: SUCCESS\n")
                        if "complexity_analysis" in result:
                            analysis = result["complexity_analysis"]
                            f.write(f"Complexity: {analysis['difficulty_level']}\n")
                            f.write(f"Functions: {analysis['function_count']}\n")
                            f.write(f"Lines: {analysis['total_lines']}\n")
                    else:
                        f.write("Status: FAILED\n")
                        f.write(f"Error: {result.get('error', 'Unknown error')}\n")
        
        logger.info(f"Batch report generated: {report_path}")


# Standalone batch conversion script
async def main():
    """Main function for standalone batch conversion."""
    import argparse
    
    parser = argparse.ArgumentParser(description="Batch convert Excel VBA files to Python")
    parser.add_argument("directory", help="Directory containing Excel files")
    parser.add_argument("-o", "--output", help="Output directory for converted files")
    parser.add_argument("-w", "--workers", type=int, default=4, help="Number of worker threads")
    
    args = parser.parse_args()
    
    directory_path = Path(args.directory)
    output_dir = Path(args.output) if args.output else None
    
    batch_converter = BatchConverter(max_workers=args.workers)
    
    print(f"Starting batch conversion of {directory_path}")
    results = await batch_converter.convert_directory(directory_path, output_dir)
    
    print("\nConversion Summary:")
    print(results["summary"])
    
    # Generate report
    if output_dir:
        batch_converter.generate_batch_report(results, Path(results["output_dir"]))


if __name__ == "__main__":
    asyncio.run(main())