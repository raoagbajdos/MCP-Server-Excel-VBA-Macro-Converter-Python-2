"""
VBA to Python Converter

Main converter class that orchestrates the conversion process.
"""

import logging
from pathlib import Path
from typing import Dict, List, Any, Optional

from vba_parser import VBAParser
from python_generator import PythonGenerator
from excel_extractor import ExcelExtractor

logger = logging.getLogger(__name__)


class VBAConverter:
    """Main class for converting VBA code to Python."""
    
    def __init__(self):
        self.parser = VBAParser()
        self.generator = PythonGenerator()
        self.extractor = ExcelExtractor()
    
    def convert_to_python(self, vba_modules: List[Dict[str, str]]) -> str:
        """
        Convert VBA modules to Python code.
        
        Args:
            vba_modules: List of VBA modules to convert
            
        Returns:
            Generated Python code string
        """
        try:
            if not vba_modules:
                return "# No VBA code found to convert"
            
            # Generate Python code
            python_code = self.generator.generate_python_code(vba_modules)
            
            # Format the code
            formatted_code = self._format_python_code(python_code)
            
            return formatted_code
        
        except Exception as e:
            logger.error(f"Error converting VBA to Python: {e}")
            return f"# Conversion failed: {str(e)}"
    
    def analyze_complexity(self, vba_modules: List[Dict[str, str]]) -> Dict[str, Any]:
        """
        Analyze the complexity of VBA code.
        
        Args:
            vba_modules: List of VBA modules to analyze
            
        Returns:
            Complexity analysis results
        """
        total_complexity = 0
        total_lines = 0
        total_functions = 0
        total_variables = 0
        all_dependencies = set()
        
        for module in vba_modules:
            structure = self.parser.parse_code(module['code'])
            metrics = self.parser.get_complexity_metrics(structure)
            dependencies = self.parser.analyze_dependencies(module['code'])
            
            total_complexity += metrics['cyclomatic_complexity']
            total_lines += metrics['lines_of_code']
            total_functions += metrics['function_count']
            total_variables += metrics['variable_count']
            all_dependencies.update(dependencies)
        
        # Calculate overall difficulty
        if total_complexity < 20:
            difficulty = "Easy"
        elif total_complexity < 50:
            difficulty = "Medium"
        elif total_complexity < 100:
            difficulty = "Hard"
        else:
            difficulty = "Very Hard"
        
        # Generate recommendations
        recommendations = self._generate_recommendations(
            total_complexity, total_functions, len(all_dependencies)
        )
        
        return {
            "complexity_score": total_complexity,
            "total_lines": total_lines,
            "function_count": total_functions,
            "variable_count": total_variables,
            "dependencies": list(all_dependencies),
            "difficulty_level": difficulty,
            "recommendations": recommendations
        }
    
    def _generate_recommendations(self, complexity: int, functions: int, deps: int) -> List[str]:
        """Generate conversion recommendations based on analysis."""
        recommendations = []
        
        if complexity > 50:
            recommendations.append("High complexity detected. Consider breaking down into smaller functions.")
        
        if functions > 20:
            recommendations.append("Many functions detected. Consider organizing into classes or modules.")
        
        if deps > 5:
            recommendations.append("Multiple dependencies detected. Review external library requirements.")
        
        if complexity < 10:
            recommendations.append("Low complexity. Conversion should be straightforward.")
        
        recommendations.extend([
            "Review Excel object model usage for pandas/openpyxl equivalents",
            "Test converted code thoroughly with sample data",
            "Consider adding type hints for better code quality",
            "Add error handling for file operations",
            "Document any manual adjustments needed"
        ])
        
        return recommendations
    
    def _format_python_code(self, code: str) -> str:
        """Format Python code with proper styling."""
        lines = code.split('\n')
        formatted_lines = []
        
        # Add header comment
        formatted_lines.extend([
            '"""',
            'Converted from VBA to Python',
            'Generated automatically - review and test before use',
            '"""',
            ''
        ])
        
        # Process each line
        for line in lines:
            # Remove excessive blank lines
            if line.strip() == '' and formatted_lines and formatted_lines[-1] == '':
                continue
            formatted_lines.append(line)
        
        # Add main execution guard
        formatted_lines.extend([
            '',
            '',
            'if __name__ == "__main__":',
            '    # Example usage',
            '    pass  # TODO: Add example usage'
        ])
        
        return '\n'.join(formatted_lines)
    
    def convert_file(self, file_path: Path, output_path: Optional[Path] = None) -> Dict[str, Any]:
        """
        Convert a single Excel file with VBA to Python.
        
        Args:
            file_path: Path to Excel file
            output_path: Optional output file path
            
        Returns:
            Conversion results
        """
        try:
            # Extract VBA code
            vba_modules = self.extractor.extract_vba_code(file_path)
            
            if not vba_modules:
                return {
                    "success": False,
                    "error": "No VBA code found in file"
                }
            
            # Convert to Python
            python_code = self.convert_to_python(vba_modules)
            
            # Determine output path
            if not output_path:
                output_path = file_path.parent / f"{file_path.stem}_converted.py"
            
            # Write Python code
            output_path.write_text(python_code, encoding='utf-8')
            
            # Analyze complexity
            analysis = self.analyze_complexity(vba_modules)
            
            return {
                "success": True,
                "input_file": str(file_path),
                "output_file": str(output_path),
                "modules_converted": len(vba_modules),
                "complexity_analysis": analysis
            }
        
        except Exception as e:
            logger.error(f"Error converting file {file_path}: {e}")
            return {
                "success": False,
                "error": str(e)
            }


# Import asyncio for the convert_file method
import asyncio