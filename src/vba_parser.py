"""
VBA Parser

This module provides functionality to parse VBA code and analyze its structure.
"""

import logging
import re
from typing import Dict, List, Any, Optional, Tuple

logger = logging.getLogger(__name__)


class VBAParser:
    """Parse and analyze VBA code structure."""
    
    def __init__(self):
        self.vba_keywords = {
            'control_flow': ['IF', 'THEN', 'ELSE', 'ELSEIF', 'END IF', 'FOR', 'NEXT', 
                           'WHILE', 'WEND', 'DO', 'LOOP', 'SELECT', 'CASE', 'WITH', 'END WITH'],
            'declarations': ['DIM', 'CONST', 'STATIC', 'GLOBAL', 'PRIVATE', 'PUBLIC'],
            'data_types': ['INTEGER', 'LONG', 'STRING', 'BOOLEAN', 'DOUBLE', 'SINGLE', 
                          'VARIANT', 'DATE', 'OBJECT', 'CURRENCY'],
            'functions': ['SUB', 'FUNCTION', 'END SUB', 'END FUNCTION', 'PROPERTY'],
            'excel_objects': ['RANGE', 'CELLS', 'WORKSHEET', 'WORKBOOK', 'APPLICATION'],
            'operators': ['AND', 'OR', 'NOT', 'MOD', 'LIKE']
        }
        
        self.function_pattern = re.compile(r'^\s*(PUBLIC|PRIVATE)?\s*(SUB|FUNCTION)\s+(\w+)', re.IGNORECASE)
        self.variable_pattern = re.compile(r'^\s*DIM\s+(\w+)', re.IGNORECASE)
        self.comment_pattern = re.compile(r"^\s*'|^\s*REM\s", re.IGNORECASE)
    
    def parse_code(self, vba_code: str) -> Dict[str, Any]:
        """
        Parse VBA code and extract structure information.
        
        Args:
            vba_code: VBA code string to parse
            
        Returns:
            Dictionary containing parsed code structure
        """
        lines = vba_code.split('\n')
        
        structure = {
            'functions': [],
            'variables': [],
            'imports': [],
            'constants': [],
            'classes': [],
            'line_count': len(lines),
            'comment_lines': 0,
            'code_lines': 0,
            'complexity_score': 0
        }
        
        current_function = None
        indent_level = 0
        
        for i, line in enumerate(lines, 1):
            original_line = line
            line = line.strip()
            
            if not line:
                continue
            
            # Count comments
            if self.comment_pattern.match(line):
                structure['comment_lines'] += 1
                continue
            
            structure['code_lines'] += 1
            
            # Parse functions and subroutines
            func_match = self.function_pattern.match(line)
            if func_match:
                if current_function:
                    current_function['end_line'] = i - 1
                
                current_function = {
                    'name': func_match.group(3),
                    'type': func_match.group(2).upper(),
                    'visibility': func_match.group(1).upper() if func_match.group(1) else 'PUBLIC',
                    'start_line': i,
                    'end_line': None,
                    'parameters': self._extract_parameters(line),
                    'variables': [],
                    'complexity': 1
                }
                structure['functions'].append(current_function)
            
            # Check for function end
            if re.match(r'^\s*END\s+(SUB|FUNCTION)', line, re.IGNORECASE):
                if current_function:
                    current_function['end_line'] = i
                    current_function = None
            
            # Parse variable declarations
            var_match = self.variable_pattern.match(line)
            if var_match:
                var_info = {
                    'name': var_match.group(1),
                    'line': i,
                    'type': self._extract_variable_type(line),
                    'scope': 'local' if current_function else 'module'
                }
                
                if current_function:
                    current_function['variables'].append(var_info)
                else:
                    structure['variables'].append(var_info)
            
            # Parse constants
            if re.match(r'^\s*CONST\s+', line, re.IGNORECASE):
                const_info = self._parse_constant(line, i)
                structure['constants'].append(const_info)
            
            # Calculate complexity
            complexity_keywords = ['IF', 'FOR', 'WHILE', 'SELECT', 'CASE']
            for keyword in complexity_keywords:
                if re.search(r'\b' + keyword + r'\b', line, re.IGNORECASE):
                    structure['complexity_score'] += 1
                    if current_function:
                        current_function['complexity'] += 1
        
        # Close last function if needed
        if current_function:
            current_function['end_line'] = len(lines)
        
        return structure
    
    def _extract_parameters(self, function_line: str) -> List[Dict[str, str]]:
        """Extract parameters from function declaration."""
        params = []
        
        # Find parameters between parentheses
        match = re.search(r'\(([^)]*)\)', function_line)
        if match:
            param_string = match.group(1).strip()
            if param_string:
                param_parts = param_string.split(',')
                for param in param_parts:
                    param = param.strip()
                    if param and len(param) > 0:
                        try:
                            param_info = self._parse_parameter(param)
                            params.append(param_info)
                        except Exception as e:
                            logger.warning(f"Failed to parse parameter '{param}': {e}")
                            # Add a default parameter info
                            params.append({
                                'name': param.split()[0] if param.split() else 'unknown',
                                'type': 'Variant',
                                'optional': False
                            })
        
        return params
    
    def _parse_parameter(self, param_str: str) -> Dict[str, str]:
        """Parse a single parameter string."""
        param_info = {
            'name': '',
            'type': 'Variant',
            'optional': False,
            'by_ref': True
        }
        
        # Check for Optional keyword
        if 'OPTIONAL' in param_str.upper():
            param_info['optional'] = True
            param_str = re.sub(r'\bOPTIONAL\b', '', param_str, flags=re.IGNORECASE).strip()
        
        # Check for ByVal/ByRef
        if 'BYVAL' in param_str.upper():
            param_info['by_ref'] = False
            param_str = re.sub(r'\bBYVAL\b', '', param_str, flags=re.IGNORECASE).strip()
        elif 'BYREF' in param_str.upper():
            param_str = re.sub(r'\bBYREF\b', '', param_str, flags=re.IGNORECASE).strip()
        
        # Extract name and type
        if ' AS ' in param_str.upper():
            parts = param_str.split(' AS ', 1)
            param_info['name'] = parts[0].strip()
            if len(parts) > 1:
                param_info['type'] = parts[1].strip()
            else:
                param_info['type'] = 'Variant'
        else:
            param_info['name'] = param_str.strip()
        
        return param_info
    
    def _extract_variable_type(self, dim_line: str) -> str:
        """Extract variable type from DIM statement."""
        if ' AS ' in dim_line.upper():
            parts = dim_line.upper().split(' AS ', 1)
            return parts[1].strip()
        return 'Variant'
    
    def _parse_constant(self, const_line: str, line_num: int) -> Dict[str, Any]:
        """Parse constant declaration."""
        const_info = {
            'name': '',
            'value': '',
            'type': 'Variant',
            'line': line_num
        }
        
        # Remove CONST keyword
        line = re.sub(r'^\s*CONST\s+', '', const_line, flags=re.IGNORECASE).strip()
        
        # Check for type declaration
        if ' AS ' in line.upper():
            parts = line.split(' AS ', 1)
            if len(parts) >= 2:
                name_value = parts[0].strip()
                const_info['type'] = parts[1].strip()
            else:
                name_value = line
        else:
            name_value = line
        
        # Extract name and value
        if '=' in name_value:
            name, value = name_value.split('=', 1)
            const_info['name'] = name.strip()
            const_info['value'] = value.strip()
        else:
            const_info['name'] = name_value
        
        return const_info
    
    def analyze_dependencies(self, vba_code: str) -> List[str]:
        """Analyze external dependencies in VBA code."""
        dependencies = []
        lines = vba_code.split('\n')
        
        for line in lines:
            line = line.strip().upper()
            
            # Look for Excel object model usage
            excel_objects = ['APPLICATION', 'WORKBOOK', 'WORKSHEET', 'RANGE', 'CELLS']
            for obj in excel_objects:
                if obj in line and obj not in dependencies:
                    dependencies.append(f"Excel.{obj}")
            
            # Look for Windows API calls
            if 'DECLARE' in line and 'LIB' in line:
                dependencies.append("Windows API")
            
            # Look for other common dependencies
            if 'FILESYSTEM' in line:
                dependencies.append("FileSystem")
            if 'SCRIPTING' in line:
                dependencies.append("Scripting")
        
        return list(set(dependencies))
    
    def get_complexity_metrics(self, structure: Dict[str, Any]) -> Dict[str, Any]:
        """Calculate complexity metrics from parsed structure."""
        metrics = {
            'cyclomatic_complexity': structure['complexity_score'],
            'function_count': len(structure['functions']),
            'variable_count': len(structure['variables']),
            'lines_of_code': structure['code_lines'],
            'comment_ratio': structure['comment_lines'] / max(structure['line_count'], 1),
            'average_function_complexity': 0,
            'max_function_complexity': 0
        }
        
        if structure['functions']:
            complexities = [func['complexity'] for func in structure['functions']]
            metrics['average_function_complexity'] = sum(complexities) / len(complexities)
            metrics['max_function_complexity'] = max(complexities)
        
        # Determine difficulty level
        if metrics['cyclomatic_complexity'] < 10:
            metrics['difficulty'] = 'Easy'
        elif metrics['cyclomatic_complexity'] < 20:
            metrics['difficulty'] = 'Medium'
        else:
            metrics['difficulty'] = 'Hard'
        
        return metrics