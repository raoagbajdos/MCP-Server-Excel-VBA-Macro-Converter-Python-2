"""
Test VBA Converter

Unit tests for the VBA to Python converter functionality.
"""

import pytest
import tempfile
from pathlib import Path
from unittest.mock import Mock, patch

from src.converter import VBAConverter
from src.vba_parser import VBAParser
from src.python_generator import PythonGenerator


class TestVBAConverter:
    """Test cases for VBAConverter class."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.converter = VBAConverter()
    
    def test_convert_simple_vba_to_python(self):
        """Test conversion of simple VBA code."""
        vba_modules = [{
            'module_name': 'TestModule',
            'code': '''
Sub TestSub()
    Dim x As Integer
    x = 5
    MsgBox "Hello World"
End Sub
            '''.strip(),
            'type': 'module'
        }]
        
        result = self.converter.convert_to_python(vba_modules)
        
        assert 'def testsub()' in result.lower()
        assert 'x: int' in result
        assert 'print(' in result
    
    def test_analyze_complexity_simple(self):
        """Test complexity analysis for simple VBA code."""
        vba_modules = [{
            'module_name': 'Simple',
            'code': '''
Sub Simple()
    Dim x As Integer
    x = 1
End Sub
            '''.strip(),
            'type': 'module'
        }]
        
        analysis = self.converter.analyze_complexity(vba_modules)
        
        assert analysis['difficulty_level'] == 'Easy'
        assert analysis['function_count'] == 1
        assert 'recommendations' in analysis
    
    def test_analyze_complexity_complex(self):
        """Test complexity analysis for complex VBA code."""
        vba_modules = [{
            'module_name': 'Complex',
            'code': '''
Sub ComplexFunction()
    Dim i As Integer, j As Integer
    For i = 1 To 10
        If i Mod 2 = 0 Then
            For j = 1 To 5
                If j > 3 Then
                    MsgBox "Complex logic"
                End If
            Next j
        Else
            While i < 20
                i = i + 1
            Wend
        End If
    Next i
End Sub
            '''.strip(),
            'type': 'module'
        }]
        
        analysis = self.converter.analyze_complexity(vba_modules)
        
        assert analysis['complexity_score'] > 5
        assert analysis['function_count'] == 1
    
    def test_convert_empty_modules(self):
        """Test conversion with empty VBA modules."""
        result = self.converter.convert_to_python([])
        
        assert "No VBA code found" in result
    
    @patch('src.converter.VBAConverter._format_python_code')
    def test_format_python_code_called(self, mock_format):
        """Test that Python code formatting is called."""
        mock_format.return_value = "formatted code"
        
        vba_modules = [{
            'module_name': 'Test',
            'code': 'Sub Test()\nEnd Sub',
            'type': 'module'
        }]
        
        self.converter.convert_to_python(vba_modules)
        
        mock_format.assert_called_once()


class TestVBAParser:
    """Test cases for VBAParser class."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.parser = VBAParser()
    
    def test_parse_function_declaration(self):
        """Test parsing of VBA function declarations."""
        vba_code = '''
Public Function TestFunction(param1 As String, param2 As Integer) As Boolean
    TestFunction = True
End Function
        '''.strip()
        
        structure = self.parser.parse_code(vba_code)
        
        assert len(structure['functions']) == 1
        func = structure['functions'][0]
        assert func['name'] == 'TestFunction'
        assert func['type'] == 'FUNCTION'
        assert func['visibility'] == 'PUBLIC'
        assert len(func['parameters']) == 2
    
    def test_parse_variable_declarations(self):
        """Test parsing of variable declarations."""
        vba_code = '''
Sub TestSub()
    Dim x As Integer
    Dim y As String
    Dim z
End Sub
        '''.strip()
        
        structure = self.parser.parse_code(vba_code)
        
        func = structure['functions'][0]
        assert len(func['variables']) == 3
        assert func['variables'][0]['type'] == 'INTEGER'
        assert func['variables'][1]['type'] == 'STRING'
        assert func['variables'][2]['type'] == 'Variant'
    
    def test_complexity_calculation(self):
        """Test complexity score calculation."""
        vba_code = '''
Sub ComplexSub()
    If x > 0 Then
        For i = 1 To 10
            Select Case i
                Case 1
                    MsgBox "One"
                Case 2
                    MsgBox "Two"
            End Select
        Next i
    End If
End Sub
        '''.strip()
        
        structure = self.parser.parse_code(vba_code)
        
        assert structure['complexity_score'] > 3
    
    def test_comment_counting(self):
        """Test counting of comment lines."""
        vba_code = '''
' This is a comment
Sub TestSub()
    ' Another comment
    Dim x As Integer ' Inline comment (counted as code)
    REM This is also a comment
End Sub
        '''.strip()
        
        structure = self.parser.parse_code(vba_code)
        
        assert structure['comment_lines'] == 3
        assert structure['code_lines'] > 0


class TestPythonGenerator:
    """Test cases for PythonGenerator class."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.generator = PythonGenerator()
    
    def test_generate_imports(self):
        """Test generation of Python imports."""
        imports = self.generator._generate_imports()
        
        assert 'import pandas as pd' in imports
        assert 'import openpyxl' in imports
        assert 'from typing import' in ' '.join(imports)
    
    def test_convert_vba_types(self):
        """Test VBA to Python type conversion."""
        assert self.generator.type_mapping['INTEGER'] == 'int'
        assert self.generator.type_mapping['STRING'] == 'str'
        assert self.generator.type_mapping['BOOLEAN'] == 'bool'
        assert self.generator.type_mapping['VARIANT'] == 'Any'
    
    def test_convert_function_signature(self):
        """Test conversion of VBA function to Python signature."""
        vba_modules = [{
            'module_name': 'Test',
            'code': '''
Function TestFunc(param1 As String, Optional param2 As Integer) As String
    TestFunc = "result"
End Function
            '''.strip(),
            'type': 'module'
        }]
        
        result = self.generator.generate_python_code(vba_modules)
        
        assert 'def testfunc(' in result.lower()
        assert 'param1: str' in result
        assert 'Optional[int]' in result
        assert '-> Any:' in result


@pytest.fixture
def sample_vba_file(tmp_path):
    """Create a sample VBA file for testing."""
    # This would create a mock Excel file with VBA
    # In practice, you'd use a library like xlwt or create actual test files
    return tmp_path / "sample.xlsm"


@pytest.fixture
def temp_output_dir(tmp_path):
    """Create a temporary output directory."""
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    return output_dir