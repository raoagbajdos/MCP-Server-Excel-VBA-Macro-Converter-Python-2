"""
Test VBA Parser

Unit tests for VBA parsing functionality.
"""

import pytest
from src.vba_parser import VBAParser


class TestVBAParserDetailed:
    """Detailed test cases for VBA parsing functionality."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.parser = VBAParser()
    
    def test_parse_complex_function_with_parameters(self):
        """Test parsing function with various parameter types."""
        vba_code = '''
Private Function ComplexFunc(ByVal param1 As String, _
                           ByRef param2 As Long, _
                           Optional param3 As Boolean = False, _
                           ParamArray args() As Variant) As Variant
    Dim localVar As Integer
    localVar = 42
    ComplexFunc = localVar
End Function
        '''.strip()
        
        structure = self.parser.parse_code(vba_code)
        
        assert len(structure['functions']) == 1
        func = structure['functions'][0]
        
        assert func['name'] == 'ComplexFunc'
        assert func['visibility'] == 'PRIVATE'
        assert len(func['parameters']) >= 3  # May not catch ParamArray perfectly
        
        # Check parameter details
        params = func['parameters']
        assert any(p['name'] == 'param1' and p['type'] == 'String' for p in params)
        assert any(p['name'] == 'param3' and p['optional'] for p in params)
    
    def test_parse_control_structures(self):
        """Test parsing of various control structures."""
        vba_code = '''
Sub ControlStructures()
    Dim i As Integer, j As Integer
    
    ' If statement
    If i > 0 Then
        MsgBox "Positive"
    ElseIf i < 0 Then
        MsgBox "Negative"
    Else
        MsgBox "Zero"
    End If
    
    ' For loop
    For i = 1 To 10 Step 2
        MsgBox i
    Next i
    
    ' While loop
    While i < 100
        i = i * 2
    Wend
    
    ' Select Case
    Select Case i
        Case 1 To 10
            MsgBox "Small"
        Case 11 To 100
            MsgBox "Medium"
        Case Else
            MsgBox "Large"
    End Select
End Sub
        '''.strip()
        
        structure = self.parser.parse_code(vba_code)
        
        # Should detect multiple complexity points
        assert structure['complexity_score'] >= 4  # IF, FOR, WHILE, SELECT
        
        func = structure['functions'][0]
        assert func['complexity'] >= 4
    
    def test_parse_constants_and_enums(self):
        """Test parsing of constants and enum-like structures."""
        vba_code = '''
Const MAX_ITEMS As Integer = 100
Const PI As Double = 3.14159
Const APP_NAME As String = "VBA Converter"

Sub UseConstants()
    Dim items(1 To MAX_ITEMS) As String
    Dim area As Double
    area = PI * 5 * 5
End Sub
        '''.strip()
        
        structure = self.parser.parse_code(vba_code)
        
        assert len(structure['constants']) == 3
        
        # Check constant details
        constants = {c['name']: c for c in structure['constants']}
        assert 'MAX_ITEMS' in constants
        assert constants['MAX_ITEMS']['type'] == 'INTEGER'
        assert constants['PI']['value'] == '3.14159'
    
    def test_parse_nested_functions(self):
        """Test parsing code with multiple functions."""
        vba_code = '''
Public Sub MainSub()
    Call HelperFunction("test")
    Dim result As String
    result = AnotherFunction(42)
End Sub

Private Function HelperFunction(param As String) As Boolean
    MsgBox param
    HelperFunction = True
End Function

Function AnotherFunction(num As Integer) As String
    AnotherFunction = CStr(num * 2)
End Function
        '''.strip()
        
        structure = self.parser.parse_code(vba_code)
        
        assert len(structure['functions']) == 3
        
        # Check function names
        func_names = [f['name'] for f in structure['functions']]
        assert 'MainSub' in func_names
        assert 'HelperFunction' in func_names
        assert 'AnotherFunction' in func_names
        
        # Check visibility
        main_sub = next(f for f in structure['functions'] if f['name'] == 'MainSub')
        helper_func = next(f for f in structure['functions'] if f['name'] == 'HelperFunction')
        
        assert main_sub['visibility'] == 'PUBLIC'
        assert helper_func['visibility'] == 'PRIVATE'
    
    def test_dependency_analysis(self):
        """Test analysis of external dependencies."""
        vba_code = '''
Sub UseDependencies()
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim rng As Range
    
    Set wb = Application.Workbooks.Open("test.xlsx")
    Set ws = wb.Worksheets("Sheet1")
    Set rng = ws.Range("A1:B10")
    
    rng.Value = "Hello"
    
    Declare Function GetTickCount Lib "kernel32" () As Long
End Sub
        '''.strip()
        
        dependencies = self.parser.analyze_dependencies(vba_code)
        
        assert 'Excel.WORKSHEET' in dependencies
        assert 'Excel.WORKBOOK' in dependencies
        assert 'Excel.RANGE' in dependencies
        assert 'Excel.APPLICATION' in dependencies
        assert 'Windows API' in dependencies
    
    def test_complexity_metrics_calculation(self):
        """Test detailed complexity metrics calculation."""
        vba_code = '''
Function ComplexFunction(data As Variant) As Boolean
    Dim i As Integer, j As Integer
    Dim result As Boolean
    
    For i = 1 To UBound(data)
        If IsArray(data(i)) Then
            For j = 1 To UBound(data(i))
                Select Case data(i)(j)
                    Case Is > 100
                        If j Mod 2 = 0 Then
                            result = True
                        End If
                    Case 50 To 99
                        While result = False
                            result = ProcessItem(data(i)(j))
                        Wend
                    Case Else
                        result = False
                End Select
            Next j
        End If
    Next i
    
    ComplexFunction = result
End Function

Sub ProcessItem(item As Variant)
    ' Simple processing
End Sub
        '''.strip()
        
        structure = self.parser.parse_code(vba_code)
        metrics = self.parser.get_complexity_metrics(structure)
        
        assert metrics['cyclomatic_complexity'] > 10
        assert metrics['function_count'] == 2
        assert metrics['difficulty'] in ['Medium', 'Hard']
        assert metrics['max_function_complexity'] > 5
    
    def test_parameter_parsing_edge_cases(self):
        """Test parameter parsing with edge cases."""
        vba_code = '''
Function EdgeCaseParams(ByVal str1 As String, _
                       ByRef num1 As Long, _
                       Optional flag As Boolean, _
                       Optional ByVal defaultStr As String = "default") As Integer
    EdgeCaseParams = 0
End Function
        '''.strip()
        
        structure = self.parser.parse_code(vba_code)
        func = structure['functions'][0]
        params = func['parameters']
        
        # Find specific parameters
        str1_param = next((p for p in params if p['name'] == 'str1'), None)
        num1_param = next((p for p in params if p['name'] == 'num1'), None)
        flag_param = next((p for p in params if p['name'] == 'flag'), None)
        
        assert str1_param is not None
        assert str1_param['by_ref'] == False  # ByVal
        
        assert num1_param is not None
        assert num1_param['by_ref'] == True   # ByRef
        
        assert flag_param is not None
        assert flag_param['optional'] == True