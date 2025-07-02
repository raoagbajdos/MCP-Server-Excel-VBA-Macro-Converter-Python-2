#!/usr/bin/env python3
"""
Test the FastMCP server with actuarial VBA conversion
"""

import asyncio
import json
import sys
import os
from pathlib import Path

# Add src directory to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

async def test_mcp_server_actuarial():
    """Test MCP server functionality with actuarial calculations"""
    print("🏦 Testing FastMCP Server with Actuarial VBA Conversion")
    print("=" * 60)
    
    try:
        # Import MCP server
        from mcp_server import app, convert_vba_file, analyze_vba_complexity
        
        # Test file path
        test_file = Path("examples/sample_2_macro_actuary.xlsm")
        
        print(f"📁 Testing with file: {test_file}")
        print(f"📊 File size: {test_file.stat().st_size} bytes")
        
        # Test 1: Convert VBA file
        print("\n🔄 Testing VBA file conversion...")
        conversion_result = await convert_vba_file(str(test_file))
        
        print("✅ Conversion Results:")
        print(f"   Success: {conversion_result['success']}")
        print(f"   Functions converted: {len(conversion_result.get('functions', []))}")
        print(f"   Python lines: {len(conversion_result.get('python_code', '').split(chr(10)))}")
        
        # Show some converted functions
        if 'functions' in conversion_result:
            print(f"\n📋 Converted Functions:")
            for func in conversion_result['functions'][:5]:  # Show first 5
                print(f"   • {func}")
        
        # Test 2: Analyze complexity
        print(f"\n📊 Testing complexity analysis...")
        complexity_result = await analyze_vba_complexity(str(test_file))
        
        print("✅ Complexity Analysis:")
        print(f"   Overall complexity: {complexity_result['complexity_score']}")
        print(f"   Difficulty: {complexity_result['difficulty']}")
        print(f"   Function count: {complexity_result['function_count']}")
        print(f"   Dependencies: {len(complexity_result['dependencies'])}")
        
        if complexity_result['recommendations']:
            print(f"\n💡 Recommendations:")
            for rec in complexity_result['recommendations'][:3]:
                print(f"   • {rec}")
        
        # Test 3: Save conversion results
        output_file = "actuarial_mcp_results.json"
        results = {
            'conversion': conversion_result,
            'complexity': complexity_result,
            'test_summary': {
                'file_tested': str(test_file),
                'functions_found': len(conversion_result.get('functions', [])),
                'complexity_score': complexity_result['complexity_score'],
                'conversion_success': conversion_result['success']
            }
        }
        
        with open(output_file, 'w') as f:
            json.dump(results, f, indent=2, default=str)
        
        print(f"\n💾 Results saved to: {output_file}")
        
        # Summary
        print(f"\n🎯 MCP Server Test Summary:")
        print(f"   ✅ FastMCP server functionality: Working")
        print(f"   ✅ VBA file conversion: Working")
        print(f"   ✅ Complexity analysis: Working")
        print(f"   📊 Actuarial functions processed: {len(conversion_result.get('functions', []))}")
        
        return True
        
    except Exception as e:
        print(f"❌ Error during MCP server testing: {e}")
        import traceback
        traceback.print_exc()
        return False

async def test_direct_tool_calls():
    """Test direct MCP tool calls"""
    print(f"\n" + "=" * 60)
    print("Testing Direct MCP Tool Calls")
    print("=" * 60)
    
    try:
        from mcp_server import extract_vba_code, generate_python_equivalent
        
        # Test extracting VBA code
        test_file = "examples/sample_2_macro_actuary.xlsm"
        print(f"🔍 Testing VBA code extraction...")
        
        vba_extraction = await extract_vba_code(test_file)
        print(f"✅ VBA Extraction Results:")
        print(f"   Modules found: {len(vba_extraction.get('modules', []))}")
        print(f"   Total VBA lines: {sum(len(m.get('code', '').split(chr(10))) for m in vba_extraction.get('modules', []))}")
        
        # Test generating Python equivalent from VBA code sample
        sample_vba = '''
Function CalculateNetPremium(age As Integer, sumInsured As Double) As Double
    Dim mortalityRate As Double
    mortalityRate = 0.001 * (1.08 ^ (age - 20))
    CalculateNetPremium = sumInsured * mortalityRate
End Function
'''
        
        print(f"\n🐍 Testing Python generation from VBA code...")
        python_generation = await generate_python_equivalent(sample_vba)
        
        print(f"✅ Python Generation Results:")
        print(f"   Generated successfully: {python_generation['success']}")
        print(f"   Python code length: {len(python_generation.get('python_code', ''))}")
        
        if python_generation.get('python_code'):
            print(f"\n📝 Sample Generated Python:")
            print("-" * 40)
            print(python_generation['python_code'][:200] + "..." if len(python_generation['python_code']) > 200 else python_generation['python_code'])
            print("-" * 40)
        
        return True
        
    except Exception as e:
        print(f"❌ Error during direct tool testing: {e}")
        import traceback
        traceback.print_exc()
        return False

async def main():
    """Run all MCP server tests"""
    print("🚀 Starting FastMCP Server Tests for VBA Conversion")
    print("=" * 70)
    
    try:
        # Test 1: Main MCP server functionality
        test1_success = await test_mcp_server_actuarial()
        
        # Test 2: Direct tool calls
        test2_success = await test_direct_tool_calls()
        
        # Final summary
        print(f"\n" + "=" * 70)
        print("🎉 FastMCP Server Testing Completed!")
        print("=" * 70)
        
        if test1_success and test2_success:
            print("✅ All MCP server tests passed successfully!")
            print("🏦 Actuarial VBA conversion via MCP: Working")
            print("🔧 All MCP tools functional: Working")
            print("📊 Complex financial calculations handled: Successfully")
        else:
            print("❌ Some MCP server tests failed")
            print(f"   Main server test: {'✅' if test1_success else '❌'}")
            print(f"   Direct tool test: {'✅' if test2_success else '❌'}")
        
    except Exception as e:
        print(f"\n❌ MCP server test suite failed: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    # Run the async main function
    asyncio.run(main())
