#!/usr/bin/env python3
"""
Test script for enhanced PPT and Excel features
Verifies that all new functionality works correctly
"""

import sys
import os
import json
from pathlib import Path

def test_imports():
    """Test that all required packages are installed"""
    print("\n" + "="*60)
    print("Testing Package Imports...")
    print("="*60)

    try:
        import openpyxl
        print("✅ openpyxl imported successfully")
    except ImportError as e:
        print(f"❌ openpyxl import failed: {e}")
        return False

    try:
        import pandas
        print("✅ pandas imported successfully")
    except ImportError as e:
        print(f"❌ pandas import failed: {e}")
        return False

    try:
        from pptx import Presentation
        print("✅ python-pptx imported successfully")
    except ImportError as e:
        print(f"❌ python-pptx import failed: {e}")
        return False

    try:
        from PIL import Image
        print("✅ Pillow imported successfully")
    except ImportError as e:
        print(f"❌ Pillow import failed: {e}")
        return False

    return True

def test_ppt_enhanced():
    """Test enhanced PPT creator"""
    print("\n" + "="*60)
    print("Testing Enhanced PowerPoint Creator...")
    print("="*60)

    try:
        sys.path.insert(0, 'professional-ppt-skill/scripts')
        from ppt_creator_enhanced import EnhancedPPTCreator

        creator = EnhancedPPTCreator()

        # Test theme listing
        themes = creator.get_available_themes()
        print(f"✅ Found {len(themes)} themes: {', '.join(themes[:5])}...")

        # Test theme preview
        preview = creator.get_theme_preview('corporate_blue')
        print(f"✅ Theme preview works: {preview['name']}")

        # Test simple presentation creation
        config = {
            "theme": "modern_minimal",
            "output_path": "test_output/test_presentation.pptx",
            "slides": [
                {
                    "type": "title",
                    "title": "Test Presentation",
                    "subtitle": "Automated Test"
                },
                {
                    "type": "content",
                    "title": "Test Slide",
                    "bullets": ["Test point 1", "Test point 2"]
                }
            ]
        }

        os.makedirs("test_output", exist_ok=True)
        result = creator.create_presentation(config)
        print(f"✅ Created test presentation: {result}")

        if os.path.exists(result):
            print(f"✅ File exists and is {os.path.getsize(result)} bytes")
            return True
        else:
            print(f"❌ File was not created")
            return False

    except Exception as e:
        print(f"❌ PPT test failed: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_excel_enhanced():
    """Test enhanced Excel creator"""
    print("\n" + "="*60)
    print("Testing Enhanced Excel Master...")
    print("="*60)

    try:
        sys.path.insert(0, 'excel-master-skill/scripts')
        from excel_master_enhanced import EnhancedExcelMaster

        master = EnhancedExcelMaster()

        # Test theme listing
        themes = master.get_available_themes()
        print(f"✅ Found {len(themes)} themes: {', '.join(themes)}")

        # Test theme preview
        preview = master.get_theme_preview('corporate_blue')
        print(f"✅ Theme preview works: {preview['name']}")

        # Test simple workbook creation
        config = {
            "theme": "corporate_blue",
            "output_path": "test_output/test_workbook.xlsx",
            "sheets": [
                {
                    "name": "Test_Data",
                    "type": "data",
                    "headers": ["Name", "Value", "Status"],
                    "data": [
                        ["Item 1", 100, "Active"],
                        ["Item 2", 200, "Pending"],
                        ["Item 3", 150, "Active"]
                    ],
                    "formats": {
                        "Value": "#,##0"
                    }
                }
            ]
        }

        os.makedirs("test_output", exist_ok=True)
        result = master.create_workbook(config)
        print(f"✅ Created test workbook: {result}")

        if os.path.exists(result):
            print(f"✅ File exists and is {os.path.getsize(result)} bytes")
            return True
        else:
            print(f"❌ File was not created")
            return False

    except Exception as e:
        print(f"❌ Excel test failed: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_example_configs():
    """Test that example configurations are valid JSON"""
    print("\n" + "="*60)
    print("Testing Example Configurations...")
    print("="*60)

    examples = [
        "professional-ppt-skill/examples/example_config.json",
        "professional-ppt-skill/examples/startup_pitch_config.json",
        "excel-master-skill/examples/financial_dashboard_config.json",
        "excel-master-skill/examples/sales_tracker_config.json"
    ]

    all_valid = True
    for example in examples:
        if os.path.exists(example):
            try:
                with open(example, 'r') as f:
                    config = json.load(f)
                print(f"✅ {os.path.basename(example)} is valid JSON")
            except json.JSONDecodeError as e:
                print(f"❌ {os.path.basename(example)} has invalid JSON: {e}")
                all_valid = False
        else:
            print(f"⚠️  {example} not found")

    return all_valid

def test_slide_types():
    """Test all PowerPoint slide types"""
    print("\n" + "="*60)
    print("Testing All PowerPoint Slide Types...")
    print("="*60)

    try:
        sys.path.insert(0, 'professional-ppt-skill/scripts')
        from ppt_creator_enhanced import EnhancedPPTCreator

        creator = EnhancedPPTCreator()

        slide_types = [
            {"type": "title", "title": "Title Slide", "subtitle": "Test"},
            {"type": "section", "title": "Section Divider"},
            {"type": "content", "title": "Content", "bullets": ["Point 1", "Point 2"]},
            {"type": "two_column", "title": "Two Column", "left_content": ["Left 1"], "right_content": ["Right 1"]},
            {"type": "comparison", "title": "Comparison", "left_title": "A", "right_title": "B"},
            {"type": "timeline", "title": "Timeline", "events": ["Q1", "Q2", "Q3"]},
            {"type": "blank"}
        ]

        config = {
            "theme": "corporate_blue",
            "output_path": "test_output/all_slide_types.pptx",
            "slides": slide_types
        }

        result = creator.create_presentation(config)
        print(f"✅ Created presentation with all {len(slide_types)} slide types")
        print(f"✅ Output: {result}")

        return os.path.exists(result)

    except Exception as e:
        print(f"❌ Slide types test failed: {e}")
        return False

def test_conditional_formatting():
    """Test Excel conditional formatting"""
    print("\n" + "="*60)
    print("Testing Excel Conditional Formatting...")
    print("="*60)

    try:
        sys.path.insert(0, 'excel-master-skill/scripts')
        from excel_master_enhanced import EnhancedExcelMaster

        master = EnhancedExcelMaster()

        config = {
            "theme": "financial_green",
            "output_path": "test_output/conditional_formatting_test.xlsx",
            "sheets": [
                {
                    "name": "CF_Test",
                    "type": "data",
                    "headers": ["Value1", "Value2", "Value3"],
                    "data": [
                        [10, 50, 100],
                        [20, 60, 90],
                        [30, 70, 80],
                        [40, 80, 70],
                        [50, 90, 60]
                    ],
                    "conditional_formatting": [
                        {
                            "type": "color_scale",
                            "range": "A2:A6"
                        },
                        {
                            "type": "data_bar",
                            "range": "B2:B6"
                        },
                        {
                            "type": "icon_set",
                            "range": "C2:C6",
                            "icon_style": "3Arrows"
                        }
                    ]
                }
            ]
        }

        result = master.create_workbook(config)
        print(f"✅ Created workbook with conditional formatting")
        print(f"   - Color scale on column A")
        print(f"   - Data bars on column B")
        print(f"   - Icon set on column C")
        print(f"✅ Output: {result}")

        return os.path.exists(result)

    except Exception as e:
        print(f"❌ Conditional formatting test failed: {e}")
        return False

def print_summary(results):
    """Print test summary"""
    print("\n" + "="*60)
    print("TEST SUMMARY")
    print("="*60)

    total = len(results)
    passed = sum(results.values())
    failed = total - passed

    for test_name, result in results.items():
        status = "✅ PASSED" if result else "❌ FAILED"
        print(f"{status} - {test_name}")

    print(f"\n{passed}/{total} tests passed ({failed} failed)")

    if failed == 0:
        print("\n🎉 All tests passed! The enhanced features are working correctly.")
    else:
        print(f"\n⚠️  {failed} test(s) failed. Please check the errors above.")

    return failed == 0

def main():
    """Run all tests"""
    print("="*60)
    print("Enhanced PPT & Excel Features - Test Suite")
    print("="*60)

    # Run tests
    results = {}

    results["Package Imports"] = test_imports()

    if results["Package Imports"]:
        results["PPT Enhanced Creator"] = test_ppt_enhanced()
        results["Excel Enhanced Master"] = test_excel_enhanced()
        results["Example Configurations"] = test_example_configs()
        results["PPT Slide Types"] = test_slide_types()
        results["Excel Conditional Formatting"] = test_conditional_formatting()
    else:
        print("\n❌ Package imports failed. Please install required packages:")
        print("   pip install python-pptx openpyxl pandas pillow")
        return False

    # Print summary
    all_passed = print_summary(results)

    # Cleanup
    print("\n" + "="*60)
    print("Test files created in: test_output/")
    print("="*60)

    return all_passed

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
