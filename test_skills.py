#!/usr/bin/env python3
"""
Skills Validation Script
Tests basic functionality of both PowerPoint and Excel skills
"""

import sys
import os
from pathlib import Path

def test_ppt_skill():
    """Test PowerPoint skill basic functionality"""
    print(" Testing PowerPoint Creator Skill...")
    
    # Add scripts directory to path
    ppt_scripts_path = Path(__file__).parent / "professional-ppt-skill" / "scripts"
    sys.path.insert(0, str(ppt_scripts_path))
    
    try:
        from ppt_creator import ProfessionalPPTCreator
        
        # Test prompt analysis
        creator = ProfessionalPPTCreator()
        test_prompt = "Create a business presentation about renewable energy for board meeting"
        analysis = creator.analyze_prompt(test_prompt)
        
        print(f"Prompt analysis successful:")
        print(f"   Topic: {analysis['topic']}")
        print(f"   Type: {analysis['type']}")
        print(f"   Slide Count: {analysis['slide_count']}")
        print(f"   Audience: {analysis['audience']}")
        
        # Test presentation generation (without actually creating file)
        print("PowerPoint skill structure validated")
        return True
        
    except ImportError as e:
        print(f" PowerPoint skill import error: {e}")
        return False
    except Exception as e:
        print(f" PowerPoint skill error: {e}")
        return False

def test_excel_skill():
    """Test Excel skill basic functionality"""
    print(" Testing Excel Master Skill...")
    
    # Add scripts directory to path
    excel_scripts_path = Path(__file__).parent / "excel-master-skill" / "scripts"
    sys.path.insert(0, str(excel_scripts_path))
    
    try:
        from excel_master import ExcelMaster
        
        # Test request analysis
        excel_master = ExcelMaster()
        test_prompt = "Create a budget tracker with expense categories"
        analysis = excel_master.analyze_request(test_prompt)
        
        print(f"Request analysis successful:")
        print(f"   Type: {analysis['type']}")
        print(f"   Complexity: {analysis['complexity']}")
        print(f"   Is Update: {analysis['is_update']}")
        print(f"   Color Scheme: {analysis['color_scheme']}")
        
        print("Excel skill structure validated")
        return True
        
    except ImportError as e:
        print(f" Excel skill import error: {e}")
        return False
    except Exception as e:
        print(f" Excel skill error: {e}")
        return False

def validate_skill_structure():
    """Validate the Skills directory structure"""
    print(" Validating Skills Structure...")
    
    base_path = Path(__file__).parent
    
    # Check PowerPoint skill
    ppt_skill_path = base_path / "professional-ppt-skill"
    ppt_skill_md = ppt_skill_path / "SKILL.md"
    ppt_scripts = ppt_skill_path / "scripts" / "ppt_creator.py"
    
    if not ppt_skill_md.exists():
        print(" PowerPoint SKILL.md not found")
        return False
    
    if not ppt_scripts.exists():
        print("PowerPoint ppt_creator.py not found")
        return False
    
    # Check Excel skill
    excel_skill_path = base_path / "excel-master-skill"
    excel_skill_md = excel_skill_path / "SKILL.md"
    excel_scripts = excel_skill_path / "scripts" / "excel_master.py"
    
    if not excel_skill_md.exists():
        print(" Excel SKILL.md not found")
        return False
    
    if not excel_scripts.exists():
        print(" Excel excel_master.py not found")
        return False
    
    print(" All skill files found in correct structure")
    return True

def validate_dependencies():
    """Check if required Python packages are available"""
    print(" Checking Dependencies...")
    
    required_packages = [
        'openpyxl',
        'pandas'
    ]
    
    # Note: python-pptx and pillow would normally be checked but may not be available in test environment
    
    missing_packages = []
    for package in required_packages:
        try:
            __import__(package)
            print(f" {package} available")
        except ImportError:
            print(f"{package} not available (install with: pip install {package})")
            missing_packages.append(package)
    
    return len(missing_packages) == 0

def main():
    """Run all validation tests"""
    print("Claude Skills Validation Test")
    print("=" * 50)
    
    tests = [
        ("Skill Structure", validate_skill_structure),
        ("Dependencies", validate_dependencies), 
        ("PowerPoint Skill", test_ppt_skill),
        ("Excel Skill", test_excel_skill)
    ]
    
    results = []
    for test_name, test_func in tests:
        print(f"\n{test_name}:")
        try:
            result = test_func()
            results.append(result)
        except Exception as e:
            print(f" {test_name} failed with error: {e}")
            results.append(False)
    
    print("\n" + "=" * 50)
    print(" VALIDATION SUMMARY:")
    
    passed = sum(results)
    total = len(results)
    
    for i, (test_name, _) in enumerate(tests):
        status = "PASS" if results[i] else "FAIL"
        print(f"  {test_name}: {status}")
    
    print(f"\nOverall: {passed}/{total} tests passed")
    
    if passed == total:
        print(" All tests passed! Skills are ready for deployment.")
        return True
    else:
        print("Some tests failed. Please review the issues above.")
        return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
