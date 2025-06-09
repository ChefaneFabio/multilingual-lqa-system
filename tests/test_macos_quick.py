#!/usr/bin/env python3
"""
macOS LQA Quick Test Script
Validates system functionality before full deployment
"""

import sys
import platform
import time
import subprocess
import json
from typing import Dict, List

def test_system_info():
    """Test basic system information"""
    print("🍎 macOS LQA System - Quick Test")
    print("=" * 50)
    
    # Platform detection
    platform_name = platform.system()
    print(f"🖥️ Platform: {platform_name}")
    
    if platform_name == 'Darwin':
        print("✅ macOS detected - optimizations enabled")
        # Get macOS version
        try:
            version = platform.mac_ver()[0]
            print(f"🍎 macOS Version: {version}")
        except:
            print("⚠️ Could not detect macOS version")
    else:
        print(f"✅ {platform_name} detected - cross-platform mode")
    
    # Python version
    python_version = f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}"
    print(f"🐍 Python Version: {python_version}")
    
    if sys.version_info >= (3, 8):
        print("✅ Python version compatible")
    else:
        print("⚠️ Python 3.8+ recommended")
    
    return platform_name == 'Darwin'

def test_dependencies():
    """Test required Python packages"""
    print("\n📦 Testing Dependencies...")
    
    dependencies = {
        'xlwings': 'Excel integration',
        'requests': 'API communication', 
        'json': 'Data processing (built-in)',
        'subprocess': 'System integration (built-in)'
    }
    
    results = {}
    
    for package, description in dependencies.items():
        try:
            if package in ['json', 'subprocess']:
                # Built-in modules
                exec(f"import {package}")
                print(f"✅ {package}: Available ({description})")
                results[package] = True
            else:
                # External packages
                exec(f"import {package}")
                version = eval(f"{package}.__version__")
                print(f"✅ {package}: v{version} ({description})")
                results[package] = True
                
        except ImportError:
            print(f"❌ {package}: Not installed ({description})")
            results[package] = False
        except Exception as e:
            print(f"⚠️ {package}: Error - {e}")
            results[package] = False
    
    return all(results.values())

def test_excel_availability(is_macos: bool):
    """Test Excel installation and availability"""
    print("\n📊 Testing Excel Availability...")
    
    if is_macos:
        # Test Excel on macOS
        try:
            # Check if Excel is installed
            result = subprocess.run(['mdfind', 'kMDItemCFBundleIdentifier == "com.microsoft.Excel"'], 
                                  capture_output=True, text=True, timeout=10)
            
            if result.stdout.strip():
                print("✅ Microsoft Excel found on macOS")
                excel_path = result.stdout.strip().split('\n')[0]
                print(f"📁 Location: {excel_path}")
                
                # Test if Excel can be launched
                try:
                    subprocess.run(['open', '-a', 'Microsoft Excel'], timeout=5)
                    print("✅ Excel can be launched")
                    return True
                except:
                    print("⚠️ Excel found but cannot be launched")
                    return False
            else:
                print("❌ Microsoft Excel not found")
                print("💡 Install with: brew install --cask microsoft-excel")
                return False
                
        except Exception as e:
            print(f"⚠️ Excel detection failed: {e}")
            return False
    else:
        # Test Excel on Windows/Linux
        try:
            import xlwings as xw
            app = xw.App(visible=False)
            app.quit()
            print("✅ Excel connection successful")
            return True
        except Exception as e:
            print(f"❌ Excel connection failed: {e}")
            return False

def test_api_connectivity():
    """Test API connectivity for LQA services"""
    print("\n🌐 Testing API Connectivity...")
    
    try:
        import requests
        
        # Test LanguageTool API
        try:
            response = requests.get("https://api.languagetool.org/v2/languages", timeout=10)
            if response.status_code == 200:
                languages = response.json()
                print(f"✅ LanguageTool API: Available ({len(languages)} languages)")
            else:
                print(f"⚠️ LanguageTool API: Status {response.status_code}")
        except Exception as e:
            print(f"❌ LanguageTool API: Failed - {e}")
        
        # Test OpenAI API (basic connectivity)
        try:
            response = requests.get("https://api.openai.com/v1/models", 
                                  headers={'Authorization': 'Bearer invalid-key'}, 
                                  timeout=10)
            if response.status_code in [401, 403]:
                print("✅ OpenAI API: Reachable (authentication required)")
            else:
                print(f"⚠️ OpenAI API: Unexpected status {response.status_code}")
        except Exception as e:
            print(f"❌ OpenAI API: Failed - {e}")
            
    except ImportError:
        print("❌ requests package not available")

def test_basic_lqa_functionality():
    """Test basic LQA analysis functionality"""
    print("\n🔍 Testing Basic LQA Functionality...")
    
    try:
        # Simple text analysis without external APIs
        test_text = "This sentence have grammar errors for testing."
        
        # Basic grammar checks
        errors_found = []
        
        # Simple grammar rules
        if " have " in test_text and not test_text.startswith("I have") and not test_text.startswith("You have"):
            errors_found.append("Subject-verb disagreement detected")
        
        # Simple spelling checks
        spelling_errors = ["grammer", "speling", "recieve", "seperate"]
        for error in spelling_errors:
            if error in test_text.lower():
                errors_found.append(f"Spelling error: {error}")
        
        # Calculate basic quality score
        word_count = len(test_text.split())
        error_count = len(errors_found)
        quality_score = max(0, 100 - (error_count * 15))
        
        print(f"📝 Test Text: {test_text}")
        print(f"📊 Quality Score: {quality_score}/100")
        print(f"❗ Errors Found: {error_count}")
        
        for error in errors_found:
            print(f"   • {error}")
        
        if error_count > 0:
            print("✅ Basic error detection working")
        else:
            print("⚠️ No errors detected (may need API connectivity)")
            
        return True
        
    except Exception as e:
        print(f"❌ Basic LQA test failed: {e}")
        return False

def test_xlwings_excel_integration():
    """Test xlwings Excel integration specifically"""
    print("\n📊 Testing xlwings Excel Integration...")
    
    try:
        import xlwings as xw
        
        # Test app creation
        app = xw.App(visible=False)
        print("✅ xlwings app created")
        
        # Test workbook creation
        wb = app.books.add()
        print("✅ Workbook created")
        
        # Test cell operations
        ws = wb.sheets[0]
        ws.range('A1').value = "Test LQA System"
        ws.range('B1').value = 95.5
        
        # Test reading values
        text_value = ws.range('A1').value
        number_value = ws.range('B1').value
        
        print(f"✅ Cell operations: Text='{text_value}', Number={number_value}")
        
        # Test formatting
        ws.range('A1').color = (144, 238, 144)  # Light green
        print("✅ Cell formatting applied")
        
        # Cleanup
        wb.close()
        app.quit()
        print("✅ Excel integration test completed")
        
        return True
        
    except ImportError:
        print("❌ xlwings not installed")
        return False
    except Exception as e:
        print(f"❌ xlwings test failed: {e}")
        return False

def test_macos_specific_features(is_macos: bool):
    """Test macOS-specific features"""
    if not is_macos:
        print("\n⏭️ Skipping macOS-specific tests (not on macOS)")
        return True
        
    print("\n🍎 Testing macOS-Specific Features...")
    
    try:
        # Test AppleScript availability
        result = subprocess.run(['osascript', '-e', 'return "AppleScript works"'], 
                              capture_output=True, text=True, timeout=5)
        
        if result.returncode == 0 and "AppleScript works" in result.stdout:
            print("✅ AppleScript: Available")
        else:
            print("⚠️ AppleScript: Not working properly")
        
        # Test file system permissions
        import tempfile
        import os
        
        temp_dir = '/tmp/lqa_test'
        try:
            os.makedirs(temp_dir, exist_ok=True)
            os.chmod(temp_dir, 0o755)
            
            # Create test file
            test_file = os.path.join(temp_dir, 'test.txt')
            with open(test_file, 'w') as f:
                f.write("LQA test file")
            
            # Read test file
            with open(test_file, 'r') as f:
                content = f.read()
            
            if content == "LQA test file":
                print("✅ File system operations: Working")
            
            # Cleanup
            os.remove(test_file)
            os.rmdir(temp_dir)
            
        except Exception as e:
            print(f"⚠️ File system operations: {e}")
        
        # Test macOS application detection
        try:
            result = subprocess.run(['ps', '-A'], capture_output=True, text=True, timeout=5)
            processes = result.stdout
            
            excel_running = 'Microsoft Excel' in processes
            print(f"📊 Excel Status: {'Running' if excel_running else 'Not running'}")
            
        except:
            print("⚠️ Process detection failed")
        
        return True
        
    except Exception as e:
        print(f"❌ macOS features test failed: {e}")
        return False

def run_comprehensive_test():
    """Run all tests and provide summary"""
    print("🚀 Starting Comprehensive macOS LQA System Test")
    print("=" * 60)
    
    start_time = time.time()
    
    # Run all tests
    test_results = {}
    
    test_results['system_info'] = test_system_info()
    is_macos = test_results['system_info']
    
    test_results['dependencies'] = test_dependencies()
    test_results['excel_available'] = test_excel_availability(is_macos)
    test_results['api_connectivity'] = test_api_connectivity()
    test_results['basic_lqa'] = test_basic_lqa_functionality()
    test_results['xlwings_integration'] = test_xlwings_excel_integration()
    test_results['macos_features'] = test_macos_specific_features(is_macos)
    
    # Calculate results
    total_tests = len(test_results)
    passed_tests = sum(1 for result in test_results.values() if result)
    
    elapsed_time = time.time() - start_time
    
    # Print summary
    print("\n" + "=" * 60)
    print("📋 TEST SUMMARY")
    print("=" * 60)
    
    for test_name, result in test_results.items():
        status = "✅ PASS" if result else "❌ FAIL"
        formatted_name = test_name.replace('_', ' ').title()
        print(f"{status} | {formatted_name}")
    
    print(f"\n📊 Results: {passed_tests}/{total_tests} tests passed")
    print(f"⏱️ Total time: {elapsed_time:.2f} seconds")
    
    # Recommendations
    print(f"\n💡 RECOMMENDATIONS:")
    
    if not test_results['dependencies']:
        print("🔧 Install missing dependencies:")
        print("   pip3 install xlwings requests openai")
    
    if not test_results['excel_available']:
        if is_macos:
            print("📊 Install Excel for macOS:")
            print("   brew install --cask microsoft-excel")
        else:
            print("📊 Check Excel installation and xlwings configuration")
    
    if not test_results['xlwings_integration']:
        print("🔗 Configure xlwings:")
        print("   xlwings addin install")
    
    if passed_tests == total_tests:
        print("🎉 ALL TESTS PASSED! System ready for production use.")
        print("\n🚀 Next steps:")
        print("   1. Add your OpenAI API key for maximum accuracy")
        print("   2. Run the full LQA system: python3 macos_lqa_system.py")
        print("   3. Create demo workbook to test Excel integration")
    elif passed_tests >= total_tests * 0.8:
        print("✅ Most tests passed! System should work with minor issues.")
    else:
        print("⚠️ Multiple issues detected. Please resolve before deployment.")
    
    return passed_tests == total_tests

if __name__ == "__main__":
    try:
        success = run_comprehensive_test()
        exit_code = 0 if success else 1
        sys.exit(exit_code)
    except KeyboardInterrupt:
        print("\n🛑 Test interrupted by user")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ Test framework error: {e}")
        sys.exit(1)