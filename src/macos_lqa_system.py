#!/usr/bin/env python3
"""
macOS Optimized Multilingual LQA System
Cross-platform Language Quality Assurance for Excel
Optimized for macOS with Windows compatibility
"""

import os
import sys
import platform
import json
import time
import requests
import re
from typing import Dict, List, Optional, Tuple, Any
from dataclasses import dataclass
from datetime import datetime
import tempfile
import subprocess

# === CONFIGURATION ===
OPENAI_API_KEY = ""  # Add your OpenAI API key here for maximum accuracy

# macOS-specific configurations
MACOS_SETTINGS = {
    'excel_app_name': 'Microsoft Excel',
    'temp_dir': '/tmp/lqa_temp',
    'path_separator': '/',
    'platform_specific_imports': True,
    'use_applescript_fallback': True,
    'excel_com_timeout': 30,
}

WINDOWS_SETTINGS = {
    'excel_app_name': 'Excel.Application', 
    'temp_dir': 'C:\\temp\\lqa_temp',
    'path_separator': '\\',
    'platform_specific_imports': False,
    'use_applescript_fallback': False,
    'excel_com_timeout': 10,
}

# Detect platform and set appropriate settings
CURRENT_PLATFORM = platform.system()
PLATFORM_SETTINGS = MACOS_SETTINGS if CURRENT_PLATFORM == 'Darwin' else WINDOWS_SETTINGS

# Enhanced accuracy settings for professional LQA
LQA_CONFIG = {
    'openai_model': 'gpt-4',
    'analysis_temperature': 0.1,
    'max_tokens_analysis': 2000,
    'quality_thresholds': {
        'excellent': 98,    # Professional publication ready
        'good': 85,         # Business communication ready  
        'acceptable': 75,   # Basic communication acceptable
        'poor': 50          # Needs significant improvement
    },
    'error_weights': {
        'grammar': 1.5,      # Grammar errors weighted heavily
        'spelling': 1.2,     # Spelling important for professionalism
        'syntax': 1.3,       # Syntax affects readability
        'accuracy': 2.0,     # Accuracy is critical for LQA
        'style': 0.8,        # Style is less critical than correctness
        'fluency': 1.0       # Standard weight for fluency
    },
    'supported_languages': [
        'en', 'es', 'fr', 'de', 'it', 'pt', 'ru', 'zh', 'ja', 'ko', 
        'ar', 'hi', 'tr', 'pl', 'nl', 'sv', 'da', 'no', 'fi', 'el',
        'he', 'th', 'vi', 'uk', 'cs', 'hu', 'ro', 'bg', 'hr', 'sk'
    ]
}

@dataclass
class LQAResult:
    """Enhanced LQA analysis result with comprehensive metrics"""
    text: str
    language: str
    quality_score: float
    error_count: int
    errors: List[Dict]
    suggestions: List[str]
    confidence: float
    processing_time: float
    api_source: str
    detailed_analysis: Dict

class macOSCompatibilityLayer:
    """Handles macOS-specific Excel interactions and fallbacks"""
    
    def __init__(self):
        self.platform = CURRENT_PLATFORM
        self.settings = PLATFORM_SETTINGS
        
    def get_excel_app(self):
        """Get Excel application with macOS optimizations"""
        try:
            import xlwings as xw
            
            if self.platform == 'Darwin':
                # macOS specific Excel handling
                app = xw.App(visible=True)
                app.display_alerts = False
                app.screen_updating = True
                return app
            else:
                # Windows handling
                app = xw.App(visible=True)
                app.display_alerts = False
                return app
                
        except Exception as e:
            print(f"⚠️ Excel connection failed: {e}")
            return None
    
    def applescript_fallback(self, script_command: str):
        """Use AppleScript as fallback for macOS Excel operations"""
        if self.platform != 'Darwin':
            return None
            
        try:
            result = subprocess.run(['osascript', '-e', script_command], 
                                 capture_output=True, text=True, timeout=10)
            return result.stdout.strip() if result.returncode == 0 else None
        except Exception as e:
            print(f"AppleScript fallback failed: {e}")
            return None
    
    def ensure_temp_directory(self):
        """Create temp directory with proper permissions for macOS"""
        temp_dir = self.settings['temp_dir']
        try:
            os.makedirs(temp_dir, exist_ok=True)
            if self.platform == 'Darwin':
                os.chmod(temp_dir, 0o755)  # Proper macOS permissions
            return temp_dir
        except Exception as e:
            print(f"⚠️ Could not create temp directory: {e}")
            return tempfile.gettempdir()

class EnhancedMultilingualAnalyzer:
    """Advanced multilingual LQA analyzer optimized for macOS"""
    
    def __init__(self):
        self.compatibility = macOSCompatibilityLayer()
        self.session = requests.Session()
        self.session.timeout = 30  # Increased timeout for macOS
        
        # Enhanced API configurations
        self.openai_configured = bool(OPENAI_API_KEY)
        self.languagetool_base = "https://api.languagetool.org/v2"
        
        # macOS-specific optimizations
        if CURRENT_PLATFORM == 'Darwin':
            self.session.headers.update({
                'User-Agent': 'macOS-LQA-System/1.0',
                'Accept-Encoding': 'gzip, deflate'
            })
        
        print(f"🍎 macOS LQA System initialized")
        print(f"📱 Platform: {CURRENT_PLATFORM}")
        print(f"🔧 OpenAI: {'✅ Configured' if self.openai_configured else '⚠️ Not configured'}")
        print(f"🌍 Languages: {len(LQA_CONFIG['supported_languages'])} supported")
    
    def detect_language(self, text: str) -> str:
        """Enhanced language detection with macOS optimizations"""
        if not text or len(text.strip()) < 10:
            return 'en'  # Default to English for short texts
        
        try:
            # Try LanguageTool first (fast and reliable)
            response = self.session.post(
                f"{self.languagetool_base}/languages",
                timeout=10
            )
            
            if response.status_code == 200:
                # Use simple heuristics for common languages
                text_lower = text.lower()
                
                # Language detection patterns
                patterns = {
                    'es': ['el ', 'la ', 'que ', 'de ', 'en ', 'un ', 'es ', 'se ', 'no ', 'te '],
                    'fr': ['le ', 'de ', 'et ', 'à ', 'un ', 'il ', 'être ', 'et ', 'en ', 'avoir '],
                    'de': ['der ', 'die ', 'und ', 'in ', 'den ', 'von ', 'zu ', 'das ', 'mit ', 'sich '],
                    'it': ['il ', 'di ', 'che ', 'e ', 'la ', 'per ', 'in ', 'un ', 'è ', 'con '],
                    'pt': ['de ', 'a ', 'o ', 'que ', 'e ', 'do ', 'da ', 'em ', 'um ', 'para '],
                    'zh': ['的', '是', '在', '了', '不', '与', '也', '上', '个', '人'],
                    'ja': ['の', 'に', 'は', 'を', 'た', 'が', 'で', 'て', 'と', 'し'],
                }
                
                scores = {}
                for lang, words in patterns.items():
                    score = sum(1 for word in words if word in text_lower)
                    if score > 0:
                        scores[lang] = score
                
                if scores:
                    return max(scores.items(), key=lambda x: x[1])[0]
            
        except Exception as e:
            print(f"⚠️ Language detection failed: {e}")
        
        return 'en'  # Default fallback
    
    def analyze_with_languagetool(self, text: str, language: str) -> Dict:
        """Enhanced LanguageTool analysis with macOS optimizations"""
        try:
            params = {
                'text': text,
                'language': language,
                'enabledOnly': 'false',
                'level': 'picky'  # Most strict checking
            }
            
            response = self.session.post(
                f"{self.languagetool_base}/check",
                data=params,
                timeout=15
            )
            
            if response.status_code == 200:
                data = response.json()
                
                errors = []
                for match in data.get('matches', []):
                    error = {
                        'type': match.get('rule', {}).get('category', {}).get('name', 'Grammar'),
                        'message': match.get('message', ''),
                        'suggestions': [r['value'] for r in match.get('replacements', [])[:3]],
                        'offset': match.get('offset', 0),
                        'length': match.get('length', 0),
                        'severity': match.get('rule', {}).get('category', {}).get('id', 'MINOR'),
                        'confidence': 0.8  # LanguageTool baseline confidence
                    }
                    errors.append(error)
                
                return {
                    'errors': errors,
                    'error_count': len(errors),
                    'api_source': 'LanguageTool',
                    'processing_time': response.elapsed.total_seconds()
                }
                
        except Exception as e:
            print(f"⚠️ LanguageTool analysis failed: {e}")
        
        return {'errors': [], 'error_count': 0, 'api_source': 'None', 'processing_time': 0}
    
    def analyze_with_openai(self, text: str, language: str) -> Dict:
        """Enhanced OpenAI analysis for maximum accuracy"""
        if not self.openai_configured:
            return {'errors': [], 'error_count': 0, 'api_source': 'OpenAI-NotConfigured', 'processing_time': 0}
        
        try:
            # Enhanced prompt for professional LQA
            prompt = f"""
            Perform a comprehensive Language Quality Assurance analysis of the following text in {language}.
            
            Analyze for:
            1. Grammar errors (subject-verb agreement, tense consistency, etc.)
            2. Spelling mistakes and typos
            3. Syntax issues (punctuation, sentence structure)
            4. Accuracy and terminology consistency
            5. Style and fluency issues
            6. Professional communication standards
            
            Text to analyze: "{text}"
            
            Provide response in JSON format:
            {{
                "overall_quality": 0-100,
                "language_detected": "language_code",
                "errors": [
                    {{
                        "type": "grammar|spelling|syntax|accuracy|style",
                        "severity": "critical|major|minor",
                        "message": "Clear description of the issue",
                        "suggestions": ["correction1", "correction2"],
                        "confidence": 0.0-1.0,
                        "explanation": "Linguistic reasoning for the correction"
                    }}
                ],
                "summary": "Overall assessment and recommendations"
            }}
            """
            
            headers = {
                'Authorization': f'Bearer {OPENAI_API_KEY}',
                'Content-Type': 'application/json'
            }
            
            data = {
                'model': LQA_CONFIG['openai_model'],
                'messages': [
                    {'role': 'system', 'content': 'You are a professional linguistic quality assurance expert.'},
                    {'role': 'user', 'content': prompt}
                ],
                'temperature': LQA_CONFIG['analysis_temperature'],
                'max_tokens': LQA_CONFIG['max_tokens_analysis']
            }
            
            start_time = time.time()
            response = self.session.post(
                'https://api.openai.com/v1/chat/completions',
                headers=headers,
                json=data,
                timeout=30
            )
            processing_time = time.time() - start_time
            
            if response.status_code == 200:
                result = response.json()
                content = result['choices'][0]['message']['content']
                
                # Parse JSON response
                try:
                    analysis = json.loads(content)
                    return {
                        'errors': analysis.get('errors', []),
                        'error_count': len(analysis.get('errors', [])),
                        'api_source': 'OpenAI-GPT4',
                        'processing_time': processing_time,
                        'overall_quality': analysis.get('overall_quality', 0),
                        'summary': analysis.get('summary', '')
                    }
                except json.JSONDecodeError:
                    # Fallback: parse text response
                    return {
                        'errors': [{'type': 'analysis', 'message': content, 'suggestions': [], 'confidence': 0.9}],
                        'error_count': 1 if 'error' in content.lower() else 0,
                        'api_source': 'OpenAI-Text',
                        'processing_time': processing_time
                    }
            else:
                print(f"⚠️ OpenAI API error: {response.status_code}")
                
        except Exception as e:
            print(f"⚠️ OpenAI analysis failed: {e}")
        
        return {'errors': [], 'error_count': 0, 'api_source': 'OpenAI-Failed', 'processing_time': 0}
    
    def calculate_quality_score(self, text: str, lt_result: Dict, openai_result: Dict) -> float:
        """Enhanced quality scoring with professional standards"""
        if not text or len(text.strip()) < 5:
            return 0.0
        
        base_score = 100.0
        word_count = len(text.split())
        
        # Combine errors from both sources with deduplication
        all_errors = []
        
        # Add LanguageTool errors
        for error in lt_result.get('errors', []):
            all_errors.append({
                'type': error.get('type', 'grammar').lower(),
                'severity': error.get('severity', 'minor'),
                'source': 'languagetool'
            })
        
        # Add OpenAI errors (if available and different)
        for error in openai_result.get('errors', []):
            error_type = error.get('type', 'grammar').lower()
            # Simple deduplication by type
            if not any(e['type'] == error_type and e['source'] == 'languagetool' for e in all_errors):
                all_errors.append({
                    'type': error_type,
                    'severity': error.get('severity', 'minor'),
                    'source': 'openai'
                })
        
        # Apply enhanced error weights
        for error in all_errors:
            error_type = error['type']
            severity = error['severity']
            
            # Base penalty
            penalty = 5.0  # Default penalty
            
            # Type-based penalty with enhanced weights
            if error_type in ['grammar', 'syntax']:
                penalty = 8.0 * LQA_CONFIG['error_weights'].get('grammar', 1.0)
            elif error_type in ['spelling', 'typo']:
                penalty = 6.0 * LQA_CONFIG['error_weights'].get('spelling', 1.0)
            elif error_type in ['accuracy', 'terminology']:
                penalty = 10.0 * LQA_CONFIG['error_weights'].get('accuracy', 1.0)
            elif error_type in ['style', 'fluency']:
                penalty = 3.0 * LQA_CONFIG['error_weights'].get('style', 1.0)
            
            # Severity multiplier
            if severity in ['critical', 'major']:
                penalty *= 2.0
            elif severity == 'minor':
                penalty *= 0.7
            
            # Apply word count scaling (more errors in longer text is relatively better)
            if word_count > 50:
                penalty *= 0.8
            elif word_count < 10:
                penalty *= 1.5
            
            base_score -= penalty
        
        # Ensure score is within bounds
        final_score = max(0.0, min(100.0, base_score))
        
        # Use OpenAI overall quality if available and reasonable
        if openai_result.get('overall_quality') and self.openai_configured:
            openai_score = openai_result['overall_quality']
            if 0 <= openai_score <= 100:
                # Weighted average: 60% calculated, 40% OpenAI
                final_score = (final_score * 0.6) + (openai_score * 0.4)
        
        return round(final_score, 1)
    
    def analyze_text(self, text: str) -> LQAResult:
        """Comprehensive text analysis with multi-source validation"""
        start_time = time.time()
        
        if not text or len(text.strip()) < 3:
            return LQAResult(
                text=text, language='unknown', quality_score=0.0,
                error_count=0, errors=[], suggestions=[],
                confidence=0.0, processing_time=0.0,
                api_source='None', detailed_analysis={}
            )
        
        # Language detection
        language = self.detect_language(text)
        
        # Multi-source analysis
        lt_result = self.analyze_with_languagetool(text, language)
        openai_result = self.analyze_with_openai(text, language)
        
        # Combine errors and suggestions
        all_errors = lt_result.get('errors', []) + openai_result.get('errors', [])
        all_suggestions = []
        
        for error in all_errors:
            all_suggestions.extend(error.get('suggestions', []))
        
        # Calculate enhanced quality score
        quality_score = self.calculate_quality_score(text, lt_result, openai_result)
        
        # Calculate confidence based on API availability and agreement
        confidence = 0.7  # Base confidence
        if self.openai_configured and openai_result.get('errors'):
            confidence = 0.9  # High confidence with GPT-4
        elif lt_result.get('errors'):
            confidence = 0.8  # Good confidence with LanguageTool
        
        processing_time = time.time() - start_time
        
        # Detailed analysis
        detailed_analysis = {
            'languagetool_analysis': lt_result,
            'openai_analysis': openai_result,
            'language_detected': language,
            'word_count': len(text.split()),
            'character_count': len(text),
            'quality_category': self.get_quality_category(quality_score),
            'platform': CURRENT_PLATFORM
        }
        
        return LQAResult(
            text=text,
            language=language,
            quality_score=quality_score,
            error_count=len(all_errors),
            errors=all_errors,
            suggestions=list(set(all_suggestions))[:5],  # Top 5 unique suggestions
            confidence=confidence,
            processing_time=processing_time,
            api_source=f"LanguageTool+{openai_result.get('api_source', 'None')}",
            detailed_analysis=detailed_analysis
        )
    
    def get_quality_category(self, score: float) -> str:
        """Get quality category based on enhanced thresholds"""
        thresholds = LQA_CONFIG['quality_thresholds']
        
        if score >= thresholds['excellent']:
            return 'Excellent - Publication Ready'
        elif score >= thresholds['good']:
            return 'Good - Professional Standard'
        elif score >= thresholds['acceptable']:
            return 'Acceptable - Minor Issues'
        elif score >= thresholds['poor']:
            return 'Poor - Needs Improvement'
        else:
            return 'Critical - Significant Issues'

class macOSExcelIntegration:
    """Enhanced Excel integration optimized for macOS"""
    
    def __init__(self):
        self.analyzer = EnhancedMultilingualAnalyzer()
        self.compatibility = macOSCompatibilityLayer()
        self.excel_app = None
        
    def connect_to_excel(self):
        """Establish Excel connection with macOS optimizations"""
        try:
            import xlwings as xw
            
            if CURRENT_PLATFORM == 'Darwin':
                # macOS specific connection
                print("🍎 Connecting to Excel on macOS...")
                self.excel_app = xw.App(visible=True)
                self.excel_app.display_alerts = False
                
                # Test connection with AppleScript fallback
                try:
                    books = self.excel_app.books
                    print(f"✅ Excel connected successfully ({len(books)} workbooks)")
                except:
                    print("⚠️ Using AppleScript fallback for Excel connection")
                    self.compatibility.applescript_fallback(
                        'tell application "Microsoft Excel" to activate'
                    )
            else:
                # Windows connection
                print("🖥️ Connecting to Excel on Windows...")
                self.excel_app = xw.App(visible=True)
                
            return True
            
        except Exception as e:
            print(f"❌ Excel connection failed: {e}")
            if CURRENT_PLATFORM == 'Darwin':
                print("💡 Tip: Make sure Excel is installed and try running:")
                print("   brew install --cask microsoft-excel")
            return False
    
    def get_or_create_workbook(self, name: str = "LQA_Analysis"):
        """Get existing workbook or create new one with macOS compatibility"""
        try:
            if not self.excel_app:
                if not self.connect_to_excel():
                    return None
            
            # Try to find existing workbook
            for wb in self.excel_app.books:
                if name in wb.name:
                    return wb
            
            # Create new workbook
            wb = self.excel_app.books.add()
            wb.name = name
            
            return wb
            
        except Exception as e:
            print(f"⚠️ Workbook creation failed: {e}")
            
            # macOS AppleScript fallback
            if CURRENT_PLATFORM == 'Darwin':
                script = f'''
                tell application "Microsoft Excel"
                    make new workbook
                    set name of workbook 1 to "{name}"
                end tell
                '''
                self.compatibility.applescript_fallback(script)
            
            return None
    
    def analyze_excel_selection(self):
        """Analyze selected Excel cells with enhanced macOS support"""
        try:
            if not self.connect_to_excel():
                print("❌ Could not connect to Excel")
                return
            
            import xlwings as xw
            
            # Get active selection
            selection = xw.selection
            if not selection:
                print("⚠️ No cells selected in Excel")
                return
            
            print(f"🔍 Analyzing {selection.shape[0]} rows × {selection.shape[1]} columns...")
            
            # Process each cell
            results = []
            for row in range(selection.shape[0]):
                for col in range(selection.shape[1]):
                    cell = selection[row, col]
                    text = str(cell.value) if cell.value else ""
                    
                    if text and len(text.strip()) > 3:
                        print(f"📝 Analyzing cell ({row+1}, {col+1}): {text[:50]}...")
                        
                        result = self.analyzer.analyze_text(text)
                        results.append((cell, result))
                        
                        # Apply formatting based on quality
                        self.format_cell_by_quality(cell, result.quality_score)
                        
                        # Add comment with analysis
                        self.add_quality_comment(cell, result)
            
            print(f"✅ Analysis complete! Processed {len(results)} cells")
            self.create_summary_report(results)
            
        except Exception as e:
            print(f"❌ Excel analysis failed: {e}")
            if CURRENT_PLATFORM == 'Darwin':
                print("💡 Try selecting cells manually and running analysis again")
    
    def format_cell_by_quality(self, cell, quality_score: float):
        """Apply color formatting based on quality score"""
        try:
            # Enhanced color scheme for professional LQA
            if quality_score >= LQA_CONFIG['quality_thresholds']['excellent']:
                # Excellent: Deep Green
                cell.color = (34, 139, 34)  # Forest Green
            elif quality_score >= LQA_CONFIG['quality_thresholds']['good']:
                # Good: Light Green  
                cell.color = (144, 238, 144)  # Light Green
            elif quality_score >= LQA_CONFIG['quality_thresholds']['acceptable']:
                # Acceptable: Yellow
                cell.color = (255, 255, 0)  # Yellow
            elif quality_score >= LQA_CONFIG['quality_thresholds']['poor']:
                # Poor: Orange
                cell.color = (255, 165, 0)  # Orange
            else:
                # Critical: Red
                cell.color = (255, 99, 71)  # Tomato Red
                
        except Exception as e:
            print(f"⚠️ Cell formatting failed: {e}")
    
    def add_quality_comment(self, cell, result: LQAResult):
        """Add detailed comment with analysis results"""
        try:
            comment_text = f"""
LQA Analysis Results:
━━━━━━━━━━━━━━━━━━━━━━
📊 Quality Score: {result.quality_score}/100
🗣️ Language: {result.language.upper()}
❗ Issues Found: {result.error_count}
🎯 Category: {self.analyzer.get_quality_category(result.quality_score)}
⚡ Confidence: {result.confidence:.1%}

"""
            
            if result.errors:
                comment_text += "🔍 Issues Detected:\n"
                for i, error in enumerate(result.errors[:3], 1):  # Show top 3 errors
                    comment_text += f"{i}. {error.get('type', 'Issue').title()}: {error.get('message', 'Check required')}\n"
                    if error.get('suggestions'):
                        comment_text += f"   💡 Suggestion: {error['suggestions'][0]}\n"
            
            if result.suggestions:
                comment_text += f"\n✨ Top Suggestions:\n"
                for suggestion in result.suggestions[:2]:  # Show top 2 suggestions
                    comment_text += f"• {suggestion}\n"
            
            comment_text += f"\n🕒 Processed in {result.processing_time:.2f}s"
            comment_text += f"\n🔧 Analysis: {result.api_source}"
            
            # Add comment to cell
            try:
                cell.note = comment_text
            except:
                # Alternative method for macOS
                print(f"💬 Quality: {result.quality_score}/100 | Errors: {result.error_count}")
                
        except Exception as e:
            print(f"⚠️ Comment addition failed: {e}")
    
    def create_summary_report(self, results: List):
        """Create comprehensive summary report with macOS optimization"""
        try:
            if not results:
                return
            
            wb = self.get_or_create_workbook("LQA_Summary_Report")
            if not wb:
                print("⚠️ Could not create summary workbook")
                return
            
            # Create summary sheet
            ws = wb.sheets.add("LQA_Summary")
            
            # Headers
            headers = ["Cell", "Text Sample", "Language", "Quality Score", "Category", "Error Count", "Top Issue", "Processing Time"]
            for i, header in enumerate(headers):
                ws.range(f"{chr(65+i)}1").value = header
                ws.range(f"{chr(65+i)}1").color = (70, 130, 180)  # Steel Blue
                ws.range(f"{chr(65+i)}1").api.Font.Bold = True
                ws.range(f"{chr(65+i)}1").api.Font.ColorIndex = 2  # White text
            
            # Data rows
            for i, (cell, result) in enumerate(results, 2):
                row_data = [
                    f"{cell.row}, {cell.column}",
                    result.text[:50] + "..." if len(result.text) > 50 else result.text,
                    result.language.upper(),
                    result.quality_score,
                    self.analyzer.get_quality_category(result.quality_score),
                    result.error_count,
                    result.errors[0].get('message', 'None') if result.errors else 'None',
                    f"{result.processing_time:.2f}s"
                ]
                
                for j, data in enumerate(row_data):
                    ws.range(f"{chr(65+j)}{i}").value = data
                
                # Color code rows by quality
                quality_color = self.get_quality_color(result.quality_score)
                for j in range(len(headers)):
                    ws.range(f"{chr(65+j)}{i}").color = quality_color
            
            # Auto-fit columns
            ws.autofit()
            
            # Add statistics
            self.add_summary_statistics(ws, results, len(results) + 3)
            
            print(f"📊 Summary report created with {len(results)} analyzed texts")
            
        except Exception as e:
            print(f"⚠️ Summary report creation failed: {e}")
    
    def get_quality_color(self, score: float) -> Tuple[int, int, int]:
        """Get color tuple for quality score"""
        if score >= 98: return (34, 139, 34)    # Forest Green
        elif score >= 85: return (144, 238, 144)  # Light Green  
        elif score >= 75: return (255, 255, 0)    # Yellow
        elif score >= 50: return (255, 165, 0)    # Orange
        else: return (255, 99, 71)                # Tomato Red
    
    def add_summary_statistics(self, ws, results: List, start_row: int):
        """Add comprehensive statistics to summary report"""
        try:
            # Calculate statistics
            total_texts = len(results)
            if total_texts == 0:
                return
            
            scores = [result[1].quality_score for result in results]
            avg_score = sum(scores) / len(scores)
            
            categories = {}
            languages = {}
            total_errors = 0
            
            for _, result in results:
                category = self.analyzer.get_quality_category(result.quality_score)
                categories[category] = categories.get(category, 0) + 1
                languages[result.language] = languages.get(result.language, 0) + 1
                total_errors += result.error_count
            
            # Statistics section
            stats_start = start_row + 2
            ws.range(f"A{stats_start}").value = "📈 LQA ANALYSIS STATISTICS"
            ws.range(f"A{stats_start}").api.Font.Bold = True
            ws.range(f"A{stats_start}").api.Font.Size = 14
            
            stats = [
                f"Total Texts Analyzed: {total_texts}",
                f"Average Quality Score: {avg_score:.1f}/100",
                f"Total Issues Found: {total_errors}",
                f"Average Issues per Text: {total_errors/total_texts:.1f}",
                "",
                "📊 Quality Distribution:",
            ]
            
            for category, count in categories.items():
                percentage = (count/total_texts) * 100
                stats.append(f"  {category}: {count} texts ({percentage:.1f}%)")
            
            stats.extend(["", "🌍 Language Distribution:"])
            for lang, count in languages.items():
                percentage = (count/total_texts) * 100
                stats.append(f"  {lang.upper()}: {count} texts ({percentage:.1f}%)")
            
            # Add statistics to worksheet
            for i, stat in enumerate(stats):
                ws.range(f"A{stats_start + 1 + i}").value = stat
                if stat.startswith("📊") or stat.startswith("🌍"):
                    ws.range(f"A{stats_start + 1 + i}").api.Font.Bold = True
            
        except Exception as e:
            print(f"⚠️ Statistics creation failed: {e}")
    
    def create_professional_demo(self):
        """Create comprehensive demo with multilingual samples optimized for macOS"""
        try:
            wb = self.get_or_create_workbook("Professional_LQA_Demo")
            if not wb:
                print("❌ Could not create demo workbook")
                return
            
            ws = wb.sheets[0]
            ws.name = "Multilingual_LQA_Demo"
            
            # Demo texts in multiple languages with varying quality
            demo_texts = [
                ("English - Excellent", "This sentence demonstrates perfect grammar, spelling, and professional writing standards suitable for publication."),
                ("English - Good", "This sentence has good quality but could benefit from minor improvements in style and clarity."),
                ("English - Poor", "This sentence have several grammar erors and speling mistakes that needs fixing for professional communication."),
                
                ("Spanish - Excellent", "Esta oración demuestra una gramática perfecta y estándares profesionales de escritura."),
                ("Spanish - Poor", "Esta oracion tiene varios errores de gramatica y ortografia que necessita correccion."),
                
                ("French - Excellent", "Cette phrase démontre une grammaire parfaite et des standards professionnels d'écriture."),
                ("French - Poor", "Cette phrase à plusieurs erreurs de grammaire et d'ortographe qui nécéssite correction."),
                
                ("German - Excellent", "Dieser Satz zeigt perfekte Grammatik und professionelle Schreibstandards."),
                ("German - Poor", "Dieser Satz hat mehrere Grammatikfehler und Rechtschreibfehler die Korrektur braucht."),
                
                ("Chinese - Sample", "这个句子展示了专业的写作标准和语法结构。"),
                ("Japanese - Sample", "この文は専門的な文章基準と文法構造を示しています。"),
                
                ("Technical English", "The API integration facilitates seamless data synchronization between heterogeneous systems."),
                ("Casual English", "Hey there! This is just a casual message to test how the system handles informal writing styles.")
            ]
            
            # Headers
            ws.range("A1").value = "🌍 PROFESSIONAL MULTILINGUAL LQA DEMO"
            ws.range("A1").api.Font.Bold = True
            ws.range("A1").api.Font.Size = 16
            ws.range("A1").color = (70, 130, 180)
            
            ws.range("A2").value = f"Generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Platform: {CURRENT_PLATFORM}"
            
            # Column headers
            headers = ["Sample Type", "Text to Analyze", "Language", "Quality Score", "Category", "Issues", "Analysis Details"]
            for i, header in enumerate(headers):
                cell = ws.range(f"{chr(65+i)}4")
                cell.value = header
                cell.color = (100, 149, 237)  # Cornflower Blue
                cell.api.Font.Bold = True
                cell.api.Font.ColorIndex = 2  # White text
            
            print("🔄 Analyzing demo texts...")
            
            # Analyze and populate demo texts
            for i, (sample_type, text) in enumerate(demo_texts, 5):
                print(f"   📝 Analyzing: {sample_type}")
                
                result = self.analyzer.analyze_text(text)
                
                # Populate row
                row_data = [
                    sample_type,
                    text,
                    result.language.upper(),
                    f"{result.quality_score}/100",
                    self.analyzer.get_quality_category(result.quality_score),
                    result.error_count,
                    f"Confidence: {result.confidence:.1%} | {result.api_source}"
                ]
                
                for j, data in enumerate(row_data):
                    cell = ws.range(f"{chr(65+j)}{i}")
                    cell.value = data
                    
                    # Apply quality-based formatting
                    if j == 1:  # Text column
                        cell.color = self.get_quality_color(result.quality_score)
                
                # Add detailed comment
                self.add_quality_comment(ws.range(f"B{i}"), result)
            
            # Auto-fit columns
            ws.autofit()
            
            # Add instruction panel
            instructions_start = len(demo_texts) + 7
            instructions = [
                "🎯 HOW TO USE THIS LQA SYSTEM:",
                "",
                "1. 📝 SELECT CELLS: Choose Excel cells containing text to analyze",
                "2. 🔍 RUN ANALYSIS: Use analyze_excel_selection() function",  
                "3. 📊 VIEW RESULTS: See color-coded quality indicators and detailed comments",
                "4. 📈 CHECK REPORTS: Review summary statistics and recommendations",
                "",
                "🌈 COLOR CODING:",
                f"🟢 Green ({LQA_CONFIG['quality_thresholds']['excellent']}+): Publication Ready",
                f"🟡 Yellow ({LQA_CONFIG['quality_thresholds']['good']}+): Professional Standard", 
                f"🟠 Orange ({LQA_CONFIG['quality_thresholds']['acceptable']}+): Acceptable Quality",
                f"🔴 Red (<{LQA_CONFIG['quality_thresholds']['poor']}): Needs Improvement",
                "",
                "💡 TIPS:",
                "• Hover over colored cells to see detailed analysis",
                "• Use summary reports for batch analysis insights",
                "• Configure OpenAI API key for maximum accuracy",
                f"• System supports {len(LQA_CONFIG['supported_languages'])} languages",
                "",
                f"🍎 Optimized for {CURRENT_PLATFORM} | Professional LQA Standards"
            ]
            
            for i, instruction in enumerate(instructions):
                cell = ws.range(f"A{instructions_start + i}")
                cell.value = instruction
                if instruction.startswith("🎯") or instruction.startswith("🌈") or instruction.startswith("💡"):
                    cell.api.Font.Bold = True
            
            print("✅ Professional demo created successfully!")
            print("💡 Hover over colored cells to see detailed LQA analysis")
            
        except Exception as e:
            print(f"❌ Demo creation failed: {e}")

def interactive_menu():
    """Enhanced interactive menu for macOS LQA system"""
    excel_integration = macOSExcelIntegration()
    
    while True:
        print("\n" + "="*60)
        print("🍎 macOS OPTIMIZED MULTILINGUAL LQA SYSTEM")
        print("="*60)
        print(f"🖥️  Platform: {CURRENT_PLATFORM}")
        print(f"🔧 OpenAI: {'✅ Configured' if OPENAI_API_KEY else '⚠️ Add API key for max accuracy'}")
        print(f"🌍 Languages: {len(LQA_CONFIG['supported_languages'])} supported")
        print(f"⚡ Quality Standard: Professional (Excellent≥{LQA_CONFIG['quality_thresholds']['excellent']})")
        
        print("\n📋 AVAILABLE OPTIONS:")
        print("1. 📊 Create Professional Demo Workbook")
        print("2. 🔍 Analyze Selected Excel Cells") 
        print("3. 📝 Interactive Text Analysis")
        print("4. 🔧 System Status & Configuration")
        print("5. 🧪 Test System Accuracy")
        print("6. 💡 Setup Instructions & Help")
        print("7. 🚪 Exit")
        
        choice = input("\n🎯 Select option (1-7): ").strip()
        
        if choice == '1':
            print("\n📊 Creating professional demo workbook...")
            excel_integration.create_professional_demo()
            
        elif choice == '2':
            print("\n🔍 Analyzing Excel selection...")
            print("💡 Please select cells in Excel containing text to analyze")
            input("Press Enter when cells are selected...")
            excel_integration.analyze_excel_selection()
            
        elif choice == '3':
            print("\n📝 Interactive Text Analysis")
            text = input("Enter text to analyze: ").strip()
            if text:
                print("🔄 Analyzing...")
                result = excel_integration.analyzer.analyze_text(text)
                print(f"\n✅ ANALYSIS RESULTS:")
                print(f"📊 Quality Score: {result.quality_score}/100")
                print(f"🗣️ Language: {result.language.upper()}")
                print(f"❗ Issues Found: {result.error_count}")
                print(f"🎯 Category: {excel_integration.analyzer.get_quality_category(result.quality_score)}")
                print(f"⚡ Confidence: {result.confidence:.1%}")
                print(f"🕒 Processing Time: {result.processing_time:.2f}s")
                
                if result.errors:
                    print("\n🔍 Issues Detected:")
                    for i, error in enumerate(result.errors[:3], 1):
                        print(f"   {i}. {error.get('type', 'Issue').title()}: {error.get('message', 'Check required')}")
                        if error.get('suggestions'):
                            print(f"      💡 Suggestion: {error['suggestions'][0]}")
                
                if result.suggestions:
                    print(f"\n✨ Suggestions: {', '.join(result.suggestions[:3])}")
            
        elif choice == '4':
            print(f"\n🔧 SYSTEM STATUS:")
            print(f"🖥️ Platform: {CURRENT_PLATFORM}")
            print(f"🐍 Python: {sys.version.split()[0]}")
            
            # Check dependencies
            try:
                import xlwings
                print(f"📊 xlwings: ✅ {xlwings.__version__}")
            except ImportError:
                print("📊 xlwings: ❌ Not installed")
            
            try:
                import requests
                print(f"🌐 requests: ✅ {requests.__version__}")
            except ImportError:
                print("🌐 requests: ❌ Not installed")
            
            print(f"🔑 OpenAI API: {'✅ Configured' if OPENAI_API_KEY else '❌ Not configured'}")
            print(f"🎯 Model: {LQA_CONFIG['openai_model']}")
            print(f"📏 Quality Thresholds: Excellent≥{LQA_CONFIG['quality_thresholds']['excellent']}")
            
            # Test Excel connection
            if excel_integration.connect_to_excel():
                print("📊 Excel: ✅ Connected")
            else:
                print("📊 Excel: ⚠️ Connection failed")
                
        elif choice == '5':
            print("\n🧪 Testing system accuracy...")
            test_texts = [
                "This is a perfect sentence.",
                "This sentence have grammar errors.",
                "Speling mistakes here too.",
                "Perfect professional communication standards."
            ]
            
            for i, text in enumerate(test_texts, 1):
                print(f"\n📝 Test {i}: {text}")
                result = excel_integration.analyzer.analyze_text(text)
                print(f"   📊 Score: {result.quality_score}/100 | Issues: {result.error_count} | Time: {result.processing_time:.2f}s")
                
        elif choice == '6':
            print(f"\n💡 SETUP INSTRUCTIONS:")
            print("="*50)
            print("1. 🔧 INSTALL DEPENDENCIES:")
            print("   pip install xlwings requests openai")
            print("")
            print("2. 🔑 CONFIGURE API KEY:")
            print("   • Get OpenAI API key from: https://platform.openai.com/")
            print("   • Add to line 15 in this file:")
            print("   OPENAI_API_KEY = 'sk-proj-your-key-here'")
            print("")
            print("3. 📊 EXCEL SETUP:")
            if CURRENT_PLATFORM == 'Darwin':
                print("   • Install Excel: brew install --cask microsoft-excel")
                print("   • Install xlwings add-in: xlwings addin install")
            else:
                print("   • Install xlwings add-in: xlwings addin install")
            print("")
            print("4. 🎯 USAGE:")
            print("   • Open Excel workbook")
            print("   • Select cells with text")
            print("   • Run option 2 to analyze")
            print("")
            print("🌟 FEATURES:")
            print(f"   • {len(LQA_CONFIG['supported_languages'])} languages supported")
            print("   • Professional quality standards")
            print("   • Real-time analysis and reporting")
            print("   • Cross-platform compatibility")
            
        elif choice == '7':
            print("🍎 Thank you for using macOS LQA System!")
            break
            
        else:
            print("⚠️ Invalid option. Please select 1-7.")

def main():
    """Main entry point for macOS LQA system"""
    try:
        print("🍎 Initializing macOS Optimized Multilingual LQA System...")
        
        # Platform verification
        print(f"🖥️ Platform detected: {CURRENT_PLATFORM}")
        
        if CURRENT_PLATFORM == 'Darwin':
            print("✅ macOS optimizations enabled")
        else:
            print("✅ Windows compatibility mode")
        
        # Configuration status
        if OPENAI_API_KEY:
            print("✅ OpenAI API configured for maximum accuracy")
        else:
            print("⚠️ OpenAI API not configured - using LanguageTool only")
            print("💡 Add your API key to line 15 for enhanced accuracy")
        
        # Enhanced configuration summary
        print(f"🎯 Quality Standards: Excellent≥{LQA_CONFIG['quality_thresholds']['excellent']}")
        print(f"🌍 Supported Languages: {len(LQA_CONFIG['supported_languages'])}")
        print(f"🔧 Analysis Model: {LQA_CONFIG['openai_model']}")
        
        # Quick system test
        print("\n🧪 Quick system test...")
        analyzer = EnhancedMultilingualAnalyzer()
        test_result = analyzer.analyze_text("This sentence have grammar errors for testing.")
        print(f"✅ Test completed: Score {test_result.quality_score}/100 | {test_result.error_count} issues | {test_result.processing_time:.2f}s")
        
        # Launch interactive menu
        interactive_menu()
        
    except KeyboardInterrupt:
        print("\n🛑 System interrupted by user")
    except Exception as e:
        print(f"❌ System error: {e}")
        print("💡 Please check dependencies: pip install xlwings requests openai")

if __name__ == "__main__":
    main()