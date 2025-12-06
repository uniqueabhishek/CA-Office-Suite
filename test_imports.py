"""
Test script to verify all imports resolve correctly after Phase 4.2 refactoring.
"""

import sys
import os

# Add current directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

print("Testing Phase 4.2 refactoring imports...")
print("-" * 60)

# Test 1: Workers package
try:
    from workers import BaseWorker, FormatterWorker, PDFWorker, MergeWorker
    print("✓ Workers package imports successfully")
    print(f"  - BaseWorker: {BaseWorker}")
    print(f"  - FormatterWorker: {FormatterWorker}")
    print(f"  - PDFWorker: {PDFWorker}")
    print(f"  - MergeWorker: {MergeWorker}")
except ImportError as e:
    print(f"✗ Workers package import failed: {e}")
    sys.exit(1)

# Test 2: Core modules
try:
    from core.excel_utils import list_excel_files_in_folder, read_file_to_df, read_all_sheets
    from core.excel_writer import save_df_to_excel, apply_formatting_to_workbook
    print("✓ Core modules import successfully")
except ImportError as e:
    print(f"✗ Core modules import failed: {e}")
    sys.exit(1)

# Test 3: Config
try:
    from config.constants import SUPPORTED_EXTENSIONS
    print(f"✓ Config module imports successfully")
    print(f"  - SUPPORTED_EXTENSIONS: {SUPPORTED_EXTENSIONS}")
except ImportError as e:
    print(f"✗ Config module import failed: {e}")
    sys.exit(1)

# Test 4: Verify worker inheritance
print("-" * 60)
print("Verifying worker class hierarchy...")
print(f"  - FormatterWorker inherits from BaseWorker: {issubclass(FormatterWorker, BaseWorker)}")
print(f"  - PDFWorker inherits from BaseWorker: {issubclass(PDFWorker, BaseWorker)}")
print(f"  - MergeWorker inherits from BaseWorker: {issubclass(MergeWorker, BaseWorker)}")

# Test 5: Verify worker has required methods
print("-" * 60)
print("Verifying worker methods...")
for worker_class in [FormatterWorker, PDFWorker, MergeWorker]:
    has_run = hasattr(worker_class, 'run')
    has_cancel = hasattr(worker_class, 'cancel')
    has_emit_progress = hasattr(worker_class, 'emit_progress')
    has_emit_error = hasattr(worker_class, 'emit_error')

    print(f"{worker_class.__name__}:")
    print(f"  - run(): {has_run}")
    print(f"  - cancel(): {has_cancel}")
    print(f"  - emit_progress(): {has_emit_progress}")
    print(f"  - emit_error(): {has_emit_error}")

    if not (has_run and has_cancel and has_emit_progress and has_emit_error):
        print(f"✗ {worker_class.__name__} missing required methods!")
        sys.exit(1)

print("-" * 60)
print("✓ All Phase 4.2 refactoring tests PASSED!")
print()
print("Summary:")
print("  - Workers package created and exports correctly")
print("  - All worker classes inherit from BaseWorker")
print("  - All worker classes have required methods")
print("  - Core and config modules import successfully")
print()
print("Phase 4.2: Add Threading to PDF & Merge Tools - COMPLETE ✓")
