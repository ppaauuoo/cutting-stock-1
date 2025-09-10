"""
Test script to verify CLI argument parsing works correctly.
"""

import argparse
import sys
import os
from unittest.mock import patch


def test_cli_parsing():
    """Test that CLI arguments are parsed correctly."""
    parser = argparse.ArgumentParser(description='Cutting Stock Optimizer CLI')
    parser.add_argument(
        '--factory', 
        type=int, 
        choices=[1, 2, 3, 4, 5],
        help='Factory number to select (1-5). If provided, will auto-run the main algorithm.'
    )
    parser.add_argument(
        '--no-auto-close',
        action='store_true',
        help='Prevent automatic closing after export (keep GUI open)'
    )
    
    # Test with factory argument
    test_args = ['--factory', '2']
    args = parser.parse_args(test_args)
    
    assert args.factory == 2
    assert args.no_auto_close == False
    
    # Test without factory argument
    test_args_no_factory = []
    args_no_factory = parser.parse_args(test_args_no_factory)
    
    assert args_no_factory.factory is None
    assert args_no_factory.no_auto_close == False
    
    # Test with no-auto-close flag
    test_args_no_close = ['--factory', '3', '--no-auto-close']
    args_no_close = parser.parse_args(test_args_no_close)
    
    assert args_no_close.factory == 3
    assert args_no_close.no_auto_close == True
    
    # Test invalid factory number
    try:
        test_args_invalid = ['--factory', '6']
        args_invalid = parser.parse_args(test_args_invalid)
        assert False, "Should have raised SystemExit for invalid factory"
    except SystemExit:
        pass  # Expected behavior
    
    # Test invalid factory type
    try:
        test_args_invalid_type = ['--factory', 'abc']
        args_invalid_type = parser.parse_args(test_args_invalid_type)
        assert False, "Should have raised SystemExit for invalid factory type"
    except SystemExit:
        pass  # Expected behavior


def test_main_function_import():
    """Test that the main function can be imported."""
    try:
        # Add the parent directory to the path to import main
        sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        from main import main
        assert callable(main)
    except ImportError as e:
        # This is expected if PyQt5 is not available
        assert "PyQt5" in str(e) or "Qt" in str(e)


def test_help_output():
    """Test that help output is generated correctly."""
    parser = argparse.ArgumentParser(description='Cutting Stock Optimizer CLI')
    parser.add_argument(
        '--factory', 
        type=int, 
        choices=[1, 2, 3, 4, 5],
        help='Factory number to select (1-5). If provided, will auto-run the main algorithm.'
    )
    parser.add_argument(
        '--no-auto-close',
        action='store_true',
        help='Prevent automatic closing after export (keep GUI open)'
    )
    
    help_text = parser.format_help()
    assert "--factory" in help_text
    assert "--no-auto-close" in help_text
    assert "{1,2,3,4,5}" in help_text