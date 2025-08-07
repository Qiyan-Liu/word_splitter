#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word Document Splitter - Main Entry Point

This is the main entry point for the Word Document Splitter application.
It imports and runs the actual implementation from the src directory.

author: QiyanLiu
date: 2025-08-06
"""

import sys
from pathlib import Path

# Add src directory to Python path
src_path = Path(__file__).parent / "src"
sys.path.insert(0, str(src_path))

# Import and run the main function from src
from app import main

if __name__ == "__main__":
    main()