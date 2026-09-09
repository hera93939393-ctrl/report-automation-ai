#!/bin/bash

# Solution script for sudoku solver
# This script calls the Python sudoku solver

# Ensure we're in the app directory
cd "$(dirname "$0")"

# Run the Python sudoku solver
python3 sudoku_solver.py
