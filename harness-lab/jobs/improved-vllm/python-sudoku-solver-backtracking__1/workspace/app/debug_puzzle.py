#!/usr/bin/env python3
"""Debug script to examine puzzle file format"""

with open('app/puzzle.txt', 'r', encoding='utf-8') as f:
    content = f.read()
    
print(f"File contents (repr): {repr(content)}")
print(f"File contents (print):")
print(content)
print("---")
print(f"Lines: {len(content.split(chr(10)))}")
