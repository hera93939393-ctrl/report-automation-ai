#!/usr/bin/env python3
"""Check what's actually in the puzzle file"""

with open('app/puzzle.txt', 'rb') as f:
    raw = f.read()

print(f"File size: {len(raw)} bytes")
print(f"Raw bytes: {raw}")
print(f"Characters:")
for i, b in enumerate(raw[:100]):
    print(f"  {i}: 0x{b:02X} = {chr(b) if b < 128 else '?'}")

# Check if it contains actual newlines
has_newline = b'\n' in raw
print(f"\nHas actual newline (0x0A): {has_newline}")

# Check for Python's escape format
has_backslash_n = b'\\n' in raw
print(f"Has literal \\n (backslash-n): {has_backslash_n}")

# Let's try parsing as JSON
import json
try:
    data = json.loads(raw.decode('utf-8'))
    print(f"Is valid JSON: {data}")
    if isinstance(data, str):
        print(f"JSON string content: {data[:100]}")
except Exception as e:
    print(f"Not valid JSON: {e}")
