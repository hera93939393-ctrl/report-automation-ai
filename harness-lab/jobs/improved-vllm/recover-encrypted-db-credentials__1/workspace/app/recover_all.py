#!/usr/bin/env python3
import glob
import struct
import base64
import os

block_dir = r'./db/.disk_blocks'

# Collect all block files (excluding garbage)
block_files = {}
for fn in glob.glob(os.path.join(block_dir, 'block_*.bin')):
    with open(fn, 'rb') as f:
        data = f.read()
    seq = struct.unpack('<I', data[0:4])[0]
    block_files[seq] = fn

print(f"Found {len(block_files)} blocks with sequence numbers: {sorted(block_files.keys())}")

# Reassemble in sequence order
reassembled = b''
# Include only blocks that seem legitimate (sequence numbers 0-15 presumably)
for seq in sorted(block_files.keys()):
    if seq <= 15:  # Assume only up to block_015 are legitimate
        print(f"  Adding block {seq}: {block_files[seq]}")
        with open(block_files[seq], 'rb') as f:
            content = f.read()
        reassembled += content[4:]  # Skip 4-byte sequence number

print(f"\\nReassembled data ({len(reassembled)} bytes):")
print(reassembled.decode('latin-1', errors='replace'))
