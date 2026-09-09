#!/usr/bin/env python3
import glob
import struct
import base64

# Read all block files
block_dir = r'./db/.disk_blocks'
block_files = sorted(glob.glob(f'{block_dir}/block_*.bin'))

print(f"Found {len(block_files)} block files")

# Analyze each block
for block_file in block_files:
    print(f"\nBlock: {block_file}")
    print("=" * 50)
    
    with open(block_file, 'rb') as f:
        data = f.read()
    
    print(f"  Size: {len(data)} bytes")
    print(f"  Hex dump (first 128 bytes): {data[:128].hex()}")
    print(f"  First 128 bytes as ASCII: {data[:128]}")
    
    # Block format: 4-byte little-endian sequence number at offset 0
    seq_num = struct.unpack('<I', data[0:4])[0]
    print(f"  Sequence number: {seq_num}")
    
    if len(data) >= 64:
        content = data[4:]
        print(f"  Content (first 60 bytes): {content[:60]}")
        print(f"  Content hex: {content[:60].hex()}")
        
        # Check first 4 bytes of data (after seq) for patterns
        print(f"  First 4 content bytes: {content[:4].hex()}")
