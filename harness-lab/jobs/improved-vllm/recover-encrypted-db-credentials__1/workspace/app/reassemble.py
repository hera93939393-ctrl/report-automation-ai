#!/usr/bin/env python3
import glob
import struct
import base64
import os

# Read all block files numbered 000-007 (in sequence order)
block_dir = r'./db/.disk_blocks'
block_files = []

for i in range(16):  # Check up to block_015
    import subprocess
    out = subprocess.call(['', f"./db/.disk_blocks/block_{i:03d}.bin"])
    if os.path.exists(os.path.join(block_dir, f'block_{i:03d}.bin')):
        block_files.append((i, os.path.join(block_dir, f'block_{i:03d}.bin')))

# Actually just sort existing numbered blocks
block_files = []
for f in glob.glob(os.path.join(block_dir, 'block_*.bin')):
    if '_garbage' not in f:  # Only numbered blocks
        block_num = int('_'.join(list(f)[-2:]))  # This won't work
        pass

# Let's do this more carefully
block_files = []
for filename in sorted(glob.glob(os.path.join(block_dir, 'block_*.bin'))):
    if '_garbage' not in filename:
        # Extract the actual sequence number from offset 0
        with open(filename, 'rb') as kf:
            seq = struct.unpack('<I', kf.read(4))[0]
            block_files.append((seq, filename))

# Wait, simpler: just use glob with pattern for block_000.bin through block_007.bin
block_files = []
for fn in sorted(glob.glob(os.path.join(block_dir, 'block_000.bin'))):
    # Just use block_000.bin through block_00N.bin
    pass

# Let's examine the files with base64 numbers properly
import glob
blockfiles = glob.glob(os.path.join(block_dir, 'block_*.bin'))
filterfn = [f for f in blockfiles if '_garbage' not in f and f.endswith('.bin')]
filterfn.sort()

print("This doesn't give us the right list...")
