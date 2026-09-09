#!/usr/bin/env python3
import os
import glob

current_dir = os.getcwd()
print(f"Current directory: {current_dir}")

# Check the db directory
db_dir = os.path.join(current_dir, 'db')
if os.path.exists(db_dir):
    print(f"\\ndb directory exists: {db_dir}")
    files = os.listdir(db_dir)
    print(f"Contents of db: {files}")

# Check for .disk_blocks
disk_blocks_path = os.path.join(db_dir, '.disk_blocks')
if os.path.exists(disk_blocks_path):
    print(f"\\n.disk_blocks exists: {disk_blocks_path}")
    block_files = sorted(glob.glob(os.path.join(disk_blocks_path, 'block_*.bin')))
    print(f"Block files found: {len(block_files)}")
    for i, f in enumerate(block_files):
        print(f"  {i}: {f}")
else:
    print(f".disk_blocks not found at {disk_blocks_path}")
