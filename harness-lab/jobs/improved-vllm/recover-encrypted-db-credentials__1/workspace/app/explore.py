#!/usr/bin/env python3
import os
import glob

workspace_root = '/app'

print(f"Current directory: {os.getcwd()}")
print(f"Workspace root: {workspace_root}")

# List files at workspace root level
try:
    files = os.listdir(workspace_root)
    print(f"Files in workspace root: {files}")
except Exception as e:
    print(f"Error listing root: {e}")

# Check for .disk_blocks directory if it exists
disk_blocks_path = '/app/db/.disk_blocks'
if os.path.exists(disk_blocks_path):
    print(f"\\nChecking {disk_blocks_path}")
    block_files = sorted(glob.glob(os.path.join(disk_blocks_path, 'block_*.bin')))
    print(f"Block files found: {len(block_files)}")
    for f in block_files[:10]:  # Show first 10
        print(f"  {f}")
    if len(block_files) > 10:
        print(f"  ... and {len(block_files) - 10} more")
else:
    print(f"{disk_blocks_path} does not exist")
