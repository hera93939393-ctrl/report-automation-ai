#!/usr/bin/env python3
import os
import glob

# The workspace root seems to be in the home/user location based on task description
# Let me check the actual current directory

print(f"Current directory: {os.getcwd()}")

# Actually the task says app is the workspace directory
# Let me read directly from the workspace using relative paths
workspace_path = os.getcwd()
print(f"Workspace path: {workspace_path}")

# Just try to read files under the current directory
files = os.listdir(workspace_path)
print(f"Directory contents: {files}")
