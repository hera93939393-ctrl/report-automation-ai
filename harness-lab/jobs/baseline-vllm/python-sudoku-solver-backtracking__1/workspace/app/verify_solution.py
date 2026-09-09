#!/usr/bin/env python3
"""Verify that the solution is correct"""

def main():
    """Verify that the sudoku board is valid"""
    # Read from solution file
    board = []
    with open('app/solution.txt', 'r') as f:
        for line in f:
            line = line.strip()
            row = []
            for char in line:
                row.append(int(char))
            board.append(row)
    
    print("Verifying solution...")
    
    # Check rows
    print("Checking rows...")
    for r in range(9):
        row_set = set()
        valid = True
        for c in range(9):
            num = board[r][c]
            if num == 0:
                continue
            if num in row_set:
                print("Error: Row", r, "has duplicate", num)
                valid = False
            row_set.add(num)
        if valid:
            print("Row", r, "OK")
    
    # Check columns
    print("Checking columns...")
    for c in range(9):
        col_set = set()
        valid = True
        for r in range(9):
            num = board[r][c]
            if num == 0:
                continue
            if num in col_set:
                print("Error: Column", c, "has duplicate", num)
                valid = False
            col_set.add(num)
        if valid:
            print("Column", c, "OK")
    
    # Check 3x3 boxes
    print("Checking 3x3 boxes...")
    for r in range(3):
        for c in range(3):
            box_set = set()
            valid = True
            for i in range(3):
                for j in range(3):
                    num = board[r*3+i][c*3+j]
                    if num == 0:
                        continue
                    if num in box_set:
                        print("Error: Box", (r,c), "has duplicate", num)
                        valid = False
                    box_set.add(num)
            if valid:
                print("Box", (r,c), "OK")
    
    print("\nVerification complete!")

if __name__ == "__main__":
    main()
