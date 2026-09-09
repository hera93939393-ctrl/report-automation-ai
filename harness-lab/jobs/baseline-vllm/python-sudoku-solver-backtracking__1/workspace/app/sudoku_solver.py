#!/usr/bin/env python3
"""
Sudoku Solver using Backtracking Algorithm
"""

def is_valid(board, row, col, num):
    """
    Check if placing num at board[row][col] is valid
    
    Rules:
    - Each row must contain 1-9 without duplicates
    - Each column must contain 1-9 without duplicates
    - Each 3x3 box must contain 1-9 without duplicates
    """
    # Check row
    for c in range(9):
        if board[row][c] == num and c != col:
            return False
    
    # Check column
    for r in range(9):
        if board[r][col] == num and r != row:
            return False
    
    # Check 3x3 box
    start_row = row // 3 * 3
    start_col = col // 3 * 3
    for r in range(start_row, start_row + 3):
        for c in range(start_col, start_col + 3):
            if board[r][c] == num and (r != row or c != col):
                return False
    
    return True


def find_empty(board):
    """
    Find next empty position
    Returns (row, col) tuple, or None if no empty positions
    """
    for r in range(9):
        for c in range(9):
            if board[r][c] == 0:
                return (r, c)
    return None


def solve_sudoku(board):
    """
    Solve sudoku using recursive backtracking
    Returns True if solved, False otherwise
    """
    root = find_empty(board)
    
    if root is None:
        # All cells filled
        return True
    
    row, col = root
    
    for num in range(1, 10):
        if is_valid(board, row, col, num):
            board[row][col] = num
            if solve_sudoku(board):
                return True
            board[row][col] = 0  # Backtrack
    
    return False


def solve_given_file():
    """
    Solve sudoku from file and write solution
    """
    puzzle_file = 'app/puzzle.txt'
    solution_file = 'app/solution.txt'
    
    try:
        print("Reading sudoku puzzle...")
        board = []
        with open(puzzle_file, 'r') as f:
            for line in f:
                line = line.strip()
                row = []
                for char in line:
                    if char.isdigit():
                        row.append(int(char))
                    else:
                        row.append(0)
                board.append(row)
        
        if not solve_sudoku(board):
            print("Error: No solution exists")
            return 1
        
        print("Sudoku solved successfully!")
        with open(solution_file, 'w') as f:
            for row in board:
                f.write(''.join(str(num) for num in row) + '\n')
        print("Solution written to: " + solution_file)
        return 0
    except FileNotFoundError as e:
        print("Error: File not found - " + str(e))
        return 1
    except Exception as e:
        print("Error - " + str(e))
        return 1


if __name__ == "__main__":
    exit(solve_given_file())
