#!/usr/bin/env python3
"""Sudoku Solver using Backtracking Algorithm"""

def is_valid(board, row, col, num):
    """
    Check if placing num at board[row][col] is valid

    Args:
        board: 9x9 Sudoku board
        row: row index
        col: column index
        num: number to check (1-9)

    Returns:
        True if valid, False otherwise
    """
    # Check row
    for j in range(9):
        if board[row][j] == num:
            return False

    # Check column
    for i in range(9):
        if board[i][col] == num:
            return False

    # Check 3x3 box
    box_row = (row // 3) * 3
    box_col = (col // 3) * 3
    for i in range(3):
        for j in range(3):
            if board[box_row + i][box_col + j] == num:
                return False

    return True


def find_empty(board):
    """
    Find an empty position (0) in the board

    Args:
        board: 9x9 Sudoku board

    Returns:
        (row, col) tuple of empty position, or None if no empty position exists
    """
    for i in range(9):
        for j in range(9):
            if board[i][j] == 0:
                return (i, j)
    return None


def solve_sudoku(board):
    """
    Solve the Sudoku board using backtracking algorithm

    Args:
        board: 9x9 Sudoku board (modified in place)

    Returns:
        True if solved, False if no solution exists
    """
    # Find empty cell
    empty = find_empty(board)

    # No empty cells means board is filled - we're done
    if empty is None:
        return True

    row, col = empty

    # Try numbers 1-9 in the empty cell
    for num in range(1, 10):
        if is_valid(board, row, col, num):
            # Place the number
            board[row][col] = num

            # Recursively try to solve the rest
            if solve_sudoku(board):
                return True

            # If that failed, backtrack (remove the number)
            board[row][col] = 0

    # If no digit works, this path is invalid
    return False


def solve_and_print(board):
    """Solve the Sudoku and print the result"""
    if not solve_sudoku(board):
        print("No solution exists for this Sudoku puzzle.")
        exit(1)
    print_board(board)


def print_board(board):
    """Print the Sudoku board in a readable format"""
    for row in board:
        print("".join(str(d) for d in row))


def main():
    """Main function to read puzzle, solve it, and write solution"""
    # Read the puzzle from file
    with open('app/puzzle.txt', 'r', encoding='utf-8') as puzzle_file:
        puzzle_lines_raw = puzzle_file.read().strip()

    # The puzzle file contains lines separated by newlines
    # Parse the board - each line is a row of digits
    board_lines = puzzle_lines_raw.split('\n')

    # Parse each line into a board row
    board = []
    for line in board_lines:
        row = []
        for char in line:
            if char.isdigit():
                row.append(int(char))
            else:
                row.append(0)  # Treat non-digits as empty
        board.append(row)

    # Solve the Sudoku
    solve_and_print(board)

    # Write the solution to output file
    with open('app/solution.txt', 'w', encoding='utf-8') as f:
        for row in board:
            f.write("".join(str(d) for d in row) + '\n')


if __name__ == '__main__':
    main()
