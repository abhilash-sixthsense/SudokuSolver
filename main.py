import csv

import openpyxl

solve_call_count = 0


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f"Hi, {name}")  # Press F9 to toggle the breakpoint.


def load_array_from_file(f_name="sudoko_problem.csv"):
    """
    Load a CSV file and return its contents as a 2D array (list of lists).
    """
    grid = []
    with open(f_name, encoding="utf-8") as csvfile:
        reader = csv.reader(csvfile)
        for row in reader:
            cleaned_row = []
            for cell in row:
                cell = cell.strip()
                if cell.isdigit():
                    cleaned_row.append(int(cell))
                else:
                    cleaned_row.append(0)
            # Ignore completely empty rows (optional safety)

            grid.append(cleaned_row)

    return grid


def load_array_from_xlsx(f_name="sudoko_problem.xlsx"):
    """
    Load an Excel file and return its contents as a 2D array (list of lists).
    """

    grid = []
    workbook = openpyxl.load_workbook(f_name)

    sheet = workbook["ActiveProblem"]

    for row in sheet.iter_rows(min_row=5, max_row=13, values_only=True):
        cleaned_row = []
        # read from G5 to O13 (inclusive)
        for cell in row[6:15]:  # G is index 6, O is index 14
            if cell is None:
                cleaned_row.append(0)
            elif isinstance(cell, (int, float)):
                cleaned_row.append(int(cell))
            elif isinstance(cell, str) and cell.strip().isdigit():
                cleaned_row.append(int(cell.strip()))
            else:
                cleaned_row.append(0)
        grid.append(cleaned_row)

    return grid


def print_grid(grid, minimal=False):
    missing_numbers_columns, missing_numbers_rows, missing_number_boxes = get_missing_numbers(grid)
    if not minimal:
        for m in missing_numbers_columns:
            print(m)
    for i, row in enumerate(grid):  # Print horizontal line every 3 rows (except at top)
        if i % 3 == 0 and i != 0:
            print("-" * 52)
        if not minimal:
            print(missing_numbers_rows[i] if missing_numbers_rows else "", end=" ")
            print((9 - len(missing_numbers_rows[i])) * "   ", end=" | ")
        for j, num in enumerate(row):
            # Print vertical line every 3 columns (except at left)
            if j % 3 == 0 and j != 0:
                print("|", end=" ")

            print(num if num != 0 else ".", end=" ")

        print()
    if not minimal:
        for i, box in enumerate(missing_number_boxes):
            print(f"Box {i}: missing numbers: {box}")


def get_size(grid):
    return len(grid), len(grid[0]) if grid else 0


def get_box_index(r, c):
    return (r // 3) * 3 + (c // 3)


def get_missing_numbers(grid):
    # Placeholder for Sudoku solving logic
    rsize, csize = get_size(grid)
    missing_numbers_rows = []
    for i in range(rsize):
        missing_in_row = set(range(1, 10)) - set(grid[i])
        missing_numbers_rows.append(missing_in_row)
    missing_numbers_columns = []
    for j in range(csize):
        col = [grid[i][j] for i in range(rsize)]
        missing_in_col = set(range(1, 10)) - set(col)
        missing_numbers_columns.append(missing_in_col)

    missing_numbers_boxes = [set(range(1, 10)) for _ in range((rsize * csize) // 9)]
    for r in range(rsize):
        for c in range(csize):
            if grid[r][c] != 0:
                box = get_box_index(r, c)
                missing_numbers_boxes[box].discard(grid[r][c])

    return missing_numbers_columns, missing_numbers_rows, missing_numbers_boxes


def clone_grid(grid):
    return [row[:] for row in grid]


def is_grid_solved(grid):
    for row in grid:
        if 0 in row:
            return False
    return True


def find_empty_position(grid):
    rsize, csize = get_size(grid)
    for r in range(rsize):
        for c in range(csize):
            if grid[r][c] == 0:
                return r, c
    return None, None


def solve_grid(grid):
    global solve_call_count
    solve_call_count += 1

    print(f"\n\n\nSolve call count: {solve_call_count} ")
    print_grid(grid, minimal=True)

    # Find the first empty cell
    r, c = find_empty_position(grid)

    # If no empty cell, puzzle is solved
    if r is None and is_grid_solved(grid):
        print("Sudoku Solved Successfully! ")
        print_grid(grid, minimal=True)
        return grid

    missing_numbers_columns, missing_numbers_rows, missing_numbers_boxes = get_missing_numbers(grid)
    box = get_box_index(r, c)
    possible_numbers = missing_numbers_rows[r].intersection(missing_numbers_columns[c]).intersection(missing_numbers_boxes[box])

    if len(possible_numbers) == 0:
        # Dead end - no valid numbers for this cell
        return None

    # Try each possible number
    for number in possible_numbers:
        test_grid = clone_grid(grid)
        test_grid[r][c] = number
        print(f"Trying {number} at ({r+1}, {c+1})")
        solved_grid = solve_grid(test_grid)
        if solved_grid is not None:
            return solved_grid

    # No valid solution found
    return None


# Press the green button in the gutter to run the script.
if __name__ == "__main__":

    loaded_grid = load_array_from_xlsx()
    print("Loaded Sudoku Grid....\n\n")

    print_grid(loaded_grid)
    solve_grid(loaded_grid)
    print("\n\n")

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
