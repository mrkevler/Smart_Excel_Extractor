import openpyxl
import csv
import os
import curses

def select_file(stdscr):
    curses.curs_set(0)
    current_row = 0

    files = [f for f in os.listdir('.') if os.path.isfile(f) and f.endswith('.xlsx')]
    if not files:
        stdscr.addstr(0, 0, "No Excel files found in the current directory.")
        stdscr.refresh()
        stdscr.getch()
        return None

    while True:
        stdscr.clear()
        stdscr.addstr(0, 0, "Select an Excel file:")
        for idx, file in enumerate(files):
            if idx == current_row:
                stdscr.attron(curses.color_pair(1))
                stdscr.addstr(idx + 1, 0, file)
                stdscr.attroff(curses.color_pair(1))
            else:
                stdscr.addstr(idx + 1, 0, file)
        key = stdscr.getch()
        if key == curses.KEY_UP and current_row > 0:
            current_row -= 1
        elif key == curses.KEY_DOWN and current_row < len(files) - 1:
            current_row += 1
        elif key == curses.KEY_ENTER or key in [10, 13]:
            return files[current_row]

def read_excel(file_path):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    data = []
    current_category = None

    for row in sheet.iter_rows(min_row=2, values_only=True):
        if all(cell is None for cell in row):
            continue
        if row[0] is not None:
            current_category = row[0]
        data.append((current_category, row[1], row[2]))

    return data

def write_to_csv(data, output_file):
    with open(output_file, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(['Category', 'Item', 'Description'])
        writer.writerows(data)

def write_to_txt(data, output_file):
    with open(output_file, mode='w', encoding='utf-8') as file:
        for row in data:
            file.write(f"Category: {row[0]}\nItem: {row[1]}\nDescription: {row[2]}\n\n")

def handle_existing_files(stdscr, base_name):
    csv_file = f"{base_name}.csv"
    txt_file = f"{base_name}.txt"
    stdscr.addstr(0, 0, f"Files with the same name already exist: {csv_file} or {txt_file}")
    stdscr.addstr(1, 0, "Do you want to overwrite them? (y/n/r for rename): ")
    stdscr.refresh()
    key = stdscr.getch()
    if key in [ord('y'), ord('Y')]:
        return "overwrite", csv_file, txt_file
    elif key in [ord('r'), ord('R')]:
        stdscr.addstr(3, 0, "Enter new file name without extension: ")
        curses.echo()
        new_base_name = stdscr.getstr(4, 0).decode('utf-8')
        curses.noecho()
        return "rename", f"{new_base_name}.csv", f"{new_base_name}.txt"
    else:
        return "cancel", None, None

def main():
    curses.wrapper(run)

def run(stdscr):
    curses.start_color()
    curses.init_pair(1, curses.COLOR_BLACK, curses.COLOR_WHITE)

    selected_file = select_file(stdscr)
    if not selected_file:
        return

    data = read_excel(selected_file)

    base_name = os.path.splitext(selected_file)[0]
    csv_file = f"{base_name}.csv"
    txt_file = f"{base_name}.txt"

    if os.path.exists(csv_file) or os.path.exists(txt_file):
        action, csv_file, txt_file = handle_existing_files(stdscr, base_name)
        if action == "cancel":
            return

    write_to_csv(data, csv_file)
    write_to_txt(data, txt_file)
    max_y, max_x = stdscr.getmaxyx()
    message = f"Files '{csv_file}' and '{txt_file}' have been created successfully."
    message_row = min(len(data) + 5, max_y - 1)
    message_col = min(0, max_x - len(message) - 1)
    stdscr.addstr(message_row, message_col, message)
    stdscr.refresh()
    stdscr.getch()

if __name__ == "__main__":
    main()