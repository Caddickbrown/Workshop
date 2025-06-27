import curses
import os
import subprocess
import platform
from datetime import datetime

JOURNAL_DIR = os.path.expanduser("~/journal/entries")

def ensure_journal_dir():
    os.makedirs(JOURNAL_DIR, exist_ok=True)

def get_entry_list():
    ensure_journal_dir()
    return sorted(os.listdir(JOURNAL_DIR))

def debug_key_codes(stdscr):
    """Debug function to identify key codes"""
    stdscr.clear()
    stdscr.addstr(0, 0, "Press keys to see their codes (ESC to exit):")
    stdscr.addstr(2, 0, "Key Code:")
    
    while True:
        key = stdscr.getch()
        if key == 27:  # ESC
            break
        stdscr.addstr(2, 10, f"{key} (0x{key:02x})")
        stdscr.refresh()

def is_enter_key(key):
    """Check if the key is any form of enter key"""
    return key in (10, 13, 459, curses.KEY_ENTER, ord('\r'))

def write_in_terminal(stdscr, title, tags):
    """Write entry content directly in the terminal"""
    curses.echo()
    curses.curs_set(1)
    
    # Create a subwindow for writing
    height, width = stdscr.getmaxyx()
    write_win = curses.newwin(height-4, width-2, 2, 1)
    write_win.addstr(0, 0, f"Writing: {title}")
    write_win.addstr(1, 0, "Start typing your entry (Ctrl+D when finished):")
    
    # Get the content
    content_lines = []
    current_line = ""
    y_pos = 2
    
    while True:
        try:
            char = write_win.getch()
            
            if char == 4:  # Ctrl+D
                if current_line:
                    content_lines.append(current_line)
                break
            elif is_enter_key(char):  # Enter (regular or numpad)
                content_lines.append(current_line)
                current_line = ""
                y_pos += 1
                if y_pos >= height-6:  # Scroll if needed
                    write_win.scroll()
                    y_pos = height-6
            elif char == 127 or char == 8:  # Backspace
                if current_line:
                    current_line = current_line[:-1]
                    write_win.addch(y_pos, len(current_line) + 1, ' ')
                    write_win.move(y_pos, len(current_line) + 1)
            else:
                current_line += chr(char)
                write_win.addch(y_pos, len(current_line), char)
                
        except KeyboardInterrupt:
            break
    
    curses.noecho()
    curses.curs_set(0)
    
    return "\n".join(content_lines)

def new_entry(stdscr, use_editor=False):
    ensure_journal_dir()
    curses.echo()
    stdscr.clear()
    stdscr.addstr(0, 0, "Entry Title: ")
    title = stdscr.getstr(0, 13).decode()
    stdscr.addstr(1, 0, "Tags (comma separated): ")
    tags = stdscr.getstr(1, 25).decode()
    curses.noecho()

    date = datetime.now().strftime("%Y-%m-%d")
    safe_title = title.lower().replace(" ", "-")
    filename = f"{date}-{safe_title}.md"
    filepath = os.path.join(JOURNAL_DIR, filename)

    if use_editor:
        # Write header and open in editor
        with open(filepath, "w") as f:
            f.write(f"# {title}\n\ntags: {tags}\n\n")
        
        stdscr.addstr(3, 0, "Press any key to open editor...")
        stdscr.getch()
        curses.endwin()
        
        # Define the editor
        editor = "nano"  # default for Linux/macOS
        if platform.system() == "Windows":
            editor = "notepad.exe"
        
        subprocess.call([editor, filepath])
    else:
        # Write content in terminal
        content = write_in_terminal(stdscr, title, tags)
        
        # Save the file
        with open(filepath, "w") as f:
            f.write(f"# {title}\n\ntags: {tags}\n\n{content}\n")
        
        stdscr.clear()
        stdscr.addstr(0, 0, f"Entry saved: {filename}")
        stdscr.addstr(2, 0, "Press any key to continue...")
        stdscr.getch()

def select_entry(stdscr, action):
    entries = get_entry_list()
    if not entries:
        stdscr.addstr(0, 0, "No entries found. Press any key to continue...")
        stdscr.getch()
        return
        
    idx = 0
    while True:
        stdscr.clear()
        stdscr.addstr(0, 0, f"{action} Entry (ESC to cancel):")
        for i, entry in enumerate(entries):
            if i == idx:
                stdscr.addstr(i+2, 2, f"> {entry}", curses.A_REVERSE)
            else:
                stdscr.addstr(i+2, 2, f"  {entry}")
        key = stdscr.getch()
        if key == curses.KEY_UP and idx > 0:
            idx -= 1
        elif key == curses.KEY_DOWN and idx < len(entries) - 1:
            idx += 1
        elif is_enter_key(key):  # Enter (regular or numpad)
            filepath = os.path.join(JOURNAL_DIR, entries[idx])
            if action == "Read":
                # Read and display file content in terminal
                curses.endwin()
                try:
                    with open(filepath, 'r', encoding='utf-8') as f:
                        content = f.read()
                    print(f"\n=== {entries[idx]} ===")
                    print(content)
                    print("\nPress Enter to continue...")
                    input()
                except Exception as e:
                    print(f"Error reading file: {e}")
                    input("Press Enter to continue...")
            else:
                # Edit in external editor
                curses.endwin()
                editor = "nano"  # default for Linux/macOS
                if platform.system() == "Windows":
                    editor = "notepad.exe"
                subprocess.call([editor, filepath])
            break
        elif key == 27:  # ESC
            break

def settings_menu(stdscr):
    """Settings submenu"""
    curses.curs_set(0)
    current_row = 0
    menu = ["Debug Keys", "Back"]
    while True:
        stdscr.clear()
        stdscr.addstr(0, 0, "Settings")
        for idx, item in enumerate(menu):
            if idx == current_row:
                stdscr.addstr(idx + 2, 2, f"> {item}", curses.A_REVERSE)
            else:
                stdscr.addstr(idx + 2, 2, f"  {item}")
        key = stdscr.getch()
        if key == curses.KEY_UP and current_row > 0:
            current_row -= 1
        elif key == curses.KEY_DOWN and current_row < len(menu) - 1:
            current_row += 1
        elif is_enter_key(key):  # Enter (regular or numpad)
            if menu[current_row] == "Debug Keys":
                debug_key_codes(stdscr)
            elif menu[current_row] == "Back":
                break
        elif key == 27:  # ESC to go back
            break

def main_menu(stdscr):
    curses.curs_set(0)
    current_row = 0
    menu = ["New Entry (Terminal)", "New Entry (Editor)", "Read Entry", "Edit Entry", "Settings", "Exit"]
    while True:
        stdscr.clear()
        stdscr.addstr(0, 0, "Terminal Journal")
        for idx, item in enumerate(menu):
            if idx == current_row:
                stdscr.addstr(idx + 2, 2, f"> {item}", curses.A_REVERSE)
            else:
                stdscr.addstr(idx + 2, 2, f"  {item}")
        key = stdscr.getch()
        if key == curses.KEY_UP and current_row > 0:
            current_row -= 1
        elif key == curses.KEY_DOWN and current_row < len(menu) - 1:
            current_row += 1
        elif is_enter_key(key):  # Enter (regular or numpad)
            if menu[current_row] == "New Entry (Terminal)":
                new_entry(stdscr, use_editor=False)
            elif menu[current_row] == "New Entry (Editor)":
                new_entry(stdscr, use_editor=True)
            elif menu[current_row] == "Read Entry":
                select_entry(stdscr, "Read")
            elif menu[current_row] == "Edit Entry":
                select_entry(stdscr, "Edit")
            elif menu[current_row] == "Settings":
                settings_menu(stdscr)
            elif menu[current_row] == "Exit":
                break
        elif key == 27:  # ESC to exit
            break

def main():
    curses.wrapper(main_menu)

if __name__ == "__main__":
    main()
