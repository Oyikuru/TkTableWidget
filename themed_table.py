import tkinter as tk
from tkinter import ttk
import pandas as pd
import numpy as np
from tkinter import messagebox
from tkinter.font import Font
import math

class Table(ttk.Frame):
    def __init__(self, parent, rows=10, cols=5, cell_width=120, cell_height=30, theme='default',spreadsheet_mode=False, **kwargs):
        super().__init__(parent)
        self.parent = parent
        self.rows = rows
        self.cols = cols
        self.theme = theme
        self.cell_width = kwargs.get('cell_width', 120)
        self.cell_height = kwargs.get('cell_height', 30)
        self.auto_scroll = True
        self.selection_mode = "single"  # Add this line
        self.default_cell_width = cell_width
        self.default_cell_height = cell_height
        self.row_heights = [self.default_cell_height] * rows
        self.font = Font(font=('Calibre', 11))
        self.spreadsheet_mode = spreadsheet_mode

        # Initialize the canvas and scrollbars
        # Configure grid weights for proper expansion
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self._setup_canvas()
        self._update_canvas_size()
        # self._setup_event_bindings()  

        # Add header style
        style = ttk.Style()
        style.configure('Header.TLabel', 
                        background='#f0f0f0',
                        foreground='black',
                        font=('Arial', 10, 'bold'))
        
        # Theme configuration
        self.themes = {
            'default': {
                'bg': 'white', 'fg': 'black',
                'select_bg': '#a6d2ff', 'select_fg': 'black',
                'even': 'white', 'odd': '#f0f0f0',
                'grid': 'gray', 'font': ('Arial', 10)
            },
            'dark': {
                'bg': '#2d2d2d', 'fg': 'white',
                'select_bg': '#3a5f8a', 'select_fg': 'white',
                'even': '#2d2d2d', 'odd': '#3d3d3d',
                'grid': '#555555', 'font': ('Arial', 10)
            },
            'blue': {
                'bg': '#e6f3ff', 'fg': 'black',
                'select_bg': '#0078d7', 'select_fg': 'white',
                'even': '#e6f3ff', 'odd': '#cce7ff',
                'grid': '#99ceff', 'font': ('Segoe UI', 10)
            },
            'green': {
                'bg': '#f0fff0', 'fg': 'black',
                'select_bg': '#2e8b57', 'select_fg': 'white',
                'even': '#f0fff0', 'odd': '#d0ffd0',
                'grid': '#a0d0a0', 'font': ('Calibri', 10)
            }
        }
        
        # Initialize data structures
        self.cells = {}
        self.merged_cells = {}
        self.selected_cells = set()
        self.undo_stack = []
        self.redo_stack = []
        self.max_undo_steps = 100
        self.selection_rect = None
        self.selection_start = None

        self.formulas = {}  # Stores formulas: {(row,col): "=A1+B2"}
        self.calculated_values = {}  # Stores computed values
        self.calculation_enabled = True  # Master switch
        
        if spreadsheet_mode:
            self._init_cell_references()
            self._setup_event_bindings()
        
        # Apply initial theme
        self.current_theme = self.themes[theme]
        self.create_grid()

        # Create initial grid
        self.resize_grid(rows, cols)
        
        # Bind events
        self.canvas.bind("<Button-1>", self.on_click)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_drag_end)
        self.bind_all("<Control-z>", self.undo)
        self.bind_all("<Control-y>", self.redo)
        self.bind_all("<Control-Shift-Z>", self.redo)
        self.bind_all("<Control-Button-1>", self.on_ctrl_click)
        self.bind_all("<Shift-Button-1>", self.on_shift_click)
        # Add to clipboard
        self.bind_all("<Control-c>", lambda e: self._copy_selected())
        self.bind_all("<Control-v>", lambda e: self._paste_to_selected())

        def _copy_selected(self):
            """Copy selected cells to clipboard"""
            if not self.selected_cells:
                return
            values = [self.get_cell(r, c)['value'] for r, c in self.selected_cells]
            self.clipboard_clear()
            self.clipboard_append("\t".join(values))
            return "break"

        def _paste_to_selected(self):
            """Paste clipboard content to selected cells"""
            if not self.selected_cells:
                return
            try:
                data = self.clipboard_get()
                values = data.split("\t")
                for (r, c), value in zip(sorted(self.selected_cells), values):
                    self.set_cell(r, c, value)
            except tk.TclError:
                pass
            return "break"

    def _setup_canvas(self):
        """Initialize canvas with optional reference headers"""
        self.grid_frame = ttk.Frame(self)
        self.grid_frame.grid(row=0, column=0, sticky='nsew')
    
        # Create headers only if in spreadsheet mode
        self.col_headers = []
        self.row_headers = []
    
        if self.spreadsheet_mode:
            # Column headers (A, B, C...)
            for col in range(self.cols):
                header = ttk.Label(
                    self.grid_frame,
                    text=chr(65 + col),
                    width=self.cell_width//7,
                    anchor='center',
                    style='Header.TLabel'
                )
                header.grid(row=0, column=col+1, sticky='nsew')
                self.col_headers.append(header)
        
            # Row headers (1, 2, 3...)
            for row in range(self.rows):
                header = ttk.Label(
                    self.grid_frame,
                    text=str(row + 1),
                    width=4,
                    anchor='e',
                    style='Header.TLabel'
                )
                header.grid(row=row+1, column=0, sticky='nsew')
                self.row_headers.append(header)
    
        # Main table canvas
        canvas_row = 1 if self.spreadsheet_mode else 0
        canvas_col = 1 if self.spreadsheet_mode else 0
        canvas_colspan = self.cols if self.spreadsheet_mode else self.cols + 1
        canvas_rowspan = self.rows if self.spreadsheet_mode else self.rows + 1
    
        self.canvas = tk.Canvas(
            self.grid_frame,
            highlightthickness=0,
            width=self.cols * self.cell_width,
            height=self.rows * self.cell_height
        )
        self.canvas.grid(
            row=canvas_row, 
            column=canvas_col, 
            columnspan=canvas_colspan, 
            rowspan=canvas_rowspan, 
            sticky='nsew'
        )
        # self.h_scroll = ttk.Scrollbar(self, orient='horizontal', command=self.canvas.xview)
        # self.v_scroll = ttk.Scrollbar(self, orient='vertical', command=self.canvas.yview)
        

    
        # Configure grid weights
        for i in range(self.rows + (1 if self.spreadsheet_mode else 0)):
            self.grid_frame.rowconfigure(i, weight=1)
        for i in range(self.cols + (1 if self.spreadsheet_mode else 0)):
            self.grid_frame.columnconfigure(i, weight=1)

        # Delegate canvas methods to the table instance
        self.yview = self.canvas.yview
        self.xview = self.canvas.xview
        self.yview_moveto = self.canvas.yview_moveto
        self.xview_moveto = self.canvas.xview_moveto
        self.configure = self.canvas.configure
        self.itemconfig = self.canvas.itemconfig
        self.create_rectangle = self.canvas.create_rectangle
        self.create_window = self.canvas.create_window
        self.create_line = self.canvas.create_line
        self.delete = self.canvas.delete  # This fixes the AttributeError

        # # Configure canvas scrolling
        # self.canvas.configure(
        #     xscrollcommand=self.h_scroll.set,
        #     yscrollcommand=self.v_scroll.set
        # )
        
        # # Grid layout
        # self.canvas.grid(row=0, column=0, sticky='nsew')
        # self.v_scroll.grid(row=0, column=1, sticky='ns')
        # self.h_scroll.grid(row=1, column=0, sticky='ew')

        # # Bind resize event
        self.canvas.bind("<Configure>", self._on_canvas_resize)
        self.parent.bind('<Configure>', lambda e: self._update_canvas_size())

    # def _setup_canvas(self):
    #     """Initialize canvas and scrollbars with proper delegation"""
    #     self.canvas = tk.Canvas(self, highlightthickness=0)
    #     self.h_scroll = ttk.Scrollbar(self, orient='horizontal', command=self.canvas.xview)
    #     self.v_scroll = ttk.Scrollbar(self, orient='vertical', command=self.canvas.yview)
        
    #     # Configure canvas scrolling
    #     self.canvas.configure(
    #         xscrollcommand=self.h_scroll.set,
    #         yscrollcommand=self.v_scroll.set
    #     )
        
    #     # Grid layout
    #     self.canvas.grid(row=0, column=0, sticky='nsew')
    #     self.v_scroll.grid(row=0, column=1, sticky='ns')
    #     self.h_scroll.grid(row=1, column=0, sticky='ew')
        
    #     # Configure resizing
    #     self.grid_rowconfigure(0, weight=1)
    #     self.grid_columnconfigure(0, weight=1)
        
    #     # Delegate canvas methods to the table instance
    #     self.yview = self.canvas.yview
    #     self.xview = self.canvas.xview
    #     self.yview_moveto = self.canvas.yview_moveto
    #     self.xview_moveto = self.canvas.xview_moveto
    #     self.configure = self.canvas.configure
    #     self.itemconfig = self.canvas.itemconfig
    #     self.create_rectangle = self.canvas.create_rectangle
    #     self.create_window = self.canvas.create_window
    #     self.create_line = self.canvas.create_line
    #     self.delete = self.canvas.delete  # This fixes the AttributeError

    #     # Bind resize event
        self.canvas.bind("<Configure>", self._on_canvas_resize)

    def auto_scroll_to_selection(self):
        """Automatically scroll to keep selection visible"""
        if not self.auto_scroll or not self.selected_cells:
            return
        
        # Get the first selected cell for scrolling
        row, col = next(iter(self.selected_cells))
        
        # Calculate visible area
        x_pos = col / self.cols
        y_pos = row / self.rows
        
        self.xview_moveto(max(0, min(x_pos, 1)))
        self.yview_moveto(max(0, min(y_pos, 1)))

    def update_selection(self, row, col):
        """Update the selected cell and focus"""
        # Truncate all previously selected cells
        for r, c in self.selected_cells.copy():
            if (r, c) in self.cells and not (r == row and c == col):
                self._truncate_text(r, c)
        
        # Skip if cell is merged (select the merge origin instead)
        if self.is_merged_cell(row, col):
            for (r, c), (span_r, span_c) in self.merged_cells.items():
                if (r <= row < r + span_r) and (c <= col < c + span_c):
                    row, col = r, c
                    break
        
        # Validate coordinates
        if not (0 <= row < self.rows and 0 <= col < self.cols):
            return
        
        # Reset previous selection
        for r, c in self.selected_cells.copy():
            cell = self.cells.get((r, c))
            if cell:
                self.canvas.itemconfig(cell['rect'], fill=cell['bg_color'])
                cell['text'].config(bg=cell['bg_color'])
        
        # Set new selection (single selection)
        self.selected_cells = {(row, col)}
        current_cell = self.cells.get((row, col))
        if current_cell:
            self.canvas.itemconfig(current_cell['rect'], 
                                 fill=self.current_theme['select_bg'])
            current_cell['text'].config(bg=self.current_theme['select_bg'])
            current_cell['text'].focus_set()
            # Set cursor to end for Text widget
            current_cell['text'].mark_set("insert", "end")
            current_cell['text'].see("insert")  # Ensure cursor is visible
            self.auto_scroll_to_selection()

    def on_click_cell(self, event, row, col):
        """Handle cell click events"""
        self.update_selection(row, col)

        if (row, col) in self.cells:
            text_widget = self.cells[(row, col)]['text']
            text_widget.focus_set()
            # Place cursor at click position or end of text
            text_widget.icursor(tk.END)

    # ... (keep all other methods the same, but ensure they use self.selected_cells)

    def draw_grid_lines(self):
        """Draw grid lines based on current row heights"""
        self.canvas.delete('grid_line')
        
        # Calculate cumulative heights
        y_positions = [0]
        for height in self.row_heights:
            y_positions.append(y_positions[-1] + height)
        
        # Horizontal lines
        for y in y_positions:
            self.canvas.create_line(
                0, y,
                self.cols * self.cell_width, y,
                fill=self.current_theme['grid'],
                tags="grid_line"
            )
        
        # Vertical lines
        for col in range(self.cols + 1):
            x = col * self.cell_width
            self.canvas.create_line(
                x, 0,
                x, y_positions[-1],
                fill=self.current_theme['grid'],
                tags="grid_line"
            )

    # def _setup_canvas(self):
    #     """Initialize canvas with resize binding"""
    #     self.canvas = tk.Canvas(self, highlightthickness=0)
    #     self.h_scroll = ttk.Scrollbar(self, orient='horizontal', command=self.canvas.xview)
    #     self.v_scroll = ttk.Scrollbar(self, orient='vertical', command=self.canvas.yview)
        
    #     self.canvas.configure(
    #         xscrollcommand=self.h_scroll.set,
    #         yscrollcommand=self.v_scroll.set
    #     )
        
    #     self.canvas.grid(row=0, column=0, sticky='nsew')
    #     self.v_scroll.grid(row=0, column=1, sticky='ns')
    #     self.h_scroll.grid(row=1, column=0, sticky='ew')
        
    #     self.grid_rowconfigure(0, weight=1)
    #     self.grid_columnconfigure(0, weight=1)
        
    #     # Bind resize event
    #     self.canvas.bind("<Configure>", self._on_canvas_resize)

    def _on_canvas_resize(self, event):
        """Handle canvas resize - retruncate all text"""
        for (row, col) in self.cells:
            self._truncate_text(row, col)


    def create_grid(self):
        """Create the grid of cells with theme support"""
        self.canvas.delete("all")
        self.cells = {}
        self.row_heights = [self.default_cell_height] * self.rows  # Reset row heights
        # Calculate positions relative to headers
        header_width = 30  # Width of row headers
        header_height = 25  # Height of column headers
        
        # Calculate cumulative heights
        cumulative_heights = [0]
        for height in self.row_heights:
            cumulative_heights.append(cumulative_heights[-1] + height)
            
        for row in range(self.rows):
            for col in range(self.cols):
                if self.is_merged_cell(row, col):
                    continue
                        
                span_rows, span_cols = self.get_merged_span(row, col)
                x1 = col * self.cell_width
                y1 = cumulative_heights[row]
                x2 = x1 + (span_cols * self.cell_width)
                y2 = y1 + sum(self.row_heights[row:row+span_rows])  # Sum of spanned rows
                
                bg_color = self.current_theme['even'] if row % 2 == 0 else self.current_theme['odd']
                
                rect = self.canvas.create_rectangle(
                    x1, y1, x2, y2,
                    fill=bg_color,
                    outline='',
                    tags=f"cell_{row}_{col}"
                )
                
                text = tk.Text(
                    self.canvas,
                    bg=bg_color,
                    fg=self.current_theme['fg'],
                    font=self.current_theme['font'],
                    relief='flat',
                    borderwidth=0,
                    height=1,
                    width=max(1, int((x2-x1)/7)),
                    wrap=tk.NONE,  # Disable wrapping
                    highlightthickness=0,
                    selectbackground=self.current_theme['select_bg'],
                    selectforeground=self.current_theme['select_fg'],
                    exportselection=0  # Important for proper selection handling
                )
                
                text_window = self.canvas.create_window(
                    x1 + 2, y1 + 2,
                    window=text,
                    anchor='nw',
                    width=x2 - x1 - 4,
                    height=y2 - y1 - 4,
                    tags=f"text_{row}_{col}"
                )
                
                self.cells[(row, col)] = {
                    'rect': rect,
                    'text': text,
                    'text_window': text_window,
                    'bg_color': bg_color,
                    'span_rows': span_rows,
                    'span_cols': span_cols
                }
                
                text.bind('<Button-1>', lambda e, r=row, c=col: self.on_click_cell(e, r, c))
                for key in ['<Up>', '<Down>', '<Left>', '<Right>', '<Tab>', '<Shift-Tab>']:
                    text.bind(key, lambda e: 'break')
                text.bind('<FocusOut>', lambda e, r=row, c=col: self.process_cell_edit(r, c))
                text.bind('<Return>', lambda e, r=row, c=col: self.process_cell_edit(r, c))
                # Add key bindings for navigation
                text.bind('<Up>', lambda e, r=row, c=col: self._navigate_cell(e, r, c, 'up'))
                text.bind('<Down>', lambda e, r=row, c=col: self._navigate_cell(e, r, c, 'down'))
                text.bind('<Left>', lambda e, r=row, c=col: self._navigate_cell(e, r, c, 'left'))
                text.bind('<Right>', lambda e, r=row, c=col: self._navigate_cell(e, r, c, 'right'))
                text.bind('<Tab>', lambda e, r=row, c=col: self._navigate_cell(e, r, c, 'tab'))
                text.bind('<Shift-Tab>', lambda e, r=row, c=col: self._navigate_cell(e, r, c, 'shift_tab'))
        
        self.draw_grid_lines()
        self.canvas.configure(
            scrollregion=(0, 0, 
                         self.cols * self.cell_width, 
                         self.rows * self.cell_height)
        )

    def _navigate_cell(self, event, current_row, current_col, direction):
        """Handle keyboard navigation between cells"""
        new_row, new_col = current_row, current_col
        
        if direction == 'up':
            new_row = max(0, current_row - 1)
        elif direction == 'down':
            new_row = min(self.rows - 1, current_row + 1)
        elif direction == 'left':
            new_col = max(0, current_col - 1)
        elif direction == 'right':
            new_col = min(self.cols - 1, current_col + 1)
        elif direction == 'tab':
            if event.state & 0x0001:  # Shift key (Shift+Tab)
                new_col = current_col - 1
                if new_col < 0:
                    new_col = self.cols - 1
                    new_row = max(0, current_row - 1)
            else:  # Regular Tab
                new_col = current_col + 1
                if new_col >= self.cols:
                    new_col = 0
                    new_row = min(self.rows - 1, current_row + 1)
        elif direction == 'shift_tab':
            new_col = current_col - 1
            if new_col < 0:
                new_col = self.cols - 1
                new_row = max(0, current_row - 1)
        
        # Skip merged cells (except the top-left cell of a merged range)
        while self.is_merged_cell(new_row, new_col) and not ((new_row, new_col) in self.merged_cells):
            # Adjust position based on direction
            if direction in ['up', 'down']:
                new_row = max(0, new_row - 1) if direction == 'up' else min(self.rows - 1, new_row + 1)
            else:
                new_col = max(0, new_col - 1) if direction in ['left', 'shift_tab'] else min(self.cols - 1, new_col + 1)
        
        # Only move if we're actually changing cells
        if (new_row != current_row) or (new_col != current_col):
            self.update_selection(new_row, new_col)
        
        return 'break'  # Prevent default behavior

    def _setup_event_bindings(self):
        """Add hover effects for references"""
        # Highlight column when hovering header
        for col, header in enumerate(self.col_headers):
            header.bind("<Enter>", lambda e, c=col: self._highlight_column(c))
            header.bind("<Leave>", lambda e: self._remove_highlight())
        
        # Highlight row when hovering header
        for row, header in enumerate(self.row_headers):
            header.bind("<Enter>", lambda e, r=row: self._highlight_row(r))
            header.bind("<Leave>", lambda e: self._remove_highlight())

    def _highlight_column(self, col):
        """Visual feedback for column reference"""
        for row in range(self.rows):
            if (row, col) in self.cells:
                self.canvas.itemconfig(
                    self.cells[(row, col)]['rect'],
                    fill='#e6f3ff'  # Light blue highlight
                )

    def _highlight_row(self, row):
        """Visual feedback for row reference"""
        for col in range(self.cols):
            if (row, col) in self.cells:
                self.canvas.itemconfig(
                    self.cells[(row, col)]['rect'],
                    fill='#e6f3ff'  # Light blue highlight
                )

    def _remove_highlight(self):
        """Remove visual highlights"""
        for (row, col), cell in self.cells.items():
            bg_color = self.current_theme['even'] if row % 2 == 0 else self.current_theme['odd']
            self.canvas.itemconfig(cell['rect'], fill=bg_color)

    def _truncate_text(self, row, col):
        """Truncate text with ellipsis if it's too long for the cell"""
        if (row, col) not in self.cells:
            return
        
        text_widget = self.cells[(row, col)]['text']
        full_text = text_widget.get("1.0", "end-1c")
        
        # Calculate available width in characters
        cell_width = text_widget.winfo_width()
        avg_char_width = self.font.measure("M")
        max_chars = max(3, int(cell_width / avg_char_width)) if avg_char_width > 0 else 20
        
        if len(full_text) > max_chars:
            # Truncate and add ellipsis
            truncated = full_text[:max_chars-3] + "..."
            text_widget.delete("1.0", "end")
            text_widget.insert("1.0", truncated)
            text_widget.tag_add("truncated", "1.0", "end")
            text_widget.tag_config("truncated", elide=True)  # Hide the visual text
            text_widget.insert("1.0", full_text)  # But keep full text stored
        else:
            text_widget.delete("1.0", "end")
            text_widget.insert("1.0", full_text)

    def on_click_cell(self, event, row, col):
        """Handle cell click - show full text when selected"""
        self.update_selection(row, col)
        
        if (row, col) in self.cells:
            text_widget = self.cells[(row, col)]['text']
            # For Text widget, set cursor at click position
            text_widget.focus_set()
            text_widget.mark_set("insert", f"@{event.x},{event.y}")
            text_widget.see("insert")


    def add_theme(self, name, config):
        """
        Add a custom theme configuration
        
        Args:
            name (str): Theme name
            config (dict): Theme configuration with keys:
                - bg: background color
                - fg: text color
                - select_bg: selection background
                - select_fg: selection text
                - even: even row color
                - odd: odd row color
                - grid: grid line color
                - font: text font
        """
        required_keys = ['bg', 'fg', 'select_bg', 'select_fg', 'even', 'odd', 'grid', 'font']
        if all(key in config for key in required_keys):
            self.themes[name] = config
        else:
            raise ValueError("Theme config must contain all required keys")

    def get_available_themes(self):
        """Return list of available theme names"""
        return list(self.themes.keys())

    def is_merged_cell(self, row, col):
        """Check if cell is part of a merged range"""
        for (r, c), (span_r, span_c) in self.merged_cells.items():
            if (r <= row < r + span_r) and (c <= col < c + span_c) and not (r == row and c == col):
                return True
        return False

    def get_merged_span(self, row, col):
        """Get merged span for cell (default 1x1 if not merged)"""
        return self.merged_cells.get((row, col), (1, 1))

    def merge_cells(self, start_row, start_col, span_rows=1, span_cols=1):
        """Merge a range of cells"""
        # Validate merge area
        self.save_state(f"Merge cells at ({start_row},{start_col})")
        if (start_row + span_rows > self.rows or start_col + span_cols > self.cols):
            raise ValueError("Merge area exceeds grid dimensions")
            
        # Clear existing merged cells in this range
        for r in range(start_row, start_row + span_rows):
            for c in range(start_col, start_col + span_cols):
                if (r, c) in self.merged_cells:
                    self.unmerge_cells(r, c)
        
        # Store merge info
        self.merged_cells[(start_row, start_col)] = (span_rows, span_cols)
        
        # Recreate grid with new merged cells
        self.create_grid()
        
        # Update selection to the merged cell
        if self.selected_cells:
            # Get the first selected cell's row
            first_selected_row = next(iter(self.selected_cells))[0]
            first_selected_col = next(iter(self.selected_cells))[1]
            self.update_selection(min(start_row, first_selected_row), 
                                min(start_col, first_selected_col))
        # After merging:
        
        self.refresh_grid()

    def merge_selected(self):
        """Merge currently selected cells"""
        if not self.selected_cells:
            return
            
        # Find bounding rectangle of selection
        rows = {r for r, c in self.selected_cells}
        cols = {c for r, c in self.selected_cells}
        
        min_row, max_row = min(rows), max(rows)
        min_col, max_col = min(cols), max(cols)
        
        # Merge the rectangular area
        self.merge_cells(
            min_row, min_col,
            max_row - min_row + 1,
            max_col - min_col + 1
        )

    def unmerge_cells(self, row, col):
        """Unmerge cells starting at row,col"""
        self.save_state(f"Unmerge cells at ({row},{col})")
        if (row, col) in self.merged_cells:
            del self.merged_cells[(row, col)]
            self.create_grid()
            self.update_selection(row, col)
        self.refresh_grid()


    def select_cell(self, row, col):
        """Select a single cell (maintaining single selection)"""
        if not (0 <= row < self.rows and 0 <= col < self.cols):
            return
            
        self.update_selection(row, col)  # Reuse the unified method

    def select_range(self, start_row, start_col, end_row, end_col):
        """Select range of cells (for multiple selection mode)"""
        if self.selection_mode == "single":
            self.update_selection(end_row, end_col)
            return
            
        # For multiple selection mode
        new_selection = set()
        for row in range(min(start_row, end_row), max(start_row, end_row) + 1):
            for col in range(min(start_col, end_col), max(start_col, end_col) + 1):
                if (row, col) in self.cells:
                    new_selection.add((row, col))
        
        # Clear previous selection
        for r, c in self.selected_cells - new_selection:
            cell = self.cells.get((r, c))
            if cell:
                self.canvas.itemconfig(cell['rect'], fill=cell['bg_color'])
                cell['text'].config(bg=cell['bg_color'])
        
        # Set new selection
        for r, c in new_selection:
            cell = self.cells.get((r, c))
            if cell:
                self.canvas.itemconfig(cell['rect'], fill=self.current_theme['select_bg'])
                cell['text'].config(bg=self.current_theme['select_bg'])
        
        self.selected_cells = new_selection
        if new_selection:
            self.auto_scroll_to_selection(*next(iter(new_selection)))

    def clear_selection(self):
        """Clear all selections using selected_cells"""
        for row, col in self.selected_cells.copy():
            cell = self.cells.get((row, col))
            if cell:
                self.canvas.itemconfig(cell['rect'], fill=cell['bg_color'])
                cell['text'].config(bg=cell['bg_color'])
        self.selected_cells.clear()

    def get_selected_cell(self):
        """Get the primary selected cell (for single selection)"""
        if not self.selected_cells:
            return None
        return next(iter(self.selected_cells))


    def on_resize(self, event):
        """Improved resize handler for clean redraw"""
        if event.width < 100 or event.height < 100:
            return
        
        # Calculate new cell sizes (ensuring whole numbers)
        new_cell_width = max(50, int(event.width / self.cols))
        new_cell_height = max(20, int(event.height / self.rows))
        
        if (new_cell_width != self.cell_width or 
            new_cell_height != self.cell_height):
            self.cell_width = new_cell_width
            self.cell_height = new_cell_height
            self.create_grid()
            self.configure_scroll()

    def on_click(self, event):
        """Handle mouse clicks on the canvas"""
        col = min(max(0, int(event.x // self.cell_width)), self.cols - 1)
        row = min(max(0, int(event.y // self.cell_height)), self.rows - 1)
        self.update_selection(row, col)

    def on_key(self, event):
        """Handle keyboard navigation"""
        row, col = self.selected_cells
        
        if event.keysym == 'Up':
            row = max(0, row - 1)
        elif event.keysym == 'Down':
            row = min(self.rows - 1, row + 1)
        elif event.keysym == 'Left':
            col = max(0, col - 1)
        elif event.keysym == 'Right':
            col = min(self.cols - 1, col + 1)
        elif event.keysym == 'Tab':
            if event.state & 0x0001:  # Shift+Tab
                col -= 1
                if col < 0:
                    col = self.cols - 1
                    row = max(0, row - 1)
            else:  # Regular Tab
                col += 1
                if col >= self.cols:
                    col = 0
                    row = min(self.rows - 1, row + 1)
        else:
            return  # Let other keys be handled normally
        
        self.update_selection(row, col)
        return 'break'

    def set_auto_scroll(self, enabled):
        """Enable/disable auto-scroll to selection"""
        self.auto_scroll = enabled

    def get_values(self):
        """Return table data in multiple formats"""
        data = []
        for row in range(self.rows):
            row_data = []
            for col in range(self.cols):
                if (row, col) in self.cells:
                    text = self.cells[(row, col)]['text']
                    content = text.get('1.0', 'end-1c')
                    row_data.append(content)
                else:
                    row_data.append("")  # For merged cells
            data.append(row_data)
        
        return {
            'dict': {f'Row {i}': row for i, row in enumerate(data)},
            'dataframe': pd.DataFrame(data),
            'array': np.array(data)
        }

    def set_values(self, data):
        """Populate table with data"""
        self.save_state("Set table values")
        for row in range(min(self.rows, len(data))):
            for col in range(min(self.cols, len(data[row]))):
                if (row, col) in self.cells:
                    self.cells[(row, col)]['text'].delete('1.0', 'end')
                    self.cells[(row, col)]['text'].insert('1.0', str(data[row][col]))

    def load_dataframe(self, df):
        """Load data from a pandas DataFrame"""
        self.resize_grid(len(df), len(df.columns))
        self.set_values(df.values.tolist())

    def get_cell(self, row, col, raw=False):
        """
        Get cell content with optional merged cell handling
        
        Args:
            row (int): Row index (0-based)
            col (int): Column index (0-based)
            raw (bool): If True, returns content of exactly this cell.
                       If False (default), returns content of merge origin for merged cells.
        
        Returns:
            dict: {
                'value': str, 
                'is_merged': bool,
                'merge_span': (rows, cols) or None,
                'bg_color': str,
                'is_selected': bool
            }
        
        Raises:
            IndexError: If coordinates are out of bounds
        """
        # Validate coordinates
        if not (0 <= row < self.rows and 0 <= col < self.cols):
            raise IndexError(f"Cell ({row},{col}) out of bounds (table size {self.rows}x{self.cols})")
        
        # Handle merged cells
        if not raw and self.is_merged_cell(row, col):
            for (r, c), (span_r, span_c) in self.merged_cells.items():
                if r <= row < r + span_r and c <= col < c + span_c:
                    row, col = r, c
                    break
        
        # Get cell data
        cell_data = {
            'value': '',
            'is_merged': False,
            'merge_span': None,
            'bg_color': self.current_theme['even'] if row % 2 == 0 else self.current_theme['odd'],
            'is_selected': (row, col) in self.selected_cells
        }
        
        if (row, col) in self.cells:
            text_widget = self.cells[(row, col)]['text']
            cell_data['value'] = text_widget.get("1.0", "end-1c")
            
            if (row, col) in self.merged_cells:
                cell_data['is_merged'] = True
                cell_data['merge_span'] = self.merged_cells[(row, col)]
                cell_data['bg_color'] = self.cells[(row, col)]['bg_color']
        
        return cell_data

    def set_cell(self, row, col, value, expand_merged=False):
        """
        Set cell content with merged cell handling
        
        Args:
            row (int): Row index (0-based)
            col (int): Column index (0-based)
            value (str): Content to set
            expand_merged (bool): If True, sets value to all merged cells.
                                 If False (default), only sets top-left cell.
        
        Returns:
            bool: True if cell was modified, False if out of bounds
        
        Examples:
            # Basic usage
            table.set_cell(0, 0, "Hello")
            
            # Force set all merged cells
            table.set_cell(1, 1, "Merged", expand_merged=True)
        """
        # Validate coordinates
        if not (0 <= row < self.rows and 0 <= col < self.cols):
            return False
        
        # Handle merged cells
        target_cells = [(row, col)]
        if not expand_merged and self.is_merged_cell(row, col):
            for (r, c), (span_r, span_c) in self.merged_cells.items():
                if r <= row < r + span_r and c <= col < c + span_c:
                    target_cells = [(r, c)]  # Only modify merge origin
                    break
        
        elif expand_merged and (row, col) in self.merged_cells:
            span_r, span_c = self.merged_cells[(row, col)]
            target_cells = [
                (r, c) 
                for r in range(row, row + span_r)
                for c in range(col, col + span_c)
                if (r, c) in self.cells
            ]
        
        # Update cells
        modified = False
        for r, c in target_cells:
            if (r, c) in self.cells:
                text_widget = self.cells[(r, c)]['text']
                text_widget.delete("1.0", "end")
                text_widget.insert("1.0", str(value))
                modified = True
                
                # Recalculate formulas if in spreadsheet mode
                if self.spreadsheet_mode and (r, c) in self.formulas:
                    self._evaluate_cell(r, c)
        
        return modified

    def get_selected_cell_value(self):
        """Convenience method to get value from first selected cell"""
        if not self.selected_cells:
            return None
        row, col = next(iter(self.selected_cells))
        return self.get_cell(row, col)['value']

    def set_selected_cell_value(self, value):
        """Convenience method to set value in all selected cells"""
        if not self.selected_cells:
            return False
        results = []
        for row, col in self.selected_cells:
            results.append(self.set_cell(row, col, value))
        return any(results)


        
    def set_selection_mode(self, mode):
        """Set selection mode: 'single' or 'multiple'"""
        if mode in ["single", "multiple"]:
            self.selection_mode = mode
        else:
            raise ValueError("Selection mode must be 'single' or 'multiple'")

    def clear_selection(self):
        """Clear all cell selections"""
        for row, col in self.selected_cells.copy():
            self.deselect_cell(row, col)
        self.selected_cells.clear()

    def select_cell(self, row, col):
        """Select a single cell"""
        if not (0 <= row < self.rows and 0 <= col < self.cols):
            return
            
        if self.selection_mode == "single":
            self.clear_selection()
            
        if (row, col) not in self.selected_cells:
            cell = self.cells.get((row, col))
            if cell:
                self.itemconfig(cell['rect'], fill = self.current_theme['select_bg'])
                self.selected_cells.add((row, col))

    def deselect_cell(self, row, col):
        """Deselect a cell"""
        cell = self.cells.get((row, col))
        if cell and (row, col) in self.selected_cells:
            self.itemconfig(cell['rect'], fill=cell['bg_color'])
            self.selected_cells.discard((row, col))

    def select_range(self, start_row, start_col, end_row, end_col):
        """Select a range of cells"""
        self.clear_selection()
        for row in range(min(start_row, end_row), max(start_row, end_row) + 1):
            for col in range(min(start_col, end_col), max(start_col, end_col) + 1):
                self.select_cell(row, col)

    def on_drag(self, event):
        """Handle mouse drag for selection"""
        if self.selection_mode != "multiple":
            return
            
        col = min(max(0, int(event.x // self.cell_width)), self.cols - 1)
        row = min(max(0, int(event.y // self.cell_height)), self.rows - 1)
        
        if not self.selection_start:
            self.selection_start = (row, col)
            return
            
        start_row, start_col = self.selection_start
        self.select_range(start_row, start_col, row, col)
        
        # Draw rubber band
        if self.selection_rect:
            self.delete(self.selection_rect)
            
        x1 = start_col * self.cell_width
        y1 = start_row * self.cell_height
        x2 = (col + 1) * self.cell_width
        y2 = (row + 1) * self.cell_height
        
        self.selection_rect = self.create_rectangle(
            x1, y1, x2, y2,
            outline='blue',
            dash=(2, 2),
            width=1,
            tags="selection_rect"
        )

    def on_drag_end(self, event):
        """Clean up after drag selection"""
        if self.selection_rect:
            self.delete(self.selection_rect)
            self.selection_rect = None
        self.selection_start = None

    def on_ctrl_click(self, event):
        """Add/remove cell from selection with Ctrl+Click"""
        col = min(max(0, int(event.x // self.cell_width)), self.cols - 1)
        row = min(max(0, int(event.y // self.cell_height)), self.rows - 1)
        
        if (row, col) in self.selected_cells:
            self.deselect_cell(row, col)
        else:
            self.select_cell(row, col)
        return "break"  # Prevent default behavior

    def on_shift_click(self, event):
        """Extend selection with Shift+Click"""
        if not self.selected_cells:
            return
            
        col = min(max(0, int(event.x // self.cell_width)), self.cols - 1)
        row = min(max(0, int(event.y // self.cell_height)), self.rows - 1)
        
        # Use first selected cell as anchor
        first_row, first_col = next(iter(self.selected_cells))
        self.select_range(first_row, first_col, row, col)
        return "break"

    def get_selected_values(self):
        """Return values from selected cells"""
        values = []
        for row, col in sorted(self.selected_cells):
            cell = self.cells.get((row, col))
            if cell:
                values.append({
                    'row': row,
                    'col': col,
                    'value': cell['text'].get("1.0", "end-1c")
                })
        return values

    def merge_selected(self):
        """Merge currently selected cells"""
        if not self.selected_cells:
            return
            
        # Find bounding rectangle of selection
        rows = {r for r, c in self.selected_cells}
        cols = {c for r, c in self.selected_cells}
        
        min_row, max_row = min(rows), max(rows)
        min_col, max_col = min(cols), max(cols)
        
        # Merge the rectangular area
        self.merge_cells(
            min_row, min_col,
            max_row - min_row + 1,
            max_col - min_col + 1
        )

    def resize_selected(self, new_width=None, new_height=None):
        """Resize selected columns/rows"""
        if not self.selected_cells:
            return
            
        # Get unique affected columns and rows
        affected_cols = {c for r, c in self.selected_cells}
        affected_rows = {r for r, c in self.selected_cells}
        
        # Resize columns
        if new_width:
            self.default_cell_width = new_width
            for col in affected_cols:
                # Update all cells in column
                for row in range(self.rows):
                    if (row, col) in self.cells:
                        self.cells[(row, col)]['text'].config(width=new_width//7)
        
        # Resize rows
        if new_height:
            self.default_cell_height = new_height
            for row in affected_rows:
                # Update all cells in row
                for col in range(self.cols):
                    if (row, col) in self.cells:
                        self.cells[(row, col)]['text'].config(height=1)


    def insert_row(self, position="below"):
        """
        Insert a new row relative to current selection
        
        Args:
            position: "above" or "below" current row
        """
        self.save_state(f"Insert row {position}")
        if not self.selected_cells:
            return
            
        ref_row = min(r for r, c in self.selected_cells)
        
        # Calculate insertion index
        insert_at = ref_row + 1 if position == "below" else ref_row
        
        # Update data structure
        new_cells = {}
        for (r, c), cell in sorted(self.cells.items()):
            if r >= insert_at:
                new_cells[(r + 1, c)] = cell
                # Update merged cell references
                if (r, c) in self.merged_cells:
                    span_r, span_c = self.merged_cells.pop((r, c))
                    self.merged_cells[(r + 1, c)] = (span_r, span_c)
            else:
                new_cells[(r, c)] = cell
        
        self.cells = new_cells
        self.rows += 1
        
        # Recreate grid
        self.create_grid()
        self._update_canvas_size()
        
        # Adjust selection
        new_selection = set()
        for r, c in self.selected_cells:
            new_r = r + 1 if r >= insert_at else r
            new_selection.add((new_r, c))
        self.selected_cells = new_selection
        self.update_selection(*next(iter(new_selection)))

        # After modifying the grid:
        self.row_heights.insert(insert_at, self.default_cell_height)
        
        self.refresh_grid()

    def delete_row(self):
        """Delete currently selected row(s)"""
        self.save_state("Delete row")
        if not self.selected_cells:
            return
            
        rows_to_delete = {r for r, c in self.selected_cells}
        
        # Update data structure
        new_cells = {}
        deleted_indices = set()
        
        for (r, c), cell in sorted(self.cells.items()):
            if r in rows_to_delete:
                deleted_indices.add(r)
                continue
                
            new_r = r - sum(1 for dr in deleted_indices if dr < r)
            new_cells[(new_r, c)] = cell
            
            # Update merged cell references
            if (r, c) in self.merged_cells:
                span_r, span_c = self.merged_cells.pop((r, c))
                self.merged_cells[(new_r, c)] = (span_r, span_c)
        
        self.cells = new_cells
        self.rows -= len(rows_to_delete)
        
        # Recreate grid
        self.create_grid()
        
        # Clear selection
        self.clear_selection()
        # After modifying the grid:
        for row in sorted(rows_to_delete, reverse=True):
            del self.row_heights[row]
        
        self.refresh_grid()

    def insert_column(self, position="right"):
        """
        Insert new column relative to current selection
        
        Args:
            position: "left" or "right" of current column
        """
        self.save_state(f"Insert column {position}")
        if not self.selected_cells:
            return
            
        ref_col = min(c for r, c in self.selected_cells)
        
        # Calculate insertion index
        insert_at = ref_col + 1 if position == "right" else ref_col
        
        # Update data structure
        new_cells = {}
        for (r, c), cell in sorted(self.cells.items()):
            if c >= insert_at:
                new_cells[(r, c + 1)] = cell
                # Update merged cell references
                if (r, c) in self.merged_cells:
                    span_r, span_c = self.merged_cells.pop((r, c))
                    self.merged_cells[(r, c + 1)] = (span_r, span_c)
            else:
                new_cells[(r, c)] = cell
        
        self.cells = new_cells
        self.cols += 1
        
        # Recreate grid
        self.create_grid()
        self._update_canvas_size()
        
        # Adjust selection
        new_selection = set()
        for r, c in self.selected_cells:
            new_c = c + 1 if c >= insert_at else c
            new_selection.add((r, new_c))
        self.selected_cells = new_selection
        self.update_selection(*next(iter(new_selection)))
        # After modifying the grid:
        self.row_heights.insert(insert_at, self.default_cell_height)
        
        self.refresh_grid()

    def delete_column(self):
        """Delete currently selected column(s)"""
        self.save_state("Delete column")
        if not self.selected_cells:
            return
            
        cols_to_delete = {c for r, c in self.selected_cells}
        
        # Update data structure
        new_cells = {}
        deleted_indices = set()
        
        for (r, c), cell in sorted(self.cells.items()):
            if c in cols_to_delete:
                deleted_indices.add(c)
                continue
                
            new_c = c - sum(1 for dc in deleted_indices if dc < c)
            new_cells[(r, new_c)] = cell
            
            # Update merged cell references
            if (r, c) in self.merged_cells:
                span_r, span_c = self.merged_cells.pop((r, c))
                self.merged_cells[(r, new_c)] = (span_r, span_c)
        
        self.cells = new_cells
        self.cols -= len(cols_to_delete)
        
        # Recreate grid
        self.create_grid()
        
        # Clear selection
        self.clear_selection()
        for row in sorted(rows_to_delete, reverse=True):
            del self.row_heights[row]
        
        self.refresh_grid()

    def move_row(self, direction="down"):
        """
        Move selected row up or down
        
        Args:
            direction: "up" or "down"
        """
        self.save_state(f"Move row {direction}")
        if not self.selected_cells:
            return
            
        rows = {r for r, c in self.selected_cells}
        if len(rows) != 1:
            messagebox.showwarning("Move Row", "Select exactly one row to move")
            return
            
        row = rows.pop()
        new_pos = row + 1 if direction == "down" else row - 1
        
        # Validate new position
        if new_pos < 0 or new_pos >= self.rows:
            return
            
        # Swap row data
        for c in range(self.cols):
            if (row, c) in self.cells and (new_pos, c) in self.cells:
                # Swap text content
                temp = self.cells[(row, c)]['text'].get("1.0", "end-1c")
                self.cells[(row, c)]['text'].delete("1.0", "end")
                self.cells[(row, c)]['text'].insert("1.0", 
                    self.cells[(new_pos, c)]['text'].get("1.0", "end-1c"))
                self.cells[(new_pos, c)]['text'].delete("1.0", "end")
                self.cells[(new_pos, c)]['text'].insert("1.0", temp)
                
                # Swap merge status if needed
                if (row, c) in self.merged_cells:
                    self.merged_cells[(new_pos, c)] = self.merged_cells.pop((row, c))
                elif (new_pos, c) in self.merged_cells:
                    self.merged_cells[(row, c)] = self.merged_cells.pop((new_pos, c))
        
        # Update selection
        self.clear_selection()
        for c in range(self.cols):
            self.select_cell(new_pos, c)
        
        self.update_selection(new_pos, 0)

        # After moving:
        self.row_heights[row], self.row_heights[new_pos] = self.row_heights[new_pos], self.row_heights[row]
        self.refresh_grid()

    def set_row_height(self, row, height):
        """Manually set row height"""
        if 0 <= row < self.rows:
            self.row_heights[row] = height
            self.refresh_grid()

    def refresh_grid(self):
        """Redraw grid while preserving content and selection"""
        content = {k: v['text'].get("1.0", "end-1c") 
                  for k, v in self.cells.items()}
        selection = list(self.selected_cells)
        
        self.create_grid()
        
        # Restore content
        for (r, c), text in content.items():
            if (r, c) in self.cells:
                self.cells[(r, c)]['text'].insert("1.0", text)
        
        # Restore selection
        self.clear_selection()
        for r, c in selection:
            if (r, c) in self.cells:
                self.select_cell(r, c)

    def move_column(self, direction="right"):
        """
        Move selected column left or right
        
        Args:
            direction: "left" or "right"
        """
        self.save_state(f"Move column {direction}")
        if not self.selected_cells:
            return
            
        cols = {c for r, c in self.selected_cells}
        if len(cols) != 1:
            messagebox.showwarning("Move Column", "Select exactly one column to move")
            return
            
        col = cols.pop()
        new_pos = col + 1 if direction == "right" else col - 1
        
        # Validate new position
        if new_pos < 0 or new_pos >= self.cols:
            return
            
        # Swap column data
        for r in range(self.rows):
            if (r, col) in self.cells and (r, new_pos) in self.cells:
                # Swap text content
                temp = self.cells[(r, col)]['text'].get("1.0", "end-1c")
                self.cells[(r, col)]['text'].delete("1.0", "end")
                self.cells[(r, col)]['text'].insert("1.0", 
                    self.cells[(r, new_pos)]['text'].get("1.0", "end-1c"))
                self.cells[(r, new_pos)]['text'].delete("1.0", "end")
                self.cells[(r, new_pos)]['text'].insert("1.0", temp)
                
                # Swap merge status if needed
                if (r, col) in self.merged_cells:
                    self.merged_cells[(r, new_pos)] = self.merged_cells.pop((r, col))
                elif (r, new_pos) in self.merged_cells:
                    self.merged_cells[(r, col)] = self.merged_cells.pop((r, new_pos))
        
        # Update selection
        self.clear_selection()
        for r in range(self.rows):
            self.select_cell(r, new_pos)
        
        self.update_selection(0, new_pos)
        # After moving:
        self.refresh_grid()


    def split_cell(self, row, col, horizontal=True, vertical=True):
        """
        Split a merged cell either horizontally, vertically, or both
        
        Args:
            row (int): Row index of the cell to split
            col (int): Column index of the cell to split
            horizontal (bool): Split horizontally (default True)
            vertical (bool): Split vertically (default True)
        """
        # Check if this cell is actually a merged cell
        self.save_state("Split grid")
        if (row, col) not in self.merged_cells:
            return False  # Not a merged cell
        
        # Get the merge span
        span_rows, span_cols = self.merged_cells[(row, col)]
        
        # Can't split single cells
        if span_rows == 1 and span_cols == 1:
            return False
        
        # Determine new spans after splitting
        new_span_rows = span_rows
        new_span_cols = span_cols
        
        if horizontal and span_rows > 1:
            new_span_rows = 1  # Split horizontally by removing row span
        
        if vertical and span_cols > 1:
            new_span_cols = 1  # Split vertically by removing column span
        
        # If we're not actually changing anything
        if new_span_rows == span_rows and new_span_cols == span_cols:
            return False
        
        # Save the cell content
        cell_content = ""
        if (row, col) in self.cells:
            cell_content = self.cells[(row, col)]['text'].get("1.0", "end-1c")
        
        # Remove the merged cell
        del self.merged_cells[(row, col)]
        
        # If we're splitting completely (both directions)
        if new_span_rows == 1 and new_span_cols == 1:
            # No need to create new merged cells
            pass
        else:
            # Create new merged cell with reduced span
            self.merged_cells[(row, col)] = (new_span_rows, new_span_cols)
        
        # Recreate the grid
        self.create_grid()
        
        # Restore content to the original cell
        if (row, col) in self.cells:
            self.cells[(row, col)]['text'].insert("1.0", cell_content)
        
        # Update selection to the original cell
        self.update_selection(row, col)
        return True

        # After merging:
        
        self.refresh_grid()

    def split_selected(self, horizontal=True, vertical=True):
        """
        Split currently selected merged cells
        
        Args:
            horizontal (bool): Split horizontally (default True)
            vertical (bool): Split vertically (default True)
        """
        self.save_state("Split selected grid")
        if not self.selected_cells:
            return False
        
        # Find all merged cells in selection
        merged_in_selection = []
        for (row, col) in self.selected_cells:
            # Check if this cell is the top-left of a merged area
            if (row, col) in self.merged_cells:
                merged_in_selection.append((row, col))
            else:
                # Check if it's part of a merged area
                for (r, c), (span_r, span_c) in self.merged_cells.items():
                    if r <= row < r + span_r and c <= col < c + span_c:
                        merged_in_selection.append((r, c))
                        break
        
        # Split each merged cell found
        for cell in set(merged_in_selection):  # Remove duplicates
            self.split_cell(cell[0], cell[1], horizontal, vertical)
        
        return len(merged_in_selection) > 0
        # After merging:
        
        self.refresh_grid()

        

    def save_state(self, description=""):
        """Save current table state to undo stack"""
        if len(self.undo_stack) >= self.max_undo_steps:
            self.undo_stack.pop(0)
            
        state = {
            'cells': self._get_cell_contents(),
            'merged': dict(self.merged_cells),
            'dimensions': (self.rows, self.cols),
            'selection': list(self.selected_cells),
            'description': description
        }
        self.undo_stack.append(state)
        self.redo_stack = []  # Clear redo stack on new action

    def _get_cell_contents(self):
        """Capture current cell contents"""
        return {
            (r, c): cell['text'].get("1.0", "end-1c")
            for (r, c), cell in self.cells.items()
        }

    def _restore_state(self, state):
        """Restore table state from saved state"""
        # Restore dimensions if changed
        if (self.rows, self.cols) != state['dimensions']:
            self.rows, self.cols = state['dimensions']
        
        # Restore cell contents
        for (r, c), content in state['cells'].items():
            if (r, c) in self.cells:
                self.cells[(r, c)]['text'].delete("1.0", "end")
                self.cells[(r, c)]['text'].insert("1.0", content)
        
        # Restore merged cells
        self.merged_cells = dict(state['merged'])
        
        # Rebuild grid if dimensions changed
        if (self.rows, self.cols) != state['dimensions']:
            self.create_grid()
        
        # Restore selection
        self.clear_selection()
        for r, c in state['selection']:
            if (r, c) in self.cells:
                self.select_cell(r, c)

    def undo(self, event=None):
        """Undo the last operation"""
        if not self.undo_stack:
            return
            
        # Save current state to redo stack
        current_state = {
            'cells': self._get_cell_contents(),
            'merged': dict(self.merged_cells),
            'dimensions': (self.rows, self.cols),
            'selection': list(self.selected_cells),
            'description': "Before undo"
        }
        self.redo_stack.append(current_state)
        
        # Restore previous state
        state = self.undo_stack.pop()
        self._restore_state(state)
        
        return "break"  # Prevent default binding

    def redo(self, event=None):
        """Redo the last undone operation"""
        if not self.redo_stack:
            return
            
        # Save current state to undo stack
        current_state = {
            'cells': self._get_cell_contents(),
            'merged': dict(self.merged_cells),
            'dimensions': (self.rows, self.cols),
            'selection': list(self.selected_cells),
            'description': "Before redo"
        }
        self.undo_stack.append(current_state)
        
        # Restore next state
        state = self.redo_stack.pop()
        self._restore_state(state)
        
        return "break"  # Prevent default binding

    # Modified existing methods to support undo/redo:

    def configure_scroll(self):
        """Configure canvas scrolling settings"""
        self.configure(
            scrollregion=(0, 0, 
                         self.cols * self.cell_width, 
                         self.rows * self.cell_height),
            highlightthickness=0
        )
        if self.auto_scroll:
            self.xview_moveto(0)
            self.yview_moveto(0)

    def resize_grid(self, new_rows=None, new_cols=None):
        """Resize grid and update references"""
        old_rows, old_cols = self.rows, self.cols
        self.rows = new_rows if new_rows else self.rows
        self.cols = new_cols if new_cols else self.cols
        
        # Update headers if in spreadsheet mode
        if self.spreadsheet_mode:
            # Update column headers
            for col, header in enumerate(self.col_headers):
                if col < self.cols:
                    header.config(text=chr(65 + col))
                else:
                    header.grid_remove()
            
            # Add/remove column headers
            while len(self.col_headers) < self.cols:
                col = len(self.col_headers)
                header = ttk.Label(
                    self.grid_frame,
                    text=chr(65 + col),
                    width=self.cell_width//7,
                    anchor='center',
                    style='Header.TLabel'
                )
                header.grid(row=0, column=col+1, sticky='nsew')
                self.col_headers.append(header)
            
            # Update row headers
            for row, header in enumerate(self.row_headers):
                if row < self.rows:
                    header.config(text=str(row + 1))
                else:
                    header.grid_remove()
            
            # Add/remove row headers
            while len(self.row_headers) < self.rows:
                row = len(self.row_headers)
                header = ttk.Label(
                    self.grid_frame,
                    text=str(row + 1),
                    width=4,
                    anchor='e',
                    style='Header.TLabel'
                )
                header.grid(row=row+1, column=0, sticky='nsew')
                self.row_headers.append(header)
        
        # Update main grid
        self.create_grid()
        
        # Update canvas size and scroll region
        self._update_canvas_size()

    def _update_canvas_size(self):
        """Update canvas dimensions and scroll region"""
        total_width = self.cols * self.cell_width
        total_height = sum(self.row_heights[:self.rows])  # Sum only active row heights
        
        # Update canvas dimensions
        self.canvas.config(
            width=min(total_width, self.parent.winfo_width()),
            height=min(total_height, self.parent.winfo_height())
        )
        
        # Update scroll region
        self.canvas.config(scrollregion=(0, 0, total_width, total_height))
        
        # Update grid weights for proper expansion
        self.grid_frame.rowconfigure(1, weight=1)
        self.grid_frame.columnconfigure(1, weight=1)

    # def resize_grid(self, new_rows=None, new_cols=None):
    #     """Resize the grid dynamically"""
    #     self.save_state("Resize grid")
    #     self.rows = new_rows if new_rows is not None else self.rows
    #     self.cols = new_cols if new_cols is not None else self.cols
        
    #     # Reset row heights for new grid size
    #     self.row_heights = [self.default_cell_height] * self.rows
        
    #     # Clear existing widgets using the delegated delete method
    #     self.delete("all")
    #     self.cells = {}
    #     self.merged_cells = {}
        
    #     # Create new grid
    #     self.create_grid()
    #     self.configure_scroll()
    #     self.update_selection(0, 0)

    def auto_fit_row(self, row):
        """Optional manual fit-to-content"""
        max_lines = 1
        for col in range(self.cols):
            if (row, col) in self.cells:
                text = self.cells[(row, col)]['text'].get("1.0", "end-1c")
                lines = text.count('\n') + 1
                max_lines = max(max_lines, lines)
        self.set_row_height(row, self.default_cell_height * max_lines)

    def configure_scroll(self):
        """Configure canvas scrolling settings"""
        self.canvas.configure(
            scrollregion=(0, 0, 
                         self.cols * self.cell_width, 
                         sum(self.row_heights)),
            highlightthickness=0
        )

    def _update_dependencies(self, changed_row, changed_col):
        """Recalculate cells that depend on the changed cell"""
        for (row, col), formula in self.formulas.items():
            if f"{chr(65+changed_col)}{changed_row+1}" in formula:
                self._evaluate_cell(row, col)

    def _init_cell_references(self):
        """Initialize A1, B1 style references"""
        self.cell_references = {}
        for row in range(self.rows):
            for col in range(self.cols):
                col_letter = chr(65 + col)
                self.cell_references[f"{col_letter}{row+1}"] = (row, col)

    def _get_cell_value(self, ref):
        """Get value from A1 reference with better type handling"""
        if ref in self.cell_references:
            row, col = self.cell_references[ref]
            if (row, col) in self.calculated_values:
                return self.calculated_values[(row, col)]
            elif (row, col) in self.cells:
                value = self.cells[(row, col)]['text'].get("1.0", "end-1c")
                try:
                    return float(value) if value else 0
                except ValueError:
                    return value
        return 0

    def _parse_range(self, range_str):
        """Parse A1:B2 style ranges"""
        if ':' in range_str:
            start, end = range_str.split(':')
            start_row, start_col = self.cell_references.get(start, (0, 0))
            end_row, end_col = self.cell_references.get(end, (0, 0))
            return [(r, c) 
                    for r in range(start_row, end_row+1)
                    for c in range(start_col, end_col+1)]
        else:
            return [self.cell_references.get(range_str, (0, 0))]

    def _calculate_formula(self, formula, trigger_cell):
        """Evaluate formula with basic operations"""
        if not formula.startswith('='):
            return formula
        
        try:
            # Extract formula parts
            expr = formula[1:].strip().upper()
            
            # Handle SUM(A1:A5) style
            if expr.startswith('SUM(') and expr.endswith(')'):
                range_str = expr[4:-1].strip()
                cells = self._parse_range(range_str)
                values = []
                for r, c in cells:
                    ref = f"{chr(65+c)}{r+1}"
                    value = self._get_cell_value(ref)
                    if isinstance(value, (int, float)):
                        values.append(value)
                return sum(values) if values else 0
            
            # Handle AVERAGE(A1:A5)
            elif expr.startswith('AVG(') and expr.endswith(')'):
                range_str = expr[4:-1].strip()
                cells = self._parse_range(range_str)
                values = []
                for r, c in cells:
                    ref = f"{chr(65+c)}{r+1}"
                    value = self._get_cell_value(ref)
                    if isinstance(value, (int, float)):
                        values.append(value)
                return sum(values)/len(values) if values else 0
            
            # Handle basic math (A1+B2*C3)
            else:
                # Replace references with values
                for ref in sorted(self.cell_references.keys(), key=len, reverse=True):
                    if ref in expr:
                        value = self._get_cell_value(ref)
                        if isinstance(value, str):
                            # Quote strings for evaluation
                            expr = expr.replace(ref, f'"{value}"')
                        else:
                            expr = expr.replace(ref, str(value))
                
                # Safe evaluation with limited operations
                allowed_names = {
                    'sum': sum,
                    'avg': lambda x: sum(x)/len(x) if x else 0,
                    'sqrt': math.sqrt,
                    'math': math
                }
                code = compile(expr, '<string>', 'eval')
                for name in code.co_names:
                    if name not in allowed_names:
                        raise NameError(f"Use of {name} not allowed")
                return eval(code, {'__builtins__': {}}, allowed_names)
        
        except Exception as e:
            return f"#ERROR: {str(e)}"

    def process_cell_edit(self, row, col):
        """Handle cell content changes"""
        content = self.cells[(row, col)]['text'].get("1.0", "end-1c")
        
        if self.spreadsheet_mode and content.startswith('='):
            self.formulas[(row, col)] = content
            self._evaluate_cell(row, col)
        else:
            if (row, col) in self.formulas:
                del self.formulas[(row, col)]
            if (row, col) in self.calculated_values:
                del self.calculated_values[(row, col)]
            self._update_dependencies(row, col)

    def _evaluate_cell(self, row, col):
        """Calculate and display formula result"""
        if not self.spreadsheet_mode:
            return
            
        formula = self.formulas.get((row, col), "")
        if formula:
            result = self._calculate_formula(formula, (row, col))
            self.calculated_values[(row, col)] = result
            self.cells[(row, col)]['text'].delete("1.0", "end")
            self.cells[(row, col)]['text'].insert("1.0", str(result))
            self._update_dependencies(row, col)

    def enable_spreadsheet_mode(self, enable=True):
        """Toggle spreadsheet functionality and references"""
        self.spreadsheet_mode = enable
        
        # Show/hide headers
        if enable:
            self._init_cell_references()
            self._setup_reference_headers()
            self.recalculate_all()
        else:
            self._remove_reference_headers()
            self.formulas.clear()
            self.calculated_values.clear()
        
        self.refresh_grid()

    def _setup_reference_headers(self):
        """Create reference headers"""
        # Column headers (A, B, C...)
        for col in range(self.cols):
            header = ttk.Label(
                self.grid_frame,
                text=chr(65 + col),
                width=self.cell_width//7,
                anchor='center',
                style='Header.TLabel'
            )
            header.grid(row=0, column=col+1, sticky='nsew')
            self.col_headers.append(header)
        
        # Row headers (1, 2, 3...)
        for row in range(self.rows):
            header = ttk.Label(
                self.grid_frame,
                text=str(row + 1),
                width=4,
                anchor='e',
                style='Header.TLabel'
            )
            header.grid(row=row+1, column=0, sticky='nsew')
            self.row_headers.append(header)
        
        # Reposition canvas
        self.canvas.grid(
            row=1, 
            column=1, 
            columnspan=self.cols, 
            rowspan=self.rows, 
            sticky='nsew'
        )

    def _remove_reference_headers(self):
        """Remove reference headers"""
        for header in self.col_headers + self.row_headers:
            header.destroy()
        self.col_headers = []
        self.row_headers = []
        
        # Reposition canvas
        self.canvas.grid(
            row=0, 
            column=0, 
            columnspan=self.cols + 1, 
            rowspan=self.rows + 1, 
            sticky='nsew'
        )

    def recalculate_all(self):
        """Force recalculation of all formulas"""
        for (row, col) in self.formulas:
            self._evaluate_cell(row, col)

    def get_cell_reference(self, row, col):
        """Convert (row,col) to A1 notation"""
        if 0 <= row < self.rows and 0 <= col < self.cols:
            return f"{chr(65 + col)}{row + 1}"
        return ""