import tkinter as tk
from table import Table
if __name__ == "__main__":
    root = tk.Tk()
    root.title("Enhanced Table Widget")
    root.geometry("800x600")
    
    # Create main container
    main_frame = tk.Frame(root)
    main_frame.pack(expand=True, fill='both', padx=10, pady=10)
    
    # Add control buttons
    control_frame = tk.Frame(main_frame)
    control_frame.pack(fill='x', pady=5)
    
    # Create table with scrollbars
    table_frame = tk.Frame(main_frame)
    table_frame.pack(expand=True, fill='both')
    
    table = Table(table_frame, rows=10, cols=10, width = 200)
    table.pack(side='left', expand=True, fill='both')
    table.set_selection_mode("multiple")
    
    yscroll = tk.Scrollbar(table_frame, orient='vertical', command=table.yview)
    xscroll = tk.Scrollbar(main_frame, orient='horizontal', command=table.xview)
    table.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
    
    yscroll.pack(side='right', fill='y')
    xscroll.pack(side='bottom', fill='x')
    
    # Add demo buttons
    def merge_action():
        # pass
        table.merge_cells(1, 1, 2, 3)  # Merge 2 rows x 3 cols starting at (1,1)
    
    def unmerge_action():
        # pass
        table.unmerge_cells(1, 1)
    
    def load_data():
        data = [
            ["Header 1", "Header 2", "Header 3"],
            ["Merged", "Data 2", "Data 3"],
            ["Row 2", "More", "Values"]
        ]
        table.set_values(data)
    
    def show_values():
        print(table.get_values()['dataframe'])
    
    tk.Button(control_frame, text="Merge Cells", command=table.merge_selected).pack(side='left', padx=2)
    tk.Button(control_frame, text="Unmerge Cells", command=unmerge_action).pack(side='left', padx=2)
    tk.Button(control_frame, text="Load Data", command=load_data).pack(side='left', padx=2)
    tk.Button(control_frame, text="Show Values", command=show_values).pack(side='left', padx=2)
    
    # Add checkbox for auto-scroll
    auto_scroll_var = tk.BooleanVar(value=True)
    tk.Checkbutton(control_frame, text="Auto-scroll", variable=auto_scroll_var,
                   command=lambda: table.set_auto_scroll(auto_scroll_var.get())).pack(side='right')
    
    root.mainloop()