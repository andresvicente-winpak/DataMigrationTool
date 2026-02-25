import customtkinter as ctk
import sys
from tkinter import messagebox

class TextRedirector:
    def __init__(self, widget, tag="stdout"):
        self.widget = widget
        self.tag = tag
    def write(self, str):
        try:
            self.widget.configure(state="normal")
            self.widget.insert("end", str, (self.tag,))
            self.widget.see("end")
            self.widget.configure(state="disabled")
            self.widget.update_idletasks()
        except: pass
    def flush(self): pass

def bind_context_help(widget, help_text):
    """
    Binds a right-click event to a widget to display a Help popup.
    Handles Windows/Linux (Button-3) and macOS (Button-2/3).
    """
    def show_help(event):
        # Prevent event propagation if necessary
        messagebox.showinfo("Context Help", help_text)
        return "break"

    # Bind standard Right-Click
    widget.bind("<Button-3>", show_help)
    
    # macOS often uses Button-2 for context menus depending on mouse setup
    if sys.platform == "darwin":
        widget.bind("<Button-2>", show_help)