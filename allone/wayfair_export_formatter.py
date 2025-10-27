"""Compatibility wrapper around :mod:`wayfair_formatter`.

This module previously contained a separate implementation of the Wayfair export
formatter UI.  The functionality now lives entirely in
``allone.wayfair_formatter`` which provides a fully localized widget.  The
legacy entry point is preserved here so existing imports continue to work.
"""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from wayfair_formatter import WayfairFormatter

__all__ = ["WayfairFormatter", "main"]


def main() -> None:
    """Run the Wayfair formatter widget in a standalone Tkinter window."""

    root = tk.Tk()
    root.title("AllOne - Wayfair Export Formatter")
    root.geometry("900x700")

    notebook = ttk.Notebook(root)
    notebook.pack(fill="both", expand=True)

    tab_frame = ttk.Frame(notebook)
    tab_frame.pack(fill="both", expand=True)
    notebook.add(tab_frame, text="Excel Tools")

    for index in range(2):
        tab_frame.rowconfigure(index, weight=1)
    tab_frame.columnconfigure(0, weight=1)

    formatter = WayfairFormatter(tab_frame)
    formatter.grid(row=0, column=0, sticky="nsew")

    root.mainloop()


if __name__ == "__main__":
    main()
