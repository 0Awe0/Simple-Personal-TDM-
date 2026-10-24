
import os
import queue
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

from timeline_marker import read_timeline, highlight_file


def main():
    root = tk.Tk()
    root.withdraw()

    try:
        files = [Path(p).resolve() for p in sys.argv[1:]]

        if len(files) != 3:
            raise ValueError(
                "Select exactly three files:\n"
                "1. Timeline (.xlsx)\n"
                "2. Detailed (.csv or .xlsx)\n"
                "3. Audio Features (.csv or .xlsx)"
            )

        if any(not p.is_file() for p in files):
            raise ValueError("One or more selected files do not exist.")

        if len({p.parent for p in files}) != 1:
            raise ValueError(
                "All three files must be in the same folder."
            )

        timeline = [
            p for p in files
            if "timeline" in p.stem.lower()
            and p.suffix.lower() == ".xlsx"
        ]

        detailed = [
            p for p in files
            if "detailed" in p.stem.lower()
            and p.suffix.lower() in (".csv", ".xlsx")
        ]

        audio = [
            p for p in files
            if "audio_features" in p.stem.lower()
            and p.suffix.lower() in (".csv", ".xlsx")
        ]

        if not all(len(group) == 1 for group in
                   (timeline, detailed, audio)):
            raise ValueError(
                "Could not recognise the three files.\n\n"
                "Their filenames must contain:\n"
                "timeline\n"
                "detailed\n"
                "audio_features"
            )

        output = files[0].parent / "Timeline_Marked"
        detailed_out = output / "detailed_marked.xlsx"
        audio_out = output / "audio_features_marked.xlsx"

        if detailed_out.exists() or audio_out.exists():
            if not messagebox.askyesno(
                "Existing files",
                "Marked Excel files already exist.\n"
                "Do you want to replace them?"
            ):
                root.destroy()
                return

    except Exception as error:
        messagebox.showerror("Timeline Marker", str(error))
        root.destroy()
        return

    root.deiconify()
    root.title("Timeline Marker")
    root.geometry("430x150")
    root.resizable(False, False)

    tk.Label(
        root,
        text="Processing selected files...",
        font=("Segoe UI", 15, "bold")
    ).pack(pady=(22, 9))

    tk.Label(
        root,
        text="Matching timestamps and highlighting Excel rows",
        font=("Segoe UI", 10)
    ).pack()

    progress = ttk.Progressbar(
        root, mode="indeterminate", length=350
    )
    progress.pack(pady=15)
    progress.start(12)

    results = queue.Queue()

    def process():
        try:
            periods = read_timeline(timeline[0])

            detailed_count, _ = highlight_file(
                detailed[0], periods, detailed_out
            )

            audio_count, _ = highlight_file(
                audio[0], periods, audio_out
            )

            results.put((
                True,
                f"Processing complete!\n\n"
                f"Detailed: {detailed_count} highlighted rows\n"
                f"Audio Features: {audio_count} highlighted rows\n\n"
                f"Open the results folder?"
            ))

        except Exception as error:
            results.put((False, str(error)))

    def check():
        try:
            success, message = results.get_nowait()
        except queue.Empty:
            root.after(150, check)
            return

        progress.stop()
        root.withdraw()

        if success:
            if messagebox.askyesno(
                "Timeline Marker - Complete", message
            ):
                os.startfile(str(output))
        else:
            messagebox.showerror(
                "Processing failed", message
            )

        root.destroy()

    threading.Thread(target=process, daemon=True).start()
    root.after(150, check)
    root.mainloop()


if __name__ == "__main__":
    main()
