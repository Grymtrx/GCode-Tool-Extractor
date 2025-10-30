import os
import re
import sys
import tempfile
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

APP_TITLE = "GCode Tool Extractor"
APP_WIDTH = 900
APP_HEIGHT = 650

# -----------------------------
# Robust parsing helpers
# -----------------------------

IGNORE_COMMENT_PATTERNS = [
    re.compile(p, re.I)
    for p in [
        r'^\s*(created by|date|time)\b',
        r'^\s*(operation|timely)\b',           # job headers we don’t want as tool text
        r'\bsend to home\b',
        r'^\s*inch\s*$',
        r'\bopen mind',
    ]
]

def is_noise_comment(txt: str) -> bool:
    s = txt.strip()
    # Remove leading letter-dash like "D-..." "A-..." etc
    s = re.sub(r'^\s*[A-Z]\s*-\s*', '', s)
    if len(s) < 3:
        return True
    for pat in IGNORE_COMMENT_PATTERNS:
        if pat.search(s):
            return True
    return False

def extract_program_name(text: str) -> str | None:
    # Prefer: O##### (NAME)
    for line in text.splitlines():
        m = re.search(r'^\s*O\d+\s*\((.+?)\)\s*$', line)
        if m:
            return m.group(1).strip()
    # Else: first standalone ( NAME ) after a %
    seen_percent = False
    for line in text.splitlines():
        if '%' in line:
            seen_percent = True
            continue
        if seen_percent:
            m = re.search(r'^\s*\((.+?)\)\s*$', line)
            if m:
                return m.group(1).strip()
    return None

def _comments_on_line(line: str):
    # return list of contents inside (...) on this line, in order they appear
    return re.findall(r'\((.+?)\)', line)

def _nearest_good_comment(lines, start_idx, search_back=25):
    # Prefer comment on the same line; else look back a bit
    # Same line, prefer the LAST comment on that line
    same = _comments_on_line(lines[start_idx])
    for c in reversed(same):
        cand = re.sub(r'^\s*[A-Z]\s*-\s*', '', c.strip())
        if not is_noise_comment(cand):
            return cand

    for j in range(start_idx - 1, max(-1, start_idx - search_back - 1), -1):
        cms = _comments_on_line(lines[j])
        for c in reversed(cms):
            cand = re.sub(r'^\s*[A-Z]\s*-\s*', '', c.strip())
            if not is_noise_comment(cand):
                return cand
    return ""

def extract_tools(text: str) -> dict[int, str]:
    """
    Returns {tool_number: description}
    - Mill style: T01, T1, T01 M06, ...
    - Lathe style: T0404 (first two digits = tool number)
    Picks the nearest sensible comment around the T-call.
    """
    lines = text.splitlines()
    tools: dict[int, str] = {}

    for idx, line in enumerate(lines):
        # 1) Lathe style: T0404 (or T1212, etc.)
        for m in re.finditer(r'T(\d{2})(\d{2})', line):
            tnum = int(m.group(1))
            desc = _nearest_good_comment(lines, idx)
            tools.setdefault(tnum, desc)

        # 2) General/Mill style: T1 or T01 (avoid capturing T123 etc.)
        for m in re.finditer(r'T(\d{1,2})(?!\d)', line):
            tnum = int(m.group(1))
            desc = _nearest_good_comment(lines, idx)
            tools.setdefault(tnum, desc)

    return tools

def format_tool_list(program_name: str | None, tools: dict[int, str]) -> str:
    lines = []
    header_name = program_name or "UNKNOWN PROGRAM"
    lines.append(f"Tool List for ( {header_name} ).")
    lines.append("")  # blank line

    if not tools:
        lines.append("(No tools found)")
        return "\n".join(lines)

    for t in sorted(tools.keys()):
        desc = tools[t].strip()
        if desc:
            lines.append(f"T{t:02d} - {desc}")
        else:
            lines.append(f"T{t:02d}")
    return "\n".join(lines)

# -----------------------------
# Tkinter App
# -----------------------------

class ToolExtractorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry(f"{APP_WIDTH}x{APP_HEIGHT}")
        self.minsize(700, 450)

        self.selected_files: list[str] = []

        # Top controls frame
        top = tk.Frame(self)
        top.pack(fill=tk.X, padx=10, pady=10)

        self.btn_select = tk.Button(top, text="Select Files…", command=self.select_files)
        self.btn_select.pack(side=tk.LEFT)

        self.btn_extract = tk.Button(top, text="Get Tool List", command=self.get_tool_list)
        self.btn_extract.pack(side=tk.LEFT, padx=8)

        self.btn_print = tk.Button(top, text="Print Tool List", command=self.print_tool_list)
        self.btn_print.pack(side=tk.LEFT, padx=8)

        self.lbl_files = tk.Label(top, text="No files selected", anchor="w")
        self.lbl_files.pack(side=tk.LEFT, padx=12)

        # Output box
        self.output = scrolledtext.ScrolledText(self, wrap=tk.WORD, font=("Consolas", 11))
        self.output.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0,10))

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="Choose one or more G-code files",
            filetypes=[
                ("G-code files", "*.nc *.NC *.has *.HAS *.min *.MIN"),
                ("All files", "*.*"),
            ],
        )
        if not files:
            return
        self.selected_files = list(files)
        display = ", ".join(os.path.basename(f) for f in self.selected_files[:3])
        if len(self.selected_files) > 3:
            display += f" … (+{len(self.selected_files)-3} more)"
        self.lbl_files.config(text=display)

    def get_tool_list(self):
        if not self.selected_files:
            messagebox.showinfo(APP_TITLE, "Please select one or more files first.")
            return

        combined = []
        for path in self.selected_files:
            try:
                with open(path, "r", errors="ignore") as f:
                    text = f.read()
            except Exception as e:
                combined.append(f"[{os.path.basename(path)}]\nError reading file: {e}\n")
                continue

            program_name = extract_program_name(text)
            tools = extract_tools(text)
            block = format_tool_list(program_name, tools)

            # If multiple files, add file header
            if len(self.selected_files) > 1:
                combined.append(f"[{os.path.basename(path)}]\n{block}\n")
            else:
                combined.append(block)

        self.output.delete("1.0", tk.END)
        self.output.insert(tk.END, "\n\n".join(combined))

    def print_tool_list(self):
        content = self.output.get("1.0", tk.END).strip()
        if not content:
            messagebox.showinfo(APP_TITLE, "Nothing to print. Click 'Get Tool List' first.")
            return

        try:
            # Save to a temp .txt and ask Windows to print it
            with tempfile.NamedTemporaryFile(delete=False, suffix=".txt", mode="w", encoding="utf-8") as tf:
                tf.write(content)
                temp_path = tf.name

            if sys.platform.startswith("win"):
                os.startfile(temp_path, "print")
                messagebox.showinfo(APP_TITLE, "Sent to printer.")
            else:
                # Fallback: just open the temp file location
                messagebox.showinfo(APP_TITLE, f"Saved printable file at:\n{temp_path}")
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Print failed:\n{e}")

def main():
    app = ToolExtractorApp()
    app.mainloop()

if __name__ == "__main__":
    main()