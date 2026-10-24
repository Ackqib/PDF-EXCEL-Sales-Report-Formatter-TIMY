"""Report Converter - pick (or drag in) PDF / Excel sales or inventory reports and get
clean Product / Quantity workbooks back."""
import os
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import converter

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    BaseTk = TkinterDnD.Tk
except Exception:  # drag & drop is optional
    DND_FILES, BaseTk = None, tk.Tk

FILE_TYPES = [('Reports', '*.pdf *.xlsx *.xlsm *.xls *.csv'), ('PDF', '*.pdf'),
              ('Excel', '*.xlsx *.xlsm *.xls'), ('CSV', '*.csv'), ('All files', '*.*')]
SUPPORTED = ('.pdf', '.xlsx', '.xlsm', '.xls', '.csv')


class App(BaseTk):
    def __init__(self):
        super().__init__()
        self.title('Report Converter - SKU / Size / Quantity')
        self.geometry('760x600')
        self.minsize(620, 480)
        self.files = []
        self.last_out = None
        self.msgs = queue.Queue()

        pad = {'padx': 10, 'pady': 4}
        top = ttk.Frame(self)
        top.pack(fill='x', **pad)
        ttk.Button(top, text='Add files…', command=self.add_files).pack(side='left')
        ttk.Button(top, text='Remove selected', command=self.remove_selected).pack(side='left', padx=6)
        ttk.Button(top, text='Clear', command=self.clear).pack(side='left')
        hint = 'Drag PDFs / Excel files here' if DND_FILES else 'Select one or more PDF / Excel files'
        ttk.Label(top, text=hint, foreground='gray').pack(side='right')

        box = ttk.Frame(self)
        box.pack(fill='both', expand=True, **pad)
        self.listbox = tk.Listbox(box, selectmode='extended', height=10)
        sb = ttk.Scrollbar(box, command=self.listbox.yview)
        self.listbox.config(yscrollcommand=sb.set)
        self.listbox.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')
        if DND_FILES:
            self.listbox.drop_target_register(DND_FILES)
            self.listbox.dnd_bind('<<Drop>>', self.on_drop)

        opts = ttk.LabelFrame(self, text='Options')
        opts.pack(fill='x', **pad)
        self.out_dir = tk.StringVar()
        row = ttk.Frame(opts)
        row.pack(fill='x', padx=8, pady=4)
        ttk.Label(row, text='Save to:').pack(side='left')
        ttk.Entry(row, textvariable=self.out_dir).pack(side='left', fill='x', expand=True, padx=6)
        ttk.Button(row, text='Browse…', command=self.pick_out).pack(side='left')
        ttk.Label(opts, text='(leave blank to save next to each original file)',
                  foreground='gray').pack(anchor='w', padx=8)

        row2 = ttk.Frame(opts)
        row2.pack(fill='x', padx=8, pady=4)
        self.combined = tk.BooleanVar(value=False)
        self.keep_zero = tk.BooleanVar(value=False)
        self.save_raw = tk.BooleanVar(value=False)
        ttk.Checkbutton(row2, text='Also make one combined file', variable=self.combined).pack(side='left')
        ttk.Checkbutton(row2, text='Keep 0 quantities', variable=self.keep_zero).pack(side='left', padx=12)
        ttk.Checkbutton(row2, text='Also save raw PDF→Excel', variable=self.save_raw).pack(side='left')

        row3 = ttk.Frame(opts)
        row3.pack(fill='x', padx=8, pady=(0, 6))
        ttk.Label(row3, text='Quantity column name (optional override):').pack(side='left')
        self.qty_col = tk.StringVar()
        ttk.Entry(row3, textvariable=self.qty_col, width=22).pack(side='left', padx=6)
        ttk.Label(row3, text='e.g. "On Hand" or "Sold"', foreground='gray').pack(side='left')

        act = ttk.Frame(self)
        act.pack(fill='x', **pad)
        self.go_btn = ttk.Button(act, text='Convert', command=self.run)
        self.go_btn.pack(side='left')
        self.open_btn = ttk.Button(act, text='Open output folder', command=self.open_out, state='disabled')
        self.open_btn.pack(side='left', padx=6)
        self.progress = ttk.Progressbar(act, mode='indeterminate')
        self.progress.pack(side='left', fill='x', expand=True, padx=6)

        self.log = tk.Text(self, height=10, state='disabled', wrap='word', font=('Consolas', 9))
        self.log.pack(fill='both', expand=True, padx=10, pady=(4, 10))
        self.after(100, self.poll)

    # -- file list
    def _add(self, paths):
        for p in paths:
            p = os.path.normpath(p)
            if os.path.splitext(p)[1].lower() in SUPPORTED and p not in self.files:
                self.files.append(p)
                self.listbox.insert('end', p)

    def add_files(self):
        self._add(filedialog.askopenfilenames(title='Choose reports', filetypes=FILE_TYPES))

    def on_drop(self, event):
        self._add(self.tk.splitlist(event.data))

    def remove_selected(self):
        for i in reversed(self.listbox.curselection()):
            self.listbox.delete(i)
            del self.files[i]

    def clear(self):
        self.listbox.delete(0, 'end')
        self.files.clear()

    def pick_out(self):
        d = filedialog.askdirectory(title='Save clean files to')
        if d:
            self.out_dir.set(d)

    # -- conversion
    def write(self, text):
        self.log.config(state='normal')
        self.log.insert('end', text + '\n')
        self.log.see('end')
        self.log.config(state='disabled')

    def poll(self):
        while not self.msgs.empty():
            kind, payload = self.msgs.get()
            if kind == 'log':
                self.write(payload)
            elif kind == 'done':
                self.progress.stop()
                self.go_btn.config(state='normal')
                if payload:
                    self.last_out = os.path.dirname(payload[0])
                    self.open_btn.config(state='normal')
                self.write(f'Finished - {len(payload)} file(s) written.\n')
        self.after(100, self.poll)

    def run(self):
        if not self.files:
            messagebox.showinfo('Report Converter', 'Add at least one PDF or Excel file first.')
            return
        out = self.out_dir.get().strip() or None
        if out and not os.path.isdir(out):
            messagebox.showerror('Report Converter', f'Folder does not exist:\n{out}')
            return
        self.go_btn.config(state='disabled')
        self.progress.start(12)
        args = (list(self.files), out, self.combined.get(), self.qty_col.get().strip() or None,
                self.keep_zero.get(), self.save_raw.get())

        def work():
            try:
                written = converter.convert(*args, log=lambda m: self.msgs.put(('log', m)))
            except Exception as e:
                self.msgs.put(('log', f'ERROR: {e}'))
                written = []
            self.msgs.put(('done', written))

        threading.Thread(target=work, daemon=True).start()

    def open_out(self):
        if self.last_out:
            os.startfile(self.last_out)


if __name__ == '__main__':
    App().mainloop()
