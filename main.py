import fnmatch
import os
import tkinter as tk
import tkinter.font as tkFont
import tkinter.ttk as ttk
import win32api
import win32file
import subprocess

mdr = 'D:\\'


class DetailsView:
    def __init__(self):
        self.result = []
        self.tree = None
        self._setup_widgets()

    def _setup_widgets(self):
        search = ttk.Frame(relief='flat', borderwidth=5)
        search.pack()
        self.entry = ttk.Entry()
        self.entry.grid(column=0, row=0, in_=search)
        self.entry.insert(0, 'mp3')
        self.entry.bind("<Return>", self.search_result)
        self.button = ttk.Button(text='Search', command=self.search_result)
        self.button.grid(column=1, row=0, in_=search)

        container = ttk.Frame(relief='flat', borderwidth=5)
        container.pack(fill='both', expand=True)
        self.tree = ttk.Treeview(columns=colum, show="headings")
        vsb = ttk.Scrollbar(orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(column=0, row=0, sticky='nsew', in_=container)
        vsb.grid(column=1, row=0, sticky='ns', in_=container)
        hsb.grid(column=0, row=1, sticky='ew', in_=container)
        container.grid_columnconfigure(0, weight=1)
        container.grid_rowconfigure(0, weight=1)

    def _build_tree(self, disks):
        for col in colum:
            self.tree.heading(col, text=col.title(), anchor=tk.W, command=lambda c=col: sort_by(self.tree, c, 0))
            self.tree.bind("<Double-1>", self.on_double_click)
            self.tree.bind("<Button-3>", self.open_file_location)
        self.tree.column('#01', stretch=tk.YES, minwidth=300)
        self.tree.column('#02', stretch=tk.YES, minwidth=60)
        self.tree.column('#03', stretch=tk.YES, minwidth=70)
        self.tree.column('#04', stretch=tk.YES, minwidth=400)

        for dr in disks:
            for item in self.find_match('*' + self.entry.get() + '*', dr):
                self.tree.insert('', 'end', values=item)

    def on_double_click(self, event):
        curItem = self.tree.focus()
        item = self.tree.item(curItem)
        itemValues = item['values']
        self.open_file(itemValues[3])
        """item = self.tree.selection()[0]
        itemValues = self.tree.item(item, "values")"""

    def open_file_location(self, event):
        curItem = self.tree.focus()
        item = self.tree.item(curItem)
        itemValues = item['values']
        subprocess.Popen(r'explorer /select,"' + itemValues[3] + '"')

    def find_match(self, patrn, path):
        findings = []
        for root, dirs, files in os.walk(path):
            for name in files:
                if fnmatch.fnmatch(name, patrn):
                    path = os.path.join(root, name)
                    extension = os.path.splitext(path)[1]
                    extension = extension.lower()
                    size = (os.stat(path).st_size) / (1024 * 1024)
                    size = "{0:.2f}".format(size)
                    # path = path.replace('\\', '\\\\')
                    findings.append([name, extension, size + " MB", path])
        return findings

    def search_result(self, event=None):
        if self.entry.get() == '':
            return
        self.tree.delete(*self.tree.get_children())
        local_disk = []
        drives = win32api.GetLogicalDriveStrings()
        drives = drives.split('\x00')[1:-1]
        # print(drives)
        for dr in drives:
            print(dr, win32file.GetDriveType(dr))
            if win32file.GetDriveType(dr) == 3 or win32file.GetDriveType(dr) == 2:
                local_disk.append(dr)
        print(local_disk)
        self._build_tree(local_disk)

    def open_file(self, path):
        mod_path = path.replace('\\', '\\\\')
        print("Opened file at PATH: " + mod_path)
        os.system('"' + mod_path + '"')


def sort_by(tree, col, descending):
    data = [(tree.set(child, col), child) for child in tree.get_children('')]
    data.sort(reverse=descending)
    for ix, item in enumerate(data):
        tree.move(item[1], '', ix)
    tree.heading(col, command=lambda col=col: sort_by(tree, col, int(not descending)))


colum = ['Name', 'Type', 'Size', 'Path']
root = tk.Tk()
root.wm_title("Document Finder")
root.minsize(width=1024, height=600)
docList = DetailsView()
root.mainloop()
