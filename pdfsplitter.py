import os
import sys
import shutil
import tkinter as tk
import webbrowser
from tkinter import filedialog

from tkinterdnd2 import *
import ttkbootstrap as ttk
from pypdf import PdfReader, PdfMerger
from ttkbootstrap.dialogs.dialogs import Messagebox


class App(Tk):

    def __init__(self):
        super().__init__()
        self.title("PDF Splitter")
        self.geometry("500x400")
        self.resizable(True, True)

        style = ttk.Style('cosmo')
        style.configure('P.TButton', justify='center')

        self.split_size_MB = tk.DoubleVar(value=10.0)
        self.split_size = 10485760
        self.split_method = tk.StringVar(value='by_size')
        self.split_pages_STR = tk.IntVar(value=1)
        self.split_pages = 1
        self.src_paths = []
        self.drop_paths = []
        self.save_dir = None
        self.can_overwrite = None
        self.dst_name_template = "{{src}} {{{num}}}"

        self.blog_url = "https://wordtips.tistory.com/entry/PDFSplitter"
        self.git_repo_url = "https://github.com/jhwpark/pdfsplitter"
        self.license_txt = "oss_notice.txt"

        try:
            self.base_path = sys._MEIPASS
        except Exception:
            self.base_path = os.path.abspath(".")

        self.menubar = tk.Menu(self)
        self.menu3 = tk.Menu(self.menubar, tearoff=False)
        self.menu3.add_command(label='Guide', command=self.open_guide)
        self.menu3.add_command(label='Git Repo', command=self.open_repo)
        self.menu3.add_command(label='OSS Notice', command=self.open_oss_notice)
        self.menubar.add_cascade(label='Help', menu=self.menu3)
        self.configure(menu=self.menubar)

        self.base = ttk.Frame(self)
        self.base.pack(fill='both', expand=True, padx=10, pady=10)

        self.panel = ttk.Frame(self.base, style='P.TFrame')
        self.panel.pack(side='top', anchor='w', fill='both')
        self.panel.columnconfigure(0, weight=1)
        self.panel.columnconfigure(1, weight=1)
        self.panel.columnconfigure(2, weight=1)
        self.panel.columnconfigure(3, weight=10)
        self.panel.columnconfigure(4, weight=1)
        self.panel.columnconfigure(5, weight=1)
        self.panel.columnconfigure(6, weight=1)
        self.radio2 = ttk.Radiobutton(self.panel, text="SIZE", value='by_size',
                                      variable=self.split_method, command=self.check_radio)
        self.radio2.grid(row=1, column=0, padx=5, pady=10, sticky='w')
        self.spin1 = ttk.Spinbox(self.panel, from_=1, to=100, increment=1, format='%3.1f', width=5,
                                 justify='right', textvariable=self.split_size_MB)
        self.spin1.grid(row=1, column=1, sticky='e')
        self.label2 = ttk.Label(self.panel, text="MB", width=5)
        self.label2.grid(row=1, column=2, sticky='w', padx=3)
        self.label3 = ttk.Label(self.panel, text="")  # 여기
        self.label3.grid(row=1, column=3, columnspan=3)
        self.button1 = ttk.Button(self.panel, text="COMMIT", width=6, bootstyle='primary', style='P.TButton',
                                  command=self.commit_split)
        self.button1.grid(row=1, column=6, rowspan=2, sticky='nsew')
        self.radio3 = ttk.Radiobutton(self.panel, text="PAGES", value='by_pages',
                                      variable=self.split_method, command=self.check_radio)
        self.radio3.grid(row=2, column=0, padx=5, pady=10, sticky='w')
        self.spin2 = ttk.Spinbox(self.panel, from_=1, to=100, increment=1, format='%1.0f', width=5,
                                 justify='right', state='disabled', textvariable=self.split_pages_STR)
        self.spin2.grid(row=2, column=1, sticky='e')
        self.separator1 = ttk.Separator(self.panel, orient='horizontal')
        self.separator1.grid(row=4, column=0, columnspan=7, pady=5, sticky='ew')
        self.button2 = ttk.Button(self.panel, text="Add Items", bootstyle='primary-outline',
                                  command=self.append_src_via_dialogbox)
        self.button2.grid(row=5, column=0, pady=(2, 7), sticky='e')
        self.button3 = ttk.Button(self.panel, text="Remove Selected", bootstyle='warning-outline',
                                  command=self.remove_selected_items)
        self.button3.grid(row=5, column=1, padx=2, pady=(2, 7))
        self.button4 = ttk.Button(self.panel, text="Remove All", bootstyle='warning-outline', command=self.remove_all_items)
        self.button4.grid(row=5, column=2, pady=(2, 7), sticky='w')
        self.treeview = ttk.Treeview(self.base, column=['#1', '#2', '#3'], padding=(0, 0, 10, 5), bootstyle='light')
        self.treeview.pack(side='bottom', fill='both', expand=True)
        self.treeview.column('#0', width=20, stretch=False)
        self.treeview.heading('#0', text='')
        self.treeview.column('#1', width=100, stretch=True)
        self.treeview.heading('#1', text='PDF File')
        self.treeview.column('#2', width=60, anchor='e', stretch=False)
        self.treeview.heading('#2', text='Size', anchor='center')
        self.treeview.column('#3', width=60, anchor='e', stretch=False)
        self.treeview.heading('#3', text='Pages')
        self.treeview.tag_configure('src', background='slategray1')
        self.treeview.tag_configure('warning', foreground='indianred1')
        self.treeview.tag_configure('task', foreground='steelblue4')
        self.scrollbar1 = ttk.Scrollbar(self.treeview, orient='vertical', command=self.treeview.yview)
        self.scrollbar1.pack(side='right', fill='y')
        self.treeview.configure(yscrollcommand=self.scrollbar1.set)
        self.scrollbar2 = ttk.Scrollbar(self.treeview, orient='horizontal', command=self.treeview.xview)
        self.scrollbar2.pack(side='bottom', fill='x')
        self.treeview.configure(xscrollcommand=self.scrollbar2.set)

        self.treeview.drop_target_register(DND_FILES)
        self.treeview.dnd_bind("<<Drop>>", self.append_src_via_drop)

    def check_radio(self):
        if self.split_method.get() == 'by_size':
            self.spin1.config(state='enabled')
            self.spin2.config(state='disabled')
        else:
            self.spin1.config(state='disabled')
            self.spin2.config(state='enabled')

    def open_guide(self):
        webbrowser.open(self.blog_url)

    def open_repo(self):
        webbrowser.open(self.git_repo_url)

    def open_oss_notice(self):
        oss_notice = ttk.Toplevel()
        oss_notice.geometry("500x400")
        oss_notice.title("OSS Notice")
        textbox = ttk.ScrolledText(oss_notice, wrap='word')
        textbox.pack(fill='both', expand=True)
        license_txt_path = os.path.join(self.base_path, self.license_txt)
        with open(license_txt_path, 'rt', encoding='utf-8') as fp:
            textbox.insert('end', fp.read())
        textbox.configure(state='disabled')

    def append_src_via_drop(self, event):
        drop_paths = list(event.widget.tk.splitlist(event.data))
        for drop_path in drop_paths:
            self.verify_and_update_src(drop_path)

    def append_src_via_dialogbox(self):
        select_paths = filedialog.askopenfilenames(initialdir=".", title="Select PDF files to split",
                                                   filetypes=(("PDF files", "*.pdf"),))
        if select_paths != "":
            for select_path in select_paths:
                self.verify_and_update_src(select_path)

    def verify_and_update_src(self, src_path):
        if src_path in self.src_paths:
            return
        if os.path.splitext(src_path)[1].lower() != '.pdf':
            return
        try:
            _, _, _ = self.get_treeview_item_info(src_path)
        except Exception as e:  # PdfReadError, EmptyFileError
            msg = f"Cannot Read \"{os.path.basename(src_path)}\".\n[{e}]"
            Messagebox.show_error(title="ERROR", message=msg)
            return
        with open(src_path, 'rb') as fp:  # Encrypted
            reader = PdfReader(fp)
            if reader.is_encrypted:
                msg = f"Cannot Load \"{os.path.basename(src_path)}\".\n[Encrypted]"
                Messagebox.show_error(title="ERROR", message=msg)
                return
        self.src_paths.append(src_path)
        self.insert_treeview_parent_item(src_path)

    def get_pages(self, pdf_path=None):
        with open(pdf_path, 'rb') as fp:
            reader = PdfReader(fp)
            pages = len(reader.pages)
        return pages

    def get_treeview_item_info(self, item_path):
        basename = os.path.basename(item_path)
        size = f"{os.path.getsize(item_path) / 1048576:.1f}MB"
        pages = self.get_pages(pdf_path=item_path)
        return basename, size, pages

    def insert_treeview_parent_item(self, parent_path, tag='src'):
        basename, size, page = self.get_treeview_item_info(parent_path)
        self.treeview.insert(parent='', index='end', iid=parent_path, open=True, text='', tags=(tag,),
                             values=(basename, size, page))
        self.treeview.update()

    def insert_treeview_child_item(self, parent_path, child_path, tag=None):
        basename, size, page = self.get_treeview_item_info(child_path)
        self.treeview.insert(parent=parent_path, index='end', text='', tags=(tag,),
                             values=(basename, size, page))
        self.treeview.update()

    def insert_treeview_message(self, parent_path, msg='', tag=None):
        self.treeview.insert(parent=parent_path, index='end', text="", tags=(tag,),
                             values=(msg,))
        self.treeview.update()

    def remove_all_items(self):
        self.src_paths = []
        self.treeview.delete(*self.treeview.get_children())

    def remove_selected_items(self):
        selected_items = self.treeview.selection()
        for selected_item in selected_items:
            if selected_item in self.src_paths:
                self.src_paths.remove(selected_item)
                self.treeview.delete(selected_item)

    def get_dst_basename(self, src_path=None, **kwargs):
        dst_name = self.dst_name_template
        if kwargs.get('src') is None:
            kwargs['src'] = os.path.basename(os.path.splitext(src_path)[0])
        for key, value in kwargs.items():
            dst_name = dst_name.replace(f"{{{{{key}}}}}", str(value))
        return f"{dst_name}.pdf"

    def can_write(self, dst_path):
        if self.can_overwrite is True:
            return True
        if os.path.exists(dst_path):
            msg = (f"\"{os.path.basename(dst_path)}\" already exists.\n"
                   f"Press Continue to allow overwriting and keep on splitting.\n"
                   f"Press Stop to quit task on its source file."
                   )
            res = Messagebox.show_question(title="File Exist", buttons=['Continue:primary', 'Stop:secondary'],
                                           message=msg)
            if res == 'Continue':
                self.can_overwrite = True
                return True
            else:
                self.can_overwrite = False
                return False
        else:
            return True

    def commit_split(self):
        if self.split_method.get() == 'by_size':
            self.split_size = float(self.split_size_MB.get()) * 1048576  # 1MB = 1024 * 1024B
            for src_path in self.src_paths:
                self.can_overwrite = None
                if os.path.getsize(src_path) <= self.split_size:
                    msg = "Not split. Size of this file is less than specified."
                    self.insert_treeview_message(src_path, msg=msg)
                    continue
                self.split_by_size(src_path)
        elif self.split_method.get() == 'by_pages':
            self.split_pages = int(self.split_pages_STR.get())
            for src_path in self.src_paths:
                self.can_overwrite = None
                if self.get_pages(src_path) <= self.split_pages:
                    msg = "Not split. Total number of pages is less than specified."
                    self.insert_treeview_message(src_path, msg=msg)
                    continue
                self.split_by_page(src_path)

    def split_by_page(self, src_path):
        interval = self.split_pages
        save_dir = self.save_dir if self.save_dir is not None else os.path.dirname(src_path)
        split_num = 1
        src_pages = self.get_pages(pdf_path=src_path)

        with open(src_path, 'rb') as fp:
            reader = PdfReader(fp)
            for chunk_head in range(0, src_pages, interval):
                chunk_tail = min(chunk_head + interval - 1, src_pages - 1)
                dst_path = os.path.join(save_dir, self.get_dst_basename(src_path=src_path, num=split_num))
                if self.can_write(dst_path):
                    try:
                        merger = PdfMerger()
                        merger.append(reader, pages=(chunk_head, chunk_tail + 1))
                        merger.write(open(dst_path, 'wb'))
                        merger.close()
                        split_num += 1
                        self.insert_treeview_child_item(src_path, dst_path)
                    except Exception as e:
                        msg = f"Cannot save \"{os.path.basename(dst_path)}\" [{e}]"
                        self.insert_treeview_message(src_path, msg=msg, tag='warning')
                else:
                    break

        self.insert_treeview_message(src_path, msg="Task done.", tag='task')

    def split_by_size(self, src_path):

        def merge_chunk():
            with open(src_path, 'rb') as fp:
                reader = PdfReader(fp)
                merger = PdfMerger()
                merger.append(reader, pages=(chunk_head, chunk_tail + 1))
                merger_path = os.path.join(self.base_path, "chunk.pdf")
                merger.write(open(merger_path, 'wb'))
                merger_size = os.path.getsize(merger_path)
            return merger_path, merger_size

        def save_at_savedir(tag=None):
            dst_path = os.path.join(save_dir, self.get_dst_basename(src_path=src_path, num=split_num))
            if self.can_write(dst_path):
                try:
                    shutil.copyfile(chunk_path, dst_path)
                    os.remove(chunk_path)
                    self.insert_treeview_child_item(src_path, dst_path, tag=tag)
                    return True
                except Exception as e:
                    msg = f"Cannot save \"{os.path.basename(dst_path)}\" [{e}]"
                    self.insert_treeview_message(src_path, msg=msg, tag='warning')
            return False

        save_dir = self.save_dir if self.save_dir is not None else os.path.dirname(src_path)
        src_size = os.path.getsize(src_path)
        src_pages = self.get_pages(pdf_path=src_path)
        approx_interval = int(self.split_size * src_pages / src_size)
        split_num = 1
        chunk_head = 0

        while chunk_head < src_pages:
            chunk_tail = min(chunk_head + approx_interval, src_pages - 1)
            chunk_path, chunk_size = merge_chunk()

            if chunk_size > self.split_size:
                while True:
                    chunk_tail -= 1
                    chunk_path, chunk_size = merge_chunk()
                    if chunk_size <= self.split_size:
                        save_at_savedir()
                        split_num += 1
                        chunk_head = chunk_tail + 1
                        break
                    if chunk_head == chunk_tail:  # when just 1 page's size exceed specified
                        save_at_savedir(tag='warning')
                        msg = f"Page {chunk_head + 1} exceeds specified size."
                        self.insert_treeview_message(src_path, tag='warning', msg=msg)
                        split_num += 1
                        chunk_head = chunk_tail + 1
                        break
                    if self.can_overwrite is False:
                        break
            elif chunk_size < self.split_size:
                while True:
                    if chunk_tail == src_pages - 1:
                        save_at_savedir()
                        chunk_head = src_pages
                        break
                    chunk_tail += 1
                    chunk_path, chunk_size = merge_chunk()
                    if chunk_size > self.split_size:
                        chunk_tail -= 1
                        chunk_path, chunk_size = merge_chunk()
                        save_at_savedir()
                        split_num += 1
                        chunk_head = chunk_tail + 1
                        break
                    if self.can_overwrite is False:
                        break
            else:
                save_at_savedir()
                split_num += 1
                chunk_head = chunk_tail + 1

            if self.can_overwrite is False:
                break

        self.insert_treeview_message(src_path, msg="Task done.", tag='task')


if __name__ == '__main__':
    app = App()
    app.mainloop()
