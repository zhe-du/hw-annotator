# -*- coding: utf-8 -*-
"""hw_annotate_gui.py —— 作业答案标注工具（图形界面版）

打开即窗口。把 .docx 试卷、answers.json 或整个文件夹拖进窗口，
点「开始转换」即可自动把答案红色加粗内联标注到每道题并导出 PDF。

复用 hw_annotate.py 的核心逻辑（解析 / 识别 / 标注 / 校验 / 导出）。
"""

import os
import re
import sys
import queue
import threading

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

import hw_annotate as hw  # noqa: E402

import tkinter as tk  # noqa: E402
from tkinter import ttk, filedialog, messagebox  # noqa: E402
from tkinterdnd2 import DND_FILES, TkinterDnD  # noqa: E402

FONT = ('Microsoft YaHei', 10)
FONT_BIG = ('Microsoft YaHei', 13)


def parse_dnd_paths(s):
    """解析 tkinterdnd2 的 DND_FILES 字符串（Windows 下形如 {路径} {路径}）。"""
    paths = re.findall(r'\{([^{}]+)\}', s)
    rest = re.sub(r'\{[^{}]+\}', ' ', s)
    paths += [p for p in rest.split() if p]
    return [p.strip() for p in paths if p.strip()]


class App:
    def __init__(self, root):
        self.root = root
        self.items = []          # [path, answer_dict, answer_src, status]
        self.answer_paths = []   # 拖入的 answers.json 路径池
        self.q = queue.Queue()
        self.working = False
        self.cfg = hw.load_config(os.path.join(BASE_DIR, 'config.yaml'))
        self._build()
        self.root.after(100, self._poll)

    # ---------------- UI ----------------
    def _build(self):
        self.root.title('作业答案标注工具')
        self.root.geometry('760x560')
        self.root.minsize(640, 480)

        drop = tk.Label(
            self.root,
            text='把 .docx 试卷 / answers.json / 文件夹\n拖到这里',
            relief='groove', bd=2, bg='#eef3f8', fg='#2f5d8a',
            font=FONT_BIG, justify='center',
        )
        drop.pack(fill='x', padx=16, pady=(14, 8), ipady=26)
        drop.drop_target_register(DND_FILES)
        drop.dnd_bind('<<Drop>>', self.on_drop)

        cols = ('file', 'answer', 'status')
        self.tree = ttk.Treeview(self.root, columns=cols, show='headings')
        self.tree.heading('file', text='文件')
        self.tree.heading('answer', text='答案来源')
        self.tree.heading('status', text='状态')
        self.tree.column('file', width=340, anchor='w')
        self.tree.column('answer', width=170, anchor='w')
        self.tree.column('status', width=150, anchor='w')
        self.tree.pack(fill='both', expand=True, padx=16, pady=4)

        bar = tk.Frame(self.root)
        bar.pack(fill='x', padx=16, pady=8)
        self.btn_run = tk.Button(bar, text='开始转换', command=self.start,
                                 width=12, bg='#2e7d32', fg='white', font=FONT)
        self.btn_run.pack(side='left', padx=4)
        tk.Button(bar, text='添加文件…', command=self.add_files, width=12, font=FONT).pack(side='left', padx=4)
        tk.Button(bar, text='打开输出目录', command=self.open_outdir, width=14, font=FONT).pack(side='left', padx=4)
        tk.Button(bar, text='清空列表', command=self.clear, width=10, font=FONT).pack(side='left', padx=4)

        self.progress = ttk.Progressbar(self.root, mode='determinate')
        self.progress.pack(fill='x', padx=16, pady=(4, 2))
        self.status = tk.Label(self.root, text='就绪。拖入文件后点「开始转换」。',
                               anchor='w', fg='#555', font=FONT)
        self.status.pack(fill='x', padx=16, pady=(0, 12))

    # ---------------- 拖拽 / 添加 ----------------
    def on_drop(self, event):
        for p in parse_dnd_paths(event.data):
            self.add_path(p)
        self.refresh()

    def add_files(self):
        paths = filedialog.askopenfilenames(
            title='选择文件',
            filetypes=[('试卷/答案', '*.docx *.json'), ('所有文件', '*.*')])
        for p in paths:
            self.add_path(p)
        self.refresh()

    def add_path(self, p):
        if not p or not os.path.exists(p):
            return
        if os.path.isdir(p):
            for f in hw.collect_inputs(p):
                self._add_docx(f)
            for root, _, files in os.walk(p):
                for f in files:
                    if f.lower().endswith('.json'):
                        self._add_answer(os.path.join(root, f))
            return
        low = p.lower()
        if low.endswith('.docx'):
            self._add_docx(p)
        elif low.endswith('.json'):
            self._add_answer(p)

    def _add_docx(self, p):
        base = os.path.basename(p)
        if base.startswith('~$') or '_含答案' in base or '.before' in base:
            return
        for item in self.items:
            if item[0] == p:
                return
        ans, src, err = self.find_answer(p)
        if ans is not None:
            self.items.append([p, ans, src or '', '待处理'])
        elif err:
            self.items.append([p, None, '', '错误：' + err])
        else:
            self.items.append([p, None, '', '无答案'])

    def _add_answer(self, p):
        if p not in self.answer_paths:
            self.answer_paths.append(p)
        # 已存在的 docx 重新尝试配对
        for item in self.items:
            if item[1] is None:
                ans, src, err = self.find_answer(item[0])
                if ans is not None:
                    item[1], item[2], item[3] = ans, src, '待处理'
                elif err:
                    item[3] = '错误：' + err

    def find_answer(self, docx_path):
        base = os.path.basename(docx_path)
        stem = os.path.splitext(base)[0]
        d = os.path.dirname(docx_path)
        read_error = None
        for cand in ('%s.json' % stem, '%s.answers.json' % stem):
            p = os.path.join(d, cand)
            if os.path.isfile(p):
                try:
                    return hw.load_answers(p), p, None
                except Exception as e:
                    read_error = '答案读取失败 %s: %s' % (os.path.basename(p), e)
        for p in self.answer_paths:
            try:
                dd = hw.load_answers(p)
            except Exception as e:
                if read_error is None:
                    read_error = '答案读取失败 %s: %s' % (os.path.basename(p), e)
                continue
            if dd.get('file') == base:
                return dd, p, None
        if len(self.answer_paths) == 1:
            try:
                return hw.load_answers(self.answer_paths[0]), self.answer_paths[0], None
            except Exception as e:
                read_error = '答案读取失败: %s' % e
        return None, None, read_error

    def refresh(self):
        self.tree.delete(*self.tree.get_children())
        for item in self.items:
            ans_src = os.path.basename(item[2]) if item[2] else ''
            self.tree.insert('', 'end', iid=item[0], values=(
                os.path.basename(item[0]), ans_src or '（无）', item[3]))
        n_ok = sum(1 for it in self.items if it[1] is not None)
        self.status.config(text='共 %d 个文件，%d 个已配对答案。' % (len(self.items), n_ok))

    # ---------------- 转换 ----------------
    def start(self):
        if self.working:
            return
        todo = [it for it in self.items if it[1] is not None]
        if not todo:
            messagebox.showinfo('提示', '没有已配对的试卷。请先拖入 .docx 及其 answers.json。')
            return
        self.working = True
        self.btn_run.config(state='disabled')
        self.progress.config(maximum=len(todo), value=0)
        for it in todo:
            it[3] = '等待…'
        self.refresh()
        threading.Thread(target=self._worker, args=(todo,), daemon=True).start()

    def _worker(self, todo):
        done = 0
        for item in todo:
            path, answer, src, _ = item
            try:
                out_dir = os.path.dirname(path)
                report, _ = hw.process_one(path, answer, self.cfg, out_dir, keep_before=True)
                if report['status'] == 'ok':
                    status = '成功'
                else:
                    status = '失败: ' + '; '.join(report['errors'])
            except Exception as e:
                status = '失败: %s' % e
            done += 1
            self.q.put(('result', path, status))
            self.q.put(('progress', done))
        self.q.put(('done', len(todo)))

    def _poll(self):
        try:
            while True:
                msg = self.q.get_nowait()
                if msg[0] == 'result':
                    _, path, status = msg
                    for it in self.items:
                        if it[0] == path:
                            it[3] = status
                elif msg[0] == 'progress':
                    self.progress['value'] = msg[1]
                elif msg[0] == 'done':
                    self.working = False
                    self.btn_run.config(state='normal')
                    ok = sum(1 for it in self.items if it[3] == '成功')
                    fail = sum(1 for it in self.items if it[3].startswith('失败'))
                    self.status.config(
                        text='转换完成：成功 %d，失败 %d。' % (ok, fail))
        except queue.Empty:
            pass
        if self.working:
            self.refresh()
        self.root.after(100, self._poll)

    # ---------------- 其他 ----------------
    def open_outdir(self):
        if self.items:
            d = os.path.dirname(self.items[0][0])
        else:
            d = BASE_DIR
        os.startfile(d)

    def clear(self):
        self.items.clear()
        self.answer_paths.clear()
        self.refresh()
        self.progress['value'] = 0
        self.status.config(text='就绪。拖入文件后点「开始转换」。')


def main():
    root = TkinterDnD.Tk()
    App(root)
    root.mainloop()


if __name__ == '__main__':
    main()
