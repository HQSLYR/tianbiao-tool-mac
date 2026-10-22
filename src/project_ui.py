# -*- coding: utf-8 -*-
"""⑨独立项目文书工作窗口：主线程UI、后台解析/生成/导出。"""
from copy import deepcopy
from pathlib import Path
import json
import os
import queue
import re
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import project_ai as ai
import project_docs as docs
import project_sources as sources
import project_export as export

EDIT_FIELDS = {
    '建设清单': ('items', [('id', '条目ID'), ('name', '建设内容'), ('category', '类别（平台建设/场地改造/其他费用）'),
                           ('spec', '规格/范围'), ('unit', '单位'), ('quantity', '数量'), ('unit_price', '单价（万元）'),
                           ('stated_amount', '原记载金额（无则留空）'), ('basis_ids', '比例费用基数ID（逗号分隔）'),
                           ('rate_percent', '计取比例（百分数）')]),
    '资金渠道': ('funding', [('name', '资金渠道'), ('amount', '金额（万元）')]),
    '绩效指标': ('performance', [('category', '一级指标'), ('name', '指标名称'), ('definition', '定义/统计口径'),
                               ('target', '目标值'), ('unit', '单位'), ('item_ids', '关联建设条目ID（逗号分隔）'),
                               ('source', '取数来源'), ('period', '考核时间'), ('owner', '责任主体')]),
    '进度安排': ('milestones', [('name', '阶段'), ('weeks', '周次/时间'), ('work', '工作内容'),
                              ('owner', '责任主体'), ('deliverable', '交付成果')]),
    '单间成本': ('room_cost', [('item_ids', '计入费用条目ID（逗号分隔）'), ('rooms', '场地数量'),
                              ('stated_cost', '原记载单间成本（万元，无则留空）')]),
    '预算上限': ('budget_limit', [('amount', '预算上限（万元，无则留空）')]),
}
NUMERIC = {'quantity', 'unit_price', 'stated_amount', 'rate_percent', 'amount', 'rooms', 'stated_cost'}


class WebSearchFailure(Exception):
    pass


def text_widget(parent, height=10):
    frame = ttk.Frame(parent)
    text = tk.Text(frame, wrap='word', height=height, undo=True, font=('Microsoft YaHei', 10))
    bar = ttk.Scrollbar(frame, orient='vertical', command=text.yview)
    text.configure(yscrollcommand=bar.set)
    text.pack(side='left', fill='both', expand=True)
    bar.pack(side='right', fill='y')
    frame.pack(fill='both', expand=True, padx=6, pady=4)
    return text


def cn_index(n):
    digits = '零一二三四五六七八九'
    if n < 10:
        return digits[n]
    tens, ones = divmod(n, 10)
    return ('' if tens == 1 else digits[tens]) + '十' + (digits[ones] if ones else '')


class ProjectWindow:
    def __init__(self, parent, out_dir):
        self.win = tk.Toplevel(parent)
        self.win.title('⑨ 项目文书生成 · 资料驱动 / Word与Excel')
        self.win.geometry('1150x820')
        self.win.minsize(900, 660)
        self.win.protocol('WM_DELETE_WINDOW', self.close)
        self.settings, warnings = ai.load_config()
        self.files, self.source_data, self.project = [], None, None
        self.signature = None
        self.results = []
        self.events = queue.Queue()
        self.busy = False
        self.closed = False
        self.cancel_event = threading.Event()
        self.selected_section = None
        self.source_revision = 0
        self._build(out_dir)
        if warnings:
            self.status.set('；'.join(warnings))
        self.after_id = self.win.after(100, self._poll)

    def _build(self, out_dir):
        head = ttk.Frame(self.win)
        head.pack(fill='x', padx=10, pady=8)
        ttk.Label(head, text='项目文书工作台', font=('Microsoft YaHei', 16, 'bold')).pack(side='left')
        ttk.Button(head, text='AI / 联网设置', command=self.configure).pack(side='right', padx=4)
        ttk.Label(head, text='源文件只读 · 模板仅参考样式 · 费用由所选服务收取').pack(side='right', padx=10)
        self.tabs = ttk.Notebook(self.win)
        self.tabs.pack(fill='both', expand=True, padx=10, pady=4)
        self.input_page = ttk.Frame(self.tabs)
        self.input_canvas = tk.Canvas(self.input_page, highlightthickness=0)
        input_scroll = ttk.Scrollbar(self.input_page, orient='vertical', command=self.input_canvas.yview)
        input_scroll.pack(side='right', fill='y')
        self.input_canvas.configure(yscrollcommand=input_scroll.set)
        self.input_canvas.pack(side='left', fill='both', expand=True)
        self.input_tab = ttk.Frame(self.input_canvas)
        input_id = self.input_canvas.create_window((0, 0), window=self.input_tab, anchor='nw')
        self.input_tab.bind('<Configure>', lambda e: self.input_canvas.configure(scrollregion=self.input_canvas.bbox('all')))
        self.input_canvas.bind('<Configure>', lambda e: self.input_canvas.itemconfigure(input_id, width=e.width))
        self.win.bind('<MouseWheel>', self._input_wheel, add='+')
        self.source_tab = ttk.Frame(self.tabs)
        self.outline_tab = ttk.Frame(self.tabs)
        self.preview_tab = ttk.Frame(self.tabs)
        self.data_tab = ttk.Frame(self.tabs)
        self.check_tab = ttk.Frame(self.tabs)
        for frame, title in ((self.input_page, '1 资料与要求'), (self.source_tab, '2 原文预览'),
                             (self.outline_tab, '3 提纲确认'), (self.preview_tab, '4 正文编辑'),
                             (self.data_tab, '5 共享表格'), (self.check_tab, '6 来源与核对')):
            self.tabs.add(frame, text=title)
        top = ttk.LabelFrame(self.input_tab, text='资料（至少一份主要依据，最多10份；支持docx/xlsx/txt/md）')
        top.pack(fill='x', padx=6, pady=4)
        self.file_tree = ttk.Treeview(top, columns=('file', 'role'), show='headings', height=4)
        self.file_tree.heading('file', text='文件')
        self.file_tree.heading('role', text='资料角色')
        self.file_tree.column('file', width=650)
        self.file_tree.column('role', width=120)
        self.file_tree.pack(fill='x', padx=4, pady=3)
        bar = ttk.Frame(top)
        bar.pack(fill='x', padx=4, pady=3)
        ttk.Button(bar, text='添加资料', command=self.add_files).pack(side='left')
        ttk.Button(bar, text='移除选中', command=self.remove_files).pack(side='left', padx=4)
        self.role = tk.StringVar(value='主要依据')
        ttk.Combobox(bar, textvariable=self.role, values=sources.ROLES, state='readonly', width=12).pack(side='left', padx=6)
        ttk.Button(bar, text='应用角色到选中', command=self.apply_role).pack(side='left')
        ttk.Button(bar, text='读取 / 预览资料（不联网）', command=self.read).pack(side='right')
        kind_bar = ttk.Frame(self.input_tab)
        kind_bar.pack(fill='x', padx=8, pady=4)
        self.kind = tk.StringVar(value=docs.KINDS[0])
        self.format = tk.StringVar(value='Word')
        ttk.Label(kind_bar, text='文书类型').pack(side='left')
        combo = ttk.Combobox(kind_bar, textvariable=self.kind, values=docs.KINDS, state='readonly', width=20)
        combo.pack(side='left', padx=5)
        combo.bind('<<ComboboxSelected>>', lambda e: self.format.set('Excel' if self.kind.get() == docs.KINDS[3] else 'Word'))
        ttk.Label(kind_bar, text='输出格式').pack(side='left', padx=(15, 0))
        ttk.Combobox(kind_bar, textvariable=self.format, values=('Word', 'Excel', 'Word＋Excel'), state='readonly', width=15).pack(side='left', padx=5)
        ttk.Button(kind_bar, text='载入示例提示词（可修改）', command=self.example).pack(side='right')
        ttk.Label(self.input_tab, text='提示词：说明写什么、指定章节、周期和篇幅（主要内容必须来自资料）；项目名称从资料中提取，每个项目都不同，示例按钮不固定题目').pack(anchor='w', padx=8)
        self.prompt = text_widget(self.input_tab, 6)
        ttk.Label(self.input_tab, text='用户确认信息：项目/学校名称、规模、预算及资金来源、周期、年度等；没有的信息留空，不套用范文').pack(anchor='w', padx=8)
        self.confirmed = text_widget(self.input_tab, 3)
        options = ttk.Frame(self.input_tab)
        options.pack(fill='x', padx=8, pady=2)
        self.online = tk.BooleanVar(value=False)
        self.estimates = tk.BooleanVar(value=False)
        self.include_notes = tk.BooleanVar(value=False)
        ttk.Checkbutton(options, text='联网查证（单独检索Key）', variable=self.online).pack(side='left')
        ttk.Checkbutton(options, text='允许有依据的测算建议', variable=self.estimates).pack(side='left', padx=8)
        ttk.Checkbutton(options, text='导出时附来源和核对事项', variable=self.include_notes).pack(side='left', padx=8)
        ttk.Label(self.input_tab, text='检索词（仅联网时发送，每行一条、最多3条；勿填个人信息或整篇资料）：').pack(anchor='w', padx=8)
        self.queries = tk.StringVar(value='教育数字化 人工智能 教育部 政策')
        ttk.Entry(self.input_tab, textvariable=self.queries).pack(fill='x', padx=8, pady=2)
        ttk.Label(self.input_tab, text='多条检索词可用分号分隔；仅采用政府和院校官网材料，查不到时由你决定是否仅用上传资料。').pack(anchor='w', padx=8)
        self.word_template = tk.StringVar()
        self.excel_template = tk.StringVar()
        for label, var, ext in (('Word样式参考（可空）', self.word_template, '.docx'), ('Excel样式参考（可空）', self.excel_template, '.xlsx')):
            row = ttk.Frame(self.input_tab)
            row.pack(fill='x', padx=8, pady=2)
            ttk.Label(row, text=label, width=23).pack(side='left')
            ttk.Entry(row, textvariable=var).pack(side='left', fill='x', expand=True)
            ttk.Button(row, text='选择', command=lambda v=var, x=ext: self.template(v, x)).pack(side='left', padx=4)
            ttk.Button(row, text='清除', command=lambda v=var: v.set('')).pack(side='left')
        self.source_text = text_widget(self.source_tab)
        ttk.Label(self.outline_tab, text='每行一个章节；子标题行以“-”开头。修改已生成提纲会清空旧正文以防错配。').pack(anchor='w', padx=8, pady=6)
        self.outline = text_widget(self.outline_tab)
        ttk.Button(self.outline_tab, text='确认提纲并生成 / 继续正文', command=self.write).pack(anchor='e', padx=8, pady=6)
        split = ttk.Panedwindow(self.preview_tab, orient='horizontal')
        split.pack(fill='both', expand=True)
        left, right = ttk.Frame(split), ttk.Frame(split)
        split.add(left, weight=1)
        split.add(right, weight=4)
        self.section_tree = ttk.Treeview(left, show='tree', selectmode='browse')
        self.section_tree.pack(fill='both', expand=True, padx=4, pady=4)
        self.section_tree.bind('<<TreeviewSelect>>', self.select_section)
        ttk.Label(right, text='小标题以（一）（二）…开头（也兼容“## ”），段落之间空一行。表格由共享数据生成，不在这里手写金额表。').pack(anchor='w', padx=6, pady=4)
        self.section_text = text_widget(right)
        preview_actions = ttk.Frame(right)
        preview_actions.pack(fill='x', padx=6, pady=4)
        ttk.Button(preview_actions, text='重新生成全部正文', command=self.regenerate).pack(side='left')
        ttk.Button(preview_actions, text='保存本章修改并重新核对', command=self.save_section).pack(side='right')
        data_bar = ttk.Frame(self.data_tab)
        data_bar.pack(fill='x', padx=8, pady=6)
        self.edit_kind = tk.StringVar(value='建设清单')
        c = ttk.Combobox(data_bar, textvariable=self.edit_kind, values=list(EDIT_FIELDS), state='readonly', width=15)
        c.pack(side='left')
        c.bind('<<ComboboxSelected>>', lambda e: self.refresh_data())
        ttk.Button(data_bar, text='编辑选中行', command=self.edit_row).pack(side='left', padx=4)
        ttk.Button(data_bar, text='新增一行', command=lambda: self.edit_row(new=True)).pack(side='left', padx=4)
        ttk.Button(data_bar, text='删除选中行', command=self.delete_row).pack(side='left', padx=4)
        self.data_tree = ttk.Treeview(self.data_tab, show='headings', height=9)
        self.data_tree.pack(fill='both', expand=True, padx=8)
        xbar = ttk.Scrollbar(self.data_tab, orient='horizontal', command=self.data_tree.xview)
        xbar.pack(fill='x', padx=8)
        self.data_tree.configure(xscrollcommand=xbar.set)
        self.data_tree.bind('<Double-1>', lambda e: self.edit_row())
        ttk.Label(self.data_tab, text='金额、小计、采购数量和关联绩效由同一清单计算；编辑是用户确认，原始资料不变。').pack(anchor='w', padx=8, pady=4)
        self.table_choice = tk.StringVar()
        self.table_combo = ttk.Combobox(self.data_tab, textvariable=self.table_choice, state='readonly')
        self.table_combo.pack(fill='x', padx=8)
        self.table_combo.bind('<<ComboboxSelected>>', lambda e: self.show_computed())
        self.computed_text = text_widget(self.data_tab, 8)
        self.check_text = text_widget(self.check_tab)
        footer = ttk.Frame(self.win)
        footer.pack(fill='x', padx=10, pady=4)
        ttk.Label(footer, text='输出目录').pack(side='left')
        self.output = tk.StringVar(value=str(out_dir))
        ttk.Entry(footer, textvariable=self.output).pack(side='left', fill='x', expand=True, padx=5)
        ttk.Button(footer, text='选择目录', command=self.choose_output).pack(side='left')
        action = ttk.Frame(self.win)
        action.pack(fill='x', padx=10, pady=4)
        self.analyse_button = ttk.Button(action, text='分析资料 / 提取提纲', command=self.analyse)
        self.analyse_button.pack(side='left')
        self.write_button = ttk.Button(action, text='生成 / 继续正文', command=self.write)
        self.write_button.pack(side='left', padx=5)
        self.cancel_button = ttk.Button(action, text='取消后续请求', command=self.cancel)
        self.cancel_button.pack(side='left')
        self.export_button = ttk.Button(action, text='校验并导出', command=self.export)
        self.export_button.pack(side='right')
        ttk.Button(action, text='打开最近结果', command=self.open_result).pack(side='right', padx=5)
        self.status = tk.StringVar(value='先添加资料并填写要求。未配置AI时仍可在本机读取资料；不会自动联网。')
        self.status_label = ttk.Label(self.win, textvariable=self.status, wraplength=850)
        self.tabs.pack_forget()
        footer.pack_forget()
        action.pack_forget()
        self.status_label.pack(side='bottom', fill='x', padx=12, pady=(2, 8))
        action.pack(side='bottom', fill='x', padx=10, pady=4)
        footer.pack(side='bottom', fill='x', padx=10, pady=4)
        self.tabs.pack(fill='both', expand=True, padx=10, pady=4)

    def _input_wheel(self, event):
        if self.tabs.select() == str(self.input_page) and not isinstance(event.widget, (tk.Text, ttk.Treeview)):
            self.input_canvas.yview_scroll(-int(event.delta / 120), 'units')

    def _guard(self):
        if self.busy:
            messagebox.showinfo('处理中', '请等待当前任务结束，或取消后续请求。', parent=self.win)
            return False
        return True

    def _replace(self, widget, text):
        widget.delete('1.0', 'end')
        widget.insert('1.0', text)
        widget.edit_modified(False)

    def _clear_preview(self):
        self.selected_section = None
        self.section_tree.delete(*self.section_tree.get_children())
        self.data_tree.delete(*self.data_tree.get_children())
        self.table_combo.configure(values=[])
        self.table_choice.set('')
        for widget in (self.outline, self.section_text, self.computed_text, self.check_text, self.source_text):
            self._replace(widget, '')

    def _refresh_files(self):
        self.file_tree.delete(*self.file_tree.get_children())
        for n, item in enumerate(self.files):
            self.file_tree.insert('', 'end', iid=str(n), values=(item['path'], item['role']))
        self.source_data = self.project = None
        self._clear_preview()
        self.source_revision += 1
        self.status.set('资料列表已变化，请重新读取。')

    def add_files(self):
        if not self._guard():
            return
        paths = filedialog.askopenfilenames(parent=self.win, title='添加项目依据或参考资料',
                                            filetypes=[('支持的资料', '*.docx *.xlsx *.txt *.md'), ('所有文件', '*.*')])
        unsupported = [Path(p).name for p in paths if Path(p).suffix.lower() not in ('.docx', '.xlsx', '.txt', '.md')]
        if unsupported:
            messagebox.showwarning('暂不支持的资料', '请将以下文件另存为docx、xlsx、txt或md；PDF或扫描件须先转为文字：\n' + '\n'.join(unsupported), parent=self.win)
            return
        if len({f['path'] for f in self.files} | set(paths)) > sources.MAX_FILES:
            messagebox.showwarning('资料过多', '每次最多10份，请先移除不需要的资料。', parent=self.win)
            return
        changed = False
        for path in paths:
            if not any(f['path'] == path for f in self.files):
                self.files.append({'path': path, 'role': self.role.get()})
                changed = True
        if changed:
            self._refresh_files()

    def remove_files(self):
        if self._guard():
            selected = {int(i) for i in self.file_tree.selection()}
            if selected:
                self.files = [f for i, f in enumerate(self.files) if i not in selected]
                self._refresh_files()

    def apply_role(self):
        if self._guard():
            selected = self.file_tree.selection()
            changed = False
            for i in selected:
                if self.files[int(i)]['role'] != self.role.get():
                    self.files[int(i)]['role'] = self.role.get()
                    changed = True
            if changed:
                self._refresh_files()

    def example(self):
        if self._guard():
            self._replace(self.prompt, docs.EXAMPLE_PROMPT)
            self.kind.set(docs.KINDS[0])

    def template(self, var, ext):
        if not self._guard():
            return
        path = filedialog.askopenfilename(parent=self.win, title='选择样式参考（不复制原正文）', filetypes=[('样式参考', '*' + ext)])
        if path:
            var.set(path)

    def choose_output(self):
        path = filedialog.askdirectory(parent=self.win)
        if path:
            self.output.set(path)

    def configure(self):
        if not self._guard():
            return
        dialog = tk.Toplevel(self.win)
        dialog.title('AI与联网服务设置')
        dialog.transient(self.win)
        dialog.grab_set()
        values = {}
        fields = [('base_url', 'AI服务HTTPS地址（通常以/v1结尾）'), ('model', '模型名称'),
                  ('api_key', 'AI API Key'), ('search_key', 'Tavily检索Key（可空）'),
                  ('input_chars', '单次输入字符上限（8000—100000）'),
                  ('output_tokens', '单次输出token上限（1000—16000）'), ('timeout', '超时秒数（15—600）')]
        for row, (key, label) in enumerate(fields):
            ttk.Label(dialog, text=label).grid(row=row, column=0, sticky='w', padx=10, pady=5)
            values[key] = tk.StringVar(value=str(self.settings.get(key, '')))
            ttk.Entry(dialog, textvariable=values[key], width=45, show='*' if 'key' in key else '').grid(row=row, column=1, padx=10, pady=5)
        remember = tk.BooleanVar(value=self.settings.get('remember_keys', False))
        ttk.Checkbutton(dialog, text='记住密钥（仅本机Windows账户可解密；默认只在本次会话使用）', variable=remember).grid(row=7, column=0, columnspan=2, padx=10, pady=8)
        ttk.Label(dialog, text='按模型真实上下文设置上限，字符数不等于token数；不同服务可能收费。\n联网只接收政府/院校官网检索结果，不代表内容已由主管部门确认。').grid(row=8, column=0, columnspan=2, padx=10, pady=5)
        def collect():
            result = {key: var.get().strip() for key, var in values.items()}
            result['remember_keys'] = remember.get()
            for key in ('input_chars', 'output_tokens', 'timeout'):
                result[key] = int(result[key])
            if result['base_url']:
                ai.endpoint(result['base_url'])
            return result
        def save():
            try:
                settings = collect()
                ai.save_config(settings)
                self.settings = settings
                dialog.destroy()
                self.status.set('服务设置已保存；密钥未写入普通配置或分发包。')
            except Exception as exc:
                messagebox.showerror('设置失败', str(exc), parent=dialog)
        def test():
            try:
                settings = collect()
                client = ai.AIClient(settings)
                if not messagebox.askyesno('连接测试', '将向配置的AI服务发送不含用户资料的测试短句，可能产生少量费用。继续？', parent=dialog):
                    return
                self.settings = settings
                dialog.destroy()
                self._start('连接测试', lambda: ai.AIClient(settings, self.cancel_event).test(), lambda result: messagebox.showinfo('连接成功', '接口已返回有效文本。', parent=self.win))
            except Exception as exc:
                messagebox.showerror('测试失败', str(exc), parent=dialog)
        bar = ttk.Frame(dialog)
        bar.grid(row=9, column=0, columnspan=2, pady=10)
        ttk.Button(bar, text='测试连接（不传资料）', command=test).pack(side='left', padx=4)
        ttk.Button(bar, text='保存设置', command=save).pack(side='left', padx=4)

    def _start(self, name, work, done):
        if self.busy:
            return
        self.busy = True
        self.cancel_event = threading.Event()
        self.status.set(name + '……')
        for button in (self.analyse_button, self.write_button, self.export_button):
            button.configure(state='disabled')
        self.cancel_button.configure(state='normal' if name not in ('保存文书', '本地读取资料') else 'disabled')
        for widget in (self.prompt, self.confirmed, self.outline, self.section_text):
            widget.configure(state='disabled')
        def worker():
            try:
                result = work()
                self.events.put(('done', done, result))
            except Exception as exc:
                self.events.put(('error', name, exc))
        threading.Thread(target=worker, daemon=True).start()

    def _poll(self):
        if self.closed:
            return
        try:
            while True:
                event = self.events.get_nowait()
                if event[0] == 'progress':
                    self.status.set(event[1])
                    continue
                self.busy = False
                for button in (self.analyse_button, self.write_button, self.export_button):
                    button.configure(state='normal')
                for widget in (self.prompt, self.confirmed, self.outline, self.section_text):
                    widget.configure(state='normal')
                if event[0] == 'done':
                    self.status.set('处理完成。')
                    event[1](event[2])
                else:
                    exc = event[2]
                    self.status.set(str(exc))
                    if self.project:
                        self.refresh_project()
                    if isinstance(exc, WebSearchFailure):
                        if messagebox.askyesno('联网查证未完成', str(exc) + '\n\n关闭联网并仅用上传资料继续？', parent=self.win):
                            self.online.set(False)
                            self.project = None
                            self.analyse()
                    elif not isinstance(exc, ai.Cancelled):
                        messagebox.showerror(event[1] + '未完成', str(exc), parent=self.win)
        except queue.Empty:
            pass
        self.after_id = self.win.after(100, self._poll)

    def _progress(self, text):
        self.events.put(('progress', text))

    def read(self):
        if not self._guard():
            return
        files = deepcopy(self.files)
        def done(result):
            self.source_data = result
            self.project = None
            self._clear_preview()
            self.source_revision += 1
            self._replace(self.source_text, '\n\n'.join('[%s] %s / %s / %s\n%s' %
                (b['id'], b['file'], b['role'], b['location'], b['text']) for b in result['blocks']))
            self.tabs.select(self.source_tab)
            self.status.set('已在本机读取%d个来源块、%d字符；尚未发送给AI。%s' %
                            (len(result['blocks']), result['characters'], '；'.join(result['warnings'])))
        self._start('本地读取资料', lambda: sources.read_sources(files), done)

    def _consent(self, phase):
        endpoint = ai.endpoint(self.settings.get('base_url', ''))
        files = '\n'.join('%s：%s' % (f['role'], Path(f['path']).name) for f in self.files)
        extra = '\n联网：Tavily（检索词：%s）' % self.queries.get() if self.online.get() else '\n联网检索：关闭'
        return messagebox.askyesno('确认发送资料', '%s\nAI服务：%s\n模型：%s\n发送资料文本、提示词和确认信息：\n%s%s\n\n不发送原文件二进制或本机路径。每个阶段可能多次调用并收费。是否继续？' %
                                  (phase, endpoint, self.settings.get('model'), files, extra), parent=self.win)

    def analyse(self):
        if not self._guard():
            return
        if self.source_data is None:
            messagebox.showinfo('先读取资料', '请先点击“读取 / 预览资料（不联网）”，确认角色和提取内容。', parent=self.win)
            return
        prompt = self.prompt.get('1.0', 'end').strip()
        confirmed = self.confirmed.get('1.0', 'end').strip()
        other = [kind for kind in docs.KINDS if kind != self.kind.get() and ('写' + kind in prompt or '生成' + kind in prompt)]
        if other and not messagebox.askyesno('类型与提示词可能不同', '当前类型为%s，提示词提到%s。仍以当前类型生成？' % (self.kind.get(), '、'.join(other)), parent=self.win):
            return
        queries = [q.strip() for q in re.split('[；;\n]', self.queries.get()) if q.strip()]
        if self.online.get() and (not queries or len(queries) > 3):
            messagebox.showwarning('检索词', '联网时请填写1至3条检索词。', parent=self.win)
            return
        try:
            ai.AIClient(self.settings)
            if not self._consent('分析资料、提取事实并生成提纲'):
                return
            signature = (self.source_revision, self.kind.get(), prompt, confirmed, self.estimates.get(), self.online.get(), tuple(queries))
            if self.project is None or signature != self.signature:
                self.project = docs.new_project(self.source_data, self.kind.get(), prompt, confirmed, self.estimates.get())
                self.signature = signature
            project = self.project
            settings, online = dict(self.settings), self.online.get()
            def work():
                client = ai.AIClient(settings, self.cancel_event)
                if online and not project.get('web_done'):
                    self._progress('检索政府与院校官网资料')
                    try:
                        blocks = ai.search_sources(queries, settings.get('search_key', ''), self.cancel_event)
                    except ai.ServiceError as exc:
                        raise WebSearchFailure(str(exc)) from None
                    project['sources']['blocks'].extend(blocks)
                    project['web_done'] = True
                return docs.analyse(project, client, self._progress)
            def done(result):
                self.refresh_project()
                self.tabs.select(self.outline_tab)
                self.status.set('资料已分析。请确认提纲和共享数据后生成正文；校情及测算仍需人工核对。')
            self._start('分析资料', work, done)
        except Exception as exc:
            messagebox.showerror('无法开始', str(exc), parent=self.win)

    def refresh_project(self):
        if not self.project:
            return
        self._replace(self.outline, '\n'.join(row['title'] + ''.join('\n- ' + s for s in row.get('subsections', []))
                                            for row in self.project['outline']))
        self.selected_section = None
        self._replace(self.section_text, '')
        self.section_tree.delete(*self.section_tree.get_children())
        for section in self.project['sections']:
            self.section_tree.insert('', 'end', iid=str(section['index']), text=section['title'])
        self.refresh_checks()
        self.refresh_data()

    def write(self, regenerate=False):
        if not self._guard():
            return
        if not self.project or not self.project.get('model'):
            messagebox.showinfo('先分析', '请先分析资料并确认提纲。', parent=self.win)
            return
        current = (self.source_revision, self.kind.get(), self.prompt.get('1.0', 'end').strip(),
                   self.confirmed.get('1.0', 'end').strip(), self.estimates.get(), self.online.get(),
                   tuple(q.strip() for q in re.split('[；;\n]', self.queries.get()) if q.strip()))
        if current != self.signature:
            messagebox.showwarning('输入已变化', '资料要求已变化，请先重新分析，以免正文混用旧事实。', parent=self.win)
            return
        try:
            if self.selected_section is not None and self.section_text.edit_modified():
                if not messagebox.askyesno('未保存的正文', '先保存当前正文修改再继续？', parent=self.win):
                    return
                if not self.save_section():
                    return
            outline = self.outline.get('1.0', 'end').strip()
            original = '\n'.join(row['title'] + ''.join('\n- ' + s for s in row.get('subsections', [])) for row in self.project['outline'])
            if outline != original and self.project['sections']:
                if not messagebox.askyesno('提纲变化', '修改提纲将清空已生成正文，是否继续？', parent=self.win):
                    return
            project = deepcopy(self.project)
            docs.update_outline(project, outline)
            ai.AIClient(self.settings)
            if not self._consent('重新生成全部正文' if regenerate else '逐章生成正文（已完成章节将跳过）'):
                return
            if regenerate:
                project['sections'] = []
                project['data_changed_after_writing'] = False
            self.project = project
            settings = dict(self.settings)
            def done(result):
                self.refresh_project()
                self.tabs.select(self.preview_tab)
                if self.section_tree.get_children():
                    self.section_tree.selection_set(self.section_tree.get_children()[0])
                self.status.set('正文处理完成，请检查内容及共享表格；有问题的草稿不等于可直接报送的正式材料。')
            self._start('逐章撰写', lambda: docs.generate_sections(project, ai.AIClient(settings, self.cancel_event), self._progress), done)
        except Exception as exc:
            messagebox.showerror('无法生成', str(exc), parent=self.win)

    def regenerate(self):
        if not self._guard() or not self.project or not self.project.get('model'):
            return
        if not messagebox.askyesno('重新生成正文', '将清空当前正文，保留资料、提纲和已确认的共享数据；再次请求AI可能收费。继续？', parent=self.win):
            return
        self.write(regenerate=True)

    def select_section(self, event=None):
        if self.busy:
            return
        selected = self.section_tree.selection()
        if not selected or not self.project:
            return
        if int(selected[0]) == self.selected_section:
            return
        if self.selected_section is not None and self.section_text.edit_modified():
            choice = messagebox.askyesnocancel('正文已修改', '切换前是否保存当前章节？选择“取消”可继续编辑。', parent=self.win)
            if choice is None or (choice and not self.save_section()):
                self.section_tree.selection_set(str(self.selected_section))
                return
        self.selected_section = int(selected[0])
        section = next(s for s in self.project['sections'] if s['index'] == self.selected_section)
        parts = []
        for n, block in enumerate(section['blocks']):
            body = '\n\n'.join(block['paragraphs'])
            if block.get('heading'):
                body = '（%s）%s\n\n%s' % (cn_index(n + 1), block['heading'], body)
            parts.append(body)
        self._replace(self.section_text, '\n\n'.join(parts))

    def save_section(self):
        if not self._guard() or self.selected_section is None or not self.project:
            return
        section = next((s for s in self.project['sections'] if s['index'] == self.selected_section), None)
        if not section:
            return
        old = {b.get('heading', ''): b for b in section['blocks']}
        blocks, heading, lines = [], '', []
        def finish():
            paras = [p.strip() for p in re.split(r'\n\s*\n', '\n'.join(lines)) if p.strip()]
            if paras:
                prior = old.get(heading, {})
                blocks.append({'heading': heading, 'paragraphs': paras,
                               'citations': prior.get('citations', []) if prior.get('paragraphs') == paras else [],
                               'table_refs': prior.get('table_refs', [])})
        head_re = re.compile(r'^（[一二三四五六七八九十]{1,3}）(\S.{0,59})$')
        for line in self.section_text.get('1.0', 'end').splitlines():
            stripped = line.strip()
            title = stripped[3:].strip() if stripped.startswith('## ') else None
            if title is None:
                matched = head_re.match(stripped)
                if matched and not matched.group(1).endswith(('。', '；', '，', '：')):
                    title = matched.group(1).strip()
            if title is not None:
                finish()
                heading, lines = title, []
            else:
                lines.append(line)
        finish()
        if not blocks:
            messagebox.showwarning('正文为空', '不能保存空章节。', parent=self.win)
            return
        section['blocks'] = blocks
        self.section_text.edit_modified(False)
        self.refresh_checks()
        self.status.set('章节修改已保存并重新核对。')
        return True

    def refresh_data(self):
        self.data_tree.delete(*self.data_tree.get_children())
        if not self.project or not self.project.get('model'):
            return
        key, fields = EDIT_FIELDS[self.edit_kind.get()]
        self.data_tree.configure(columns=[f[0] for f in fields])
        for field, title in fields:
            self.data_tree.heading(field, text=title)
            self.data_tree.column(field, width=140, stretch=False)
        rows = self.project['model'].get(key, [])
        rows = [rows] if isinstance(rows, dict) else rows
        for n, row in enumerate(rows):
            self.data_tree.insert('', 'end', iid=str(n), values=[', '.join(map(str, row.get(field, []))) if isinstance(row.get(field), list) else ('【待补充】' if row.get(field) is None else row.get(field, '')) for field, _ in fields])
        tables = docs.shared_tables(self.project)
        titles = [t['title'] for t in tables.values()]
        self.table_combo.configure(values=titles)
        if self.table_choice.get() not in titles:
            self.table_choice.set(titles[0] if titles else '')
        self.show_computed()

    def show_computed(self):
        if not self.project or not self.project.get('model'):
            return
        table = next((t for t in docs.shared_tables(self.project).values() if t['title'] == self.table_choice.get()), None)
        if table:
            self._replace(self.computed_text, ' | '.join(table['headers']) + '\n\n' + '\n'.join(' | '.join(map(str, row)) for row in table['rows']))

    def edit_row(self, new=False):
        if not self._guard() or not self.project or not self.project.get('model'):
            return
        key, fields = EDIT_FIELDS[self.edit_kind.get()]
        data = self.project['model'].get(key, [])
        selected = self.data_tree.selection()
        if not new and not selected:
            return
        index = int(selected[0]) if selected and not new else None
        row = {} if new else deepcopy(data if isinstance(data, dict) else data[index])
        if key in ('room_cost', 'budget_limit') and isinstance(data, dict):
            row = deepcopy(data)
        dialog = tk.Toplevel(self.win)
        dialog.title('确认共享数据：' + self.edit_kind.get())
        dialog.transient(self.win)
        dialog.grab_set()
        variables = {}
        for n, (field, label) in enumerate(fields):
            ttk.Label(dialog, text=label).grid(row=n, column=0, sticky='w', padx=8, pady=4)
            value = row.get(field, '')
            value = ','.join(map(str, value)) if isinstance(value, list) else str(value) if value is not None else ''
            variables[field] = tk.StringVar(value=value)
            ttk.Entry(dialog, textvariable=variables[field], width=60).grid(row=n, column=1, padx=8, pady=4)
        def save():
            try:
                changed = deepcopy(row)
                for field, var in variables.items():
                    value = var.get().strip()
                    if field in NUMERIC:
                        if value and docs.number(value) is None:
                            raise ValueError('数值必须是非负有限数字，不确定时留空。')
                        changed[field] = value or None
                    elif field in ('item_ids', 'basis_ids'):
                        changed[field] = [t.strip() for t in re.split('[,，、]', value) if t.strip()]
                    else:
                        changed[field] = value
                if key == 'items':
                    if changed.get('category') not in ('平台建设', '场地改造', '其他费用'):
                        raise ValueError('费用类别须为平台建设、场地改造或其他费用。')
                    ident = changed.get('id')
                    if not ident or any(i.get('id') == ident for n, i in enumerate(data) if n != index):
                        raise ValueError('建设条目ID不能为空或重复。')
                    changed['money_unit'] = '万元'
                changed['status'] = '用户确认'
                if key in ('room_cost', 'budget_limit'):
                    self.project['model'][key] = changed
                elif new:
                    data.append(changed)
                else:
                    data[index] = changed
                block_id = 'USER-EDIT-%d' % (len(self.project['facts']) + 1)
                quote = json.dumps(changed, ensure_ascii=False)
                self.project['sources']['blocks'].append({'id': block_id, 'file': '用户共享数据确认',
                    'role': '用户确认', 'location': self.edit_kind.get(), 'text': quote})
                self.project['facts'].append({'id': 'FACT-%d' % (len(self.project['facts']) + 1),
                    'key': self.edit_kind.get(), 'value': quote, 'status': '用户确认',
                    'citations': [{'source_id': block_id, 'quote': quote}]})
                if self.project['sections']:
                    self.project['data_changed_after_writing'] = True
                dialog.destroy()
                self.refresh_data()
                self.refresh_checks()
                self.status.set('共享数据已更新；可重新生成正文，或逐章复查已有正文。')
            except Exception as exc:
                messagebox.showerror('不能保存', str(exc), parent=dialog)
        ttk.Button(dialog, text='确认此数据并重算', command=save).grid(row=len(fields), column=0, columnspan=2, pady=10)

    def delete_row(self):
        if not self._guard() or not self.project or not self.project.get('model'):
            return
        key, _ = EDIT_FIELDS[self.edit_kind.get()]
        selected = self.data_tree.selection()
        if not selected or key in ('room_cost', 'budget_limit'):
            return
        if messagebox.askyesno('删除共享条目', '删除后引用该ID的绩效或费用须同步修正，是否继续？', parent=self.win):
            del self.project['model'][key][int(selected[0])]
            if self.project['sections']:
                self.project['data_changed_after_writing'] = True
            self.refresh_data()
            self.refresh_checks()

    def refresh_checks(self):
        if not self.project:
            return
        issues = docs.validate(self.project)
        lines = ['核对结果（来源匹配不代表独立证实；待确认稿不可免审核直接报送）']
        lines += ['[%s] %s' % (i['level'], i['message']) for i in issues]
        lines.append('\n来源事实：')
        for fact in self.project['facts']:
            lines.append('[%s / %s] %s：%s' % (fact['id'], fact['status'], fact['key'], fact['value']))
            lines += ['  %s：%s' % (c['source_id'], c['quote']) for c in fact['citations']]
        self._replace(self.check_text, '\n'.join(lines))

    def export(self):
        if not self._guard() or not self.project:
            return
        if self.selected_section is not None and self.section_text.edit_modified():
            if messagebox.askyesno('未保存的正文', '先保存当前正文修改再导出？', parent=self.win):
                if not self.save_section():
                    return
            else:
                return
        issues = docs.validate(self.project)
        errors = [i for i in issues if i['level'] == '错误']
        draft = False
        if errors:
            self.refresh_checks()
            self.tabs.select(self.check_tab)
            if not messagebox.askyesno('有核对错误', '发现%d项错误：多数可在“共享表格”里修正（如比例费用的计取基数、资金合计），未完成正文可继续生成。\n\n仍要导出为“问题草稿”吗？草稿会附全部核对事项并标明草稿，仅供内部修改，不得直接报送。' % len(errors), parent=self.win):
                return
            draft = True
        elif issues and not messagebox.askyesno('导出待确认草稿', '仍有%d项待核对内容，将明确标为待确认草稿。是否导出？' % len(issues), parent=self.win):
            return
        formats = ('Word', 'Excel') if self.format.get() == 'Word＋Excel' else (self.format.get(),)
        word, excel = self.word_template.get().strip() or None, self.excel_template.get().strip() or None
        output, notes, project = self.output.get().strip(), self.include_notes.get(), deepcopy(self.project)
        if not output:
            messagebox.showwarning('输出目录', '请选择输出目录。', parent=self.win)
            return
        def work():
            warnings = []
            for enabled, template, reader in (('Word' in formats, word, sources.word_style),
                                               ('Excel' in formats, excel, sources.excel_style)):
                if enabled and template:
                    try:
                        warnings += reader(template)[1]
                    except Exception:
                        warnings.append('样式参考无法读取，对应格式的失败原因见保存结果；不影响另一格式。')
            result = export.export_project(project, output, formats, word, excel, notes, allow_draft=draft)
            result['warnings'] = warnings
            return result
        def done(result):
            self.results = result['paths']
            text = '\n'.join(result['paths'] + result['errors'] + result['warnings'])
            self.status.set('已生成%d个文件；%d个格式保存失败。' % (len(result['paths']), len(result['errors'])))
            messagebox.showinfo('导出结果' if result['paths'] else '未生成文件', text, parent=self.win)
        self._start('保存文书', work, done)

    def open_result(self):
        if self.results:
            try:
                os.startfile(self.results[0])
            except OSError as exc:
                messagebox.showerror('无法打开', str(exc), parent=self.win)
        else:
            messagebox.showinfo('暂无结果', '成功导出后可打开结果。', parent=self.win)

    def cancel(self):
        if self.busy:
            self.cancel_event.set()
            self.status.set('已请求取消；等待当前请求返回或超时，不再提交后续章节。已提交请求仍可能计费。')

    def close(self):
        if self.busy:
            if not messagebox.askyesno('任务处理中', '关闭会停止后续请求，已提交请求可能计费。仍要关闭？', parent=self.win):
                return
            self.cancel_event.set()
        self.closed = True
        self.win.after_cancel(self.after_id)
        self.win.destroy()
