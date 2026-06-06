import os
import re
import json
import random
import time
import threading
import sqlite3
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import xml.etree.ElementTree as ET
import zipfile

# 核心依赖
import requests
from docx import Document
from pypdf import PdfReader

# ==========================================
# ⚙️ 环境配置与数据库初始化
# ==========================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, "exam_config.json")
DB_FILE = os.path.join(BASE_DIR, "exam_db_v2.sqlite") 

def load_config():
    default_config = {
        "api_key": "", 
        "api_url": "https://api.siliconflow.cn/v1/chat/completions", 
        "model_name": "deepseek-ai/DeepSeek-V3"
    }
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f: return {**default_config, **json.load(f)}
        except: pass
    return default_config

def save_config(data):
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS courses (id INTEGER PRIMARY KEY AUTOINCREMENT, course_name TEXT UNIQUE NOT NULL)''')
    c.execute('''CREATE TABLE IF NOT EXISTS course_chapters (id INTEGER PRIMARY KEY AUTOINCREMENT, course TEXT NOT NULL, chapter_name TEXT NOT NULL, weight REAL DEFAULT 0, UNIQUE(course, chapter_name))''')
    c.execute('''CREATE TABLE IF NOT EXISTS questions (id INTEGER PRIMARY KEY AUTOINCREMENT, course TEXT NOT NULL, chapter_name TEXT DEFAULT '未分类章节', kp_name TEXT, q_type TEXT, question TEXT NOT NULL, answer TEXT, source TEXT, q_year INTEGER, doc_type TEXT, UNIQUE(course, question))''')
    c.execute('''CREATE TABLE IF NOT EXISTS materials (id INTEGER PRIMARY KEY AUTOINCREMENT, course TEXT NOT NULL, raw_text TEXT NOT NULL, doc_type TEXT DEFAULT '教材/讲义大纲')''')
    c.execute('''CREATE TABLE IF NOT EXISTS papers (id INTEGER PRIMARY KEY AUTOINCREMENT, course TEXT NOT NULL, paper_name TEXT NOT NULL, created_at DATETIME DEFAULT CURRENT_TIMESTAMP)''')
    c.execute('''CREATE TABLE IF NOT EXISTS paper_questions (paper_id INTEGER, question_id INTEGER, FOREIGN KEY(paper_id) REFERENCES papers(id), FOREIGN KEY(question_id) REFERENCES questions(id))''')
    
    try: c.execute("ALTER TABLE questions ADD COLUMN chapter_name TEXT DEFAULT '未分类章节'")
    except: pass
    try: c.execute("INSERT OR IGNORE INTO courses (course_name) SELECT DISTINCT course FROM questions")
    except: pass
    conn.commit(); conn.close()

# ==========================================
# 📄 物理层文件解析引擎 
# ==========================================
def extract_text_from_file(filepath):
    ext = filepath.lower().split('.')[-1]
    text = ""
    try:
        if ext in ['md', 'txt']:
            with open(filepath, 'r', encoding='utf-8') as f: text = f.read()
        elif ext == 'pdf':
            reader = PdfReader(filepath)
            for page in reader.pages: text += page.extract_text() + "\n"
        elif ext == 'docx':
            doc = Document(filepath)
            for para in doc.paragraphs:
                if para.text.strip(): text += para.text.strip() + "\n"
            for table in doc.tables:
                for row in table.rows:
                    row_data = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                    if row_data: text += " | ".join(row_data) + "\n"
    except Exception as e: return f"[提取失败: {str(e)}]"
    return text

# ==========================================
# 🤖 AI 引擎集合
# ==========================================
def call_ai_dual_extractor(config, text_chunk, mode="试卷", existing_chapters=None):
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {config['api_key']}"}
    ch_str = f"[{', '.join(existing_chapters)}]" if existing_chapters else "[]"
    
    if mode == "试卷":
        prompt = f"""你是一位细致的大学教授。请无遗漏地提取出物理文本中所有的考试题目。
【强制指令】：
1. 提取所有考题，无答案统一填“无”。
2. 推断该题所属的【章标题】。⭐【铁律】：你只能从以下现有章节列表中选择归属：{ch_str}。如果列表为空，或题目不属于列表中任何一章，必须一律填入 "未分类章节"，绝对不可自己发明新章节名！
3. 提取该题的【核心知识点】。
4. source固定填"真题"."""
    else:
        prompt = f"""你是一位教研专家。请分析教材或讲义文本。
【强制指令】：
1. 推断文本所属的【章标题】。⭐【铁律】：你只能从以下现有章节列表中选择归属：{ch_str}。如果列表为空，或文本不属于列表中任何一章，必须一律填入 "未分类章节"，绝对不可自己发明新章节名！
2. 提取核心【知识点】，并针对每个知识点原创生成2-3道考题。
3. source固定填"AI生成"。自带练习题提取标记"真题"."""

    json_format = """必须输出合法的JSON，严格遵循以下一维数组结构：
{"questions": [{"chapter_name": "现有章名 或 未分类章节", "point_name": "核心知识点名称", "type": "题型", "question": "完整题目", "answer": "答案或无", "source": "真题 或 AI生成"}]}"""
    
    payload = {"model": config['model_name'], "messages": [{"role": "system", "content": prompt + "\n" + json_format}, {"role": "user", "content": f"【待处理文本】：\n{text_chunk}"}], "temperature": 0.1, "response_format": {"type": "json_object"}}
    try:
        resp = requests.post(config['api_url'], headers=headers, json=payload, timeout=120)
        resp.raise_for_status()
        cleaned_text = re.sub(r'^```json\s*', '', resp.json()['choices'][0]['message']['content'], flags=re.MULTILINE)
        cleaned_text = re.sub(r'```\s*$', '', cleaned_text, flags=re.MULTILINE)
        match = re.search(r'\{[\s\S]*\}', cleaned_text)
        if match: return json.loads(match.group(0)), resp.json()['choices'][0]['message']['content']
    except Exception as e: return {"error": str(e)}, str(e)
    return {"questions": []}, "解析失败"

def call_ai_answer_generator(config, course, question, context_text):
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {config['api_key']}"}
    prompt = f"你是一位教授【{course}】的高校教师。请为以下考题撰写【参考答案】。\n"
    if context_text: prompt += f"【指令】：优先检索以下讲义文本总结答案：\n===核心讲义库===\n{context_text[:15000]}\n======\n"
    else: prompt += f"【指令】：未找到讲义，请纯粹运用学科知识解答。\n"
    prompt += f"\n【考试题目】：{question}\n\n直接输出答案文本，不加废话。"
    payload = {"model": config['model_name'], "messages": [{"role": "system", "content": prompt}], "temperature": 0.3}
    try:
        resp = requests.post(config['api_url'], headers=headers, json=payload, timeout=120)
        resp.raise_for_status()
        return resp.json()['choices'][0]['message']['content'].strip()
    except Exception as e: return f"[AI生成阻断]: {str(e)}"

def call_ai_reclassify_kps(config, course, chapters, kps_data):
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {config['api_key']}"}
    prompt = f"""作为【{course}】的教研总监，请将下列【知识点】强行归类到最匹配的【目标章节】。
【绝对指令】：必须且只能从以下列表中挑选一个作为归属章节：[{', '.join(chapters)}]。若无匹配填 "未分类章节"。
严格输出合法的 JSON 格式：
{{"mapping": [{{"kp_name": "知识点名称", "chapter_name": "选出的目标章节"}}]}}"""
    payload = {"model": config['model_name'], "messages": [{"role": "system", "content": prompt}, {"role": "user", "content": f"【知识点及题目】：\n{json.dumps(kps_data, ensure_ascii=False)}"}], "temperature": 0.1, "response_format": {"type": "json_object"}}
    try:
        resp = requests.post(config['api_url'], headers=headers, json=payload, timeout=120)
        resp.raise_for_status()
        txt = resp.json()['choices'][0]['message']['content']
        cleaned = re.sub(r'^```json\s*', '', txt, flags=re.MULTILINE); cleaned = re.sub(r'```\s*$', '', cleaned, flags=re.MULTILINE)
        match = re.search(r'\{[\s\S]*\}', cleaned)
        if match: return json.loads(match.group(0))
    except Exception as e: return {"error": str(e)}
    return {"mapping": []}

def call_ai_extract_chapters_api(config, course, context_text):
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {config['api_key']}"}
    prompt = f"""你是【{course}】的教研专家。请阅读教材大纲内容，提取所有的【章标题】。
【提取铁律】：必须是如“第一章 导论”这样带有“第X章”的标准表述。绝不要提取节标题。
输出合法 JSON：{{"chapters": ["第一章 导论", "第二章 数据收集"]}}"""
    payload = {"model": config['model_name'], "messages": [{"role": "system", "content": prompt}, {"role": "user", "content": f"【讲义内容】：\n{context_text[:12000]}"}], "temperature": 0.1, "response_format": {"type": "json_object"}}
    try:
        resp = requests.post(config['api_url'], headers=headers, json=payload, timeout=120)
        resp.raise_for_status()
        txt = resp.json()['choices'][0]['message']['content']
        cleaned = re.sub(r'^```json\s*', '', txt, flags=re.MULTILINE); cleaned = re.sub(r'```\s*$', '', cleaned, flags=re.MULTILINE)
        match = re.search(r'\{[\s\S]*\}', cleaned)
        if match: return json.loads(match.group(0))
    except Exception as e: return {"error": str(e)}
    return {"chapters": []}

def sync_chapter_weights(course):
    conn = sqlite3.connect(DB_FILE); c = conn.cursor()
    c.execute("SELECT chapter_name, COUNT(*) FROM questions WHERE course=? GROUP BY chapter_name", (course,))
    actual_counts = dict(c.fetchall())
    if not actual_counts: conn.close(); return

    c.execute("SELECT chapter_name, weight FROM course_chapters WHERE course=?", (course,))
    existing_weights = dict(c.fetchall())
    total_q = sum(actual_counts.values())
    for ch, count in actual_counts.items():
        if ch not in existing_weights:
            default_w = round((count / total_q) * 100, 2) if total_q > 0 else 0
            c.execute("INSERT INTO course_chapters (course, chapter_name, weight) VALUES (?, ?, ?)", (course, ch, default_w))
    conn.commit(); conn.close()

# ==========================================
# 🖥️ GUI 主控面板
# ==========================================
class ExamBuilderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("智能期末试卷工程系统 - 教学研究专业版")
        self.root.geometry("1400x900") 
        self.current_selected_course = "" 
        init_db()
        self.config = load_config()
        self.build_ui()
        self.refresh_courses()

    def build_ui(self):
        style = ttk.Style(); style.theme_use('clam')
        top_frame = tk.Frame(self.root, bg="#2c3e50", height=50)
        top_frame.pack(side=tk.TOP, fill=tk.X)
        tk.Label(top_frame, text="📝 试卷工程系统 (顶层设计预设版)", fg="white", bg="#2c3e50", font=("Arial", 16, "bold")).pack(side=tk.LEFT, padx=20, pady=10)
        tk.Button(top_frame, text="⚙️ 设置", command=self.open_settings).pack(side=tk.RIGHT, padx=20)

        # ⭐ 核心修复：优先将黑色日志状态栏强行“钉”在屏幕最底部
        self.txt_log = tk.Text(self.root, height=8, bg="black", fg="#00ff00", font=("Consolas", 10))
        self.txt_log.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=5)

        # 然后再让选项卡占据上方剩余的所有空间，绝不会再发生挤压
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)

        self.tab_import = ttk.Frame(self.notebook); self.notebook.add(self.tab_import, text="📥 AI双轨抽取入库")
        self.build_import_tab(self.tab_import)
        
        self.tab_kb = ttk.Frame(self.notebook); self.notebook.add(self.tab_kb, text="📚 本地全栈知识库管理")
        self.build_kb_tab(self.tab_kb)

        self.tab_manage = ttk.Frame(self.notebook); self.notebook.add(self.tab_manage, text="🛠️ 题库结构与权重总控")
        self.build_manage_tab(self.tab_manage)

        self.tab_generate = ttk.Frame(self.notebook); self.notebook.add(self.tab_generate, text="🖨️ 分层结构化组卷")
        self.build_generate_tab(self.tab_generate)

    def log(self, msg):
        self.root.after(0, lambda: self.txt_log.insert(tk.END, msg + "\n"))
        self.root.after(0, lambda: self.txt_log.see(tk.END))

    def on_tab_changed(self, event):
        selected_tab = self.notebook.tab(self.notebook.select(), "text")
        if "题库结构与权重总控" in selected_tab: self.load_manage_courses()
        elif "分层结构化" in selected_tab: self.refresh_courses()
        elif "本地全栈知识库管理" in selected_tab: self.load_kb_data()

    def refresh_courses(self):
        try:
            conn = sqlite3.connect(DB_FILE); c = conn.cursor()
            c.execute("SELECT course_name FROM courses ORDER BY course_name")
            all_courses = [r[0] for r in c.fetchall()]
            conn.close()
            for combo in ['cb_import_course', 'cb_gen_course', 'cb_kb_course', 'cb_manage_course']:
                if hasattr(self, combo):
                    getattr(self, combo)['values'] = all_courses
                    if all_courses and not getattr(self, combo).get(): getattr(self, combo).current(0)
            self.load_manage_courses(); self.load_kb_data()
        except: pass

    # ⭐ 核心升级：独立弹窗专属建课舱 (支持批量预设章节)
    def add_course(self):
        win = tk.Toplevel(self.root)
        win.title("新增课程及大纲结构")
        win.geometry("450x380")
        win.grab_set() 

        tk.Label(win, text="课程名称 (*必填):", font=("Arial", 11, "bold")).pack(pady=(15, 5))
        ent_course = tk.Entry(win, width=30, font=("Arial", 11)); ent_course.pack(pady=5)

        tk.Label(win, text="预设章节大纲 (可选，每行粘贴一章):", font=("Arial", 10)).pack(pady=(15, 5))
        txt_chapters = tk.Text(win, height=8, width=40, font=("Arial", 10))
        txt_chapters.pack(pady=5)
        # 提供默认格式参考
        txt_chapters.insert(tk.END, "第一章 导论\n第二章 核心理论\n第三章 实证分析\n")

        def save_new_course():
            c_name = ent_course.get().strip()
            if not c_name:
                messagebox.showwarning("警告", "课程名称不能为空！", parent=win)
                return

            ch_text = txt_chapters.get(1.0, tk.END).strip()
            # 以换行符分割，忽略空行
            chapters = [line.strip() for line in ch_text.split('\n') if line.strip()]

            try:
                conn = sqlite3.connect(DB_FILE); c = conn.cursor()
                c.execute("INSERT OR IGNORE INTO courses (course_name) VALUES (?)", (c_name,))
                for ch in chapters:
                    c.execute("INSERT OR IGNORE INTO course_chapters (course, chapter_name, weight) VALUES (?, ?, 0)", (c_name, ch))
                conn.commit(); conn.close()

                self.log(f"✅ 成功创建课程【{c_name}】，并预设了 {len(chapters)} 个大纲章节。")
                self.refresh_courses()
                
                # 自动选中刚刚创建的课
                if hasattr(self, 'cb_import_course'): self.cb_import_course.set(c_name)
                win.destroy()
            except Exception as e:
                messagebox.showerror("错误", f"保存失败: {e}", parent=win)

        tk.Button(win, text="🚀 确定创建", font=("Arial", 11, "bold"), bg="#27ae60", fg="white", command=save_new_course).pack(pady=15)

    # ================= 业务流 1：资料双轨抽取 =================
    def build_import_tab(self, parent):
        f_in = tk.Frame(parent); f_in.pack(fill=tk.X, pady=10)
        tk.Label(f_in, text="目标课程:").pack(side=tk.LEFT, padx=10)
        self.cb_import_course = ttk.Combobox(f_in, state="readonly", width=18); self.cb_import_course.pack(side=tk.LEFT)
        tk.Button(f_in, text="+建课与大纲预设舱", font=("Arial", 9, "bold"), fg="#8e44ad", command=self.add_course).pack(side=tk.LEFT, padx=5)
        tk.Label(f_in, text="资料性质:").pack(side=tk.LEFT, padx=20)
        self.cb_doc_type = ttk.Combobox(f_in, values=["往年真题试卷", "教材/讲义大纲"], state="readonly", width=15); self.cb_doc_type.pack(side=tk.LEFT); self.cb_doc_type.current(0)
        f_file = tk.Frame(parent); f_file.pack(fill=tk.X, pady=10)
        tk.Button(f_file, text="📁 选择本地资料", command=self.select_files).pack(side=tk.LEFT, padx=10)
        self.lbl_files = tk.Label(f_file, text="未选择", fg="blue"); self.lbl_files.pack(side=tk.LEFT)
        self.selected_files = []
        tk.Button(parent, text="🚀 启动 AI 扫描提取", font=("Arial", 11, "bold"), height=2, command=self.start_extraction).pack(fill=tk.X, padx=20, pady=20)

    def select_files(self):
        files = filedialog.askopenfilenames(filetypes=[("文档", "*.docx *.pdf *.md *.txt")])
        if files: self.selected_files = list(files); self.lbl_files.config(text=f"已选 {len(files)} 个文件")

    def start_extraction(self):
        course = self.cb_import_course.get().strip(); mode = "试卷" if "试卷" in self.cb_doc_type.get() else "讲义"
        if not course or not self.selected_files: return messagebox.showwarning("警告", "请选择课程和文件！")
        threading.Thread(target=self.thread_extract, args=(course, mode), daemon=True).start()

    def thread_extract(self, course, mode):
        conn = sqlite3.connect(DB_FILE); c = conn.cursor()
        current_year = int(time.strftime("%Y")); inserted_count = 0 
        
        c.execute("SELECT chapter_name FROM course_chapters WHERE course=? AND chapter_name!='未分类章节'", (course,))
        existing_chapters = [r[0] for r in c.fetchall()]
        
        for filepath in self.selected_files:
            filename = os.path.basename(filepath); self.log(f"\n>>> 物理读取：{filename}")
            f_year = int((re.search(r'(20\d{2})', filename) or re.search(r'(20\d{2})', str(current_year))).group(1))
            full_text = extract_text_from_file(filepath)
            if len(full_text.strip()) < 50: self.log(f"   [警告] 提取字数极少。"); continue
            
            c.execute("SELECT id FROM materials WHERE course=? AND raw_text=?", (course, full_text))
            if not c.fetchone():
                try: 
                    c.execute("INSERT INTO materials (course, raw_text, doc_type) VALUES (?, ?, ?)", (course, full_text, mode))
                    self.log(f"   [文库同步] 该{mode}全文已作为 {mode} 类型永久挂载入知识库。")
                except Exception as e: self.log(f"   [挂载异常] {e}")

            chunk_size = 1500; overlap = 300; chunks = []
            start = 0
            while start < len(full_text):
                end = min(start + chunk_size, len(full_text)); chunks.append(full_text[start:end]); start += chunk_size - overlap
            for idx, chunk in enumerate(chunks):
                if len(chunk.strip()) < 50: continue
                self.log(f"-> AI解析模块 {idx+1}/{len(chunks)}...")
                res, raw_text = call_ai_dual_extractor(self.config, chunk, mode, existing_chapters)
                if "error" in res: continue
                questions = res.get("questions", [])
                if not questions: continue
                for q in questions:
                    if not q.get('question'): continue
                    try:
                        c.execute('''INSERT INTO questions (course, chapter_name, kp_name, q_type, question, answer, source, q_year, doc_type) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                                  (course, q.get('chapter_name', '未分类章节'), q.get('point_name', '未定'), q.get('type','未定'), q['question'], q.get('answer','无'), q.get('source','真题'), f_year, mode))
                        inserted_count += 1
                    except: pass
                conn.commit()
        conn.close()
        sync_chapter_weights(course); self.log(f"\n✅ 【{course}】入库完毕。新增试题: {inserted_count} 道。"); self.root.after(0, self.refresh_courses)

    # ================= 业务流 2：本地库全栈管理 =================
    def build_kb_tab(self, parent):
        f_top = tk.Frame(parent); f_top.pack(fill=tk.X, pady=10)
        tk.Label(f_top, text="归属课程:").pack(side=tk.LEFT, padx=10); self.cb_kb_course = ttk.Combobox(f_top, state="readonly", width=18); self.cb_kb_course.pack(side=tk.LEFT)
        tk.Button(f_top, text="📁 导入外部资料至该课文库", font=("Arial", 10, "bold"), command=self.import_pure_kb).pack(side=tk.LEFT, padx=20)
        tk.Button(f_top, text="🔄 刷新列表", command=self.load_kb_data).pack(side=tk.RIGHT, padx=10)
        paned = ttk.PanedWindow(parent, orient=tk.HORIZONTAL); paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        f_tree = tk.Frame(paned); paned.add(f_tree, weight=2)
        self.tree_kb = ttk.Treeview(f_tree, columns=("id"), show="tree headings"); self.tree_kb.heading("#0", text="📚 知识库源文件列表 (全生命周期)"); self.tree_kb.heading("id", text="库ID"); self.tree_kb.column("id", width=50, anchor="center")
        self.tree_kb.pack(fill=tk.BOTH, expand=True); self.tree_kb.bind("<<TreeviewSelect>>", self.on_kb_select)
        f_edit = ttk.LabelFrame(paned, text="📄 文档物理原文 (知识库基座)"); paned.add(f_edit, weight=3)
        self.lbl_kb_id = tk.Label(f_edit, text="当前未选择", fg="gray"); self.lbl_kb_id.pack(anchor=tk.W, padx=5, pady=2)
        self.txt_kb = tk.Text(f_edit, font=("Arial", 11), wrap=tk.WORD); scroll_kb = ttk.Scrollbar(f_edit, orient="vertical", command=self.txt_kb.yview); self.txt_kb.configure(yscrollcommand=scroll_kb.set); scroll_kb.pack(side=tk.RIGHT, fill=tk.Y); self.txt_kb.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        f_btn = tk.Frame(f_edit); f_btn.pack(fill=tk.X, pady=5)
        tk.Button(f_btn, text="💾 保存编辑", font=("Arial", 10, "bold"), command=self.save_kb_edit).pack(side=tk.LEFT, padx=10)
        tk.Button(f_btn, text="🗑️ 毁灭文档", font=("Arial", 10, "bold"), fg="#c0392b", command=self.delete_kb_doc).pack(side=tk.RIGHT, padx=10)

    def import_pure_kb(self):
        course = self.cb_kb_course.get().strip()
        if not course: return messagebox.showwarning("警告", "请先选择归属课程！")
        files = filedialog.askopenfilenames(filetypes=[("文档", "*.docx *.pdf *.md *.txt")])
        if files: threading.Thread(target=self.thread_import_pure_kb, args=(course, files), daemon=True).start()

    def thread_import_pure_kb(self, course, files):
        conn = sqlite3.connect(DB_FILE); c = conn.cursor(); success = 0
        for fp in files:
            text = extract_text_from_file(fp)
            if len(text.strip()) > 10:
                c.execute("SELECT id FROM materials WHERE course=? AND raw_text=?", (course, text))
                if not c.fetchone(): c.execute("INSERT INTO materials (course, raw_text, doc_type) VALUES (?, ?, ?)", (course, text, '手动挂载资料')); success += 1; self.log(f"📘 [纯净挂载]: {os.path.basename(fp)} 成功。")
        conn.commit(); conn.close(); self.log(f"✅ 知识库新增 {success} 份。"); self.root.after(0, self.load_kb_data)

    def load_kb_data(self):
        for i in self.tree_kb.get_children(): self.tree_kb.delete(i)
        self.lbl_kb_id.config(text="当前未选择"); self.txt_kb.delete(1.0, tk.END)
        try:
            conn = sqlite3.connect(DB_FILE); c = conn.cursor(); c.execute("SELECT course_name FROM courses ORDER BY course_name"); courses = [r[0] for r in c.fetchall()]
            for course in courses:
                c_node = self.tree_kb.insert("", tk.END, text=course, open=True, tags=('course',))
                c.execute("SELECT id, doc_type, SUBSTR(raw_text, 1, 20) FROM materials WHERE course=?", (course,))
                for doc in c.fetchall():
                    dt = doc[1] if doc[1] else "讲义大纲"
                    self.tree_kb.insert(c_node, tk.END, text=f"[{dt}] {doc[2].replace(chr(10),'')}...", values=(doc[0],), tags=('doc',))
            conn.close()
        except: pass

    def on_kb_select(self, event):
        sel = self.tree_kb.selection()
        if not sel: return
        item = self.tree_kb.item(sel[0])
        if 'doc' in item.get("tags", ""):
            doc_id = item['values'][0]
            conn = sqlite3.connect(DB_FILE); c = conn.cursor(); c.execute("SELECT raw_text FROM materials WHERE id=?", (doc_id,)); row = c.fetchone(); conn.close()
            if row: self.lbl_kb_id.config(text=str(doc_id)); self.txt_kb.delete(1.0, tk.END); self.txt_kb.insert(tk.END, row[0])

    def save_kb_edit(self):
        doc_id = self.lbl_kb_id.cget("text"); new_text = self.txt_kb.get(1.0, tk.END).strip()
        if doc_id == "当前未选择": return
        try: conn = sqlite3.connect(DB_FILE); c = conn.cursor(); c.execute("UPDATE materials SET raw_text=? WHERE id=?", (new_text, doc_id)); conn.commit(); conn.close(); messagebox.showinfo("成功", "保存成功！")
        except Exception as e: messagebox.showerror("失败", str(e))

    def delete_kb_doc(self):
        doc_id = self.lbl_kb_id.cget("text")
        if doc_id == "当前未选择": return
        if messagebox.askyesno("警告", "物理删除这份核心材料？"):
            conn = sqlite3.connect(DB_FILE); c = conn.cursor(); c.execute("DELETE FROM materials WHERE id=?", (doc_id,)); conn.commit(); conn.close(); self.load_kb_data()

    # ================= ⭐ 核心重构：三列联动瀑布流控制台 =================
    def build_manage_tab(self, parent):
        f_top = tk.Frame(parent); f_top.pack(fill=tk.X, pady=5)
        tk.Label(f_top, text="全景课程选择:").pack(side=tk.LEFT, padx=10)
        self.cb_manage_course = ttk.Combobox(f_top, state="readonly", width=20); self.cb_manage_course.pack(side=tk.LEFT)
        self.cb_manage_course.bind("<<ComboboxSelected>>", lambda e: self.load_manage_courses())
        tk.Button(f_top, text="🔄 全局重载", command=self.load_manage_courses).pack(side=tk.LEFT, padx=20)

        main_v_paned = ttk.PanedWindow(parent, orient=tk.VERTICAL)
        main_v_paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        top_h_paned = ttk.PanedWindow(main_v_paned, orient=tk.HORIZONTAL)
        main_v_paned.add(top_h_paned, weight=5)

        # 【第一列：章节与控制】
        f_col1_wrap = tk.Frame(top_h_paned); top_h_paned.add(f_col1_wrap, weight=2)
        f_col1 = ttk.LabelFrame(f_col1_wrap, text="1. 章节体系 (点击展示内含知识点)")
        f_col1.pack(fill=tk.BOTH, expand=True)
        self.tree_m_ch = ttk.Treeview(f_col1, columns=("weight"), show="tree headings")
        self.tree_m_ch.heading("#0", text="章名称"); self.tree_m_ch.heading("weight", text="权重%")
        self.tree_m_ch.column("weight", width=50, anchor="center")
        self.tree_m_ch.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        self.tree_m_ch.bind("<<TreeviewSelect>>", self.on_m_ch_select)
        
        f_ctrl = ttk.LabelFrame(f_col1_wrap, text="🕹️ 第一列：章节中枢")
        f_ctrl.pack(fill=tk.X, pady=2)
        
        f_c1 = tk.Frame(f_ctrl); f_c1.pack(fill=tk.X, pady=2)
        self.ent_m_ch = tk.Entry(f_c1, width=10); self.ent_m_ch.pack(side=tk.LEFT, padx=2)
        tk.Label(f_c1, text="权重:").pack(side=tk.LEFT); self.ent_m_w = tk.Entry(f_c1, width=3); self.ent_m_w.pack(side=tk.LEFT)
        tk.Button(f_c1, text="💾修改/增章", font=("Arial", 9), command=self.save_or_add_chapter).pack(side=tk.LEFT, padx=2)
        
        f_c2 = tk.Frame(f_ctrl); f_c2.pack(fill=tk.X, pady=2)
        tk.Button(f_c2, text="➡️ 归并整章", font=("Arial", 9, "bold"), fg="#e67e22", command=self.m_move_ch_dialog).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=1)
        tk.Button(f_c2, text="🗑️ 删整章", font=("Arial", 9, "bold"), fg="#c0392b", command=self.m_del_ch).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=1)

        f_c3 = tk.Frame(f_ctrl); f_c3.pack(fill=tk.X, pady=2)
        self.btn_extract_ch = tk.Button(f_c3, text="🤖 1.专门提取大纲(初次建章)", font=("Arial", 9, "bold"), fg="#2980b9", command=self.ai_extract_chapters)
        self.btn_extract_ch.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=1)
        self.btn_ai_reclass = tk.Button(f_c3, text="🤖 2.打包装车", font=("Arial", 9, "bold"), fg="#8e44ad", command=self.ai_reclassify)
        self.btn_ai_reclass.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=1)

        # 【第二列：知识点】
        f_col2_wrap = tk.Frame(top_h_paned); top_h_paned.add(f_col2_wrap, weight=2)
        f_col2 = ttk.LabelFrame(f_col2_wrap, text="2. 知识点聚合 (可Ctrl/Shift多选)")
        f_col2.pack(fill=tk.BOTH, expand=True)
        self.tree_m_kp = ttk.Treeview(f_col2, columns=("count"), show="tree headings")
        self.tree_m_kp.heading("#0", text="知识点名称"); self.tree_m_kp.heading("count", text="题量")
        self.tree_m_kp.column("count", width=40, anchor="center")
        self.tree_m_kp.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        self.tree_m_kp.bind("<<TreeviewSelect>>", self.on_m_kp_select)
        
        f_kp_ctrl = tk.LabelFrame(f_col2_wrap, text="🕹️ 第二列：知识点操作"); f_kp_ctrl.pack(fill=tk.X, pady=2)
        f_k1 = tk.Frame(f_kp_ctrl); f_k1.pack(fill=tk.X, pady=2)
        tk.Button(f_k1, text="✏️ 改名", font=("Arial", 9), command=self.m_edit_kp).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=1)
        tk.Button(f_k1, text="🔗 合并", font=("Arial", 9, "bold"), fg="#2980b9", command=self.m_merge_kp).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=1)
        tk.Button(f_k1, text="🗑️ 删点", font=("Arial", 9), fg="#c0392b", command=self.m_del_kp).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=1)
        tk.Button(f_kp_ctrl, text="➡️ 移至它章 (支持多选)", font=("Arial", 9, "bold"), fg="#e67e22", command=self.m_move_kp_dialog).pack(fill=tk.X, pady=2, padx=1)

        # 【第三列：题目列表】
        f_col3 = ttk.LabelFrame(top_h_paned, text="3. 试题清单 (双击载入底置编辑器)")
        top_h_paned.add(f_col3, weight=4)
        self.tree_m_q = ttk.Treeview(f_col3, columns=("id", "source", "type", "q"), show="headings")
        self.tree_m_q.heading("id", text="ID"); self.tree_m_q.heading("source", text="来源"); self.tree_m_q.heading("type", text="题型"); self.tree_m_q.heading("q", text="题干预览")
        self.tree_m_q.column("id", width=40, anchor="center"); self.tree_m_q.column("source", width=60, anchor="center"); self.tree_m_q.column("type", width=60, anchor="center"); self.tree_m_q.column("q", width=300)
        scroll_q = ttk.Scrollbar(f_col3, orient="vertical", command=self.tree_m_q.yview)
        self.tree_m_q.configure(yscrollcommand=scroll_q.set); scroll_q.pack(side=tk.RIGHT, fill=tk.Y); self.tree_m_q.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        self.tree_m_q.bind("<<TreeviewSelect>>", self.on_m_q_select)

        # ====== 下方：底置编辑台 ======
        f_edit = ttk.LabelFrame(main_v_paned, text="✏️ 试题人工覆写台")
        main_v_paned.add(f_edit, weight=3)
        f_meta = tk.Frame(f_edit); f_meta.pack(fill=tk.X, pady=2, padx=5)
        tk.Label(f_meta, text="题目ID:").pack(side=tk.LEFT)
        self.lbl_id = tk.Label(f_meta, text="未选择", fg="blue", width=5); self.lbl_id.pack(side=tk.LEFT)
        tk.Label(f_meta, text=" | 题型:").pack(side=tk.LEFT)
        self.ent_type = tk.Entry(f_meta, width=12); self.ent_type.pack(side=tk.LEFT, padx=5)

        f_text = tk.Frame(f_edit); f_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=2)
        f_t_q = tk.Frame(f_text); f_t_q.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0,5))
        tk.Label(f_t_q, text="题目原文:").pack(anchor=tk.W); self.txt_q = tk.Text(f_t_q, font=("Arial", 11)); self.txt_q.pack(fill=tk.BOTH, expand=True)

        f_t_a = tk.Frame(f_text); f_t_a.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5,0))
        f_a_top = tk.Frame(f_t_a); f_a_top.pack(fill=tk.X)
        tk.Label(f_a_top, text="标准答案:").pack(side=tk.LEFT)
        self.btn_ai = tk.Button(f_a_top, text="🤖 RAG 文库提取作答", font=("Arial", 9, "bold"), fg="#2980b9", command=self.generate_ai_answer)
        self.btn_ai.pack(side=tk.RIGHT)
        self.txt_a = tk.Text(f_t_a, font=("Arial", 11)); self.txt_a.pack(fill=tk.BOTH, expand=True)

        f_btn = tk.Frame(f_edit); f_btn.pack(fill=tk.X, pady=5)
        tk.Button(f_btn, text="💾 保存题目修改", font=("Arial", 11, "bold"), fg="#27ae60", command=self.save_edit).pack(side=tk.LEFT, padx=20)
        tk.Button(f_btn, text="🗑️ 删除此题", font=("Arial", 11, "bold"), fg="#c0392b", command=self.delete_q).pack(side=tk.RIGHT, padx=20)

    # --- 数据加载机制 ---
    def load_manage_courses(self):
        for i in self.tree_m_ch.get_children(): self.tree_m_ch.delete(i)
        for i in self.tree_m_kp.get_children(): self.tree_m_kp.delete(i)
        for i in self.tree_m_q.get_children(): self.tree_m_q.delete(i)
        self.clear_editor()
        
        course = self.cb_manage_course.get()
        if not course: return
        
        sync_chapter_weights(course)
        conn = sqlite3.connect(DB_FILE); c = conn.cursor()
        c.execute("SELECT chapter_name, weight FROM course_chapters WHERE course=?", (course,))
        for r in c.fetchall():
            self.tree_m_ch.insert("", tk.END, text=r[0], values=(f"{r[1]}",), tags=('ch',))
        conn.close()

    def on_m_ch_select(self, event):
        sel = self.tree_m_ch.selection()
        if not sel: return
        ch_name = self.tree_m_ch.item(sel[0])['text']
        self.ent_m_ch.delete(0, tk.END); self.ent_m_ch.insert(0, ch_name)
        self.ent_m_w.delete(0, tk.END); self.ent_m_w.insert(0, self.tree_m_ch.item(sel[0])['values'][0].replace('%', ''))
        
        course = self.cb_manage_course.get()
        for i in self.tree_m_kp.get_children(): self.tree_m_kp.delete(i)
        for i in self.tree_m_q.get_children(): self.tree_m_q.delete(i)
        self.clear_editor()

        conn = sqlite3.connect(DB_FILE); c = conn.cursor()
        c.execute("SELECT kp_name, COUNT(*) FROM questions WHERE course=? AND chapter_name=? GROUP BY kp_name", (course, ch_name))
        for r in c.fetchall(): self.tree_m_kp.insert("", tk.END, text=r[0], values=(r[1],), tags=('kp',))
        conn.close()

    def on_m_kp_select(self, event):
        sel_kps = self.tree_m_kp.selection()
        sel_ch = self.tree_m_ch.selection()
        if not sel_kps or not sel_ch: return
        kp_name = self.tree_m_kp.item(sel_kps[0])['text'] 
        ch_name = self.tree_m_ch.item(sel_ch[0])['text']
        course = self.cb_manage_course.get()

        for i in self.tree_m_q.get_children(): self.tree_m_q.delete(i)
        self.clear_editor()

        conn = sqlite3.connect(DB_FILE); c = conn.cursor()
        c.execute("SELECT id, source, q_type, question FROM questions WHERE course=? AND chapter_name=? AND kp_name=?", (course, ch_name, kp_name))
        for r in c.fetchall():
            source = r[1] if r[1] else "未知"
            preview = r[3].replace('\n', ' ')[:40] + "..." if len(r[3]) > 40 else r[3].replace('\n', ' ')
            self.tree_m_q.insert("", tk.END, values=(r[0], source, r[2], preview))
        conn.close()

    def on_m_q_select(self, event):
        sel = self.tree_m_q.selection()
        if not sel: return
        q_id = self.tree_m_q.item(sel[0])['values'][0]
        conn = sqlite3.connect(DB_FILE); c = conn.cursor()
        c.execute("SELECT q_type, question, answer FROM questions WHERE id=?", (q_id,))
        row = c.fetchone(); conn.close()
        if row:
            self.lbl_id.config(text=str(q_id)); self.ent_type.delete(0, tk.END); self.ent_type.insert(0, row[0])
            self.txt_q.delete(1.0, tk.END); self.txt_q.insert(tk.END, row[1])
            self.txt_a.delete(1.0, tk.END); self.txt_a.insert(tk.END, row[2])

    def clear_editor(self):
        self.lbl_id.config(text="未选择"); self.ent_type.delete(0, tk.END)
        self.txt_q.delete(1.0, tk.END); self.txt_a.delete(1.0, tk.END)

    # --- 章控制 (第一列) ---
    def save_or_add_chapter(self):
        course = self.cb_manage_course.get()
        if not course: return
        new_ch = self.ent_m_ch.get().strip()
        try: new_w = float(self.ent_m_w.get() or 0)
        except: return messagebox.showerror("错误", "权重必须数字。")
        if not new_ch: return
        
        sel = self.tree_m_ch.selection()
        conn = sqlite3.connect(DB_FILE); c = conn.cursor()
        if sel: 
            old_ch = self.tree_m_ch.item(sel[0])['text']
            if old_ch != "未分类章节":
                c.execute("UPDATE course_chapters SET chapter_name=?, weight=? WHERE course=? AND chapter_name=?", (new_ch, new_w, course, old_ch))
                c.execute("UPDATE questions SET chapter_name=? WHERE course=? AND chapter_name=?", (new_ch, course, old_ch))
            else:
                c.execute("INSERT OR REPLACE INTO course_chapters (course, chapter_name, weight) VALUES (?, ?, ?)", (course, new_ch, new_w))
        else:
            c.execute("INSERT OR REPLACE INTO course_chapters (course, chapter_name, weight) VALUES (?, ?, ?)", (course, new_ch, new_w))
        conn.commit(); conn.close(); self.load_manage_courses()

    def m_del_ch(self):
        sel_ch = self.tree_m_ch.selection()
        if not sel_ch: return messagebox.showwarning("提示", "请选中要删除的章！")
        ch = self.tree_m_ch.item(sel_ch[0])['text']
        course = self.cb_manage_course.get()
        if messagebox.askyesno("危险", f"彻底删除章节【{ch}】下的全部知识点和试题？"):
            conn = sqlite3.connect(DB_FILE); c = conn.cursor()
            c.execute("DELETE FROM questions WHERE course=? AND chapter_name=?", (course, ch))
            c.execute("DELETE FROM course_chapters WHERE course=? AND chapter_name=?", (course, ch))
            conn.commit(); conn.close(); self.load_manage_courses()

    def m_move_ch_dialog(self):
        sel_ch = self.tree_m_ch.selection()
        if not sel_ch: return messagebox.showwarning("提示", "请在左侧选中要归并的【章节】！")
        old_ch = self.tree_m_ch.item(sel_ch[0])['text']
        course = self.cb_manage_course.get()

        conn = sqlite3.connect(DB_FILE); c = conn.cursor()
        c.execute("SELECT chapter_name FROM course_chapters WHERE course=?", (course,))
        chapters = [r[0] for r in c.fetchall() if r[0] != old_ch]
        conn.close()

        if not chapters: return messagebox.showwarning("提示", "没有其他章节可供合并！")

        win = tk.Toplevel(self.root); win.title("整章合并"); win.geometry("350x200"); win.grab_set()
        tk.Label(win, text=f"将【{old_ch}】的全部内容\n倾泻并入下面哪一章？", font=("Arial", 11)).pack(pady=15)
        cb = ttk.Combobox(win, values=chapters, state="readonly", font=("Arial", 11), width=20); cb.pack(pady=5); cb.set(chapters[0])

        def do_move():
            new_ch = cb.get()
            if new_ch:
                if not messagebox.askyesno("二次确认", f"合并后原【{old_ch}】将被删除。\n确认合并入【{new_ch}】吗？"): return
                conn = sqlite3.connect(DB_FILE); cur = conn.cursor()
                cur.execute("UPDATE questions SET chapter_name=? WHERE course=? AND chapter_name=?", (new_ch, course, old_ch))
                cur.execute("DELETE FROM course_chapters WHERE course=? AND chapter_name=?", (course, old_ch))
                conn.commit(); conn.close()
                self.load_manage_courses(); self.log(f"✅ 整章合并完成：【{old_ch}】已全部并入【{new_ch}】。")
            win.destroy()
        tk.Button(win, text="🚀 确定合并", command=do_move, font=("Arial", 11, "bold"), bg="#c0392b", fg="white").pack(pady=15)

    # --- 知识点控制 (第二列) ---
    def m_edit_kp(self):
        sel_kp = self.tree_m_kp.selection()
        sel_ch = self.tree_m_ch.selection()
        if not sel_kp or not sel_ch: return
        old_kp = self.tree_m_kp.item(sel_kp[0])['text']; ch = self.tree_m_ch.item(sel_ch[0])['text']; course = self.cb_manage_course.get()
        new_kp = simpledialog.askstring("修改名称", "新知识点名称:", initialvalue=old_kp)
        if new_kp and new_kp.strip() and new_kp.strip() != old_kp:
            conn = sqlite3.connect(DB_FILE); c = conn.cursor()
            c.execute("UPDATE questions SET kp_name=? WHERE course=? AND chapter_name=? AND kp_name=?", (new_kp.strip(), course, ch, old_kp))
            conn.commit(); conn.close(); self.on_m_ch_select(None) 

    def m_del_kp(self):
        sel_kps = self.tree_m_kp.selection()
        sel_ch = self.tree_m_ch.selection()
        if not sel_kps or not sel_ch: return
        ch = self.tree_m_ch.item(sel_ch[0])['text']
        course = self.cb_manage_course.get()
        kp_names = [self.tree_m_kp.item(item)['text'] for item in sel_kps]

        if messagebox.askyesno("警告", f"物理删除选中的 {len(kp_names)} 个知识点下所有试题？"):
            conn = sqlite3.connect(DB_FILE); c = conn.cursor()
            for kp in kp_names:
                c.execute("DELETE FROM questions WHERE course=? AND chapter_name=? AND kp_name=?", (course, ch, kp))
            conn.commit(); conn.close(); self.on_m_ch_select(None)

    def m_move_kp_dialog(self):
        sel_kps = self.tree_m_kp.selection()
        sel_ch = self.tree_m_ch.selection()
        if not sel_kps or not sel_ch: return messagebox.showwarning("提示", "请先在列表中选中要转移的【知识点】（按住Ctrl/Shift可多选）！")
        old_ch = self.tree_m_ch.item(sel_ch[0])['text']
        course = self.cb_manage_course.get()
        kp_names = [self.tree_m_kp.item(item)['text'] for item in sel_kps]

        conn = sqlite3.connect(DB_FILE); c = conn.cursor()
        c.execute("SELECT chapter_name FROM course_chapters WHERE course=?", (course,))
        chapters = [r[0] for r in c.fetchall()]
        conn.close()

        win = tk.Toplevel(self.root); win.title("知识点大挪移"); win.geometry("350x180"); win.grab_set() 
        lbl_text = f"将选定的 {len(kp_names)} 个知识点移至：" if len(kp_names)>1 else f"将【{kp_names[0]}】移至："
        tk.Label(win, text=lbl_text, font=("Arial", 11)).pack(pady=15)
        cb = ttk.Combobox(win, values=chapters, state="readonly", font=("Arial", 11), width=20); cb.pack(pady=5)
        if chapters: cb.set(chapters[0])

        def do_move():
            new_ch = cb.get()
            if new_ch and new_ch != old_ch:
                conn = sqlite3.connect(DB_FILE); cur = conn.cursor()
                for kp in kp_names:
                    cur.execute("UPDATE questions SET chapter_name=? WHERE course=? AND chapter_name=? AND kp_name=?", (new_ch, course, old_ch, kp))
                conn.commit(); conn.close()
                self.load_manage_courses()
                self.log(f"✅ 成功将 {len(kp_names)} 个知识点空降至【{new_ch}】。")
            win.destroy()
        tk.Button(win, text="🚀 确定转移", command=do_move, font=("Arial", 11, "bold"), bg="#e67e22", fg="white").pack(pady=15)

    def m_merge_kp(self):
        sel_kps = self.tree_m_kp.selection()
        sel_ch = self.tree_m_ch.selection()
        if not sel_kps or not sel_ch: return messagebox.showwarning("提示", "请按住 Ctrl 或 Shift 选择 2 个以上要合并的知识点！")
        if len(sel_kps) < 2: return messagebox.showwarning("提示", "合并操作至少需要选中 2 个知识点！")

        ch = self.tree_m_ch.item(sel_ch[0])['text']; course = self.cb_manage_course.get()
        kp_names = [self.tree_m_kp.item(item)['text'] for item in sel_kps]

        win = tk.Toplevel(self.root); win.title("聚合合并知识点"); win.geometry("380x200"); win.grab_set()
        tk.Label(win, text=f"将选定的 {len(kp_names)} 个知识点统一合并为：\n(从下拉框选一个正统名称，或手动输入新名字)", font=("Arial", 11)).pack(pady=15)
        cb = ttk.Combobox(win, values=kp_names, font=("Arial", 11), width=25); cb.pack(pady=5); cb.set(kp_names[0]) 

        def do_merge():
            target_kp = cb.get().strip()
            if not target_kp: return
            conn = sqlite3.connect(DB_FILE); c = conn.cursor()
            placeholders = ','.join('?' for _ in kp_names)
            query = f"UPDATE questions SET kp_name=? WHERE course=? AND chapter_name=? AND kp_name IN ({placeholders})"
            params = [target_kp, course, ch] + kp_names
            c.execute(query, params)
            conn.commit(); conn.close()
            self.on_m_ch_select(None) 
            self.log(f"✅ 成功将 {len(kp_names)} 个知识点强行聚合为【{target_kp}】。")
            win.destroy()
        tk.Button(win, text="🔗 确定聚合", command=do_merge, font=("Arial", 11, "bold"), bg="#27ae60", fg="white").pack(pady=15)

    # --- AI 全局干预引擎 ---
    def ai_extract_chapters(self):
        course = self.cb_manage_course.get()
        if not course: return
        if not messagebox.askyesno("提取大纲", "AI将阅读本地文库，为你全自动建立符合规范的【标准章节骨架】？"): return
        self.btn_extract_ch.config(state=tk.DISABLED, text="⏳ AI提取中...")
        threading.Thread(target=self.thread_ai_extract_chapters, args=(course,), daemon=True).start()

    def thread_ai_extract_chapters(self, course):
        self.log(f"\n>>> 启动 AI 文库提取大纲：{course}")
        try:
            conn = sqlite3.connect(DB_FILE); c = conn.cursor()
            c.execute("SELECT raw_text FROM materials WHERE course=? AND doc_type='教材/讲义大纲'", (course,))
            combined_context = "\n\n".join([r[0] for r in c.fetchall()])[:15000]
            conn.close()
            
            if not combined_context: 
                self.log("❌ 知识库为空，无法提取。"); messagebox.showerror("中断", "知识库内没有该课程的讲义原稿！"); return
            
            res = call_ai_extract_chapters_api(self.config, course, combined_context)
            chapters = res.get("chapters", [])
            if not chapters: 
                self.log("❌ 未能提取到明确的章标题。"); messagebox.showwarning("抱歉", "AI未发现'第一章'格式的内容。"); return
            
            conn = sqlite3.connect(DB_FILE); c = conn.cursor(); success = 0
            for ch in chapters:
                try: c.execute("INSERT INTO course_chapters (course, chapter_name, weight) VALUES (?, ?, 0)", (course, ch)); success += 1
                except: pass
            conn.commit(); conn.close()
            self.log(f"✅ 成功建立了 {success} 个标准章节！")
            messagebox.showinfo("大功告成", f"成功抓取了 {success} 个章标题！")
            self.root.after(0, self.load_manage_courses)
        except Exception as e:
            self.log(f"❌ 提取失败: {e}")
        finally:
            self.root.after(0, lambda: self.btn_extract_ch.config(state=tk.NORMAL, text="🤖 1.专门提取大纲(初次建章)"))

    def ai_reclassify(self):
        course = self.cb_manage_course.get()
        if not course: return
        if not messagebox.askyesno("AI重组", f"召唤 AI 将【{course}】所有的知识点打包装车，自动匹配左侧章节？"): return
        self.btn_ai_reclass.config(state=tk.DISABLED, text="⏳ AI重组中...")
        threading.Thread(target=self.thread_ai_reclassify, args=(course,), daemon=True).start()

    def thread_ai_reclassify(self, course):
        self.log(f"\n>>> 启动 AI 自动章节归类大重组：{course}")
        try:
            conn = sqlite3.connect(DB_FILE); c = conn.cursor()
            c.execute("SELECT chapter_name FROM course_chapters WHERE course=? AND chapter_name!='未分类章节'", (course,))
            chapters = [r[0] for r in c.fetchall()]
            if len(chapters) < 1: 
                self.log("❌ 左侧无正规章节。"); conn.close(); messagebox.showwarning("错误", "没有章节，无法重组！")
                return
                
            c.execute("SELECT kp_name, question FROM questions WHERE course=? GROUP BY kp_name", (course,))
            all_kps = [{"kp_name": r[0], "question_sample": r[1][:200]} for r in c.fetchall()]
            conn.close()
            if not all_kps: return
            
            batch_size = 20; total_batches = (len(all_kps) // batch_size) + 1
            self.log(f"   发现 {len(all_kps)} 个知识点，分 {total_batches} 批交由 AI 分析...")
            
            conn = sqlite3.connect(DB_FILE); c = conn.cursor(); success_count = 0
            for i in range(0, len(all_kps), batch_size):
                batch = all_kps[i:i+batch_size]; self.log(f"   装载第 {i//batch_size + 1}/{total_batches} 批次...")
                res = call_ai_reclassify_kps(self.config, course, chapters, batch)
                for m in res.get("mapping", []):
                    kp, ch = m.get("kp_name"), m.get("chapter_name")
                    if kp and ch and ch in chapters:
                        c.execute("UPDATE questions SET chapter_name=? WHERE course=? AND kp_name=?", (ch, course, kp))
                        success_count += 1
                conn.commit(); time.sleep(1) 
            conn.close()
            self.log(f"✅ AI 大重组完毕！强行匹配了 {success_count} 个知识点。")
            messagebox.showinfo("重组完毕", f"完美！AI已成功帮你将 {success_count} 个知识点塞进对应章节。")
            self.root.after(0, self.load_manage_courses)
        except Exception as e:
            self.log(f"❌ 重组失败: {e}")
        finally:
            self.root.after(0, lambda: self.btn_ai_reclass.config(state=tk.NORMAL, text="🤖 2.打包装车"))

    # --- 第三列及底置编辑器操作 ---
    def save_edit(self):
        q_id = self.lbl_id.cget("text")
        if q_id == "未选择": return
        qt = self.ent_type.get().strip(); q = self.txt_q.get(1.0, tk.END).strip(); a = self.txt_a.get(1.0, tk.END).strip()
        try:
            conn = sqlite3.connect(DB_FILE); c = conn.cursor()
            c.execute("UPDATE questions SET q_type=?, question=?, answer=? WHERE id=?", (qt, q, a, q_id))
            conn.commit(); conn.close(); messagebox.showinfo("成功", "保存成功！"); self.on_m_kp_select(None)
        except Exception as e: messagebox.showerror("失败", str(e))

    def delete_q(self):
        q_id = self.lbl_id.cget("text")
        if q_id == "未选择": return
        if messagebox.askyesno("警告", "物理删除此题？"):
            conn = sqlite3.connect(DB_FILE); c = conn.cursor(); c.execute("DELETE FROM questions WHERE id=?", (q_id,)); conn.commit(); conn.close(); self.on_m_kp_select(None) 

    def generate_ai_answer(self):
        q_id = self.lbl_id.cget("text")
        if q_id == "未选择": return
        course = self.cb_manage_course.get(); question = self.txt_q.get(1.0, tk.END).strip()
        self.btn_ai.config(text="⏳ RAG检索中...", state=tk.DISABLED); self.txt_a.delete(1.0, tk.END); self.txt_a.insert(tk.END, "挂载库请求中...\n")
        threading.Thread(target=self.thread_rag_answer, args=(course, question), daemon=True).start()

    def thread_rag_answer(self, course, question):
        try:
            conn = sqlite3.connect(DB_FILE); c = conn.cursor()
            c.execute("SELECT raw_text FROM materials WHERE course=? AND doc_type IN ('教材/讲义大纲', '往年真题试卷', '手动挂载资料')", (course,))
            combined_context = "\n\n".join([r[0] for r in c.fetchall()])[:15000]
            conn.close()
            ans = call_ai_answer_generator(self.config, course, question, combined_context)
            self.root.after(0, self.update_answer_field, ans)
        except Exception as e: self.root.after(0, self.update_answer_field, f"异常: {e}")

    def update_answer_field(self, ans_text):
        self.txt_a.delete(1.0, tk.END); self.txt_a.insert(tk.END, ans_text); self.btn_ai.config(text="🤖 RAG 文库提取作答", state=tk.NORMAL)

    # ================= 业务流 4：分层结构化抽样组卷 =================
    def build_generate_tab(self, parent):
        f_top = tk.Frame(parent); f_top.pack(fill=tk.X, pady=10)
        tk.Label(f_top, text="选择课程:").pack(side=tk.LEFT, padx=10); self.cb_gen_course = ttk.Combobox(f_top, state="readonly"); self.cb_gen_course.pack(side=tk.LEFT)
        f_rule = ttk.LabelFrame(parent, text="参数设置"); f_rule.pack(fill=tk.X, padx=10, pady=10)
        tk.Label(f_rule, text="总题数:").grid(row=0, column=0, padx=5, pady=5); self.ent_total = tk.Entry(f_rule, width=10); self.ent_total.insert(0, "20"); self.ent_total.grid(row=0, column=1)
        tk.Label(f_rule, text="(算法基于'题库结构与权重总控'页面的比例分配配额)").grid(row=0, column=2, padx=20)
        tk.Button(parent, text="✨ 执行分层权重抽样组卷", font=("Arial", 11, "bold"), height=2, command=self.start_gen).pack(fill=tk.X, padx=50, pady=20)

    def start_gen(self):
        course = self.cb_gen_course.get()
        if not course: return messagebox.showwarning("警告", "请选择出卷课程！")
        threading.Thread(target=self.thread_gen, args=(course,), daemon=True).start()

    def get_chapter_quotas(self, total_q, weights_dict):
        quotas = {}; remainders = {}; allocated = 0; total_w = sum(weights_dict.values())
        if total_w == 0: return {ch: 0 for ch in weights_dict}
        for ch, w in weights_dict.items():
            exact = total_q * (w / total_w); base = int(exact); quotas[ch] = base; remainders[ch] = exact - base; allocated += base
        for ch in sorted(remainders.keys(), key=lambda k: remainders[k], reverse=True)[:total_q - allocated]: quotas[ch] += 1
        return quotas

    def generate_paper_text(self, course, name, qs):
        text = f"{course} - {name}\n\n一、题目\n"
        for i, q in enumerate(qs): text += f"{i+1}. 【{q['q_type']}】 {q['question']}\n"
        text += "\n二、答案\n"
        for i, q in enumerate(qs): text += f"{i+1}. {q['answer']} ({q['chapter_name']}-{q['kp_name']})\n"
        return text

    def thread_gen(self, course):
        self.log(f"\n>>> 检索【{course}】执行抽题...")
        conn = sqlite3.connect(DB_FILE); conn.row_factory = sqlite3.Row; c = conn.cursor()
        curr_yr = int(time.strftime("%Y"))
        c.execute("SELECT * FROM course_chapters WHERE course=?", (course,)); weights = {r['chapter_name']: r['weight'] for r in c.fetchall()}
        if not weights or sum(weights.values()) == 0: self.log("❌ 权重未设置！"); conn.close(); return
        try: target_total = int(self.ent_total.get())
        except: self.log("❌ 题数错误"); conn.close(); return
        
        quotas = self.get_chapter_quotas(target_total, weights)
        c.execute("SELECT * FROM questions WHERE course=? AND source='真题' AND q_year BETWEEN ? AND ?", (course, curr_yr-2, curr_yr)); p_restr = [dict(r) for r in c.fetchall()]
        c.execute("SELECT * FROM questions WHERE course=? AND (source!='真题' OR q_year < ?)", (course, curr_yr-2)); p_free = [dict(r) for r in c.fetchall()]

        def draw(pool_re, pool_fr, count):
            sel = []; max_re = max(1, int(count * 0.1)); num_re = min(max_re, len(pool_re))
            if num_re > 0: d = random.sample(pool_re, num_re); sel.extend(d); [pool_re.remove(q) for q in d]
            num_fr = count - len(sel)
            if num_fr > len(pool_fr): raise ValueError()
            d = random.sample(pool_fr, num_fr); sel.extend(d); [pool_fr.remove(q) for q in d]
            return sel

        paper_A = []; paper_B = []
        try:
            for ch, count in quotas.items():
                if count == 0: continue
                ch_r_A = [q for q in p_restr if q['chapter_name'] == ch]; ch_f_A = [q for q in p_free if q['chapter_name'] == ch]
                ch_r_B = ch_r_A.copy(); ch_f_B = ch_f_A.copy()
                try: paper_A.extend(draw(ch_r_A, ch_f_A, count)); paper_B.extend(draw(ch_r_B, ch_f_B, count))
                except: self.log(f"❌ 【{ch}】题库不足阻断！"); return
            random.shuffle(paper_A); random.shuffle(paper_B); ts = time.strftime("%H%M%S")
            self.export(course, f"期末A卷_{ts}", paper_A); self.export(course, f"期末B卷_{ts}", paper_B)
            
            text_A = self.generate_paper_text(course, f"期末A卷_{ts}", paper_A)
            text_B = self.generate_paper_text(course, f"期末B卷_{ts}", paper_B)
            c.execute("INSERT INTO materials (course, raw_text, doc_type) VALUES (?, ?, ?)", (course, text_A, '生成的试卷'))
            c.execute("INSERT INTO materials (course, raw_text, doc_type) VALUES (?, ?, ?)", (course, text_B, '生成的试卷'))
            conn.commit()
            self.log(f"✅ 基于 {len(quotas)} 个章节权重的 A/B 卷装配成功！(已同步归档至知识库)")
        except Exception as e: self.log(f"❌ 抽样报错: {e}")
        finally: conn.close()

    def export(self, course, name, qs):
        doc = Document(); doc.add_heading(f"{course} - {name}", 0); doc.add_heading("一、题目", 1)
        for i, q in enumerate(qs): doc.add_paragraph(f"{i+1}. 【{q['q_type']}】 {q['question']}")
        doc.add_page_break(); doc.add_heading("二、答案", 1)
        for i, q in enumerate(qs): doc.add_paragraph(f"{i+1}. {q['answer']} ({q['chapter_name']}-{q['kp_name']})")
        doc.save(f"{course}_{name}.docx")

    def open_settings(self):
        win = tk.Toplevel(self.root); win.title("API设置"); win.geometry("400x200")
        tk.Label(win, text="API Key:").pack(); e_key = tk.Entry(win, width=40); e_key.insert(0, self.config['api_key']); e_key.pack()
        def s(): self.config['api_key']=e_key.get(); save_config(self.config); win.destroy()
        tk.Button(win, text="保存", command=s).pack(pady=20)

if __name__ == "__main__":
    root = tk.Tk(); app = ExamBuilderApp(root); root.mainloop()