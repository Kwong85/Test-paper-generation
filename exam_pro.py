import os
import re
import json
import random
import time
import threading
import sqlite3
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog


# 核心依赖
import requests
import docx
from docx import Document
from docx.shared import Pt, Inches, RGBColor  # 确保 Pt 和 Inches 也是全局的
from docx.oxml.ns import qn          # 🚀 核心修复：将 qn 提升至全局最顶部导入

try:
    import pdfplumber
except ImportError:
    os.system("/usr/local/bin/python3 -m pip install pdfplumber")
    import pdfplumber

# ==========================================
# ⚙️ 环境配置与数据库初始化
# ==========================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, "exam_config.json")
DB_FILE = os.path.join(BASE_DIR, "exam_db_v2.sqlite") 

# ==========================================
# 🤖 2026 顶层设计：主流大模型平台与模型矩阵字典
# ==========================================
AI_PLATFORM_MANIFEST = {
    "硅基流动 (SiliconFlow)": {
        "url": "https://api.siliconflow.cn/v1/chat/completions",
        "models": ["deepseek-ai/DeepSeek-V4-Pro", "deepseek-ai/DeepSeek-V4-Flash", "deepseek-ai/DeepSeek-R1", "Qwen/Qwen3.5-397B-A17B", "THUDM/GLM-5", "moonshotai/Kimi-K2.6"]
    },
    "DeepSeek 官方平台": {
        "url": "https://api.deepseek.com/chat/completions",
        "models": ["deepseek-v4-flash", "deepseek-v4-pro", "deepseek-chat", "deepseek-reasoner"]
    },
    "阿里云百炼平台 (ModelStudio)": {
        "url": "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions",
        "models": ["deepseek-v4-pro", "qwen-max-2026", "qwen-plus", "qwen-turbo"]
    },
    "通义千问官方 (DashScope)": {
        "url": "https://dashscope.aliyuncs.com/api/v1/services/aigc/text-generation/generation",
        "models": ["qwen-max", "qwen-plus"]
    }
}

def load_config():
    default_config = {
        "platform_name": "硅基流动 (SiliconFlow)",
        "api_key": "", 
        "api_url": "https://api.siliconflow.cn/v1/chat/completions", 
        "model_name": "deepseek-ai/DeepSeek-V4-Pro"
    }
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f: 
                return {**default_config, **json.load(f)}
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
    c.execute('''CREATE TABLE IF NOT EXISTS questions (id INTEGER PRIMARY KEY AUTOINCREMENT, course TEXT NOT NULL, chapter_name TEXT DEFAULT '未分类章节', kp_name TEXT, kp_stars INTEGER DEFAULT 3, q_type TEXT, question TEXT NOT NULL, answer TEXT, source TEXT, q_year INTEGER, doc_type TEXT, UNIQUE(course, question))''')
    c.execute('''CREATE TABLE IF NOT EXISTS materials (id INTEGER PRIMARY KEY AUTOINCREMENT, course TEXT NOT NULL, raw_text TEXT NOT NULL, doc_type TEXT DEFAULT '教材/讲义大纲')''')
    c.execute('''CREATE TABLE IF NOT EXISTS papers (id INTEGER PRIMARY KEY AUTOINCREMENT, course TEXT NOT NULL, paper_name TEXT NOT NULL, created_at DATETIME DEFAULT CURRENT_TIMESTAMP)''')
    c.execute('''CREATE TABLE IF NOT EXISTS paper_questions (paper_id INTEGER, question_id INTEGER, score REAL DEFAULT 0, FOREIGN KEY(paper_id) REFERENCES papers(id), FOREIGN KEY(question_id) REFERENCES questions(id))''')
    c.execute('''CREATE TABLE IF NOT EXISTS system_q_types (id INTEGER PRIMARY KEY AUTOINCREMENT, type_name TEXT UNIQUE NOT NULL)''')
    
    default_types = ["单选题", "多选题", "判断题", "填空题", "名词解释", "简答题", "论述题", "材料分析题"]
    for t in default_types:
        try: c.execute("INSERT OR IGNORE INTO system_q_types (type_name) VALUES (?)", (t,))
        except: pass

    try: c.execute("ALTER TABLE questions ADD COLUMN chapter_name TEXT DEFAULT '未分类章节'")
    except: pass
    try: c.execute("ALTER TABLE questions ADD COLUMN kp_stars INTEGER DEFAULT 3")
    except: pass
    try: c.execute("ALTER TABLE paper_questions ADD COLUMN score REAL DEFAULT 0")
    except: pass
    try: c.execute("ALTER TABLE paper_questions ADD COLUMN prev_year INTEGER")
    except: pass
    try: c.execute("INSERT OR IGNORE INTO courses (course_name) SELECT DISTINCT course FROM questions")
    except: pass
    conn.commit(); conn.close()

def get_standard_q_types():
    conn = sqlite3.connect(DB_FILE); c = conn.cursor()
    c.execute("SELECT type_name FROM system_q_types ORDER BY id")
    types = [r[0] for r in c.fetchall()]
    conn.close()
    if not types: return ["单选题", "多选题", "名词解释", "简答题", "论述题", "材料分析题"]
    return types

def normalize_question_type(raw_type, standard_types):
    if not raw_type: return standard_types[0]
    cleaned = raw_type.strip()
    mapping_rules = {
        r'.*单项选择.*': '单选题',
        r'^单选$': '单选题',
        r'^选择题$': '单选题',
        r'.*多项选择.*': '多选题',
        r'^多选$': '多选题',
        r'.*(判断|对错|辨析|是非).*': '判断题',
        r'^判断$': '判断题',
        r'.*名词解释.*': '名词解释',
        r'^问答题$': '简答题',
        r'.*简答.*': '简答题',
        r'.*论述.*': '论述题',
        r'.*(材料分析|资料分析|案例分析).*': '材料分析题',
        r'.*填空.*': '填空题'
        
    }
    for pattern, target in mapping_rules.items():
        if re.match(pattern, cleaned) and target in standard_types:
            return target
    if cleaned in standard_types: return cleaned
    return standard_types[0]

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
            with pdfplumber.open(filepath) as pdf:
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text: text += page_text + "\n"
            cleaned_check = re.sub(r'[\s\w\u4e00-\u9fa5，。？！、：；“”‘’（）【】]', '', text)
            if len(text) > 0 and (len(cleaned_check) / len(text)) > 0.3:
                return "[乱码拦截] 该 PDF 文件存在内嵌字体编码错位或属于图片扫描件，请先转换为 Word 格式后导入解析。"
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
        prompt = f"""你是一位的高校教授。请无遗漏地提取出物理文本中所有的考试题目。
【强制指令】：
1. 提取所有考题，题目主体必须保持绝对完整。
2. 💡【新规：严防混淆判断题与单选题】：
   - 如果题目属于判断对错题（例如题干带有“( )”、“[ ]”、或明确要求判断是非对错，且没有提供 A.B.C.D 选项），你必须将 type 统一标记为 "判断题"。绝对不可将其误判为单选题！
   - 如果题目本身是选择题，你必须把题干和所有的选项（如A. B. C. D.）整体打包提取合并在 question 字段中，绝对不可遗漏选项！
3. 无答案统一填“无”。
4. 推断该题所属的【章标题】。⭐【铁律】：你只能从以下现有章节列表中选择归属：{ch_str}。如果列表为空，或题目不属于列表中任何一章，必须一律填入 "未分类章节"，绝对不可自己发明新章节名！
5. 提取该题的【核心知识点】。
6. source固定填"真题"."""
    else:
        prompt = f"""你是一位教研专家。请分析教材或讲义文本。
【强制指令】：
1. 推断文本所属的【章标题】。⭐【铁律】：你只能从以下现有章节列表中选择归属：{ch_str}。如果列表为空，或文本不属于列表中任何一章，必须一律填入 "未分类章节"，绝对不可自己发明新章节名！
2. 提取核心【知识点】，并针对每个知识点原创生成2-3道考题（包含单选、多选或判断对错题）。
3. source固定填"AI生成"。自带练习题提取标记"真题"."""

    json_format = """必须输出合法的JSON，严格遵循以下一维数组结构：
{"questions": [{"chapter_name": "现有章名 或 未分类章节", "point_name": "核心知识点名称", "type": "题型(单选题/多选题/判断题/简答题等)", "question": "完整题干(如果是选择题必须包含ABCD选项预览，如果是判断题必须保留括号)", "answer": "答案(如对/错、正确/错误、A/B/C/D或无)", "source": "真题 或 AI生成"}]}"""
    
    # ... 下方保持你原有的 requests.post 逻辑完全不变 ...    
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
        self.root.title("智能期末试卷管理系统 - 教学研究专业版")
        self.root.geometry("1450x950") 
        init_db()
        self.config = load_config()
        self.type_rows = [] 
        self.build_ui()
        self.refresh_courses()
        self.init_context_menus()

    # ====================================================================
    # 🖥️ GUI 主控面板顶层架构（已重构：注入教研级说明书按钮与全扁平美化）
    # ====================================================================
    def build_ui(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        # --- 顶部全局总控管理栏 ---
        top_frame = tk.Frame(self.root, bg="#2c3e50", height=50)
        top_frame.pack(side=tk.TOP, fill=tk.X)
        
        tk.Label(top_frame, text="📝 智能期末试卷管理系统 (顶层设计智能化专业版)", 
                 fg="white", bg="#2c3e50", font=("Microsoft YaHei", 14, "bold")).pack(side=tk.LEFT, padx=20, pady=10)
        
        # 🚀 右侧控制点 1：系统底层设置按钮（扁平化橙色）
        tk.Button(top_frame, text="⚙️ 系统底层设置", font=("Microsoft YaHei", 10, "bold"), 
                  bg="#e67e22", fg="white", activebackground="#d35400", activeforeground="white",
                  relief="flat", cursor="hand2", bd=0, padx=15, pady=4,
                  command=self.open_settings).pack(side=tk.RIGHT, padx=20, pady=10)
                  
        # 🚀 右侧控制点 2：新注入的【使用说明】按钮（扁平化紫色，置于设置左侧）
        tk.Button(top_frame, text="📖 系统使用说明书", font=("Microsoft YaHei", 10, "bold"), 
                  bg="#9b59b6", fg="white", activebackground="#8e44ad", activeforeground="white",
                  relief="flat", cursor="hand2", bd=0, padx=15, pady=4,
                  command=self.open_help_readme).pack(side=tk.RIGHT, padx=5, pady=10)

        # --- 底部全局实时日志控制台 ---
        self.txt_log = tk.Text(self.root, height=7, bg="#1c2833", fg="#00ff00", font=("Consolas", 10, "bold"))
        self.txt_log.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=5)

        # --- 核心多轨选项卡中枢 ---
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)

        # 挂载各业务舱卡片
        self.tab_import = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_import, text="📥 AI双轨抽取入库")
        self.build_import_tab(self.tab_import)
        
        self.tab_kb = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_kb, text="📚 本地全栈知识库管理")
        self.build_kb_tab(self.tab_kb)

        self.tab_manage = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_manage, text="🛠️ 题库结构与权重总控")
        self.build_manage_tab(self.tab_manage)

        self.tab_browse = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_browse, text="🔎 库中试题全景浏览")
        self.build_browse_tab(self.tab_browse)

        self.tab_generate = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_generate, text="🖨️ 分层结构化组卷")
        self.build_generate_tab(self.tab_generate)

    # ==========================================
    # 🖱️ 鼠标右键菜单分层注入引擎
    # ==========================================
    def init_context_menus(self):
        """为不同的交互界面定制专属的右键快捷菜单"""
        # 兼容系统判定：Windows/Linux 使用 <Button-3>，macOS 使用 <Button-2> 或 <Control-Button-1>
        right_click_event = "<Button-2>" if self.root._windowingsystem == "aqua" else "<Button-3>"

        # ---- 1. 知识库左侧树的右键菜单 ----
        self.kb_tree_menu = tk.Menu(self.root, tearoff=0)
        self.tree_kb.bind(right_click_event, self.popup_kb_tree_menu)

        # ---- 2. 试卷编辑器考题大表的右键菜单 ----
        self.paper_grid_menu = tk.Menu(self.root, tearoff=0)
        self.paper_grid_menu.add_command(label="💾 永久保存选中修改 (单选修改/多选归一)", command=self.action_paper_save)
        self.paper_grid_menu.add_command(label="⭐ 批量覆写考点重要度星级", command=self.action_paper_batch_stars)
        self.paper_grid_menu.add_separator()
        self.paper_grid_menu.add_command(label="↩️ 从当前试卷中移除选中题 (不删题库)", command=self.action_paper_remove)
        self.paper_grid_menu.add_command(label="🗑️ 物理永久从全库清空选中题", command=self.action_paper_delete, foreground="red")
        self.grid_paper_q.bind(right_click_event, self.popup_paper_grid_menu)

        # ---- 3. 全景试题浏览大表的右键菜单 ----
        self.browse_grid_menu = tk.Menu(self.root, tearoff=0)
        self.browse_grid_menu.add_command(label="💾 永久保存选中修改 (单选覆盖/多选归一)", command=self.action_browse_save_batch_or_single)
        self.browse_grid_menu.add_command(label="⭐ 批量覆写考点重要度星级", command=self.action_browse_batch_stars)
        self.browse_grid_menu.add_command(label="🔀 跨章转移与同义词深度合并", command=self.action_browse_transfer_and_merge_dialog)
        self.browse_grid_menu.add_separator()
        self.browse_grid_menu.add_command(label="🗑️ 物理永久从全库清空选中题", command=self.action_browse_delete_batch, foreground="red")
        self.grid_browse_q.bind(right_click_event, self.popup_browse_grid_menu)

        # ---- 4. 🛠️ 题库结构总控 - 章节表右键菜单 ----
        self.manage_ch_menu = tk.Menu(self.root, tearoff=0)
        self.manage_ch_menu.add_command(label="💾 修改选中章名/权重 (基于输入框)", command=self.save_or_add_chapter)
        self.manage_ch_menu.add_command(label="➡️ 归并整章内容至其它章节", command=self.m_move_ch_dialog)
        self.manage_ch_menu.add_separator()
        self.manage_ch_menu.add_command(label="🗑️ 物理彻底删除整章及旗下试题", command=self.m_del_ch, foreground="red")
        self.tree_m_ch.bind(right_click_event, self.popup_manage_ch_menu)

        # ---- 5. 🛠️ 题库结构总控 - 知识点表右键菜单 ----
        self.manage_kp_menu = tk.Menu(self.root, tearoff=0)
        self.manage_kp_menu.add_command(label="✏️ 核心考点重命名", command=self.m_edit_kp)
        self.manage_kp_menu.add_command(label="🔗 聚合合并多个知识点 (多选生效)", command=self.m_merge_kp)
        self.manage_kp_menu.add_command(label="➡️ 跨章迁移此考点 (支持多选)", command=self.m_move_kp_dialog)
        self.manage_kp_menu.add_separator()
        self.manage_kp_menu.add_command(label="⭐ 批量重设考点重要度星级", command=self.change_kp_stars_action)
        self.manage_kp_menu.add_command(label="🗑️ 物理删除选中点下所有试题", command=self.m_del_kp, foreground="red")
        self.tree_m_kp.bind(right_click_event, self.popup_manage_kp_menu)

        # ---- 6. 🛠️ 题库结构总控 - 试题清单表右键菜单 ----
        self.manage_q_menu = tk.Menu(self.root, tearoff=0)
        self.manage_q_menu.add_command(label="🤖 启动 RAG 云端文库智能作答", command=self.generate_ai_answer_action)
        self.manage_q_menu.add_separator()
        self.manage_q_menu.add_command(label="🗑️ 物理彻底删除选中考题 (支持多选)", command=self.delete_selected_questions_from_list, foreground="red")
        self.tree_m_q.bind(right_click_event, self.popup_manage_q_menu)

        # ---- 7. AI 抽取监测台专属右键菜单 ----
        self.extract_monitor_menu = tk.Menu(self.root, tearoff=0)
        self.extract_monitor_menu.add_command(label="🗑️ 抽取不理想，立即将此题从库中物理抹除", command=self.action_monitor_delete)
        
        # 绑定右键事件
        right_click_event = "<Button-2>" if self.root._windowingsystem == "aqua" else "<Button-3>"
        self.grid_extract_success.bind(right_click_event, self.popup_extract_monitor_menu)

        # ---- 8. 🛠️ 题库结构总控 - 底置输入文本框专属右键菜单 ----
        self.text_editor_menu = tk.Menu(self.root, tearoff=0)
        self.text_editor_menu.add_command(label="💾 题目校准完成，立即保存落盘", command=self.save_edit)
        
        # 绑定到题目和答案两个物理文本框
        right_click_event = "<Button-2>" if self.root._windowingsystem == "aqua" else "<Button-3>"
        self.txt_q.bind(right_click_event, self.popup_text_editor_menu)
        self.txt_a.bind(right_click_event, self.popup_text_editor_menu)

    # ---- 右键菜单弹出动态触发器 ----
    def popup_kb_tree_menu(self, event):
        """知识库大树右键智能判断：根据选中节点类型动态生成菜单"""
        # 自动选中鼠标右键悬停的节点
        iid = self.tree_kb.identify_row(event.y)
        if not iid: return
        self.tree_kb.selection_set(iid)
        self.on_kb_select(None) # 触发左侧树与右侧编辑器的联动

        item = self.tree_kb.item(iid)
        tags = item.get("tags", [])

        # 清空旧菜单项，防止选项重叠错乱
        self.kb_tree_menu.delete(0, tk.END)

        if 'paper' in tags:
            self.kb_tree_menu.add_command(label="📤 使用并输出 Word 文档 (标记年份)", command=self.use_and_export_paper)
            self.kb_tree_menu.add_command(label="↩️ 撤销该试卷题目年份标记", command=self.revoke_paper_usage)
            self.kb_tree_menu.add_separator()
            # 🚀 右键菜单级联注入：在此处加入红色的毁灭试卷按钮
            self.kb_tree_menu.add_command(label="🗑️ 物理彻底毁灭这套生成的试卷", command=self.delete_kb_doc, foreground="red")
            self.kb_tree_menu.post(event.x_root, event.y_root)
        elif 'doc' in tags:
            self.kb_tree_menu.add_command(label="🗑️ 毁灭此源文档", command=self.delete_kb_doc, foreground="red")
            self.kb_tree_menu.post(event.x_root, event.y_root)

    def popup_paper_grid_menu(self, event):
        """试卷编辑器考题表右键：定位并弹出"""
        iid = self.grid_paper_q.identify_row(event.y)
        if not iid: return
        # 如果当前右键的行不在已选中的多选列表里，则将选中项切换为当前单行
        current_sels = self.grid_paper_q.selection()
        if iid not in current_sels:
            self.grid_paper_q.selection_set(iid)
        self.on_paper_grid_row_selected(None) # 联动载入控制台
        self.paper_grid_menu.post(event.x_root, event.y_root)

    def popup_browse_grid_menu(self, event):
        """全景试题全景表右键：定位并弹出"""
        iid = self.grid_browse_q.identify_row(event.y)
        if not iid: return
        current_sels = self.grid_browse_q.selection()
        if iid not in current_sels:
            self.grid_browse_q.selection_set(iid)
        self.on_browse_grid_row_selected(None) # 联动载入控制台
        self.browse_grid_menu.post(event.x_root, event.y_root)

    def popup_manage_ch_menu(self, event):
        """结构总控 - 章节表右键：定位、聚焦并弹出"""
        iid = self.tree_m_ch.identify_row(event.y)
        if not iid: return
        self.tree_m_ch.selection_set(iid)
        self.on_m_ch_select(None) # 联动载入考点和输入框数值
        self.manage_ch_menu.post(event.x_root, event.y_root)

    def popup_manage_kp_menu(self, event):
        """结构总控 - 知识点表右键：兼容单选与多选智能识别"""
        iid = self.tree_m_kp.identify_row(event.y)
        if not iid: return
        
        current_sels = self.tree_m_kp.selection()
        # 如果当前右键行不在已选择的集合中，强制切换为单选当前行
        if iid not in current_sels:
            self.tree_m_kp.selection_set(iid)
            
        self.on_m_kp_select(None) # 联动载入右侧题目清单
        self.manage_kp_menu.post(event.x_root, event.y_root)

    def popup_manage_q_menu(self, event):
        """结构总控 - 试题清单表右键：聚焦联动底层人工覆写台"""
        iid = self.tree_m_q.identify_row(event.y)
        if not iid: return
        
        current_sels = self.tree_m_q.selection()
        if iid not in current_sels:
            self.tree_m_q.selection_set(iid)
            
        self.on_m_q_select(None) # 联动将当前题加载进最下方的编辑区
        self.manage_q_menu.post(event.x_root, event.y_root)    

    def popup_text_editor_menu(self, event):
        """底置人工覆写台文本框右键菜单：提供当场秒级落盘能力"""
        self.text_editor_menu.post(event.x_root, event.y_root)    

    def log(self, msg):
        self.root.after(0, lambda: self.txt_log.insert(tk.END, msg + "\n"))
        self.root.after(0, lambda: self.txt_log.see(tk.END))

    def on_tab_changed(self, event):
        selected_tab = self.notebook.tab(self.notebook.select(), "text")
        if "题库结构与权重总控" in selected_tab: 
            self.load_manage_courses()
        elif "分层结构化" in selected_tab: 
            self.refresh_courses()
            self.on_gen_course_changed() 
        elif "本地全栈知识库管理" in selected_tab: 
            self.load_kb_data()
        elif "库中试题全景浏览" in selected_tab:
            self.refresh_browse_tab()

    def refresh_courses(self):
        try:
            conn = sqlite3.connect(DB_FILE); c = conn.cursor()
            c.execute("SELECT course_name FROM courses ORDER BY course_name")
            all_courses = [r[0] for r in c.fetchall()]
            conn.close()
            for combo in ['cb_import_course', 'cb_gen_course', 'cb_kb_course', 'cb_manage_course', 'cb_browse_course']:
                if hasattr(self, combo):
                    getattr(self, combo)['values'] = all_courses
                    if all_courses and not getattr(self, combo).get(): getattr(self, combo).current(0)
            self.load_manage_courses(); self.load_kb_data()
        except: pass

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
        txt_chapters.insert(tk.END, "第一章 导论\n第二章 核心理论\n第三章 实证分析\n")

        def save_new_course():
            c_name = ent_course.get().strip()
            if not c_name:
                messagebox.showwarning("警告", "课程名称不能为空！", parent=win)
                return

            ch_text = txt_chapters.get(1.0, tk.END).strip()
            chapters = [line.strip() for line in ch_text.split('\n') if line.strip()]

            try:
                conn = sqlite3.connect(DB_FILE); c = conn.cursor()
                c.execute("INSERT OR IGNORE INTO courses (course_name) VALUES (?)", (c_name,))
                for ch in chapters:
                    c.execute("INSERT OR IGNORE INTO course_chapters (course, chapter_name, weight) VALUES (?, ?, 0)", (c_name, ch))
                conn.commit(); conn.close()
                self.log(f"✅ 成功创建课程【{c_name}】，并预设了 {len(chapters)} 个大纲章节。")
                self.refresh_courses()
                if hasattr(self, 'cb_import_course'): self.cb_import_course.set(c_name)
                win.destroy()
            except Exception as e: messagebox.showerror("错误", f"保存失败: {e}", parent=win)

        tk.Button(win, text="🚀 确定创建", font=("Arial", 11, "bold"), bg="#27ae60", fg="white", command=save_new_course).pack(pady=15)

    # ================= 业务流 1：资料双轨抽取 =================
    # ==================================================
    # 📥 业务流 1：资料双轨抽取（已重构：无缝融入实时监测舱）
    # ==================================================
    def build_import_tab(self, parent):
        # ---- 1. 原有的：课程选择与建课大纲预设舱 (保持不变) ----
        f_in = tk.Frame(parent); f_in.pack(fill=tk.X, pady=10)
        tk.Label(f_in, text="目标课程:").pack(side=tk.LEFT, padx=10)
        self.cb_import_course = ttk.Combobox(f_in, state="readonly", width=18); self.cb_import_course.pack(side=tk.LEFT)
        tk.Button(f_in, text="+建课与大纲预设舱", font=("Arial", 9, "bold"), fg="#8e44ad", command=self.add_course).pack(side=tk.LEFT, padx=5)
        
        self.lbl_auto_hint = tk.Label(f_in, text="💡 智能识别盾：系统已激活文件名与前文本矩阵校验，自动模糊清洗非标准题型代称。", fg="#2980b9", font=("Arial", 10, "italic"))
        self.lbl_auto_hint.pack(side=tk.LEFT, padx=30)

        # ---- 2. 原有的：文件选择区域 (保持不变) ----
        f_file = tk.Frame(parent); f_file.pack(fill=tk.X, pady=10)
        tk.Button(f_file, text="📁 选择本地资料(支持批量多选)", command=self.select_files).pack(side=tk.LEFT, padx=10)
        self.lbl_files = tk.Label(f_file, text="未选择", fg="blue"); self.lbl_files.pack(side=tk.LEFT)
        self.selected_files = []

        # ---- 3. 原有的：启动按钮 (注意：将其包装进一个 Frame 中，不再单独撑满全屏，为下方空出位置) ----
        f_btn_run = tk.Frame(parent)
        f_btn_run.pack(fill=tk.X, pady=5)
        tk.Button(f_btn_run, text="🚀 启动 AI 扫描提取", font=("Arial", 11, "bold"), height=2, bg="#2980b9", fg="white", command=self.start_extraction).pack(fill=tk.X, padx=10)

        # --------------------------------------------------
        # 📊 核心优化：在按钮下方注入“即时抽取结果监控舱”
        # --------------------------------------------------
        lbl_monitor_title = tk.Label(parent, text="📋 本次抽取导入实时监测清单", font=("Arial", 11, "bold"), fg="#2c3e50")
        lbl_monitor_title.pack(anchor=tk.W, padx=15, pady=(15, 2))

        # 使用 PanedWindow 实现上下灵活拉伸的内部监测台
        pane_monitor = ttk.PanedWindow(parent, orient=tk.VERTICAL)
        pane_monitor.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)

        # 【实时表 1】: 最新成功入库的试题
        f_success = ttk.LabelFrame(pane_monitor, text="✅ 成功导入/合规落盘的新增试题 (右键可快速物理抹除)")
        pane_monitor.add(f_success, weight=1)

        cols_success = ("id", "ch", "kp", "type", "stem", "ans", "stars")
        self.grid_extract_success = ttk.Treeview(f_success, columns=cols_success, show="headings")
        self.grid_extract_success.heading("id", text="ID"); self.grid_extract_success.heading("ch", text="章节归属")
        self.grid_extract_success.heading("kp", text="核心考点"); self.grid_extract_success.heading("type", text="识别题型")
        self.grid_extract_success.heading("stem", text="题干预览"); self.grid_extract_success.heading("ans", text="参考答案")
        self.grid_extract_success.heading("stars", text="星级")

        self.grid_extract_success.column("id", width=50, anchor="center")
        self.grid_extract_success.column("ch", width=110, anchor="w")
        self.grid_extract_success.column("kp", width=110, anchor="w")
        self.grid_extract_success.column("type", width=80, anchor="center")
        self.grid_extract_success.column("stem", width=450, anchor="w")
        self.grid_extract_success.column("ans", width=200, anchor="w")
        self.grid_extract_success.column("stars", width=60, anchor="center")

        vsb_s = ttk.Scrollbar(f_success, orient="vertical", command=self.grid_extract_success.yview)
        self.grid_extract_success.configure(yscrollcommand=vsb_s.set)
        vsb_s.pack(side=tk.RIGHT, fill=tk.Y)
        self.grid_extract_success.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

        # 【实时表 2】: 因重复被拦截的试题
        f_dup = ttk.LabelFrame(pane_monitor, text="⚠️ 因查重高度雷同未重复导入的试题 (系统已自动去重合并)")
        pane_monitor.add(f_dup, weight=1)

        cols_dup = ("type", "stem", "reason")
        self.grid_extract_dup = ttk.Treeview(f_dup, columns=cols_dup, show="headings")
        self.grid_extract_dup.heading("type", text="识别题型")
        self.grid_extract_dup.heading("stem", text="查重拦截题干原样")
        self.grid_extract_dup.heading("reason", text="系统拦截状态说明")

        self.grid_extract_dup.column("type", width=90, anchor="center")
        self.grid_extract_dup.column("stem", width=680, anchor="w")
        self.grid_extract_dup.column("reason", width=180, anchor="center")

        vsb_d = ttk.Scrollbar(f_dup, orient="vertical", command=self.grid_extract_dup.yview)
        self.grid_extract_dup.configure(yscrollcommand=vsb_d.set)
        vsb_d.pack(side=tk.RIGHT, fill=tk.Y)
        self.grid_extract_dup.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

    def select_files(self):
        files = filedialog.askopenfilenames(filetypes=[("文档", "*.docx *.pdf *.md *.txt")])
        if files: self.selected_files = list(files); self.lbl_files.config(text=f"已选 {len(files)} 个文件")

    def start_extraction(self):
        course = self.cb_import_course.get().strip()
        if not course or not self.selected_files: return messagebox.showwarning("警告", "请选择课程和文件！")
        threading.Thread(target=self.thread_extract, args=(course,), daemon=True).start()

    # ==========================================================
    # ⚙️ 核心逻辑改写：引入柔性感知切片与即时监测回传的扫描引擎
    # ==========================================================
    def thread_extract(self, course):
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        current_year = int(time.strftime("%Y"))
        
        # 本次批量扫描的“战果监测容器”
        success_imported_ids = []   # 记录成功写入的新增试题自增ID
        duplicate_blocked_list = []  # 记录高度雷同被拦截的试题元数据
        
        c.execute("SELECT chapter_name FROM course_chapters WHERE course=? AND chapter_name!='未分类章节'", (course,))
        existing_chapters = [r[0] for r in c.fetchall()]
        standard_types = get_standard_q_types()
        exam_keywords = re.compile(r'(学年|学期|期末|期中|考试|试卷|A卷|B卷|试题|真题|补考|测验)')

        for filepath in self.selected_files:
            filename = os.path.basename(filepath)
            self.log(f"\n>>> 物理扫描读取：{filename}")
            f_year = int((re.search(r'(20\d{2})', filename) or re.search(r'(20\d{2})', str(current_year))).group(1))
            full_text = extract_text_from_file(filepath)
            
            if full_text.startswith("[乱码拦截]"):
                self.log(f"   ❌ {full_text}")
                self.root.after(0, lambda m=full_text: messagebox.showerror("文件格式限制", m))
                continue
                
            if len(full_text.strip()) < 50: 
                self.log(f"   [警告] 提取字数极少。")
                continue
            
            front_text_sample = full_text[:300]
            if exam_keywords.search(filename) or exam_keywords.search(front_text_sample):
                detected_mode = "试卷"
                detected_doc_type = "往年真题试卷"
                default_stars = 4 
                self.log(f"   [🔬 智能防错安全盾]: 该材料为【{detected_doc_type}】，自动标记考点为高优4星")
            else:
                detected_mode = "讲义"
                detected_doc_type = "教材/讲义大纲"
                default_stars = 3 
                self.log(f"   [🔬 智能防错安全盾]: 该材料为【{detected_doc_type}】，自动标记考点为常规3星")

            # 挂载资料到文库
            c.execute("SELECT id FROM materials WHERE course=? AND raw_text=?", (course, full_text))
            if not c.fetchone():
                try: 
                    c.execute("INSERT INTO materials (course, raw_text, doc_type) VALUES (?, ?, ?)", (course, full_text, detected_doc_type))
                    self.log(f"   [文库同步] 该全文已作为 {detected_doc_type} 类型安全挂载入知识库。")
                except Exception as e: 
                    self.log(f"   [挂载异常] {e}")

            # 段落感知柔性切片策略
            lines = full_text.split('\n')
            chunks = []
            current_chunk = ""
            
            for line in lines:
                if len(current_chunk) + len(line) > 1600:
                    chunks.append(current_chunk)
                    current_chunk = current_chunk[-500:] + "\n" + line if len(current_chunk) >= 500 else line
                else:
                    if current_chunk: current_chunk += "\n" + line
                    else: current_chunk = line
            if current_chunk.strip():
                chunks.append(current_chunk)

            # 轮询调用 AI 并在本地库落盘
            for idx, chunk in enumerate(chunks):
                if len(chunk.strip()) < 50: continue
                self.log(f"-> AI解析模块 {idx+1}/{len(chunks)}...")
                res, raw_text = call_ai_dual_extractor(self.config, chunk, detected_mode, existing_chapters)
                if "error" in res: continue
                questions = res.get("questions", [])
                if not questions: continue
                
                for q in questions:
                    if not q.get('question'): continue
                    norm_type = normalize_question_type(q.get('type', '单选题'), standard_types)
                    
                    # --- 核心改写：在这里进行高频库内去重拦截校验 ---
                    q_text = q['question'].strip()
                    c.execute("SELECT id FROM questions WHERE course=? AND question=?", (course, q_text))
                    exists = c.fetchone()
                    
                    if exists:
                        # 触发雷同：塞入去重拦截监控容器，直接跳过不予重复入库
                        duplicate_blocked_list.append({"type": norm_type, "stem": q_text})
                        continue
                    
                    # --- 没重复则执行物理写入 ---
                    try:
                        c.execute('''INSERT INTO questions (course, chapter_name, kp_name, kp_stars, q_type, question, answer, source, q_year, doc_type) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                                  (course, q.get('chapter_name', '未分类章节'), q.get('point_name', '未定'), default_stars, norm_type, q_text, q.get('answer','无'), q.get('source','真题'), f_year, detected_doc_type))
                        
                        # 核心：捞取刚刚落盘成功的题目自增主键主ID
                        success_imported_ids.append(c.lastrowid)
                    except: 
                        pass
                conn.commit()
        conn.close()
        
        # 权重校调
        sync_chapter_weights(course)
        self.log(f"\n✅ 【{course}】解析入库完毕。新清洗抓取合规试题: {len(success_imported_ids)} 道。拦截高度雷同老题: {len(duplicate_blocked_list)} 道。")
        
        # ----------------------------------------------------------
        # 🚀 核心改写：安全切回主线程，将本次战果实时吐给 GUI 监控台大表
        # ----------------------------------------------------------
        self.root.after(0, lambda: self.refresh_extract_monitor_grids(success_imported_ids, duplicate_blocked_list))
        self.root.after(0, self.refresh_courses)

    def refresh_extract_monitor_grids(self, success_ids, dup_list):
        """实时渲染引擎：将本次扫描抽取出来的数据精准打入 AI 导入界面的大表中"""
        # 1. 干净清空历史监测遗留
        for i in self.grid_extract_success.get_children(): self.grid_extract_success.delete(i)
        for i in self.grid_extract_dup.get_children(): self.grid_extract_dup.delete(i)

        # 2. 批量捞取刚刚写入成功的题目完整信息并展示
        if success_ids:
            try:
                conn = sqlite3.connect(DB_FILE)
                c = conn.cursor()
                placeholders = ",".join("?" for _ in success_ids)
                query = f"""
                    SELECT id, chapter_name, kp_name, q_type, question, answer, kp_stars 
                    FROM questions 
                    WHERE id IN ({placeholders})
                    ORDER BY id DESC
                """
                c.execute(query, success_ids)
                rows = c.fetchall()
                conn.close()

                for row in rows:
                    # 消除字符串中的硬换行，使其在大表中单行预览更干净
                    clean_q = row[4].replace('\n', ' ')
                    clean_a = row[5].replace('\n', ' ')
                    stars_str = "⭐" * int(row[6]) if row[6] else "⭐"
                    
                    self.grid_extract_success.insert("", tk.END, values=(
                        row[0], row[1], row[2], row[3], clean_q, clean_a, stars_str
                    ))
            except Exception as e:
                self.log(f"❌ 渲染导入成功监控表失败: {str(e)}")

        # 3. 将被查重合并、高度雷同未录入的题目塞进下方的“拦截名单表”
        for dup_item in dup_list:
            self.grid_extract_dup.insert("", tk.END, values=(
                dup_item.get("type", "未知"),
                dup_item.get("stem", "").replace('\n', ' '),
                "⚠️ 题库中已存在(高雷同)"
            )) 

    def popup_extract_monitor_menu(self, event):
        iid = self.grid_extract_success.identify_row(event.y)
        if not iid: return
        self.grid_extract_success.selection_set(iid)
        self.extract_monitor_menu.post(event.x_root, event.y_root)

    def action_monitor_delete(self):
        sel = self.grid_extract_success.selection()
        if not sel: return
        vals = self.grid_extract_success.item(sel[0])['values']
        q_id = vals[0] # 拿到数据库自增 ID
        
        if messagebox.askyesno("秒删确认", f"确定在题库中彻底抹除 ID 为 [{q_id}] 的这道刚录入的题吗？"):
            conn = sqlite3.connect(DB_FILE); c = conn.cursor()
            c.execute("DELETE FROM questions WHERE id=?", (q_id,))
            conn.commit(); conn.close()
            # 界面同步移除
            self.grid_extract_success.delete(sel[0])
            self.refresh_courses() # 全局刷新
            self.log(f"🗑️ 已成功对刚录入的试题 ID: {q_id} 执行了物理应急熔断抹除。")

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

        # ---- 文档视图（选中文档时显示）----
        self.f_doc_view = tk.Frame(f_edit)
        self.f_doc_view.pack(fill=tk.BOTH, expand=True)
        self.txt_kb = tk.Text(self.f_doc_view, font=("Arial", 11), wrap=tk.WORD)
        scroll_kb = ttk.Scrollbar(self.f_doc_view, orient="vertical", command=self.txt_kb.yview)
        self.txt_kb.configure(yscrollcommand=scroll_kb.set)
        scroll_kb.pack(side=tk.RIGHT, fill=tk.Y)
        self.txt_kb.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # ---- 试卷编辑器视图（选中试卷时显示）----
        self.f_paper_view = tk.Frame(f_edit)
        paned_paper = ttk.PanedWindow(self.f_paper_view, orient=tk.HORIZONTAL)
        paned_paper.pack(fill=tk.BOTH, expand=True)

        # 左侧：题型分类
        f_left = ttk.LabelFrame(paned_paper, text="📁 试卷题型结构")
        paned_paper.add(f_left, weight=1)
        self.tree_paper_types = ttk.Treeview(f_left, show="tree headings")
        self.tree_paper_types.heading("#0", text="题型结构清单")
        self.tree_paper_types.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        self.tree_paper_types.bind("<<TreeviewSelect>>", self.on_paper_type_selected)
        tk.Button(f_left, text="➕ 添加新题型到试卷", font=("Arial", 9, "bold"), fg="#27ae60", command=self.action_paper_add_type).pack(fill=tk.X, padx=2, pady=(0, 2))

        # 右侧：题目列表 + 编辑区
        f_right = tk.Frame(paned_paper)
        paned_paper.add(f_right, weight=4)
        paned_right = ttk.PanedWindow(f_right, orient=tk.VERTICAL)
        paned_right.pack(fill=tk.BOTH, expand=True)

        # 右上：题目列表
        f_grid = ttk.LabelFrame(paned_right, text="📋 试卷考题清单")
        paned_right.add(f_grid, weight=3)
        columns_def = ("id", "question", "answer", "chapter", "kp_name", "stars", "score")
        self.grid_paper_q = ttk.Treeview(f_grid, columns=columns_def, show="headings")
        self.grid_paper_q.heading("id", text="ID")
        self.grid_paper_q.heading("question", text="题干预览")
        self.grid_paper_q.heading("answer", text="标准答案")
        self.grid_paper_q.heading("chapter", text="所属章标题")
        self.grid_paper_q.heading("kp_name", text="映射考点")
        self.grid_paper_q.heading("stars", text="考点星级")
        self.grid_paper_q.heading("score", text="分值")
        self.grid_paper_q.column("id", width=40, anchor="center")
        self.grid_paper_q.column("question", width=220)
        self.grid_paper_q.column("answer", width=100)
        self.grid_paper_q.column("chapter", width=90)
        self.grid_paper_q.column("kp_name", width=90)
        self.grid_paper_q.column("stars", width=60, anchor="center")
        self.grid_paper_q.column("score", width=40, anchor="center")
        scroll_p_y = ttk.Scrollbar(f_grid, orient="vertical", command=self.grid_paper_q.yview)
        self.grid_paper_q.configure(yscrollcommand=scroll_p_y.set)
        scroll_p_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.grid_paper_q.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        self.grid_paper_q.bind("<<TreeviewSelect>>", self.on_paper_grid_row_selected)

        # 右下：编辑区
        f_actions = ttk.LabelFrame(paned_right, text="🕹️ 选中试题二次开发/校准控制台")
        paned_right.add(f_actions, weight=2)
        f_r1 = tk.Frame(f_actions); f_r1.pack(fill=tk.X, pady=3, padx=5)
        tk.Label(f_r1, text="题目ID:").pack(side=tk.LEFT)
        self.lbl_p_id = tk.Label(f_r1, text="无", fg="purple", font=("Arial", 10, "bold"), width=8)
        self.lbl_p_id.pack(side=tk.LEFT)
        tk.Label(f_r1, text="题型:").pack(side=tk.LEFT)
        self.ent_p_type = tk.Entry(f_r1, width=12); self.ent_p_type.pack(side=tk.LEFT, padx=5)
        tk.Label(f_r1, text=" | 批调考点重要度:").pack(side=tk.LEFT)
        self.cb_p_star = ttk.Combobox(f_r1, values=["1", "2", "3", "4", "5"], state="readonly", width=4)
        self.cb_p_star.pack(side=tk.LEFT, padx=2); self.cb_p_star.set("3")
        tk.Button(f_r1, text="⭐ 覆写星级", font=("Arial", 9), command=self.action_paper_batch_stars).pack(side=tk.LEFT, padx=3)

        f_r2 = tk.Frame(f_actions); f_r2.pack(fill=tk.BOTH, expand=True, pady=3, padx=5)
        f_r2_l = tk.Frame(f_r2); f_r2_l.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=2)
        tk.Label(f_r2_l, text="题目文本编辑区:").pack(anchor=tk.W)
        self.txt_p_q = tk.Text(f_r2_l, font=("Arial", 10), height=4); self.txt_p_q.pack(fill=tk.BOTH, expand=True)
        f_r2_r = tk.Frame(f_r2); f_r2_r.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=2)
        tk.Label(f_r2_r, text="参考答案编辑区:").pack(anchor=tk.W)
        self.txt_p_a = tk.Text(f_r2_r, font=("Arial", 10), height=4); self.txt_p_a.pack(fill=tk.BOTH, expand=True)

        f_r3 = tk.Frame(f_actions); f_r3.pack(fill=tk.X, pady=4, padx=5)
        tk.Button(f_r3, text="💾 永久保存选中修改", font=("Arial", 10, "bold"), fg="#27ae60", command=self.action_paper_save).pack(side=tk.LEFT, padx=10)
        tk.Button(f_r3, text="➕ 向本题型添加题目", font=("Arial", 10, "bold"), fg="#2980b9", command=self.action_paper_add_questions).pack(side=tk.LEFT, padx=10)
        tk.Button(f_r3, text="🗑️ 从试卷中移除选中题", font=("Arial", 10, "bold"), fg="#c0392b", command=self.action_paper_remove).pack(side=tk.LEFT, padx=10)
        tk.Button(f_r3, text="❌ 物理永久清空选中题", font=("Arial", 10, "bold"), fg="#8e44ad", command=self.action_paper_delete).pack(side=tk.RIGHT, padx=10)

        # 底部按钮栏（固定在f_edit底部）
        f_btn = tk.Frame(f_edit); f_btn.pack(fill=tk.X, pady=5)
        tk.Button(f_btn, text="💾 保存编辑", font=("Arial", 10, "bold"), command=self.save_kb_edit).pack(side=tk.LEFT, padx=10)
        self.btn_use_export = tk.Button(f_btn, text="📤 使用与输出Word", font=("Arial", 10, "bold"), fg="#2980b9", command=self.use_and_export_paper)
        self.btn_use_export.pack(side=tk.LEFT, padx=10)
        self.btn_use_export.config(state=tk.DISABLED)
        self.btn_revoke = tk.Button(f_btn, text="↩️ 撤销使用", font=("Arial", 10, "bold"), fg="#e67e22", command=self.revoke_paper_usage)
        self.btn_revoke.pack(side=tk.LEFT, padx=10)
        self.btn_revoke.config(state=tk.DISABLED)
        tk.Button(f_btn, text="🗑️ 毁灭文档", font=("Arial", 10, "bold"), fg="#c0392b", command=self.delete_kb_doc).pack(side=tk.RIGHT, padx=10)

    def import_pure_kb(self):
        course = self.cb_kb_course.get().strip()
        if not course: return messagebox.showwarning("警告", "请先选择归属课程！")
        files = filedialog.askopenfilenames(filetypes=[("文档", "*.docx *.pdf *.md *.txt")])
        if files: threading.Thread(target=self.thread_import_pure_kb, args=(course, files), daemon=True).start()

    def thread_import_pure_kb(self, course, files):
        conn = sqlite3.connect(DB_FILE); c = conn.cursor()
        success = 0
        for fp in files:
            text = extract_text_from_file(fp)
            if text.startswith("[乱码拦截]"): continue
            if len(text.strip()) > 10:
                c.execute("SELECT id FROM materials WHERE course=? AND raw_text=?", (course, text))
                if not c.fetchone(): c.execute("INSERT INTO materials (course, raw_text, doc_type) VALUES (?, ?, ?)", (course, text, '手动挂载资料')); success += 1; self.log(f"📘 [纯净挂载]: {os.path.basename(fp)} 成功。")
        conn.commit(); conn.close(); self.log(f"✅ 知识库新增 {success} 份。"); self.root.after(0, self.load_kb_data)

    def load_kb_data(self):
        for i in self.tree_kb.get_children(): self.tree_kb.delete(i)
        self.lbl_kb_id.config(text="当前未选择"); self.txt_kb.delete(1.0, tk.END)
        self.btn_use_export.config(state=tk.DISABLED)
        try:
            conn = sqlite3.connect(DB_FILE); c = conn.cursor(); c.execute("SELECT course_name FROM courses ORDER BY course_name"); courses = [r[0] for r in c.fetchall()]
            for course in courses:
                c_node = self.tree_kb.insert("", tk.END, text=course, open=True, tags=('course',))
                # 源文件库
                src_node = self.tree_kb.insert(c_node, tk.END, text="📁 源文件库", open=False, tags=('folder',))
                c.execute("SELECT id, doc_type, SUBSTR(raw_text, 1, 20) FROM materials WHERE course=? AND doc_type!='生成的试卷'", (course,))
                for doc in c.fetchall():
                    dt = doc[1] if doc[1] else "讲义大纲"
                    self.tree_kb.insert(src_node, tk.END, text=f"[{dt}] {doc[2].replace(chr(10),'')}...", values=(doc[0],), tags=('doc',))
                # 试卷库
                paper_node = self.tree_kb.insert(c_node, tk.END, text="📝 试卷库", open=False, tags=('folder',))
                c.execute("SELECT id, paper_name, created_at FROM papers WHERE course=? ORDER BY created_at DESC", (course,))
                for p in c.fetchall():
                    self.tree_kb.insert(paper_node, tk.END, text=f"{p[1]} ({p[2]})", values=(p[0],), tags=('paper',))
            conn.close()
        except Exception as e:
            self.log(f"加载知识库错误: {e}")

    def on_kb_select(self, event):
        sel = self.tree_kb.selection()
        if not sel: return
        item = self.tree_kb.item(sel[0])
        tags = item.get("tags", [])
        if 'doc' in tags:
            doc_id = item['values'][0]
            self.btn_use_export.config(state=tk.DISABLED)
            self.btn_revoke.config(state=tk.DISABLED)
            self._show_doc_view()
            conn = sqlite3.connect(DB_FILE); c = conn.cursor(); c.execute("SELECT raw_text FROM materials WHERE id=?", (doc_id,)); row = c.fetchone(); conn.close()
            if row: self.lbl_kb_id.config(text=str(doc_id)); self.txt_kb.delete(1.0, tk.END); self.txt_kb.insert(tk.END, row[0])
        elif 'paper' in tags:
            paper_id = item['values'][0]
            self.btn_use_export.config(state=tk.NORMAL)
            self.btn_revoke.config(state=tk.NORMAL)
            self._show_paper_view(paper_id)
        else:
            self.btn_use_export.config(state=tk.DISABLED)
            self.btn_revoke.config(state=tk.DISABLED)
            self._show_doc_view()

    def save_kb_edit(self):
        doc_id = self.lbl_kb_id.cget("text"); new_text = self.txt_kb.get(1.0, tk.END).strip()
        if doc_id == "当前未选择": return
        try: conn = sqlite3.connect(DB_FILE); c = conn.cursor(); c.execute("UPDATE materials SET raw_text=? WHERE id=?", (new_text, doc_id)); conn.commit(); conn.close(); messagebox.showinfo("成功", "保存成功！")
        except Exception as e: messagebox.showerror("失败", str(e))

    # ==========================================================
    # 🗑️ 核心修复：具备智能身份判别与多表级联销毁的“毁灭文档”引擎
    # ==========================================================
    def delete_kb_doc(self):
        sel = self.tree_kb.selection()
        if not sel: 
            return messagebox.showwarning("提示", "当前未在大树中选中任何要毁灭的目标！")
            
        item = self.tree_kb.item(sel[0])
        tags = item.get("tags", [])
        
        # --------------------------------------------------
        # 🔗 分流机制 A：如果选中的是【试卷库】中的试卷节点
        # --------------------------------------------------
        if 'paper' in tags:
            paper_id = item['values'][0]
            paper_name = item['text']
            
            if messagebox.askyesno("教研级物理毁灭确认", f"⚠️ 警告：您正在物理毁灭期末试卷【{paper_name}】！\n\n该操作将同时清空该试卷的题目组装架构与生成的答案文档，是否彻底销毁？"):
                try:
                    conn = sqlite3.connect(DB_FILE)
                    c = conn.cursor()
                    
                    # 1. 开启本地事务，先根据试卷名称反查并物理删除 materials 文库中生成的答案试卷备份
                    # 截取纯粹的试卷名（去掉大树节点后面自动带的日期尾巴）
                    clean_paper_name = paper_name.split(" (")[0] if " (" in paper_name else paper_name
                    c.execute("DELETE FROM materials WHERE doc_type='生成的试卷' AND raw_text LIKE ?", (f"%{clean_paper_name}%",))
                    
                    # 2. 级联切断并删除 paper_questions 表里的考题绑定关系链条
                    c.execute("DELETE FROM paper_questions WHERE paper_id=?", (paper_id,))
                    
                    # 3. 彻底抹除 papers 表中的试卷核心框架记录
                    c.execute("DELETE FROM papers WHERE id=?", (paper_id,))
                    
                    conn.commit()
                    conn.close()
                    
                    self.log(f"🗑️ [试卷全生命周期销毁]: 成功物理毁灭试卷框架 ID: {paper_id} 及其旗下的所有关联细目。")
                    messagebox.showinfo("毁灭成功", f"期末试卷【{clean_paper_name}】已成功从全库中物理毁灭！")
                    
                    # 刷新树状图并归位视图
                    self.load_kb_data()
                except Exception as e:
                    messagebox.showerror("毁灭失败", f"销毁试卷时发生数据库冲突: {str(e)}")
            return

        # --------------------------------------------------
        # 📘 分流机制 B：如果选中的是【源文件库】中的教材讲义文档
        # --------------------------------------------------
        doc_id = self.lbl_kb_id.cget("text")
        if doc_id == "当前未选择" or "试卷" in doc_id: 
            return messagebox.showwarning("提示", "当前选中的目标无法通过此通道毁灭，请确保您选中了具体的源文档或试卷节点。")
            
        if messagebox.askyesno("物理毁灭确认", "确定要在全库中物理删除这份核心源资料吗？\n(此操作不可逆，将导致与之相关的AI RAG检索失效)"):
            try:
                conn = sqlite3.connect(DB_FILE)
                c = conn.cursor()
                c.execute("DELETE FROM materials WHERE id=?", (doc_id,))
                conn.commit()
                conn.close()
                
                self.log(f"🗑️ [源文档抹除]: 成功物理毁灭文库核心资料，库ID: {doc_id}")
                messagebox.showinfo("成功", "该核心源资料已被物理毁灭。")
                
                # 刷新并回归初始状态
                self.load_kb_data()
            except Exception as e:
                messagebox.showerror("失败", f"物理删除文档失败: {str(e)}")

    # ================= 🛠️ 业务流 3:: 题库结构与权重总控 =================
    # ====================================================================
    # 🛠️ 业务流 3：题库结构与权重总控（完整美化、排版重组与防呆拦截集成版）
    # ====================================================================
    def build_manage_tab(self, parent):
        f_top = tk.Frame(parent)
        f_top.pack(fill=tk.X, pady=5)
        tk.Label(f_top, text="全景课程选择:", font=("Microsoft YaHei", 10, "bold")).pack(side=tk.LEFT, padx=10)
        
        self.cb_manage_course = ttk.Combobox(f_top, state="readonly", width=20, font=("Microsoft YaHei", 10))
        self.cb_manage_course.pack(side=tk.LEFT)
        self.cb_manage_course.bind("<<ComboboxSelected>>", lambda e: self.load_manage_courses())
        
        tk.Button(f_top, text="🔄 强制重载刷新", font=("Microsoft YaHei", 10), 
                  bg="#7f8c8d", fg="white", activebackground="#95a5a6", activeforeground="white",
                  relief="flat", cursor="hand2", bd=0, padx=12, pady=3).pack(side=tk.LEFT, padx=20)

        main_v_paned = ttk.PanedWindow(parent, orient=tk.VERTICAL)
        main_v_paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        top_h_paned = ttk.PanedWindow(main_v_paned, orient=tk.HORIZONTAL)
        main_v_paned.add(top_h_paned, weight=5)

        # ---- 第一列：章节分值中枢 ----
        f_col1_wrap = tk.Frame(top_h_paned)
        top_h_paned.add(f_col1_wrap, weight=2)
        f_col1 = ttk.LabelFrame(f_col1_wrap, text="1. 章节分值占比 (点击看考点)")
        f_col1.pack(fill=tk.BOTH, expand=True)
        
        self.tree_m_ch = ttk.Treeview(f_col1, columns=("ch_name", "weight"), show="headings")
        self.tree_m_ch.heading("ch_name", text="章名称预览"); self.tree_m_ch.heading("weight", text="分值占比(权重%)")
        self.tree_m_ch.column("ch_name", width=140); self.tree_m_ch.column("weight", width=80, anchor="center")
        self.tree_m_ch.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        self.tree_m_ch.bind("<<TreeviewSelect>>", self.on_m_ch_select)
        
        f_ctrl = ttk.LabelFrame(f_col1_wrap, text="🕹️ 第一列：章节中枢")
        f_ctrl.pack(fill=tk.X, pady=2)
        
        f_c1 = tk.Frame(f_ctrl); f_c1.pack(fill=tk.X, pady=2)
        self.ent_m_ch = tk.Entry(f_c1, width=10, font=("Microsoft YaHei", 10))
        self.ent_m_ch.pack(side=tk.LEFT, padx=2)
        tk.Label(f_c1, text="占比%:", font=("Microsoft YaHei", 9)).pack(side=tk.LEFT)
        self.ent_m_w = tk.Entry(f_c1, width=3, font=("Microsoft YaHei", 10))
        self.ent_m_w.pack(side=tk.LEFT)
        
        tk.Button(f_c1, text="💾 修改/增章", font=("Microsoft YaHei", 10, "bold"), 
                  bg="#27ae60", fg="white", relief="flat", cursor="hand2", bd=0, padx=6, pady=2,
                  command=self.save_or_add_chapter).pack(side=tk.LEFT, padx=4)
        
        f_c2 = tk.Frame(f_ctrl); f_c2.pack(fill=tk.X, pady=2)
        tk.Button(f_c2, text="➡️ 归并整章", font=("Microsoft YaHei", 10, "bold"), 
                  bg="#e67e22", fg="white", relief="flat", cursor="hand2", bd=0, pady=3,
                  command=self.m_move_ch_dialog).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        tk.Button(f_c2, text="🗑️ 删整章", font=("Microsoft YaHei", 10, "bold"), 
                  bg="#c0392b", fg="white", relief="flat", cursor="hand2", bd=0, pady=3,
                  command=self.m_del_ch).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)

        f_c3 = tk.Frame(f_ctrl); f_c3.pack(fill=tk.X, pady=2)
        self.btn_extract_ch = tk.Button(f_c3, text="🤖 1.专门提取大纲", font=("Microsoft YaHei", 10, "bold"), 
                  bg="#2980b9", fg="white", relief="flat", cursor="hand2", bd=0, pady=3,
                  command=self.ai_extract_chapters)
        self.btn_extract_ch.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        self.btn_ai_reclass = tk.Button(f_c3, text="🤖 2.打包装车", font=("Microsoft YaHei", 10, "bold"), 
                  bg="#9b59b6", fg="white", relief="flat", cursor="hand2", bd=0, pady=3,
                  command=self.ai_reclassify)
        self.btn_ai_reclass.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)

        # ---- 第二列：考点星级调控 ----
        f_col2_wrap = tk.Frame(top_h_paned)
        top_h_paned.add(f_col2_wrap, weight=3)
        f_col2 = ttk.LabelFrame(f_col2_wrap, text="2. 知识点与考点星级分级 (可多选)")
        f_col2.pack(fill=tk.BOTH, expand=True)
        
        self.tree_m_kp = ttk.Treeview(f_col2, columns=("kp_name", "stars", "count"), show="headings")
        self.tree_m_kp.heading("kp_name", text="知识点考点"); self.tree_m_kp.heading("stars", text="重要度星级"); self.tree_m_kp.heading("count", text="题量")
        self.tree_m_kp.column("kp_name", width=120); self.tree_m_kp.column("stars", width=80, anchor="center"); self.tree_m_kp.column("count", width=40, anchor="center")
        self.tree_m_kp.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        self.tree_m_kp.bind("<<TreeviewSelect>>", self.on_m_kp_select)
        
        f_kp_ctrl = tk.LabelFrame(f_col2_wrap, text="🕹️ 第二列：考点重要度调控")
        f_kp_ctrl.pack(fill=tk.X, pady=2)
        
        f_star_panel = tk.Frame(f_kp_ctrl)
        f_star_panel.pack(fill=tk.X, pady=2)
        tk.Label(f_star_panel, text="重设考点星级:", font=("Microsoft YaHei", 9, "bold")).pack(side=tk.LEFT, padx=2)
        
        self.cb_star_setter = ttk.Combobox(f_star_panel, values=["1星 (边缘抽考)", "2星 (一般考点)", "3星 (常规重点)", "4星 (真题级重要)", "5星 (极高核心优先)"], state="readonly", width=15, font=("Microsoft YaHei", 9))
        self.cb_star_setter.pack(side=tk.LEFT, padx=2); self.cb_star_setter.current(2)
        
        tk.Button(f_star_panel, text="⭐ 确认修改", font=("Microsoft YaHei", 10, "bold"), 
                  bg="#2980b9", fg="white", relief="flat", cursor="hand2", bd=0, padx=8, pady=1,
                  command=self.change_kp_stars_action).pack(side=tk.LEFT, padx=5)

        f_k1 = tk.Frame(f_kp_ctrl); f_k1.pack(fill=tk.X, pady=2)
        tk.Button(f_k1, text="✏️ 改名", font=("Microsoft YaHei", 10), 
                  bg="#7f8c8d", fg="white", relief="flat", cursor="hand2", bd=0, pady=2,
                  command=self.m_edit_kp).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        tk.Button(f_k1, text="🔗 合并", font=("Microsoft YaHei", 10, "bold"), 
                  bg="#27ae60", fg="white", relief="flat", cursor="hand2", bd=0, pady=2,
                  command=self.m_merge_kp).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        tk.Button(f_k1, text="🗑️ 删点", font=("Microsoft YaHei", 10), 
                  bg="#c0392b", fg="white", relief="flat", cursor="hand2", bd=0, pady=2,
                  command=self.m_del_kp).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        
        tk.Button(f_kp_ctrl, text="➡️ 移至它章 (支持多选)", font=("Microsoft YaHei", 10, "bold"), 
                  bg="#e67e22", fg="white", relief="flat", cursor="hand2", bd=0, pady=4,
                  command=self.m_move_kp_dialog).pack(fill=tk.X, pady=2, padx=2)

        # ---- 第三列：试题清单 ----
        f_col3_wrap = tk.Frame(top_h_paned)
        top_h_paned.add(f_col3_wrap, weight=4)
        f_col3 = ttk.LabelFrame(f_col3_wrap, text="3. 试题清单 (双击载入底置编辑器)")
        f_col3.pack(fill=tk.BOTH, expand=True)
        
        self.tree_m_q = ttk.Treeview(f_col3, columns=("id", "source", "stars", "type", "q"), show="headings")
        self.tree_m_q.heading("id", text="ID"); self.tree_m_q.heading("source", text="来源"); self.tree_m_q.heading("stars", text="星级"); self.tree_m_q.heading("type", text="题型"); self.tree_m_q.heading("q", text="题干预览")
        self.tree_m_q.column("id", width=40, anchor="center"); self.tree_m_q.column("source", width=60, anchor="center"); self.tree_m_q.column("stars", width=60, anchor="center"); self.tree_m_q.column("type", width=60, anchor="center"); self.tree_m_q.column("q", width=260)
        
        scroll_q = ttk.Scrollbar(f_col3, orient="vertical", command=self.tree_m_q.yview)
        self.tree_m_q.configure(yscrollcommand=scroll_q.set); scroll_q.pack(side=tk.RIGHT, fill=tk.Y); self.tree_m_q.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        self.tree_m_q.bind("<<TreeviewSelect>>", self.on_m_q_select)
        
        f_q_list_ctrl = tk.Frame(f_col3_wrap)
        f_q_list_ctrl.pack(fill=tk.X, pady=2)
        tk.Button(f_q_list_ctrl, text="🗑  物理删除选中考题 (支持列表多选)", font=("Microsoft YaHei", 10, "bold"), 
                  bg="#c0392b", fg="white", relief="flat", cursor="hand2", bd=0, pady=4,
                  command=self.delete_selected_questions_from_list).pack(fill=tk.X, padx=2, pady=1)

        # ====================================================================
        # ✏️ 结构重构：精心设计的试题人工覆写控制台（按钮一律上移，绝不漏看漏存）
        # ====================================================================
        f_edit = ttk.LabelFrame(main_v_paned, text="✏️ 试题人工覆写控制台")
        main_v_paned.add(f_edit, weight=3)
        
        # --- 顶栏元数据与核心行动核心整合区 ---
        f_meta = tk.Frame(f_edit)
        f_meta.pack(fill=tk.X, pady=4, padx=5)
        
        tk.Label(f_meta, text="题目ID:", font=("Microsoft YaHei", 10, "bold")).pack(side=tk.LEFT, padx=2)
        self.lbl_id = tk.Label(f_meta, text="未选择", fg="blue", font=("Microsoft YaHei", 10, "bold"), width=6)
        self.lbl_id.pack(side=tk.LEFT)
        
        tk.Label(f_meta, text=" | 题型:", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT, padx=2)
        self.ent_type = tk.Entry(f_meta, width=12, font=("Microsoft YaHei", 10))
        self.ent_type.pack(side=tk.LEFT, padx=2)
        
        # 🚀 黄金交互线：把保存修改、RAG和物理删除按钮整体打包上移，紧贴着输入框！
        tk.Button(f_meta, text="💾 永久保存题目修改", font=("Microsoft YaHei", 10, "bold"), 
                  bg="#27ae60", fg="white", relief="flat", cursor="hand2", bd=0, padx=15, pady=2,
                  command=self.save_edit).pack(side=tk.LEFT, padx=15)
                  
        self.btn_ai = tk.Button(f_meta, text="🤖 RAG 讲义检索智能作答", font=("Microsoft YaHei", 10, "bold"), 
                  bg="#2980b9", fg="white", relief="flat", cursor="hand2", bd=0, padx=10, pady=2,
                  command=self.generate_ai_answer_action)
        self.btn_ai.pack(side=tk.LEFT, padx=5)
        
        tk.Button(f_meta, text="🗑️ 删除此题", font=("Microsoft YaHei", 10, "bold"), 
                  bg="#c0392b", fg="white", relief="flat", cursor="hand2", bd=0, padx=10, pady=2,
                  command=self.delete_q_action).pack(side=tk.RIGHT, padx=5)

        # --- 下栏左右并排大输入框文本区 ---
        f_text = tk.Frame(f_edit)
        f_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=2)
        
        f_t_q = tk.Frame(f_text)
        f_t_q.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0,5))
        tk.Label(f_t_q, text="「题目原文编辑区」", font=("Microsoft YaHei", 9, "italic"), fg="gray").pack(anchor=tk.W)
        self.txt_q = tk.Text(f_t_q, font=("Microsoft YaHei", 10), height=5)
        self.txt_q.pack(fill=tk.BOTH, expand=True)

        f_text_r_container = tk.Frame(f_text)
        f_text_r_container.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5,0))
        tk.Label(f_text_r_container, text="「标准答案编辑区」", font=("Microsoft YaHei", 9, "italic"), fg="gray").pack(anchor=tk.W)
        self.txt_a = tk.Text(f_text_r_container, font=("Microsoft YaHei", 10), height=5)
        self.txt_a.pack(fill=tk.BOTH, expand=True)

    def load_manage_courses(self):
        for i in self.tree_m_ch.get_children(): self.tree_m_ch.delete(i)
        for i in self.tree_m_kp.get_children(): self.tree_m_kp.delete(i)
        for i in self.tree_m_q.get_children(): self.tree_m_q.delete(i)
        self.clear_editor()
        course = self.cb_manage_course.get()
        if not course or course == "":
            vals = self.cb_manage_course['values']
            if vals:
                self.cb_manage_course.current(0)
                course = self.cb_manage_course.get()
            else: return
                
        sync_chapter_weights(course)
        conn = sqlite3.connect(DB_FILE); c = conn.cursor()
        c.execute("SELECT chapter_name, weight FROM course_chapters WHERE course=? ORDER BY id", (course,))
        rows = c.fetchall()
        conn.close()
        for r in rows: self.tree_m_ch.insert("", tk.END, values=(r[0], f"{r[1]}%"))

    def on_m_ch_select(self, event):
        sel = self.tree_m_ch.selection()
        if not sel: return
        ch_name = self.tree_m_ch.item(sel[0])['values'][0]
        self.ent_m_ch.delete(0, tk.END); self.ent_m_ch.insert(0, ch_name)
        self.ent_m_w.delete(0, tk.END); self.ent_m_w.insert(0, self.tree_m_ch.item(sel[0])['values'][1].replace('%', ''))
        course = self.cb_manage_course.get()
        for i in self.tree_m_kp.get_children(): self.tree_m_kp.delete(i)
        for i in self.tree_m_q.get_children(): self.tree_m_q.delete(i)
        self.clear_editor()
        conn = sqlite3.connect(DB_FILE); c = conn.cursor()
        c.execute("SELECT kp_name, ROUND(AVG(kp_stars)), COUNT(*) FROM questions WHERE course=? AND chapter_name=? GROUP BY kp_name", (course, ch_name))
        for r in c.fetchall(): 
            star_str = "⭐" * int(r[1]) if r[1] else "⭐"
            self.tree_m_kp.insert("", tk.END, values=(r[0], star_str, r[2]))
        conn.close()

    def on_m_kp_select(self, event):
        sel_kps = self.tree_m_kp.selection()
        sel_ch = self.tree_m_ch.selection()
        if not sel_kps or not sel_ch: return
        kp_name = self.tree_m_kp.item(sel_kps[0])['values'][0] 
        ch_name = self.tree_m_ch.item(sel_ch[0])['values'][0]
        course = self.cb_manage_course.get()
        for i in self.tree_m_q.get_children(): self.tree_m_q.delete(i)
        self.clear_editor()
        conn = sqlite3.connect(DB_FILE); c = conn.cursor()
        c.execute("SELECT id, source, kp_stars, q_type, question FROM questions WHERE course=? AND chapter_name=? AND kp_name=?", (course, ch_name, kp_name))
        for r in c.fetchall():
            source = r[1] if r[1] else "未知"
            star_indicator = "⭐" * int(r[2]) if r[2] else "⭐"
            preview = r[4].replace('\n', ' ')[:40] + "..." if len(r[4]) > 40 else r[4].replace('\n', ' ')
            self.tree_m_q.insert("", tk.END, values=(r[0], source, star_indicator, r[3], preview))
        conn.close()

    # ==========================================================
    # 🔎 核心修复：具备“未保存主动拦截盾”的试题清单切换引擎
    # ==========================================================
    def on_m_q_select(self, event):
        sel = self.tree_m_q.selection()
        if not sel: return
        
        # 🛡️ 【拦截核心】：检查覆写台当前是否有正在编辑且未保存的文本
        current_loaded_id = self.lbl_id.cget("text")
        if current_loaded_id != "未选择" and current_loaded_id != "":
            try:
                # 物理捞取数据库中的原样文本
                conn = sqlite3.connect(DB_FILE)
                c = conn.cursor()
                c.execute("SELECT q_type, question, answer FROM questions WHERE id=?", (current_loaded_id,))
                db_row = c.fetchone()
                conn.close()
                
                if db_row:
                    # 获取当前文本框里老师修改过的最新文字
                    current_ui_q = self.txt_q.get(1.0, tk.END).strip()
                    current_ui_a = self.txt_a.get(1.0, tk.END).strip()
                    current_ui_t = self.ent_type.get().strip()
                    
                    # 比对数据库文本，一旦发现不一致，说明老师改了字但没点保存
                    if current_ui_q != db_row[1].strip() or current_ui_a != db_row[2].strip() or current_ui_t != db_row[0].strip():
                        # 强行弹窗拦截切换事件，把选择权交还给老师
                        if messagebox.askyesno("⚠️ 检测到未保存的修改", 
                                               f"您对当前题目 [ID: {current_loaded_id}] 进行了人工校准，但尚未点击保存！\n\n"
                                               f"如果直接切换题目，您刚才修改的内容将会全部消失。\n\n"
                                               f"是否立即为您自动执行安全落盘保存？"):
                            self.save_edit() # 老师选“是”：自动帮老师存盘
                        else:
                            # 老师选“否”：恢复大表原本的高亮聚焦，中断此次切换
                            self.tree_m_q.selection_remove(sel[0])
                            for item in self.tree_m_q.get_children():
                                if str(self.tree_m_q.item(item)['values'][0]) == str(current_loaded_id):
                                    self.tree_m_q.selection_set(item)
                                    break
                            return # 彻底斩断覆盖流
            except Exception as e:
                self.log(f"⚠️ 查重安全盾自检时发生异常: {str(e)}")

        # --- 通过检查后，才允许安全加载新题目的文本到覆写台 ---
        q_id = self.tree_m_q.item(sel[0])['values'][0]
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute("SELECT q_type, question, answer FROM questions WHERE id=?", (q_id,))
        row = c.fetchone()
        conn.close()
        
        if row:
            self.lbl_id.config(text=str(q_id))
            self.ent_type.delete(0, tk.END)
            self.ent_type.insert(0, row[0])
            self.txt_q.delete(1.0, tk.END)
            self.txt_q.insert(tk.END, row[1])
            self.txt_a.delete(1.0, tk.END)
            self.txt_a.insert(tk.END, row[2])

    def change_kp_stars_action(self):
        sel_kps = self.tree_m_kp.selection()
        sel_ch = self.tree_m_ch.selection()
        if not sel_kps or not sel_ch: return messagebox.showwarning("提示", "请选择要手动调控星级的知识点考点！")
        course = self.cb_manage_course.get(); ch_name = self.tree_m_ch.item(sel_ch[0])['values'][0]
        kp_names = [self.tree_m_kp.item(item)['values'][0] for item in sel_kps]
        chosen_star_text = self.cb_star_setter.get(); target_stars = int(chosen_star_text[0])
        try:
            conn = sqlite3.connect(DB_FILE); c = conn.cursor()
            for kp in kp_names: c.execute("UPDATE questions SET kp_stars=? WHERE course=? AND chapter_name=? AND kp_name=?", (target_stars, course, ch_name, kp))
            conn.commit(); conn.close(); self.on_m_ch_select(None)
        except Exception as e: messagebox.showerror("错误", str(e))

    def save_or_add_chapter(self):
        course = self.cb_manage_course.get()
        if not course: return
        new_ch = self.ent_m_ch.get().strip()
        try: new_w = float(self.ent_m_w.get() or 0)
        except: return messagebox.showerror("错误", "分值权重必须为数字。")
        if not new_ch: return
        sel = self.tree_m_ch.selection(); conn = sqlite3.connect(DB_FILE); c = conn.cursor()
        if sel: 
            old_ch = self.tree_m_ch.item(sel[0])['values'][0]
            if old_ch != "未分类章节":
                c.execute("UPDATE course_chapters SET chapter_name=?, weight=? WHERE course=? AND chapter_name=?", (new_ch, new_w, course, old_ch))
                c.execute("UPDATE questions SET chapter_name=? WHERE course=? AND chapter_name=?", (new_ch, course, old_ch))
            else: c.execute("INSERT OR REPLACE INTO course_chapters (course, chapter_name, weight) VALUES (?, ?, ?)", (course, new_ch, new_w))
        else: c.execute("INSERT OR REPLACE INTO course_chapters (course, chapter_name, weight) VALUES (?, ?, ?)", (course, new_ch, new_w))
        conn.commit(); conn.close(); self.load_manage_courses()

    def m_del_ch(self):
        sel_ch = self.tree_m_ch.selection()
        if not sel_ch: return messagebox.showwarning("提示", "请选中要删除的章！")
        ch = self.tree_m_ch.item(sel_ch[0])['values'][0]; course = self.cb_manage_course.get()
        if messagebox.askyesno("危险", f"彻底删除章节【{ch}】下的全部知识点和试题？"):
            conn = sqlite3.connect(DB_FILE); c = conn.cursor(); c.execute("DELETE FROM questions WHERE course=? AND chapter_name=?", (course, ch)); c.execute("DELETE FROM course_chapters WHERE course=? AND chapter_name=?", (course, ch)); conn.commit(); conn.close(); self.load_manage_courses()

    def m_move_ch_dialog(self):
        sel_ch = self.tree_m_ch.selection()
        if not sel_ch: return messagebox.showwarning("提示", "请在左侧选中要归并的【章节】！")
        old_ch = self.tree_m_ch.item(sel_ch[0])['values'][0]; course = self.cb_manage_course.get()
        conn = sqlite3.connect(DB_FILE); c = conn.cursor(); c.execute("SELECT chapter_name FROM course_chapters WHERE course=?", (course,)); chapters = [r[0] for r in c.fetchall() if r[0] != old_ch]; conn.close()
        if not chapters: return messagebox.showwarning("提示", "没有其他章节可供合并！")
        win = tk.Toplevel(self.root); win.title("整章合并"); win.geometry("350x200"); win.grab_set()
        tk.Label(win, text=f"将【{old_ch}】的全部内容并入哪一章？", font=("Arial", 11)).pack(pady=15)
        cb = ttk.Combobox(win, values=chapters, state="readonly", font=("Arial", 11), width=20); cb.pack(pady=5); cb.set(chapters[0])
        def do_move():
            new_ch = cb.get()
            if new_ch:
                conn = sqlite3.connect(DB_FILE); cur = conn.cursor(); cur.execute("UPDATE questions SET chapter_name=? WHERE course=? AND chapter_name=?", (new_ch, course, old_ch)); cur.execute("DELETE FROM course_chapters WHERE course=? AND chapter_name=?", (course, old_ch)); conn.commit(); conn.close(); self.load_manage_courses()
            win.destroy()
        tk.Button(win, text="🚀 确定合并", command=do_move, font=("Arial", 11, "bold"), bg="#c0392b", fg="white").pack(pady=15)

    def m_edit_kp(self):
        sel_kp = self.tree_m_kp.selection(); sel_ch = self.tree_m_ch.selection()
        if not sel_kp or not sel_ch: return
        old_kp = self.tree_m_kp.item(sel_kp[0])['values'][0]; ch = self.tree_m_ch.item(sel_ch[0])['values'][0]; course = self.cb_manage_course.get()
        new_kp = simpledialog.askstring("修改名称", "新知识点名称:", initialvalue=old_kp)
        if new_kp and new_kp.strip() and new_kp.strip() != old_kp:
            conn = sqlite3.connect(DB_FILE); c = conn.cursor(); c.execute("UPDATE questions SET kp_name=? WHERE course=? AND chapter_name=? AND kp_name=?", (new_kp.strip(), course, ch, old_kp)); conn.commit(); conn.close(); self.on_m_ch_select(None) 

    def m_del_kp(self):
        sel_kps = self.tree_m_kp.selection(); sel_ch = self.tree_m_ch.selection()
        if not sel_kps or not sel_ch: return
        ch = self.tree_m_ch.item(sel_ch[0])['values'][0]; course = self.cb_manage_course.get(); kp_names = [self.tree_m_kp.item(item)['values'][0] for item in sel_kps]
        if messagebox.askyesno("警告", f"物理删除选中的 {len(kp_names)} 个知识点下所有试题？"):
            conn = sqlite3.connect(DB_FILE); c = conn.cursor()
            for kp in kp_names: c.execute("DELETE FROM questions WHERE course=? AND chapter_name=? AND kp_name=?", (course, ch, kp))
            conn.commit(); conn.close(); self.on_m_ch_select(None)

    def m_move_kp_dialog(self):
        sel_kps = self.tree_m_kp.selection(); sel_ch = self.tree_m_ch.selection()
        if not sel_kps or not sel_ch: return messagebox.showwarning("提示", "请选择要转移的知识点！")
        old_ch = self.tree_m_ch.item(sel_ch[0])['values'][0]; course = self.cb_manage_course.get(); kp_names = [self.tree_m_kp.item(item)['values'][0] for item in sel_kps]
        conn = sqlite3.connect(DB_FILE); c = conn.cursor(); c.execute("SELECT chapter_name FROM course_chapters WHERE course=?", (course,)); chapters = [r[0] for r in c.fetchall()]; conn.close()
        win = tk.Toplevel(self.root); win.title("知识点转移"); win.geometry("350x180"); win.grab_set() 
        tk.Label(win, text=f"将选定的 {len(kp_names)} 个知识点移至：", font=("Arial", 11)).pack(pady=15)
        cb = ttk.Combobox(win, values=chapters, state="readonly", font=("Arial", 11), width=20); cb.pack(pady=5)
        if chapters: cb.set(chapters[0])
        def do_move():
            new_ch = cb.get()
            if new_ch and new_ch != old_ch:
                conn = sqlite3.connect(DB_FILE); cur = conn.cursor()
                for kp in kp_names: cur.execute("UPDATE questions SET chapter_name=? WHERE course=? AND chapter_name=? AND kp_name=?", (new_ch, course, old_ch, kp))
                conn.commit(); conn.close(); self.load_manage_courses()
            win.destroy()
        tk.Button(win, text="🚀 确定转移", command=do_move, font=("Arial", 11, "bold"), bg="#e67e22", fg="white").pack(pady=15)

    def m_merge_kp(self):
        sel_kps = self.tree_m_kp.selection(); sel_ch = self.tree_m_ch.selection()
        if not sel_kps or len(sel_kps) < 2: return messagebox.showwarning("提示", "合并操作至少需要选中 2 个知识点！")
        ch = self.tree_m_ch.item(sel_ch[0])['values'][0]; course = self.cb_manage_course.get(); kp_names = [self.tree_m_kp.item(item)['values'][0] for item in sel_kps]
        win = tk.Toplevel(self.root); win.title("聚合合并知识点"); win.geometry("380x200"); win.grab_set()
        tk.Label(win, text=f"将选定的 {len(kp_names)} 个知识点统一合并为：", font=("Arial", 11)).pack(pady=15)
        cb = ttk.Combobox(win, values=kp_names, font=("Arial", 11), width=25); cb.pack(pady=5); cb.set(kp_names[0]) 
        def do_merge():
            target_kp = cb.get().strip()
            if target_kp:
                conn = sqlite3.connect(DB_FILE); c = conn.cursor(); placeholders = ','.join('?' for _ in kp_names); query = f"UPDATE questions SET kp_name=? WHERE course=? AND chapter_name=? AND kp_name IN ({placeholders})"; params = [target_kp, course, ch] + kp_names; c.execute(query, params); conn.commit(); conn.close(); self.on_m_ch_select(None) 
            win.destroy()
        tk.Button(win, text="🔗 确定聚合", command=do_merge, font=("Arial", 11, "bold"), bg="#27ae60", fg="white").pack(pady=15)

    def ai_extract_chapters(self):
        course = self.cb_manage_course.get()
        if not course: return
        self.btn_extract_ch.config(state=tk.DISABLED, text="⏳ AI提取中...")
        threading.Thread(target=self.thread_ai_extract_chapters, args=(course,), daemon=True).start()

    def thread_ai_extract_chapters(self, course):
        try:
            conn = sqlite3.connect(DB_FILE); c = conn.cursor(); c.execute("SELECT raw_text FROM materials WHERE course=? AND doc_type='教材/讲义大纲'", (course,)); combined_context = "\n\n".join([r[0] for r in c.fetchall()])[:15000]; conn.close()
            if not combined_context: messagebox.showerror("中断", "无原稿！"); return
            res = call_ai_extract_chapters_api(self.config, course, combined_context); chapters = res.get("chapters", [])
            conn = sqlite3.connect(DB_FILE); c = conn.cursor()
            for ch in chapters:
                try: c.execute("INSERT INTO course_chapters (course, chapter_name, weight) VALUES (?, ?, 0)", (course, ch))
                except: pass
            conn.commit(); conn.close(); self.root.after(0, self.load_manage_courses)
        except Exception as e: self.log(f"❌ 错误: {e}")
        finally: self.root.after(0, lambda: self.btn_extract_ch.config(state=tk.NORMAL, text="🤖 1.专门提取大纲"))

    def ai_reclassify(self):
        course = self.cb_manage_course.get()
        if not course: return
        self.btn_ai_reclass.config(state=tk.DISABLED, text="⏳ 重组中...")
        threading.Thread(target=self.thread_ai_reclassify, args=(course,), daemon=True).start()

    def thread_ai_reclassify(self, course):
        try:
            conn = sqlite3.connect(DB_FILE); c = conn.cursor(); c.execute("SELECT chapter_name FROM course_chapters WHERE course=? AND chapter_name!='未分类章节'", (course,)); chapters = [r[0] for r in c.fetchall()]; c.execute("SELECT kp_name, question FROM questions WHERE course=? GROUP BY kp_name", (course,)); all_kps = [{"kp_name": r[0], "question_sample": r[1][:200]} for r in c.fetchall()]; conn.close()
            if not chapters or not all_kps: return
            batch_size = 20; conn = sqlite3.connect(DB_FILE); c = conn.cursor()
            for i in range(0, len(all_kps), batch_size):
                batch = all_kps[i:i+batch_size]; res = call_ai_reclassify_kps(self.config, course, chapters, batch)
                for m in res.get("mapping", []):
                    kp, ch = m.get("kp_name"), m.get("chapter_name")
                    if kp and ch and ch in chapters: c.execute("UPDATE questions SET chapter_name=? WHERE course=? AND kp_name=?", (ch, course, kp))
                conn.commit(); time.sleep(1)
            conn.close(); self.root.after(0, self.load_manage_courses)
        except Exception as e: self.log(f"❌ 失败: {e}")
        finally: self.root.after(0, lambda: self.btn_ai_reclass.config(state=tk.NORMAL, text="🤖 2.打包装车"))

    def delete_selected_questions_from_list(self):
        selections = self.tree_m_q.selection()
        if not selections: return messagebox.showwarning("提示", "请多选或单选题目！")
        if messagebox.askyesno("物理彻底删除确认", f"确定清空选中的 {len(selections)} 道题目吗？"):
            try:
                conn = sqlite3.connect(DB_FILE); c = conn.cursor()
                for item in selections: c.execute("DELETE FROM questions WHERE id=?", (self.tree_m_q.item(item)['values'][0],))
                conn.commit(); conn.close(); self.on_m_kp_select(None) 
            except Exception as e: messagebox.showerror("异常", str(e))

    # ==========================================================
    # 💾 核心修复：具备强制视图重载与状态锁定的“保存题目修改”引擎
    # ==========================================================
    # ==========================================================
    # 💾 核心修复：具备强制视图重载与状态锁定的“保存题目修改”引擎
    # ==========================================================
    def save_edit(self):
        q_id = self.lbl_id.cget("text")
        if q_id == "未选择": 
            return messagebox.showwarning("提示", "当前未在清单中选中任何要修改的题目！")
            
        qt = self.ent_type.get().strip()
        q = self.txt_q.get(1.0, tk.END).strip()
        a = self.txt_a.get(1.0, tk.END).strip()
        
        # 顶层设计：题型规范化清洗，防止老师输入非标准题型破坏组卷大纲
        standard_types = get_standard_q_types()
        norm_type = normalize_question_type(qt, standard_types)
        
        try:
            # 1. 持久化落盘：物理更新本地 SQLite 数据库
            conn = sqlite3.connect(DB_FILE)
            c = conn.cursor()
            c.execute("UPDATE questions SET q_type=?, question=?, answer=? WHERE id=?", (norm_type, q, a, q_id))
            conn.commit()
            conn.close()
            
            self.log(f"✅ [试题人工覆写落盘]: 成功永久保存试题 ID: {q_id} 的最新教研校准文本。")
            messagebox.showinfo("成功", "本地题目人工校准覆写成功，已安全落盘！")
            
            # 2. 核心联动修复：获取当前知识点和章节的选择状态，强制重载试题大表
            sel_kp = self.tree_m_kp.selection()  # 统一使用单数 sel_kp
            sel_ch = self.tree_m_ch.selection()
            
            if sel_kp and sel_ch:                # 🚀 核心修复：这里修改为 sel_kp，彻底根除未定义报错
                # 显式提取当前锚定的考点和章标题，绕过事件漂移漏洞
                kp_name = self.tree_m_kp.item(sel_kp[0])['values'][0] 
                ch_name = self.tree_m_ch.item(sel_ch[0])['values'][0]
                course = self.cb_manage_course.get()
                
                # 清空旧大表
                for i in self.tree_m_q.get_children(): self.tree_m_q.delete(i)
                
                # 重新物理捞取该考点下的最新试题状况进行渲染
                conn = sqlite3.connect(DB_FILE)
                c = conn.cursor()
                c.execute("SELECT id, source, kp_stars, q_type, question FROM questions WHERE course=? AND chapter_name=? AND kp_name=?", (course, ch_name, kp_name))
                for r in c.fetchall():
                    source = r[1] if r[1] else "未知"
                    star_indicator = "⭐" * int(r[2]) if r[2] else "⭐"
                    preview = r[4].replace('\n', ' ')[:40] + "..." if len(r[4]) > 40 else r[4].replace('\n', ' ')
                    self.tree_m_q.insert("", tk.END, values=(r[0], source, star_indicator, r[3], preview))
                conn.close()
                
                # 3. 状态归位：将刚刚修改过的题目在试题大表中重新高亮选中，并同步更新最下方的覆写台
                for item in self.tree_m_q.get_children():
                    if str(self.tree_m_q.item(item)['values'][0]) == str(q_id):
                        self.tree_m_q.selection_set(item)
                        self.tree_m_q.see(item)  # 滚动条自动追踪聚焦
                        break
            else:
                # 保底全景刷新
                self.load_manage_courses()
                
        except Exception as e:
            messagebox.showerror("失败", f"保存题目修改时发生错误: {str(e)}")

    def delete_q_action(self):
        q_id = self.lbl_id.cget("text")
        if q_id == "未选择": return
        if messagebox.askyesno("警告", "物理删除此题？"):
            conn = sqlite3.connect(DB_FILE); c = conn.cursor(); c.execute("DELETE FROM questions WHERE id=?", (q_id,)); conn.commit(); conn.close(); self.on_m_kp_select(None) 

    def generate_ai_answer_action(self):
        q_id = self.lbl_id.cget("text")
        if q_id == "未选择": return
        course = self.cb_manage_course.get(); question = self.txt_q.get(1.0, tk.END).strip()
        self.btn_ai.config(text="⏳ RAG检索中...", state=tk.DISABLED); self.txt_a.delete(1.0, tk.END); self.txt_a.insert(tk.END, "云端库检索请求中...\n")
        threading.Thread(target=self.thread_rag_answer, args=(course, question), daemon=True).start()

    def thread_rag_answer(self, course, question):
        try:
            conn = sqlite3.connect(DB_FILE); c = conn.cursor(); c.execute("SELECT raw_text FROM materials WHERE course=? AND doc_type IN ('教材/讲义大纲', '往年真题试卷', '手动挂载资料')", (course,)); combined_context = "\n\n".join([r[0] for r in c.fetchall()])[:15000]; conn.close()
            ans = call_ai_answer_generator(self.config, course, question, combined_context); self.root.after(0, self.update_answer_field, ans)
        except Exception as e: self.root.after(0, self.update_answer_field, f"异常: {e}")

    def update_answer_field(self, ans_text):
        self.txt_a.delete(1.0, tk.END); self.txt_a.insert(tk.END, ans_text); self.btn_ai.config(text="🤖 RAG 文库提取作答", state=tk.NORMAL)

    def clear_editor(self):
        self.lbl_id.config(text="未选择"); self.ent_type.delete(0, tk.END)
        self.txt_q.delete(1.0, tk.END); self.txt_a.delete(1.0, tk.END)

    # ================= 🔎 库中试题全景浏览 =================
    def build_browse_tab(self, parent):
        f_top = tk.Frame(parent); f_top.pack(fill=tk.X, pady=8, padx=5)
        tk.Label(f_top, text="当前检索课程:", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=5)
        self.cb_browse_course = ttk.Combobox(f_top, state="readonly", width=22); self.cb_browse_course.pack(side=tk.LEFT, padx=5)
        self.cb_browse_course.bind("<<ComboboxSelected>>", lambda e: self.refresh_browse_tab())
        tk.Button(f_top, text="🔄 强制刷新总表", command=self.refresh_browse_tab).pack(side=tk.RIGHT, padx=10)

        paned_main = ttk.PanedWindow(parent, orient=tk.HORIZONTAL)
        paned_main.pack(fill=tk.BOTH, expand=True, padx=5, pady=2)

        f_left = ttk.LabelFrame(paned_main, text="📁 库中题型分类大纲")
        paned_main.add(f_left, weight=1)
        
        self.tree_browse_types = ttk.Treeview(f_left, show="tree headings")
        self.tree_browse_types.heading("#0", text="题型结构清单")
        self.tree_browse_types.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        self.tree_browse_types.bind("<<TreeviewSelect>>", self.on_browse_type_selected)

        f_right_wrap = tk.Frame(paned_main)
        paned_main.add(f_right_wrap, weight=5)

        paned_right_vertical = ttk.PanedWindow(f_right_wrap, orient=tk.VERTICAL)
        paned_right_vertical.pack(fill=tk.BOTH, expand=True)

        f_grid_frame = ttk.LabelFrame(paned_right_vertical, text="📋 全景考题清单 (支持多选一键批量合并题型)")
        paned_right_vertical.add(f_grid_frame, weight=5)

        columns_def = ("id", "question", "answer", "doc_type", "year", "chapter", "kp_name", "stars")
        self.grid_browse_q = ttk.Treeview(f_grid_frame, columns=columns_def, show="headings")
        self.grid_browse_q.heading("id", text="ID")
        self.grid_browse_q.heading("question", text="题干预览")
        self.grid_browse_q.heading("answer", text="标准答案")
        self.grid_browse_q.heading("doc_type", text="题目来源")
        self.grid_browse_q.heading("year", text="开考年份")
        self.grid_browse_q.heading("chapter", text="所属章标题")
        self.grid_browse_q.heading("kp_name", text="映射考点")
        self.grid_browse_q.heading("stars", text="考点星级")

        self.grid_browse_q.column("id", width=40, anchor="center")
        self.grid_browse_q.column("question", width=240)
        self.grid_browse_q.column("answer", width=140)
        self.grid_browse_q.column("doc_type", width=90, anchor="center")
        self.grid_browse_q.column("year", width=60, anchor="center")
        self.grid_browse_q.column("chapter", width=110)
        self.grid_browse_q.column("kp_name", width=110)
        self.grid_browse_q.column("stars", width=70, anchor="center")

        scroll_b_y = ttk.Scrollbar(f_grid_frame, orient="vertical", command=self.grid_browse_q.yview)
        self.grid_browse_q.configure(yscrollcommand=scroll_b_y.set)
        scroll_b_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.grid_browse_q.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        self.grid_browse_q.bind("<<TreeviewSelect>>", self.on_browse_grid_row_selected)

        f_actions_panel = ttk.LabelFrame(paned_right_vertical, text="🕹️ 选中试题集成化二次开发/校准控制台")
        paned_right_vertical.add(f_actions_panel, weight=4)

        f_r1 = tk.Frame(f_actions_panel); f_r1.pack(fill=tk.X, pady=3, padx=5)
        tk.Label(f_r1, text="题目ID:").pack(side=tk.LEFT)
        self.lbl_b_id = tk.Label(f_r1, text="无", fg="purple", font=("Arial", 10, "bold"), width=15); self.lbl_b_id.pack(side=tk.LEFT)
        tk.Label(f_r1, text="题型:").pack(side=tk.LEFT)
        self.ent_b_type = tk.Entry(f_r1, width=12); self.ent_b_type.pack(side=tk.LEFT, padx=5)
        
        tk.Label(f_r1, text=" | 批调考点重要度:").pack(side=tk.LEFT)
        self.cb_b_star = ttk.Combobox(f_r1, values=["1", "2", "3", "4", "5"], state="readonly", width=4)
        self.cb_b_star.pack(side=tk.LEFT, padx=2); self.cb_b_star.set("3")
        tk.Button(f_r1, text="⭐ 覆写星级", font=("Arial", 9), command=self.action_browse_batch_stars).pack(side=tk.LEFT, padx=3)

        tk.Button(f_r1, text="🔀 转移与同义词深度合并", font=("Arial", 9, "bold"), fg="#e67e22", command=self.action_browse_transfer_and_merge_dialog).pack(side=tk.RIGHT, padx=10)

        f_r2 = tk.Frame(f_actions_panel); f_r2.pack(fill=tk.BOTH, expand=True, pady=3, padx=5)
        f_r2_l = tk.Frame(f_r2); f_r2_l.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=2)
        tk.Label(f_r2_l, text="题目文本编辑区(多选时自动锁定):").pack(anchor=tk.W)
        self.txt_b_q = tk.Text(f_r2_l, font=("Arial", 10), height=4); self.txt_b_q.pack(fill=tk.BOTH, expand=True)

        f_r2_r = tk.Frame(f_r2); f_r2_r.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=2)
        tk.Label(f_r2_r, text="参考答案编辑区(多选时自动锁定):").pack(anchor=tk.W)
        self.txt_b_a = tk.Text(f_r2_r, font=("Arial", 10), height=4); self.txt_b_a.pack(fill=tk.BOTH, expand=True)

        f_r3 = tk.Frame(f_actions_panel); f_r3.pack(fill=tk.X, pady=4, padx=5)
        tk.Button(f_r3, text="💾 永久保存选中修改 (单选保存全部，多选一键合并题型)", font=("Arial", 10, "bold"), fg="#27ae60", command=self.action_browse_save_batch_or_single).pack(side=tk.LEFT, padx=20)
        tk.Button(f_r3, text="🗑️ 物理永久清空选中题目", font=("Arial", 10, "bold"), fg="#c0392b", command=self.action_browse_delete_batch).pack(side=tk.RIGHT, padx=20)

    def refresh_browse_tab(self):
        for i in self.tree_browse_types.get_children(): self.tree_browse_types.delete(i)
        for i in self.grid_browse_q.get_children(): self.grid_browse_q.delete(i)
        course = self.cb_browse_course.get()
        if not course: return

        conn = sqlite3.connect(DB_FILE); c = conn.cursor()
        c.execute("SELECT q_type, COUNT(*) FROM questions WHERE course=? AND q_type IS NOT NULL AND q_type!='' GROUP BY q_type", (course,))
        rows = c.fetchall()
        conn.close()

        root_node = self.tree_browse_types.insert("", tk.END, text=f"📚 {course} 题型分类", open=True)
        for r in rows: self.tree_browse_types.insert(root_node, tk.END, text=f"{r[0]} ({r[1]}道)", values=(r[0],))

    def on_browse_type_selected(self, event):
        sel = self.tree_browse_types.selection()
        if not sel: return
        node_text = self.tree_browse_types.item(sel[0])['text']
        if "📚" in node_text: return 
        
        q_type = self.tree_browse_types.item(sel[0])['values'][0]
        course = self.cb_browse_course.get()
        for i in self.grid_browse_q.get_children(): self.grid_browse_q.delete(i)
        
        conn = sqlite3.connect(DB_FILE); c = conn.cursor()
        c.execute("SELECT id, question, answer, doc_type, q_year, chapter_name, kp_name, kp_stars FROM questions WHERE course=? AND q_type=?", (course, q_type))
        rows = c.fetchall()
        conn.close()

        for r in rows:
            stars_repr = "⭐" * int(r[7]) if r[7] else "⭐"
            q_preview = r[1].replace('\n', ' ')
            a_preview = r[2].replace('\n', ' ')
            self.grid_browse_q.insert("", tk.END, values=(r[0], q_preview, a_preview, r[3], r[4], r[5], r[6], stars_repr))

    def on_browse_grid_row_selected(self, event):
        sel = self.grid_browse_q.selection()
        if not sel: return
        
        if len(sel) > 1:
            self.lbl_b_id.config(text=f"[已多选 {len(sel)} 道题]")
            self.txt_b_q.delete(1.0, tk.END); self.txt_b_q.insert(tk.END, "[多条题目已选中，文本编辑区已禁用，可在上方直接批量覆写题型]")
            self.txt_b_a.delete(1.0, tk.END); self.txt_b_a.insert(tk.END, "[多条题目已选中，文本编辑区已禁用，可在上方直接批量覆写题型]")
            first_row_id = self.grid_browse_q.item(sel[0])['values'][0]
            conn = sqlite3.connect(DB_FILE); c = conn.cursor()
            c.execute("SELECT q_type FROM questions WHERE id=?", (first_row_id,))
            r = c.fetchone(); conn.close()
            if r: self.ent_b_type.delete(0, tk.END); self.ent_b_type.insert(0, r[0])
            return

        main_row_id = self.grid_browse_q.item(sel[0])['values'][0]
        conn = sqlite3.connect(DB_FILE); c = conn.cursor()
        c.execute("SELECT id, q_type, question, answer, kp_stars FROM questions WHERE id=?", (main_row_id,))
        r = c.fetchone(); conn.close()
        if r:
            self.lbl_b_id.config(text=str(r[0]))
            self.ent_b_type.delete(0, tk.END); self.ent_b_type.insert(0, r[1])
            self.txt_b_q.delete(1.0, tk.END); self.txt_b_q.insert(tk.END, r[2])
            self.txt_b_a.delete(1.0, tk.END); self.txt_b_a.insert(tk.END, r[3])
            self.cb_b_star.set(str(r[4]))

    def action_browse_save_batch_or_single(self):
        selections = self.grid_browse_q.selection()
        if not selections: return messagebox.showwarning("提示", "当前未在大表中选中任何要修改的目标题目！")
        
        qt = self.ent_b_type.get().strip()
        standard_types = get_standard_q_types()
        norm_type = normalize_question_type(qt, standard_types) 
        
        try:
            conn = sqlite3.connect(DB_FILE); c = conn.cursor()
            if len(selections) > 1:
                for item in selections:
                    db_id = self.grid_browse_q.item(item)['values'][0]
                    c.execute("UPDATE questions SET q_type=? WHERE id=?", (norm_type, db_id))
                conn.commit(); conn.close()
                self.log(f"✅ [题型物理归一化]：成功批量将 {len(selections)} 道旧题目合并规范为【{norm_type}】")
                messagebox.showinfo("批量合并成功", f"已成功将选中的 {len(selections)} 道题目一键合并至标准的【{norm_type}】题型。")
            else:
                q_id = self.lbl_b_id.cget("text")
                if q_id == "无": return
                q_text = self.txt_b_q.get(1.0, tk.END).strip()
                a_text = self.txt_b_a.get(1.0, tk.END).strip()
                c.execute("UPDATE questions SET q_type=?, question=?, answer=? WHERE id=?", (norm_type, q_text, a_text, q_id))
                conn.commit(); conn.close()
                messagebox.showinfo("成功", "本地修改已清洗规范并安全覆盖。")
            self.refresh_browse_tab()
        except Exception as e: messagebox.showerror("重写错误", str(e))

    def action_browse_delete_batch(self):
        selections = self.grid_browse_q.selection()
        if not selections: return messagebox.showwarning("提示", "请选择题目！")
        if messagebox.askyesno("全景清洗确认", f"确定清空选中的 {len(selections)} 道题目吗？"):
            try:
                conn = sqlite3.connect(DB_FILE); c = conn.cursor()
                for item in selections: c.execute("DELETE FROM questions WHERE id=?", (self.grid_browse_q.item(item)['values'][0],))
                conn.commit(); conn.close(); self.refresh_browse_tab()
            except Exception as e: messagebox.showerror("错误", str(e))

    def action_browse_batch_stars(self):
        selections = self.grid_browse_q.selection()
        if not selections: return
        target_star = int(self.cb_b_star.get())
        try:
            conn = sqlite3.connect(DB_FILE); c = conn.cursor()
            for item in selections: c.execute("UPDATE questions SET kp_stars=? WHERE id=?", (target_star, self.grid_browse_q.item(item)['values'][0]))
            conn.commit(); conn.close(); self.on_browse_type_selected(None)
        except Exception as e: messagebox.showerror("错误", str(e))

    def action_browse_transfer_and_merge_dialog(self):
        selections = self.grid_browse_q.selection()
        if not selections: return messagebox.showwarning("提示", "请先在大表中选中需要转移的试题！")
        course = self.cb_browse_course.get()
        conn = sqlite3.connect(DB_FILE); c = conn.cursor()
        c.execute("SELECT chapter_name FROM course_chapters WHERE course=?", (course,))
        all_chapters = [r[0] for r in c.fetchall()]
        c.execute("SELECT DISTINCT kp_name FROM questions WHERE course=? AND kp_name IS NOT NULL AND kp_name!=''", (course,))
        all_kps = [r[0] for r in c.fetchall()]
        conn.close()
        if not all_chapters: return messagebox.showwarning("提示", "本课程还没有章节大纲，无法转移。")

        win = tk.Toplevel(self.root); win.title("试题跨章转移与合并中枢舱"); win.geometry("420x260"); win.grab_set()
        tk.Label(win, text=f"正在重构：正在合并转移选中的 {len(selections)} 道题目", font=("Arial", 10, "bold"), fg="red").pack(pady=10)
        tk.Label(win, text="1. 选择转移至的目标章节:").pack(anchor=tk.W, padx=30)
        cb_ch = ttk.Combobox(win, values=all_chapters, state="readonly", width=28); cb_ch.pack(pady=3); cb_ch.set(all_chapters[0])
        tk.Label(win, text="2. 选择/输入目标考点名称(支持同义自动重组):").pack(anchor=tk.W, padx=30)
        cb_kp = ttk.Combobox(win, values=all_kps, width=28); cb_kp.pack(pady=3)
        if all_kps: cb_kp.set(all_kps[0])

        def do_execute_transfer():
            target_ch = cb_ch.get(); target_kp = cb_kp.get().strip()
            if not target_kp: return messagebox.showwarning("警告", "目标知识点名称不能为空！", parent=win)
            try:
                conn_ex = sqlite3.connect(DB_FILE); cur_ex = conn_ex.cursor()
                for item in selections: cur_ex.execute("UPDATE questions SET chapter_name=?, kp_name=? WHERE id=?", (target_ch, target_kp, self.grid_browse_q.item(item)['values'][0]))
                conn_ex.commit(); conn_ex.close(); win.destroy(); self.refresh_browse_tab()
            except Exception as e: messagebox.showerror("错误", str(e), parent=win)
        tk.Button(win, text="🚀 确认发生级联转移与同义合并", font=("Arial", 10, "bold"), bg="#27ae60", fg="white", command=do_execute_transfer).pack(pady=20)

    # ================= 🖨️ 业务流 4：分层结构化抽样组卷 =================
    def build_generate_tab(self, parent):
        f_top = tk.Frame(parent); f_top.pack(fill=tk.X, pady=10)
        tk.Label(f_top, text="第一步 选择目标课程:").pack(side=tk.LEFT, padx=10)
        self.cb_gen_course = ttk.Combobox(f_top, state="readonly", width=25); self.cb_gen_course.pack(side=tk.LEFT)
        self.cb_gen_course.bind("<<ComboboxSelected>>", lambda e: self.on_gen_course_changed())

        self.f_blueprint = ttk.LabelFrame(parent, text="第二步 规划期末大纲蓝图 (设置题型数量及单题分值)")
        self.f_blueprint.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.canvas = tk.Canvas(self.f_blueprint, borderwidth=0, highlightthickness=0)
        self.scroll_y = ttk.Scrollbar(self.f_blueprint, orient="vertical", command=self.canvas.yview)
        self.f_rows_container = tk.Frame(self.canvas)
        
        self.canvas.configure(yscrollcommand=self.scroll_y.set)
        self.scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.canvas.create_window((0,0), window=self.f_rows_container, anchor="nw", tags="frame")
        self.f_rows_container.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        f_actions = tk.Frame(parent); f_actions.pack(fill=tk.X, pady=15)
        tk.Button(f_actions, text="✨ 开启全自动双向细目寻优组卷", font=("Arial", 12, "bold"), bg="#27ae60", fg="white", height=2, command=self.start_gen).pack(fill=tk.X, padx=50)

    def on_gen_course_changed(self):
        self.available_types = get_standard_q_types()
        for row in self.type_rows: row['frame'].destroy()
        self.type_rows.clear()
        self.add_type_row()

    def add_type_row(self):
        row_idx = len(self.type_rows) + 1
        f_row = tk.Frame(self.f_rows_container, pady=5)
        f_row.pack(fill=tk.X, anchor=tk.W)

        lbl_no = tk.Label(f_row, text=f"结构部分 {row_idx}:", font=("Arial", 10, "bold"), width=12, anchor=tk.W)
        lbl_no.pack(side=tk.LEFT, padx=5)

        cb_type = ttk.Combobox(f_row, values=self.available_types, state="readonly", width=15)
        cb_type.pack(side=tk.LEFT, padx=5)
        
        if row_idx == 1 and "单选题" in self.available_types: cb_type.set("单选题")
        elif row_idx == 2 and "多选题" in self.available_types: cb_type.set("多选题")
        elif row_idx == 3 and "名词解释" in self.available_types: cb_type.set("名词解释")
        elif row_idx == 4 and "判断题" in self.available_types: cb_type.set("判断题")
        elif row_idx == 5 and "简答题" in self.available_types: cb_type.set("简答题")
        elif row_idx == 6 and "论述题" in self.available_types: cb_type.set("论述题")
        else:
            if self.available_types: cb_type.current(min(row_idx-1, len(self.available_types)-1))

        tk.Label(f_row, text="题目数量:").pack(side=tk.LEFT, padx=2)
        ent_num = tk.Entry(f_row, width=6)
        if row_idx == 1: ent_num.insert(0, "10")
        else: ent_num.insert(0, "5")
        ent_num.pack(side=tk.LEFT, padx=5)

        tk.Label(f_row, text="单题分值(分):").pack(side=tk.LEFT, padx=2)
        ent_score = tk.Entry(f_row, width=6)
        
        t_name = cb_type.get()
        if "单选" in t_name: ent_score.insert(0, "1")
        elif "多选" in t_name: ent_score.insert(0, "2")
        elif "判断" in t_name: ent_score.insert(0, "1")
        elif "名词解释" in t_name: ent_score.insert(0, "3")
        elif "简答" in t_name: ent_score.insert(0, "5")
        elif "论述" in t_name: ent_score.insert(0, "10")
        elif "资料" in t_name or "材料" in t_name: ent_score.insert(0, "15")
        else: ent_score.insert(0, "2")
        ent_score.pack(side=tk.LEFT, padx=5)

        def on_type_reselected(event, es=ent_score, cb=cb_type):
            name = cb.get()
            es.delete(0, tk.END)
            if "单选" in name: es.insert(0, "1")
            elif "多选" in name: es.insert(0, "2")
            elif "判断" in name: es.insert(0, "1")
            elif "名词解释" in name: es.insert(0, "3")
            elif "简答" in name: es.insert(0, "5")
            elif "论述" in name: es.insert(0, "10")
            elif "资料" in name or "材料" in name: es.insert(0, "15")
            else: es.insert(0, "2")

        cb_type.bind("<<ComboboxSelected>>", on_type_reselected)

        btn_add = tk.Button(f_row, text=" ➕ ", font=("Arial", 9, "bold"), fg="green", command=self.add_type_row)
        btn_add.pack(side=tk.LEFT, padx=2)
        btn_del = tk.Button(f_row, text=" ➖ ", font=("Arial", 9), fg="red", command=lambda: self.remove_type_row(f_row))
        btn_del.pack(side=tk.LEFT, padx=2)
        if row_idx == 1: btn_del.config(state=tk.DISABLED)

        self.type_rows.append({
            'frame': f_row, 'label': lbl_no, 'cb_type': cb_type, 'ent_num': ent_num, 'ent_score': ent_score, 'btn_del': btn_del
        })
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def remove_type_row(self, frame_target):
        for row in self.type_rows:
            if row['frame'] == frame_target:
                row['frame'].destroy()
                self.type_rows.remove(row)
                break
        for idx, row in enumerate(self.type_rows):
            row['label'].config(text=f"结构部分 {idx+1}:")
            if idx == 0: row['btn_del'].config(state=tk.DISABLED)

    def generate_paper_text(self, course, name, paper_parts):
        text = f"{course} - {name}\n\n"
        text += "一、 考试题目\n"
        chinese_numbers = ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十"]
        for idx, part in enumerate(paper_parts):
            ch_num = chinese_numbers[idx] if idx < len(chinese_numbers) else str(idx + 1)
            text += f"{ch_num}、{part['q_type']}（每小题{part['single_score']}分，共{part['part_total_score']:.1f}分）\n"
            for q_idx, q in enumerate(part['questions']): text += f"  {q_idx + 1}. {q['question']}\n"
            text += "\n"
        text += "二、 参考答案及命题双向细目表\n"
        for idx, part in enumerate(paper_parts):
            ch_num = chinese_numbers[idx] if idx < len(chinese_numbers) else str(idx + 1)
            text += f"{ch_num}、{part['q_type']}部分参考答案\n"
            for q_idx, q in enumerate(part['questions']):
                text += f"  {q_idx + 1}. 【标准答案】: {q['answer']} (星级:{q['kp_stars']}⭐ | 归属:{q['chapter_name']})\n"
            text += "\n"
        return text

    # ================= 组卷核心辅助函数 =================
    def _random_sample_paper(self, all_pool, blueprint, exclude_ids=None):
        """随机加权不放回抽样生成一套试卷"""
        if exclude_ids is None:
            exclude_ids = set()
        parts = []
        flat_qs = []
        used_ids = set()
        for part in blueprint:
            qt = part['q_type']
            p = all_pool[qt]
            q_pool = p['pool']
            needed = p['needed']
            score = p['score']
            # 优先从exclude之外选取
            available = [q for q in q_pool if q['id'] not in exclude_ids and q['id'] not in used_ids]
            if len(available) < needed:
                extra = [q for q in q_pool if q['id'] in exclude_ids and q['id'] not in used_ids]
                available.extend(extra)
            if len(available) < needed:
                return None, None, None
            weights_vector = [float(q.get('kp_stars', 3)) ** 3 for q in available]
            sampled = []
            temp_pool = available.copy()
            temp_weights = weights_vector.copy()
            for _ in range(needed):
                if not temp_pool:
                    break
                total_w = sum(temp_weights)
                if total_w <= 0:
                    idx = random.randint(0, len(temp_pool) - 1)
                else:
                    r = random.uniform(0, total_w)
                    cumsum = 0
                    idx = 0
                    for i, w in enumerate(temp_weights):
                        cumsum += w
                        if cumsum >= r:
                            idx = i
                            break
                sampled.append(temp_pool[idx])
                del temp_pool[idx]
                del temp_weights[idx]
            if len(sampled) < needed:
                return None, None, None
            part_qs = []
            for q in sampled:
                qc = q.copy()
                qc['score'] = score
                part_qs.append(qc)
                flat_qs.append(qc)
                used_ids.add(q['id'])
            parts.append({'q_type': qt, 'single_score': score, 'part_total_score': needed * score, 'questions': part_qs})
        return parts, flat_qs, used_ids

    def _calc_paper_fitness(self, flat_qs, chapter_target_scores, min_single_score):
        """计算试卷适应度和章节分布"""
        ch_scores = {}
        total_stars = 0.0
        for q in flat_qs:
            ch = q.get('chapter_name', '未分类章节')
            ch_scores[ch] = ch_scores.get(ch, 0.0) + q['score']
            total_stars += float(q.get('kp_stars', 3)) * 10.0
        variance_error = 0.0
        for ch, target in chapter_target_scores.items():
            actual = ch_scores.get(ch, 0.0)
            if target <= min_single_score:
                if actual < min_single_score - 0.01:
                    variance_error += 200.0
            else:
                variance_error += ((actual - target) ** 2) * 3.0
        # 鼓励章节覆盖
        coverage_bonus = 0
        for ch, target in chapter_target_scores.items():
            if target > 0 and ch_scores.get(ch, 0) > 0:
                coverage_bonus += 50
        fitness = total_stars - variance_error + coverage_bonus
        return fitness, ch_scores

    def _flatten_paper(self, paper_parts):
        """将paper_parts扁平化为题目列表"""
        return [q for part in paper_parts for q in part['questions']]

    def _greedy_generate_paper(self, all_pool, blueprint, chapter_target_scores, min_single_score, exclude_paper=None):
        """贪心保底算法生成一套试卷"""
        exclude_ids = set()
        if exclude_paper:
            for part in exclude_paper:
                for q in part['questions']:
                    exclude_ids.add(q['id'])
        parts = []
        chapter_current = {ch: 0.0 for ch in chapter_target_scores}
        for part in blueprint:
            qt = part['q_type']
            p = all_pool[qt]
            q_pool = p['pool']
            needed = p['needed']
            score = p['score']
            available = [q for q in q_pool if q['id'] not in exclude_ids]
            fallback = [q for q in q_pool if q['id'] in exclude_ids]
            selected = []
            working = available.copy()
            for _ in range(needed):
                if not working:
                    if fallback:
                        working = fallback.copy()
                        fallback = []
                    else:
                        break
                best_q = None
                best_val = -float('inf')
                for q in working:
                    ch = q.get('chapter_name', '未分类章节')
                    stars = float(q.get('kp_stars', 3))
                    target = chapter_target_scores.get(ch, 0)
                    current = chapter_current.get(ch, 0)
                    gap = max(0, target - current)
                    # 综合价值：星级质量 + 章节缺口补偿 - 同章重复惩罚
                    val = stars * 15 + gap * 8
                    same_ch_count = sum(1 for sq in selected if sq.get('chapter_name') == ch)
                    val -= same_ch_count * 2
                    if val > best_val:
                        best_val = val
                        best_q = q
                if best_q is None:
                    break
                qc = best_q.copy()
                qc['score'] = score
                selected.append(qc)
                ch = best_q.get('chapter_name', '未分类章节')
                chapter_current[ch] = chapter_current.get(ch, 0.0) + score
                working = [q for q in working if q['id'] != best_q['id']]
            if len(selected) < needed:
                return None
            parts.append({'q_type': qt, 'single_score': score, 'part_total_score': needed * score, 'questions': selected})
        return parts

    def _save_paper_to_db(self, course, paper_name, paper_parts):
        """保存试卷到papers/paper_questions/materials"""
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute("INSERT INTO papers (course, paper_name) VALUES (?, ?)", (course, paper_name))
        paper_id = c.lastrowid
        for part in paper_parts:
            for q in part['questions']:
                # 记录该题当前的 q_year，用于撤销时恢复
                c.execute("SELECT q_year FROM questions WHERE id=?", (q['id'],))
                row = c.fetchone()
                prev_year = row[0] if row and row[0] is not None else None
                c.execute("INSERT INTO paper_questions (paper_id, question_id, score, prev_year) VALUES (?, ?, ?, ?)",
                          (paper_id, q['id'], part['single_score'], prev_year))
        paper_text = self.generate_paper_text(course, paper_name, paper_parts)
        c.execute("INSERT INTO materials (course, raw_text, doc_type) VALUES (?, ?, ?)",
                  (course, paper_text, '生成的试卷'))
        conn.commit()
        conn.close()
        return paper_id

    def start_gen(self):
        course = self.cb_gen_course.get()
        if not course: return messagebox.showwarning("警告", "请选择出卷课程！")
        blueprint = []
        total_paper_score = 0
        min_single_score = 999.0  
        for row in self.type_rows:
            q_type = row['cb_type'].get().strip()
            try:
                num = int(row['ent_num'].get().strip())
                score = float(row['ent_score'].get().strip())
            except: return messagebox.showerror("格式错误", f"题型的数量及单题分值必须是纯数字！")
            if num <= 0: continue
            blueprint.append({'q_type': q_type, 'count': num, 'single_score': score})
            total_paper_score += (num * score)
            if score < min_single_score: min_single_score = score

        if not blueprint: return messagebox.showwarning("警告", "请配置好各部分题型的抽样配额！")
        self.log(f"\n>>> 试卷蓝图规划完毕。总设计分值：{total_paper_score} 分。最小单题分值: {min_single_score} 分.")
        threading.Thread(target=self.thread_gen_by_score_weight, args=(course, blueprint, total_paper_score, min_single_score), daemon=True).start()

    # ==========================================================================
    # 🖨️ 组卷核心内核：基于分值权重与近3年（2024/2025）查重强固拦截的智能组卷引擎
    # ==========================================================================
    def thread_gen_by_score_weight(self, course, blueprint, total_paper_score, min_single_score):
        current_gen_year = int(time.strftime("%Y")) # 自动获取当前年份（2026）
        ban_years = (current_gen_year - 1, current_gen_year - 2) # 锁定需屏蔽的往年：2025, 2024
        
        self.log(f"\n>>> 启动【{course}】教研级“近3年禁重策略”高级组卷内核...")
        self.log(f"保护盾状态: 激活。当前组卷年份 {current_gen_year} 年。物理隔离屏蔽 {ban_years[0]} 年与 {ban_years[1]} 年的所有已用真题...")

        conn = sqlite3.connect(DB_FILE)
        conn.row_factory = sqlite3.Row
        c = conn.cursor()

        # ------------------------------------------------------------------
        # 1. 加载章节权重（无权重时自适应使用均匀分布）
        # ------------------------------------------------------------------
        c.execute("SELECT chapter_name, weight FROM course_chapters WHERE course=?", (course,))
        weights_rows = c.fetchall()
        weights = {r['chapter_name']: r['weight'] for r in weights_rows}
        if not weights or sum(weights.values()) == 0:
            self.log("⚠️ 提示：未检测到各章节权重配额，系统将自动切入【全章节均匀分布】策略。")
            weights = {}
            
        total_weight = sum(weights.values()) if weights else 0
        chapter_target_scores = {}
        for ch, w in weights.items():
            if total_weight > 0:
                raw_target = (w / total_weight) * total_paper_score
            else:
                raw_target = 0
            if 0 < raw_target < min_single_score:
                chapter_target_scores[ch] = min_single_score
            else:
                chapter_target_scores[ch] = raw_target

        # ------------------------------------------------------------------
        # 2. 物理层去重过滤：加载各题型题库，严格拦截 2024/2025 年题目
        # ------------------------------------------------------------------
        all_pool = {}
        for part in blueprint:
            qt = part['q_type']
            
            # 使用 SQL 复合条件语句：当 q_year 为空，或者不在禁考年份（2024, 2025）内时，才允许进入抽样池
            query_secure = """
                SELECT * FROM questions 
                WHERE course=? AND q_type=? 
                  AND (q_year IS NULL OR q_year NOT IN (?, ?))
                ORDER BY kp_stars DESC
            """
            c.execute(query_secure, (course, qt, ban_years[0], ban_years[1]))
            raw_list = [dict(r) for r in c.fetchall()]
            
            # 查重剔除后的可用题库库存熔断检查
            if len(raw_list) < part['count']:
                self.log(f"❌ [库容熔断]: 题型【{qt}】在剔除2024/2025年已考题目后，剩余安全库存仅剩 {len(raw_list)} 道，无法满足本次出卷需要的 {part['count']} 道！")
                conn.close()
                # 切回主线程弹出警告，优雅熔断退出
                self.root.after(0, lambda: messagebox.showerror("可用题库枯竭",
                    f"因执行【近3年禁重】教研铁律，题型【{qt}】的历史重复题目已被安全隔离！\n\n"
                    f"隔离后该题型可用库存仅剩 {len(raw_list)} 道，无法满足需要的 {part['count']} 道！\n\n"
                    f"建议处理办法：\n1. 前往「📥 AI双轨抽取入库」为该课程补充新材料真题。\n2. 临时调低本次期末组卷中【{qt}】的抽样配额。"))
                return
                
            all_pool[qt] = {'pool': raw_list, 'score': part['single_score'], 'needed': part['count'], 'q_type': qt}
        conn.close()

        # ------------------------------------------------------------------
        # 3. 双轨细目矩阵寻优：400 次随机加权不放回抽样
        # ------------------------------------------------------------------
        best_A, best_B = None, None
        best_fit_A, best_fit_B = -float('inf'), -float('inf')
        best_dist_A, best_dist_B = None, None

        for attempt in range(400):
            # 抽取 A 卷
            parts_A, flat_A, used_ids_A = self._random_sample_paper(all_pool, blueprint)
            if not parts_A:
                continue
            # 抽取 B 卷（传入 used_ids_A，确保 A、B 卷之间完全不重题）
            parts_B, flat_B, _ = self._random_sample_paper(all_pool, blueprint, exclude_ids=used_ids_A)
            if not parts_B:
                continue
                
            # 计算适应度与章节覆盖误差
            fit_A, dist_A = self._calc_paper_fitness(flat_A, chapter_target_scores, min_single_score)
            fit_B, dist_B = self._calc_paper_fitness(flat_B, chapter_target_scores, min_single_score)
            
            if fit_A > best_fit_A:
                best_fit_A = fit_A; best_A = parts_A; best_dist_A = dist_A
            if fit_B > best_fit_B:
                best_fit_B = fit_B; best_B = parts_B; best_dist_B = dist_B

        # ------------------------------------------------------------------
        # 4. 贪心算法保底机制（当前期去重导致池子变窄时，贪心算法可提供最优保底）
        # ------------------------------------------------------------------
        greedy_A = self._greedy_generate_paper(all_pool, blueprint, chapter_target_scores, min_single_score)
        greedy_B = self._greedy_generate_paper(all_pool, blueprint, chapter_target_scores, min_single_score, exclude_paper=greedy_A)

        if greedy_A:
            gfit_A, gdist_A = self._calc_paper_fitness(self._flatten_paper(greedy_A), chapter_target_scores, min_single_score)
            if gfit_A > best_fit_A:
                best_A = greedy_A; best_dist_A = gdist_A; best_fit_A = gfit_A
        if greedy_B:
            gfit_B, gdist_B = self._calc_paper_fitness(self._flatten_paper(greedy_B), chapter_target_scores, min_single_score)
            if gfit_B > best_fit_B:
                best_B = greedy_B; best_dist_B = gdist_B; best_fit_B = gfit_B

        # ------------------------------------------------------------------
        # 5. 组卷最终合规性校验
        # ------------------------------------------------------------------
        if not best_A or not best_B:
            self.log("❌ 组卷严重失败：当前安全题库极度匮乏，无法拼凑出满足大纲结构的完整试卷。")
            self.root.after(0, lambda: messagebox.showerror("组卷失败", "当前可用安全题库极度匮乏，无法完成组卷！请补充题库后再试。"))
            return

        # ------------------------------------------------------------------
        # 6. 吐出可视化双向细目章节分值分布报告
        # ------------------------------------------------------------------
        self.log(f"\n{'='*60}")
        self.log(f"📊 A卷【{sum(best_dist_A.values()):.1f}分】双向细目章节分值分布报告：")
        for ch in sorted(chapter_target_scores.keys()):
            target = chapter_target_scores[ch]
            actual = best_dist_A.get(ch, 0)
            self.log(f"   * {ch}: 目标分值 {target:.1f} 分 | 实际抽取 {actual:.1f} 分 | 偏差 {actual-target:+.1f} 分")
            
        self.log(f"📊 B卷【{sum(best_dist_B.values()):.1f}分】双向细目章节分值分布报告：")
        for ch in sorted(chapter_target_scores.keys()):
            target = chapter_target_scores[ch]
            actual = best_dist_B.get(ch, 0)
            self.log(f"   * {ch}: 目标分值 {target:.1f} 分 | 实际抽取 {actual:.1f} 分 | 偏差 {actual-target:+.1f} 分")
        self.log(f"{'='*60}")

        # ------------------------------------------------------------------
        # 7. 自动将生成的 A/B 试卷持久化收录入库（此步暂不刷新试题的使用年份，直至用户导出Word时才正式标记）
        # ------------------------------------------------------------------
        ts = time.strftime("%Y%m%d_%H%M%S")
        pid_A = self._save_paper_to_db(course, f"期末A卷_{ts}", best_A)
        pid_B = self._save_paper_to_db(course, f"期末B卷_{ts}", best_B)

        self.log(f"✅ 双轨组卷大功告成！A卷 (ID:{pid_A}) / B卷 (ID:{pid_B}) 已无损收录至本地试卷库。")
        self.root.after(0, lambda: messagebox.showinfo("出卷成功",
            f"遵照【近3年禁重】铁律，两套期末 A/B 卷已成功并行生成并收录入库！\n\n"
            f"请切换到「📚 本地全栈知识库管理」查看试卷详情。\n选中后点击「📤 使用与输出Word」即可秒级导出完美排版的真题卷及细目答案。"))
        
        # 刷新文库树
        self.root.after(0, self.load_kb_data)

    # ==========================================================
    # 🖨️ 组卷业务流：教研级 Word 试卷高保真排版引擎（宋体优化版）
    # ==========================================================
    def export_structured_docx(self, course, name, paper_parts):
        doc = Document()
        
        # --------------------------------------------------
        # ⚙️ 顶层设计：全局正统宋体样式基座配置
        # --------------------------------------------------
        style_normal = doc.styles['Normal']
        font = style_normal.font
        font.name = 'Times New Roman'  # 西文与数字默认使用 Times New Roman
        style_normal._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')  # 中文严格锁定为正统宋体
        font.size = Pt(10.5)  # 默认正文大小：五号字

        # --- 试卷大标题 (小二号，加粗，居中) ---
        title_p = doc.add_paragraph()
        title_p.alignment = 1  # 居中对齐
        title_run = title_p.add_run(f"{course} - {name}")
        title_run.font.size = Pt(18)  # 小二号
        title_run.bold = True
        title_run.font.name = 'Times New Roman'
        title_p.paragraph_format.space_after = Pt(20)

        # --------------------------------------------------
        # 📝 第一部分：考试题目部分
        # --------------------------------------------------
        h1_exam = doc.add_paragraph()
        run_h1_exam = h1_exam.add_run("一、 考试题目部分")
        run_h1_exam.font.size = Pt(14)  # 四号字
        run_h1_exam.bold = True
        h1_exam.paragraph_format.space_before = Pt(12)
        h1_exam.paragraph_format.space_after = Pt(6)

        chinese_numbers = ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十"]
        
        for idx, part in enumerate(paper_parts):
            ch_num = chinese_numbers[idx] if idx < len(chinese_numbers) else str(idx + 1)
            header_text = f"（{ch_num}）{part['q_type']}（每小题 {part['single_score']} 分，共 {part['part_total_score']:.1f} 分）"
            
            p_part = doc.add_paragraph()
            run_part = p_part.add_run(header_text)
            run_part.font.size = Pt(12)  # 小四号
            run_part.bold = True
            p_part.paragraph_format.space_before = Pt(8)
            p_part.paragraph_format.space_after = Pt(4)
            
            for q_idx, q in enumerate(part['questions']):
                p_q = doc.add_paragraph()
                p_q.paragraph_format.left_indent = Inches(0.25)  # 缩进排版，错落有致
                p_q.paragraph_format.space_after = Pt(3)
                
                # 清洗题干，确保多行题目在 Word 里也能完美呈现
                q_text = str(q['question']).strip()
                p_q.add_run(f"{q_idx + 1}.  {q_text}")

        # --- 强行插入教研级换页符，使试卷与答案物理隔离，方便单独打印 ---
        doc.add_page_break()

        # --------------------------------------------------
        # 📊 第二部分：参考答案及命题双向细目溯源报告
        # --------------------------------------------------
        h1_ans = doc.add_paragraph()
        run_h1_ans = h1_ans.add_run("二、 参考答案及命题双向细目溯源报告")
        run_h1_ans.font.size = Pt(14)  # 四号字
        run_h1_ans.bold = True
        h1_ans.paragraph_format.space_before = Pt(12)
        h1_ans.paragraph_format.space_after = Pt(6)

        for idx, part in enumerate(paper_parts):
            ch_num = chinese_numbers[idx] if idx < len(chinese_numbers) else str(idx + 1)
            header_text = f"（{ch_num}）{part['q_type']} 部分标准答案"
            
            p_part_ans = doc.add_paragraph()
            run_part_ans = p_part_ans.add_run(header_text)
            run_part_ans.font.size = Pt(12)  # 小四号
            run_part_ans.bold = True
            p_part_ans.paragraph_format.space_before = Pt(8)
            p_part_ans.paragraph_format.space_after = Pt(4)
            
            for q_idx, q in enumerate(part['questions']):
                p_a = doc.add_paragraph()
                p_a.paragraph_format.left_indent = Inches(0.25)
                p_a.paragraph_format.space_after = Pt(4)
                
                # 标准答案头部加粗
                run_label = p_a.add_run(f"{q_idx + 1}. 【参考答案】：")
                run_label.bold = True
                
                # 注入参考答案正文
                ans_text = str(q['answer']).strip()
                p_a.add_run(f"{ans_text}\n")
                
                # 命题细目双向溯源小字（使用斜体、灰色、小五号字，极具学术严谨性）
                stars_repr = "★" * int(q.get('kp_stars', 3))
                trace_text = f"   [命题细目溯源] ➔ 考点归属：{q['chapter_name']} | 考点名称：{q['kp_name']} | 考点重要度：{stars_repr}"
                run_trace = p_a.add_run(trace_text)
                run_trace.font.size = Pt(9)  # 小五号
                run_trace.italic = True
                
                # 🚀 核心改写：这里直接使用顶层导入的 RGBColor，不再加 docx.shared 前缀，彻底绝缘报错
                run_trace.font.color.rgb = RGBColor(127, 127, 127)  # 教研专业灰色

        # 核心持久化安全落盘
        doc.save(f"{course}_{name}.docx")

    def use_and_export_paper(self):
        """选中试卷后：输出Word文档 + 标记题目使用年份"""
        sel = self.tree_kb.selection()
        if not sel:
            return messagebox.showwarning("提示", "请先选中一份试卷！")
        item = self.tree_kb.item(sel[0])
        if 'paper' not in item.get("tags", []):
            return messagebox.showwarning("提示", "请选中「试卷库」下的试卷节点！")

        paper_id = item['values'][0]
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute("SELECT course, paper_name FROM papers WHERE id=?", (paper_id,))
        row = c.fetchone()
        if not row:
            conn.close()
            return messagebox.showerror("错误", "试卷记录不存在！")
        course, paper_name = row

        # 读取试卷题目及分值
        c.execute("""
            SELECT q.id, q.q_type, q.question, q.answer, q.chapter_name, q.kp_name, q.kp_stars, pq.score
            FROM questions q JOIN paper_questions pq ON q.id = pq.question_id
            WHERE pq.paper_id=? ORDER BY q.q_type, q.id
        """, (paper_id,))
        rows = c.fetchall()
        conn.close()

        if not rows:
            return messagebox.showwarning("提示", "该试卷没有关联题目！")

        # 组织成paper_parts格式
        type_map = {}
        for r in rows:
            qt = r[1]
            if qt not in type_map:
                type_map[qt] = []
            type_map[qt].append({
                'id': r[0], 'q_type': r[1], 'question': r[2], 'answer': r[3],
                'chapter_name': r[4], 'kp_name': r[5], 'kp_stars': r[6],
                'score': r[7]
            })

        paper_parts = []
        for qt, questions in type_map.items():
            score = questions[0]['score'] if questions else 0
            paper_parts.append({
                'q_type': qt,
                'single_score': score,
                'part_total_score': len(questions) * score,
                'questions': questions
            })

        # 输出Word
        try:
            self.export_structured_docx(course, paper_name, paper_parts)
        except Exception as e:
            return messagebox.showerror("导出失败", f"Word文档生成失败：{str(e)}")

        # 标记题目使用年份
        current_year = int(time.strftime("%Y"))
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        updated = 0
        for r in rows:
            c.execute("UPDATE questions SET q_year=? WHERE id=?", (current_year, r[0]))
            updated += 1
        conn.commit()
        conn.close()

        self.log(f"✅ 试卷【{paper_name}】Word文档已输出，共{updated}道题目标记使用年份为{current_year}。")
        messagebox.showinfo("导出成功",
            f"试卷【{paper_name}】Word文档已输出！\n"
            f"共 {updated} 道题目已标记使用年份为 {current_year}。")


    def _show_doc_view(self):
        """切换到文档视图"""
        self.f_paper_view.pack_forget()
        self.f_doc_view.pack(fill=tk.BOTH, expand=True)

    def _show_paper_view(self, paper_id):
        """切换到试卷编辑器视图并加载数据"""
        self.f_doc_view.pack_forget()
        self.f_paper_view.pack(fill=tk.BOTH, expand=True)
        self.lbl_kb_id.config(text=f"试卷ID:{paper_id}")
        self.current_paper_id = paper_id
        # 加载题型树
        for i in self.tree_paper_types.get_children(): self.tree_paper_types.delete(i)
        for i in self.grid_paper_q.get_children(): self.grid_paper_q.delete(i)
        self.lbl_p_id.config(text="无")
        self.ent_p_type.delete(0, tk.END)
        self.txt_p_q.delete(1.0, tk.END)
        self.txt_p_a.delete(1.0, tk.END)
        conn = sqlite3.connect(DB_FILE); c = conn.cursor()
        c.execute("SELECT paper_name, course FROM papers WHERE id=?", (paper_id,))
        row = c.fetchone()
        paper_name = row[0] if row else "未知试卷"
        course = row[1] if row else ""
        # 已有题型及数量
        c.execute("""
            SELECT q.q_type, COUNT(*) FROM questions q
            JOIN paper_questions pq ON q.id = pq.question_id
            WHERE pq.paper_id=? GROUP BY q.q_type
        """, (paper_id,))
        existing = dict(c.fetchall())
        # 课程下所有题型（含0道）
        c.execute("SELECT DISTINCT q_type FROM questions WHERE course=? AND q_type IS NOT NULL AND q_type!=''", (course,))
        all_types = [r[0] for r in c.fetchall()]
        conn.close()
        root_node = self.tree_paper_types.insert("", tk.END, text=f"📝 {paper_name}", open=True)
        for qt in all_types:
            count = existing.get(qt, 0)
            self.tree_paper_types.insert(root_node, tk.END, text=f"{qt} ({count}道)", values=(qt,))

    def on_paper_type_selected(self, event):
        """选中试卷中的题型，加载题目列表"""
        sel = self.tree_paper_types.selection()
        if not sel: return
        node_text = self.tree_paper_types.item(sel[0])['text']
        if "📝" in node_text: return
        q_type = self.tree_paper_types.item(sel[0])['values'][0]
        paper_id = getattr(self, 'current_paper_id', None)
        if not paper_id: return
        for i in self.grid_paper_q.get_children(): self.grid_paper_q.delete(i)
        conn = sqlite3.connect(DB_FILE); c = conn.cursor()
        c.execute("""
            SELECT q.id, q.question, q.answer, q.chapter_name, q.kp_name, q.kp_stars, pq.score
            FROM questions q JOIN paper_questions pq ON q.id = pq.question_id
            WHERE pq.paper_id=? AND q.q_type=? ORDER BY q.id
        """, (paper_id, q_type))
        rows = c.fetchall()
        conn.close()
        for r in rows:
            stars_repr = "⭐" * int(r[5]) if r[5] else "⭐"
            q_preview = r[1].replace('\n', ' ')[:50]
            a_preview = r[2].replace('\n', ' ')[:30]
            self.grid_paper_q.insert("", tk.END, values=(r[0], q_preview, a_preview, r[3], r[4], stars_repr, r[6]))

    def on_paper_grid_row_selected(self, event):
        """选中试卷中的题目，加载到编辑区"""
        sel = self.grid_paper_q.selection()
        if not sel: return
        if len(sel) > 1:
            self.lbl_p_id.config(text=f"[已多选 {len(sel)} 道]")
            self.txt_p_q.delete(1.0, tk.END); self.txt_p_q.insert(tk.END, "[多选锁定，仅可批量覆写题型或星级]")
            self.txt_p_a.delete(1.0, tk.END); self.txt_p_a.insert(tk.END, "[多选锁定]")
            first_id = self.grid_paper_q.item(sel[0])['values'][0]
            conn = sqlite3.connect(DB_FILE); c = conn.cursor()
            c.execute("SELECT q_type FROM questions WHERE id=?", (first_id,))
            r = c.fetchone(); conn.close()
            if r: self.ent_p_type.delete(0, tk.END); self.ent_p_type.insert(0, r[0])
            return
        q_id = self.grid_paper_q.item(sel[0])['values'][0]
        conn = sqlite3.connect(DB_FILE); c = conn.cursor()
        c.execute("SELECT q_type, question, answer, kp_stars FROM questions WHERE id=?", (q_id,))
        r = c.fetchone(); conn.close()
        if r:
            self.lbl_p_id.config(text=str(r[0]))
            self.ent_p_type.delete(0, tk.END); self.ent_p_type.insert(0, r[0])
            self.txt_p_q.delete(1.0, tk.END); self.txt_p_q.insert(tk.END, r[1])
            self.txt_p_a.delete(1.0, tk.END); self.txt_p_a.insert(tk.END, r[2])
            self.cb_p_star.set(str(r[3]))

    def action_paper_save(self):
        """保存试卷中题目的修改"""
        sel = self.grid_paper_q.selection()
        if not sel:
            return messagebox.showwarning("提示", "请先在列表中选中要修改的题目！")
        if len(sel) > 1:
            # 批量修改题型
            qt = self.ent_p_type.get().strip()
            standard_types = get_standard_q_types()
            norm_type = normalize_question_type(qt, standard_types)
            conn = sqlite3.connect(DB_FILE); c = conn.cursor()
            for item in sel:
                qid = self.grid_paper_q.item(item)['values'][0]
                c.execute("UPDATE questions SET q_type=? WHERE id=?", (norm_type, qid))
            conn.commit(); conn.close()
            self.log(f"✅ 批量将 {len(sel)} 道题目合并为【{norm_type}】")
            messagebox.showinfo("成功", f"已批量将 {len(sel)} 道题目修改为【{norm_type}】")
            self.on_paper_type_selected(None)
            return
        q_id = self.lbl_p_id.cget("text")
        if q_id == "无": return
        qt = self.ent_p_type.get().strip()
        q_text = self.txt_p_q.get(1.0, tk.END).strip()
        a_text = self.txt_p_a.get(1.0, tk.END).strip()
        try:
            conn = sqlite3.connect(DB_FILE); c = conn.cursor()
            c.execute("UPDATE questions SET q_type=?, question=?, answer=? WHERE id=?", (qt, q_text, a_text, q_id))
            conn.commit(); conn.close()
            messagebox.showinfo("成功", "题目修改已保存！")
            self.on_paper_type_selected(None)
        except Exception as e: messagebox.showerror("失败", str(e))

    def action_paper_remove(self):
        """从当前试卷中移除选中的题目（不删除题库）"""
        sel = self.grid_paper_q.selection()
        if not sel:
            return messagebox.showwarning("提示", "请选中要移除的题目！")
        paper_id = getattr(self, 'current_paper_id', None)
        if not paper_id:
            return messagebox.showwarning("提示", "未选中有效试卷！")
        if messagebox.askyesno("确认移除", f"确定从当前试卷中移除选中的 {len(sel)} 道题目？（题库中仍保留）"):
            conn = sqlite3.connect(DB_FILE); c = conn.cursor()
            for item in sel:
                qid = self.grid_paper_q.item(item)['values'][0]
                c.execute("DELETE FROM paper_questions WHERE paper_id=? AND question_id=?", (paper_id, qid))
            conn.commit(); conn.close()
            self.log(f"↩️ 已从试卷中移除 {len(sel)} 道题目。")
            self.on_paper_type_selected(None)
            # 刷新题型树
            self._show_paper_view(paper_id)

    def action_paper_delete(self):
        """物理永久删除选中的题目（从题库和试卷中同时删除）"""
        sel = self.grid_paper_q.selection()
        if not sel:
            return messagebox.showwarning("提示", "请选中要删除的题目！")
        if messagebox.askyesno("危险确认", f"确定物理永久清空选中的 {len(sel)} 道题目？（将从题库和所有试卷中彻底删除）"):
            conn = sqlite3.connect(DB_FILE); c = conn.cursor()
            for item in sel:
                qid = self.grid_paper_q.item(item)['values'][0]
                c.execute("DELETE FROM paper_questions WHERE question_id=?", (qid,))
                c.execute("DELETE FROM questions WHERE id=?", (qid,))
            conn.commit(); conn.close()
            self.log(f"🗑️ 已物理永久删除 {len(sel)} 道题目。")
            self.on_paper_type_selected(None)
            paper_id = getattr(self, 'current_paper_id', None)
            if paper_id:
                self._show_paper_view(paper_id)


    def action_paper_add_type(self):
        """向当前试卷添加一个新题型（从课程题库中引入）"""
        paper_id = getattr(self, 'current_paper_id', None)
        if not paper_id:
            return messagebox.showwarning("提示", "请先选中一份试卷！")
        conn = sqlite3.connect(DB_FILE); c = conn.cursor()
        c.execute("SELECT course FROM papers WHERE id=?", (paper_id,))
        row = c.fetchone()
        if not row:
            conn.close(); return messagebox.showerror("错误", "试卷不存在！")
        course = row[0]
        # 获取课程下所有题型
        c.execute("SELECT DISTINCT q_type FROM questions WHERE course=? AND q_type IS NOT NULL AND q_type!=''", (course,))
        all_types = {r[0] for r in c.fetchall()}
        conn.close()
        # 获取试卷中已有题型
        existing = set()
        for child in self.tree_paper_types.get_children():
            item = self.tree_paper_types.item(child)
            if item.get('values'):
                existing.add(item['values'][0])
        available = sorted(all_types - existing)
        if not available:
            return messagebox.showinfo("提示", "该课程下所有题型已在试卷中，无需添加！")
        win = tk.Toplevel(self.root); win.title("添加新题型到试卷"); win.geometry("350x300"); win.grab_set()
        tk.Label(win, text=f"试卷中还没有的题型（共{len(available)}种）：", font=("Arial", 11, "bold")).pack(pady=10)
        lb = tk.Listbox(win, font=("Arial", 11), selectmode=tk.SINGLE)
        for t in available: lb.insert(tk.END, t)
        lb.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        def do_add():
            sel = lb.curselection()
            if not sel:
                return messagebox.showwarning("提示", "请选择一个题型！", parent=win)
            new_type = lb.get(sel[0])
            # 在左侧树中插入新节点（0道）
            root_node = self.tree_paper_types.get_children()[0]
            self.tree_paper_types.insert(root_node, tk.END, text=f"{new_type} (0道)", values=(new_type,))
            win.destroy()
            self.log(f"➕ 已向试卷添加题型【{new_type}】")
        tk.Button(win, text="🚀 确定添加", font=("Arial", 11, "bold"), bg="#27ae60", fg="white", command=do_add).pack(pady=10)

    def action_paper_add_questions(self):
        """向当前试卷的当前题型添加题目"""
        paper_id = getattr(self, 'current_paper_id', None)
        if not paper_id:
            return messagebox.showwarning("提示", "请先选中一份试卷！")
        # 获取当前选中的题型
        sel_type = self.tree_paper_types.selection()
        if not sel_type:
            return messagebox.showwarning("提示", "请先在左侧选中一个题型！")
        node_text = self.tree_paper_types.item(sel_type[0])['text']
        if "📝" in node_text:
            return messagebox.showwarning("提示", "请选中具体的题型节点！")
        q_type = self.tree_paper_types.item(sel_type[0])['values'][0]
        conn = sqlite3.connect(DB_FILE); c = conn.cursor()
        c.execute("SELECT course FROM papers WHERE id=?", (paper_id,))
        row = c.fetchone()
        if not row:
            conn.close(); return messagebox.showerror("错误", "试卷不存在！")
        course = row[0]
        # 查询该课程下该题型中不在当前试卷中的题目
        c.execute("""
            SELECT q.id, q.question, q.answer, q.chapter_name, q.kp_name, q.kp_stars
            FROM questions q
            WHERE q.course=? AND q.q_type=? AND q.id NOT IN (
                SELECT question_id FROM paper_questions WHERE paper_id=?
            )
            ORDER BY q.kp_stars DESC, q.id
        """, (course, q_type, paper_id))
        available = c.fetchall()
        conn.close()
        if not available:
            return messagebox.showinfo("提示", f"题型【{q_type}】下没有可添加的新题目（题库中该题型已全部在试卷中）！")
        win = tk.Toplevel(self.root); win.title(f"向【{q_type}】添加题目"); win.geometry("650x500"); win.grab_set()
        tk.Label(win, text=f"课程【{course}】中【{q_type}】的可用题目（共{len(available)}道）：", font=("Arial", 11, "bold")).pack(pady=5)
        # Treeview
        cols = ("id", "question", "chapter", "stars")
        tree = ttk.Treeview(win, columns=cols, show="headings", height=12)
        tree.heading("id", text="ID"); tree.column("id", width=40, anchor="center")
        tree.heading("question", text="题干预览"); tree.column("question", width=320)
        tree.heading("chapter", text="章节"); tree.column("chapter", width=100)
        tree.heading("stars", text="星级"); tree.column("stars", width=60, anchor="center")
        scroll = ttk.Scrollbar(win, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        scroll.pack(side=tk.RIGHT, fill=tk.Y); tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        for r in available:
            stars = "⭐" * int(r[5]) if r[5] else "⭐"
            preview = r[1].replace('\n', ' ')[:55]
            tree.insert("", tk.END, values=(r[0], preview, r[3], stars))
        # 分值设置
        f_score = tk.Frame(win); f_score.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(f_score, text="添加后每题分值:", font=("Arial", 10)).pack(side=tk.LEFT)
        ent_score = tk.Entry(f_score, width=6); ent_score.insert(0, "2"); ent_score.pack(side=tk.LEFT, padx=5)
        def do_add():
            selected = tree.selection()
            if not selected:
                return messagebox.showwarning("提示", "请至少选择一道题目！", parent=win)
            try:
                score = float(ent_score.get().strip())
            except:
                return messagebox.showerror("错误", "分值必须是数字！", parent=win)
            conn = sqlite3.connect(DB_FILE); c = conn.cursor()
            added = 0
            for item in selected:
                qid = int(tree.item(item)['values'][0])
                c.execute("INSERT INTO paper_questions (paper_id, question_id, score) VALUES (?, ?, ?)",
                          (paper_id, qid, score))
                added += 1
            conn.commit(); conn.close()
            win.destroy()
            self.log(f"➕ 已向试卷【{q_type}】添加 {added} 道题目（每题{score}分）。")
            messagebox.showinfo("添加成功", f"已向【{q_type}】添加 {added} 道题目！")
            self._show_paper_view(paper_id)
            # 自动选中原题型
            for child in self.tree_paper_types.get_children():
                for sub in self.tree_paper_types.get_children(child):
                    if self.tree_paper_types.item(sub, 'values') and self.tree_paper_types.item(sub, 'values')[0] == q_type:
                        self.tree_paper_types.selection_set(sub)
                        self.tree_paper_types.see(sub)
                        break
            self.on_paper_type_selected(None)
        tk.Button(win, text="🚀 确定添加选中题目", font=("Arial", 11, "bold"), bg="#2980b9", fg="white", command=do_add).pack(pady=10)

    def action_paper_batch_stars(self):
        """批量修改试卷中题目的星级"""
        sel = self.grid_paper_q.selection()
        if not sel:
            return messagebox.showwarning("提示", "请选中要修改星级的题目！")
        target_star = int(self.cb_p_star.get())
        conn = sqlite3.connect(DB_FILE); c = conn.cursor()
        for item in sel:
            qid = self.grid_paper_q.item(item)['values'][0]
            c.execute("UPDATE questions SET kp_stars=? WHERE id=?", (target_star, qid))
        conn.commit(); conn.close()
        self.log(f"⭐ 已批量将 {len(sel)} 道题目的星级修改为 {target_star}。")
        self.on_paper_type_selected(None)

    def revoke_paper_usage(self):
        """撤销该试卷关联题目的年份标记（清空q_year）"""
        sel = self.tree_kb.selection()
        if not sel:
            return messagebox.showwarning("提示", "请先选中一份试卷！")
        item = self.tree_kb.item(sel[0])
        if 'paper' not in item.get("tags", []):
            return messagebox.showwarning("提示", "请选中「试卷库」下的试卷节点！")

        paper_id = item['values'][0]
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute("SELECT paper_name FROM papers WHERE id=?", (paper_id,))
        row = c.fetchone()
        if not row:
            conn.close()
            return messagebox.showerror("错误", "试卷记录不存在！")
        paper_name = row[0]

        # 读取关联的题目ID
        c.execute("SELECT question_id FROM paper_questions WHERE paper_id=?", (paper_id,))
        q_ids = [r[0] for r in c.fetchall()]
        conn.close()

        if not q_ids:
            return messagebox.showwarning("提示", "该试卷没有关联题目！")

        if not messagebox.askyesno("确认撤销", f"确定撤销试卷【{paper_name}】的 {len(q_ids)} 道题目的使用年份标记？"):
            return

        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        # 读取每道题保存时的原始年份
        c.execute("""
            SELECT q.id, pq.prev_year
            FROM questions q JOIN paper_questions pq ON q.id = pq.question_id
            WHERE pq.paper_id=?
        """, (paper_id,))
        rows = c.fetchall()
        updated = 0
        for qid, prev_year in rows:
            c.execute("UPDATE questions SET q_year=? WHERE id=?", (prev_year, qid))
            updated += 1
        conn.commit()
        conn.close()

        self.log(f"↩️ 试卷【{paper_name}】的使用标记已撤销，共{updated}道题目已恢复为原始年份。")
        messagebox.showinfo("撤销成功",
            f"试卷【{paper_name}】的使用标记已撤销！\n"
            f"共 {updated} 道题目已恢复为原始年份。")

    def open_settings(self):
        win = tk.Toplevel(self.root)
        win.title("系统底层设计与题型字典规范舱")
        win.geometry("620x560")
        win.grab_set()

        # 1. 大模型控制台组件架设
        f_api = ttk.LabelFrame(win, text="🔗 云端大语言模型 API 与平台路由调配 (密钥已安全加密遮蔽)")
        f_api.pack(fill=tk.X, padx=15, pady=8, ipady=5)

        # 平台选择
        tk.Label(f_api, text="大模型平台:").grid(row=0, column=0, padx=8, pady=5, sticky=tk.W)
        cb_platform = ttk.Combobox(f_api, values=list(AI_PLATFORM_MANIFEST.keys()), state="readonly", width=30)
        cb_platform.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        cb_platform.set(self.config.get("platform_name", "硅基流动 (SiliconFlow)"))

        # API URL (支持跟随平台自动变动，也支持手动微调)
        tk.Label(f_api, text="接口URL:").grid(row=1, column=0, padx=8, pady=5, sticky=tk.W)
        e_url = tk.Entry(f_api, width=45)
        e_url.grid(row=1, column=1, padx=5, pady=5, columnspan=2, sticky=tk.W)
        e_url.insert(0, self.config['api_url'])

        # 推荐模型选择
        tk.Label(f_api, text="目标大模型:").grid(row=2, column=0, padx=8, pady=5, sticky=tk.W)
        cb_model = ttk.Combobox(f_api, width=30)
        cb_model.grid(row=2, column=1, padx=5, pady=5, sticky=tk.W)

        # API Key (默认 show="*" 密文保障安全)
        tk.Label(f_api, text="API Key:").grid(row=3, column=0, padx=8, pady=5, sticky=tk.W)
        e_key = tk.Entry(f_api, width=32, show="*")
        e_key.grid(row=3, column=1, padx=5, pady=5, sticky=tk.W)
        e_key.insert(0, self.config['api_key'])

        # 眼睛按钮：控制密文明文切换
        self.key_visible = False
        def toggle_key_visibility():
            if self.key_visible:
                e_key.config(show="*")
                btn_eye.config(text="👁 显")
                self.key_visible = False
            else:
                e_key.config(show="")
                btn_eye.config(text="🔒 隐")
                self.key_visible = True

        btn_eye = tk.Button(f_api, text="👁 显", font=("Arial", 9), width=4, command=toggle_key_visibility)
        btn_eye.grid(row=3, column=1, padx=(245, 0), pady=5, sticky=tk.W)

        # 联动响应引擎
        def on_platform_changed(event):
            selected_plat = cb_platform.get()
            manifest = AI_PLATFORM_MANIFEST.get(selected_plat, {})
            # 自动刷新路由URL
            e_url.delete(0, tk.END)
            e_url.insert(0, manifest.get("url", ""))
            # 自动刷新绑定的推荐模型下拉菜单
            models = manifest.get("models", [])
            cb_model['values'] = models
            if models:
                cb_model.set(models[0])

        cb_platform.bind("<<ComboboxSelected>>", on_platform_changed)

        # 首次开启舱门时初始化模型下拉列表状态
        current_plat = cb_platform.get()
        if current_plat in AI_PLATFORM_MANIFEST:
            cb_model['values'] = AI_PLATFORM_MANIFEST[current_plat]["models"]
            cb_model.set(self.config.get("model_name", AI_PLATFORM_MANIFEST[current_plat]["models"][0]))

        # 2. 合规正统题型字典控制中枢 (保持原业务逻辑完整)
        f_dict = ttk.LabelFrame(win, text="⚙️ 顶层设计：合规正统题型字典控制中枢")
        f_dict.pack(fill=tk.BOTH, expand=True, padx=15, pady=8)
        
        f_dict_top = tk.Frame(f_dict); f_dict_top.pack(fill=tk.X, pady=5, padx=5)
        tk.Label(f_dict_top, text="追加合规新题型:").pack(side=tk.LEFT)
        ent_new_type = tk.Entry(f_dict_top, width=15); ent_new_type.pack(side=tk.LEFT, padx=5)
        
        def refresh_settings_type_list():
            for i in tree_types.get_children(): tree_types.delete(i)
            current_standards = get_standard_q_types()
            for idx, t in enumerate(current_standards): tree_types.insert("", tk.END, values=(idx+1, t))

        def add_type_to_dict():
            nt = ent_new_type.get().strip()
            if not nt: return messagebox.showwarning("警告", "题型名称不能为空！", parent=win)
            try:
                conn = sqlite3.connect(DB_FILE); c = conn.cursor()
                c.execute("INSERT INTO system_q_types (type_name) VALUES (?)", (nt,))
                conn.commit(); conn.close()
                ent_new_type.delete(0, tk.END); refresh_settings_type_list()
            except: messagebox.showwarning("提示", "该题型在系统标准字典中已存在！", parent=win)

        tk.Button(f_dict_top, text="➕ 录入字典", font=("Arial", 9, "bold"), fg="green", command=add_type_to_dict).pack(side=tk.LEFT, padx=2)

        def delete_type_from_dict():
            sel_t = tree_types.selection()
            if not sel_t: return messagebox.showwarning("提示", "请先在下方列表中选中需要注销的题型！", parent=win)
            t_name = tree_types.item(sel_t[0])['values'][1]
            if messagebox.askyesno("注销确认", f"确定在底层标准字典中将【{t_name}】剔除吗？", parent=win):
                conn = sqlite3.connect(DB_FILE); c = conn.cursor(); c.execute("DELETE FROM system_q_types WHERE type_name=?", (t_name,)); conn.commit(); conn.close(); refresh_settings_type_list()

        tk.Button(f_dict_top, text="🗑️ 从字典注销", font=("Arial", 9), fg="red", command=delete_type_from_dict).pack(side=tk.RIGHT, padx=5)

        tree_types = ttk.Treeview(f_dict, columns=("idx", "name"), show="headings", height=6)
        tree_types.heading("idx", text="序号"); tree_types.heading("name", text="系统合规正统题型名称")
        tree_types.column("idx", width=60, anchor="center"); tree_types.column("name", width=380, anchor="center")
        tree_types.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        refresh_settings_type_list()

        # 3. 闭舱持久化落盘
        def master_save():
            self.config['platform_name'] = cb_platform.get()
            self.config['api_url'] = e_url.get().strip()
            self.config['model_name'] = cb_model.get().strip()
            self.config['api_key'] = e_key.get().strip()
            
            save_config(self.config)
            self.log(f"✅ 大模型热变更落盘成功。当前激活：[{self.config['platform_name']}] -> {self.config['model_name']}")
            win.destroy()
            
        tk.Button(win, text="💾 永久保存底层配置并闭舱", font=("Arial", 11, "bold"), bg="#27ae60", fg="white", height=2, command=master_save).pack(fill=tk.X, padx=30, pady=12)

    # ====================================================================
    # 📖 业务流扩展：零基础看图说话级“系统操作指南说明书”舱（强固防白屏版）
    # ====================================================================
    def open_help_readme(self):
        win = tk.Toplevel(self.root)
        win.title("智能期末试卷管理系统 - 零基础操作指南")
        win.geometry("700x650")
        
        # 🚀 核心修复 1：强制让操作系统立即刷新该子窗口的几何布局与图形缓冲区
        win.update_idletasks()
        win.grab_set() # 锁定模态窗口
        
        # 顶层标题框样式
        f_top = tk.Frame(win, bg="#34495e", height=45)
        f_top.pack(fill=tk.X)
        tk.Label(f_top, text="📝 智能期末试卷系统零基础使用说明手册", fg="white", bg="#34495e", 
                 font=("Microsoft YaHei", 12, "bold")).pack(pady=10)
                 
        # 核心只读文本区配置
        f_text = tk.Frame(win)
        f_text.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)
        
        txt_readme = tk.Text(f_text, font=("Microsoft YaHei", 11), wrap=tk.WORD, bg="#f8f9f9", fg="#2c3e50")
        scroll = ttk.Scrollbar(f_text, orient="vertical", command=txt_readme.yview)
        txt_readme.configure(yscrollcommand=scroll.set)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        txt_readme.pack(fill=tk.BOTH, expand=True)
        
        # 🚀 核心修复 2：确保文本框此时处于可写状态（NORMAL）
        txt_readme.config(state=tk.NORMAL)
        txt_readme.delete(1.0, tk.END)
        
        # 植入详细说明书文本
        readme_content = """【📌 准备工作：系统的开启与底层配置】
1. 启动系统：双击运行本程序，系统会弹出一个深蓝色顶栏的现代化控制面板。
2. 配置大模型（首次使用必做）：
   - 点击右上角的“⚙️ 系统底层设置”按钮。
   - 在弹出的窗口中，选择您想使用的模型平台（如：硅基流动 或 DeepSeek官方平台）。
   - 在 API Key 输入框中填入您的密钥（输入时会自动显示为“*”号遮蔽保护）。
   - 点击“永久保存底层配置并闭舱”保存。

【第一步：创建课程与章节大纲预设】
1. 切换到第一个标签页“📥 AI双轨抽取入库”。
2. 点击目标课程旁的“+建课与大纲预设舱”按钮。
3. 在弹出的窗口中输入课程名称（如：中外教育史），并在下方文本框中每行输入一个章标题（如：第一章 原始时期的教育，换行，第二章 古代学校教育）。
4. 点击“确定创建”。

【第二步：导入教学材料，AI 自动清洗入库】
1. 在“目标课程”下拉菜单中，选择您刚刚创建的课程。
2. 点击“📁 选择本地资料 (支持批量多选)”按钮，在电脑中选中您的教材或真题文档。
3. 点击巨大的“🚀 启动 AI 扫描提取”大按钮。
4. 实时监控：下方监测台会自动刷新。“成功导入”表里会流式列出 AI 抓取出的标准题目；有重复的题目会出现在下方的“因查重未导入”表格中，系统已自动去重。
5. 快捷盲操：若对某道题不满意，直接在此行上【点击鼠标右键】➔【立即物理抹除】即可当场击落。

【第三步：题库结构微调与考点清洗（精细化教研）】
1. 前往“🛠️ 题库结构与权重总控”标签页。
2. 调整章节分值：双击某章节，可以在下方修改其“占比%”。系统组卷时将严格按比例分配分值。
3. 近义考点合并：若两个知识点高度同义（如科举制与科举制度），可按住 Ctrl 键将它们多选，然后【点击鼠标右键】➔【🔗 聚合合并多个知识点】进行一键全题库归并。
4. AI 智能补全答案：双击某道题加载到下方覆写台，【点击鼠标右键】➔【🤖 启动 RAG 讲义检索智能作答】，AI 将翻阅你刚才上传的所有资料给出权威参考答案。修改后点击“💾 永久保存题目修改”。

【第四步：一键规划蓝图，双轨细目寻优组卷】
1. 切换到“🖨️ 分层结构化组卷”标签页并选择课程。
2. 规划期末大纲蓝图：在结构清单中自由配置单选题、多选题、判断题或简答题的数量与分值。
3. 点击下方的“✨ 开启全自动双向细目寻优组卷”战略大按钮。
4. 查重拦截铁律：当前年份为 2026 年，系统在组卷抽样时将【绝对物理隔离屏蔽】2025 年与 2024 年考过的所有老题，确保近 3 年题目绝不重复！若隔离后数量不足将自动熔断并弹窗告警。
5. 数学内核会自动进行 400 次拼图寻优，同时并行产出考点分布完全错开、难度高度均衡的期末 A、B 双轨卷。

【第五步：试卷库管理与一键高保真导出 Word】
1. 切换到“📚 本地全栈知识库管理”标签页。
2. 展开左侧大树的“📝 试卷库”文件夹，选中刚刚生成的真题卷。
3. 导出 Word：在左侧树状图的试卷节点上【点击鼠标右键】➔【📤 使用并输出 Word 文档】。
4. 格式规范：系统会自动在根目录生成 .docx 文件。中文字体严格锁定为【正统宋体】，西文数字使用【Times New Roman】，且试卷与参考答案换页物理隔离，答案底部自动附加小字命题细目溯源报告，并自动对所有抽中题目打上使用年份标记。
5. 撤销与毁灭：若考卷今年决定不用，右键试卷节点选【↩️ 撤销该试卷题目年份标记】可无痕释放题库；点击红色的【🗑️ 物理彻底毁灭这套生成的试卷】可以把试卷残余和关联表从全库彻底抹除。
"""
        txt_readme.insert(tk.END, readme_content)
        
        # 🚀 核心修复 3：文字全部安全灌入后，最后一步锁死为只读（DISABLED）
        txt_readme.config(state=tk.DISABLED)
        
        # 底部关闭按钮
        tk.Button(win, text="💾 锁闭手册舱", font=("Microsoft YaHei", 10, "bold"), 
                  bg="#27ae60", fg="white", relief="flat", cursor="hand2", bd=0, padx=25, pady=4,
                  command=win.destroy).pack(pady=10)

if __name__ == "__main__":
    root = tk.Tk(); app = ExamBuilderApp(root); root.mainloop()