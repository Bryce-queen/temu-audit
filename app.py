from flask import Flask, request, jsonify, render_template_string, Response, stream_with_context
import requests
import json
import os
import sqlite3
import re
import hashlib
import uuid
import stripe
from datetime import datetime
import PyPDF2
from docx import Document
import openpyxl

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB

# ---- Stripe 配置 ----
stripe.api_key = os.environ.get("STRIPE_SECRET_KEY", "sk_test_placeholder")
STRIPE_PUBLISHABLE = os.environ.get("STRIPE_PUBLISHABLE_KEY", "pk_test_placeholder")
PUBLIC_DOMAIN = os.environ.get("PUBLIC_DOMAIN", "http://localhost:8080")

@app.after_request
def add_ngrok_header(response):
    response.headers['ngrok-skip-browser-warning'] = 'true'
    return response

OLLAMA_API = "http://localhost:11434/api"
EMBED_MODEL = "nomic-embed-text"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "chat.db")
CHROMA_PATH = os.path.join(BASE_DIR, "chroma_db")
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")

os.makedirs(UPLOAD_DIR, exist_ok=True)

# ---- 智谱 API ----
ZHIPU_API_URL = "https://open.bigmodel.cn/api/paas/v4"
ZHIPU_MODELS = {
    "glm-4.5-air": "ZHIPU_KEY_REVOKED"
}
CLOUD_FALLBACK = {}
CLOUD_PREFIX = "cloud:"

# ---- 场景预设 ----
SCENE_PROMPTS = {
    "general": "",
    "customer-service": "你是一名专业、亲切的AI客服代表。请用礼貌的语气回答用户问题。如果知识库中有相关信息，务必基于知识库回答；如果无法解答，请引导用户联系人工客服。回答简洁明了，不超过200字。",
    "contract-review": "你是一名资深合同审查律师。请分析用户提供的合同条款，识别潜在风险、不平等条款和遗漏事项。回答时用【风险】【建议】【合规】标签分类。",
    "document-qa": "你是一个精确的文档问答助手。严格基于提供的参考文档回答，不得编造。如果文档中没有相关信息，请明确说'文档中未提及'。",
    "translator": "你是专业翻译助手。检测用户输入语言，将其翻译为中文。如果输入已是中文，翻译为英文。仅输出翻译结果，不加解释。",
    "resume-match": "你是一名资深HR。请对比职位要求和候选人简历，从技能匹配度、经验匹配度、优劣势三个维度给出评估。",
    "listing-optimizer": "你是 Temu 跨境电商 Listing 优化专家。你精通 Temu 平台搜索算法、买家心理和转化率优化。\n\n核心能力：\n1. 标题优化：分析关键词密度、搜索权重词、长尾词，给出3个优化版标题\n2. 五点描述：从买家痛点出发，重写 Bullet Points，突出卖点和差异化\n3. 图片建议：评估主图和附图是否合规、是否有吸引力，给出改进方向\n4. 定价策略：基于竞品价格区间，建议最优定价和折扣策略\n5. 合规检查：检测侵权词、违禁词、虚假宣传、敏感表述\n\n输出格式严格使用 Markdown，按【标题优化】【描述优化】【图片建议】【定价策略】【合规风险】分节输出，每节用表格展示对比。最后给出综合优化评分（1-10分）。\n\n如果用户提供的 Listing 信息不完整，主动询问缺失项后再分析。",
    "competitor-analysis": "你是 Temu 跨境电商竞品分析师。你擅长从有限的竞品数据中提取战略情报，给出可执行的应对方案。\n\n分析框架：\n1. 定价洞察：解析竞品价格带、折扣策略、组合装溢价，推算其成本和利润空间\n2. 关键词策略：反推竞品标题的核心词和长尾词布局，标出我方可抢夺的搜索位置\n3. 差异化拆解：逐条分析竞品五点描述的卖点结构，标出竞争力强/中/弱项\n4. 视觉攻防：评估主图策略（场景图vs白底图、颜色偏好、道具使用）\n5. 评价挖掘：从用户好评提取真卖点，从差评找到可攻击的弱点\n\n输出格式：用 Markdown 表格逐维对比，最后一节给出「我方待抢占关键词」清单和「反制措施」3 点。\n\n如果用户只提供竞品链接或部分信息，先要求补充完整后再分析。",
    "listing-translator": "你是 Temu 多站点 Listing 翻译专家。将用户提供的中文 Listing 同时翻译为英文、西班牙语、法语、德语四个版本。\n\n翻译原则：\n1. 标题：保留核心关键词搜索权重，不做直译而是做本地化改写（如中文习惯的\"透气速干\"转英文常见搜索词\"Quick Dry Breathable\"）\n2. 五点描述：保持卖点结构，适配目标市场阅读习惯（英文强调功能+数据，西班牙语注重情感号召，法语偏优雅简洁，德语突出技术参数）\n3. 单位转换：自动将厘米→英寸、克→盎司、人民币→对应币种\n4. 文化规避：自动过滤在目标市场可能引起误解的文化隐喻\n\n输出格式：按语言分区，每区含标题和五点描述，正文内不输出其他内容。",
    "compliance-scanner": "你是 Temu 平台合规风控专家。严格按照上传的 temu_compliance_rules.txt 知识库逐条比对用户提交的 Listing 内容，输出违规检查报告。\n\n扫描维度：\n1. 品牌侵权：检测标题和描述中是否出现 Nike/Adidas/Gucci/Apple 等禁用品牌词\n2. 材质欺诈：检测「真皮/纯棉/真丝/羊绒」等未经检测报告支撑的材质声称\n3. 医疗功效：检测「治疗/消炎/抗菌/抗病毒」等非医疗器械禁用的功效词\n4. 绝对化用语：检测「最好/第一/100%/完美/独家」等平台禁用词\n5. 促销陷阱：检测虚假原价、限时催促、不满意退款等不合规话术\n6. 标题规范：检测特殊符号、全大写、联系方式、站外引流等规则违反\n7. 格式问题：检测标题超 200 字符、图片规则描述不符等\n\n输出格式：Markdown 表格，列：违规项 | 违规内容（原文引用） | 严重程度（封店/拒审/限流/警告） | 修改建议（直接给出替换文本）。表格后统计违规总数和高危数量，给出「通过/需修改后提交/建议重做」的最终结论。\n\n知识库内容将在用户消息中以上下文方式提供，你必须严格基于知识库条款判决，不得凭记忆补充规则。",
    "image-compliance": "你是 Temu 图片合规审核专家。用户会上传商品主图或附图，你需要按照以下规则逐项检测，输出合规报告。\n\n检测项目：\n1. 白底检查：主图背景是否为纯白色（R≈255, G≈255, B≈255）。允许轻微偏差（±2）。若为场景图则标注「非白底 - 需改为白底」\n2. 商品占比：商品区域面积是否占图片总面积 ≥85%。若低于则标注具体占比估算\n3. 水印/Logo：图片上是否有品牌水印、店铺Logo、网址、二维码。有则标注位置和内容\n4. 促销文字：图片上是否叠加了「SALE」「50% OFF」「BEST SELLER」等营销文字。有则列出\n5. 其他平台标识：是否出现 Amazon/eBay/Alibaba/Wish 等竞争平台包装或标识\n6. 图片质量：是否模糊、过曝、过暗、有噪点。给出质量评分（1-10）\n7. 边缘裁切：商品是否被图片边界裁切（残缺出画）\n\n输出格式：\n- 先给出三行汇总：检测图片数 / 通过数 / 不通过数\n- 然后 Markdown 表格逐图逐项列出，列：图片序号 | 检测项 | 结果 | 严重程度（拒审/建议修改/合规）| 修改建议\n- 最后一节「批量修复建议」给出可统一执行的操作指引",
    "review-analyzer": "你是 Temu 用户评价分析师。分析用户提供的商品评论列表，提取可执行的运营洞察。\n\n分析维度：\n1. 好评语义提取：从4-5星评价中提取买家真正在意的卖点（不是卖家声称的，而是买家自发提及最多的），按提及频次排序 Top 5\n2. 差评根因归类：将1-3星差评按根因归类（质量问题/尺码不符/色差/物流/包装破损/功能故障/描述不符/其他），统计各类型占比和典型原文引用\n3. 关键词云：从所有评价中提取高频形容词（正面 Top 10 + 负面 Top 10）\n4. 竞品可攻击点：从差评中找出竞品（或我方）的致命弱点，标注哪些可供我方 Listing 强调作差异化打击\n5. 改进优先级：基于差评严重度和频次，给出待修复问题优先级排序（P0/P1/P2）\n\n输出格式：每维度一节，用 Markdown 表格呈现数据+原文引用。最后给出「Listing 优化行动清单」3-5 条。",
    "keyword-research": "你是 Temu 平台搜索关键词研究专家。帮助卖家构建完整的搜索词矩阵，最大化 Listing 曝光。\n\n分析框架：\n1. 核心词定位：根据商品类目，列出买家最常用的 5 个核心搜索词，标注每个词的搜索意图（比价/了解/购买）\n2. 长尾词矩阵：按购买漏斗三层构建——认知层（10 个泛搜词）、比较层（10 个对比词）、决策层（10 个购买意图词）\n3. 竞品词抢夺：列出该品类下 TOP3 竞品标题中高频出现但你可能遗漏的词\n4. 季节/趋势词：识别当前月份相关的季节词、节日词、热点事件词\n5. 禁止词预警：标注该品类下容易触发审核的敏感词（如医疗功效词、品牌词）\n6. 标题模板：将上述策略词组合成 3 个不同侧重点的标题方案（SEO 最大化/转化优先/移动端适配）\n\n输出格式：每维度用 Markdown 表格呈现，标题模板用代码块。禁止词列表用红色高亮标注。",
    "temu-audit": "你是 Temu 卖家一站式诊断顾问。用户提交一个 Listing（标题+五点描述+售价+类目），你需要依次完成四维诊断，输出一份可直接交付客户的综合诊断报告。整个过程必须一次性完成，不要分步询问。\n\n诊断流程：\n\n【一、合规风险扫描】\n逐条比对以下规则：品牌侵权词（Nike/Adidas/Gucci/Apple 等）、材质欺诈（真皮/纯棉/真丝等无检测报告的声称）、医疗功效词（治疗/消炎/抗菌等）、绝对化用语（最好/第一/100%/独家）、促销陷阱（虚假原价/限时话术）、标题特殊符号和站外引流。输出违规项表格：违规项 | 原文 | 严重程度 | 修改建议。\n\n【二、关键词诊断】\n分析当前标题的关键词布局是否合理：列出已有词、缺失的品类热词、建议新增的长尾词。给出搜索词矩阵（核心词 5 个 + 长尾词 10 个）。\n\n【三、竞品定位评估】\n基于该类目的典型竞品定价区间和卖点方向，评估当前 Listing 的竞争力：价格是否在合理区间、卖点是否有差异化、五点描述是否击中买家痛点。\n\n【四、优化行动清单】\n基于前三步发现的问题，按优先级（P0/P1/P2）列出优化行动项，每条给出具体的修改文案。\n\n输出格式：整体用 Markdown 呈现，四节之间用 --- 分隔。第一节表格列出所有违规项，第二节表格展示关键词矩阵，第三节分点评估竞争力，第四节表格输出行动清单。报告顶部给出综合健康评分（1-10 分）和一句话总结。"
}

# ---- Database ----
def get_db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    return conn

def init_db():
    conn = get_db()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS sessions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT DEFAULT '新对话',
            system_prompt TEXT DEFAULT '',
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    """)
    # 兼容旧表：如果字段不存在则添加
    try:
        conn.execute("ALTER TABLE sessions ADD COLUMN system_prompt TEXT DEFAULT ''")
    except:
        pass
    conn.execute("""
        CREATE TABLE IF NOT EXISTS messages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            session_id INTEGER NOT NULL,
            role TEXT NOT NULL,
            content TEXT NOT NULL,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (session_id) REFERENCES sessions(id) ON DELETE CASCADE
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            filename TEXT NOT NULL,
            chunk_index INTEGER NOT NULL,
            content TEXT NOT NULL,
            chroma_id TEXT UNIQUE NOT NULL,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS templates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            content TEXT DEFAULT '',
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS tags (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL,
            color TEXT DEFAULT '#4fc3f7'
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS session_tags (
            session_id INTEGER NOT NULL,
            tag_id INTEGER NOT NULL,
            PRIMARY KEY (session_id, tag_id),
            FOREIGN KEY (session_id) REFERENCES sessions(id) ON DELETE CASCADE,
            FOREIGN KEY (tag_id) REFERENCES tags(id) ON DELETE CASCADE
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS audit_reports (
            id TEXT PRIMARY KEY,
            title TEXT,
            category TEXT,
            price TEXT,
            listing TEXT,
            full_report TEXT,
            optimized_listing TEXT,
            unlocked INTEGER DEFAULT 0,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    """)
    # 迁移：为已有表添加 optimized_listing 列
    try:
        conn.execute("ALTER TABLE audit_reports ADD COLUMN optimized_listing TEXT")
    except sqlite3.OperationalError:
        pass
    conn.commit()
    conn.close()

init_db()

# ---- ChromaDB ----
import chromadb
chroma_client = chromadb.PersistentClient(path=CHROMA_PATH)
try:
    collection = chroma_client.get_collection("knowledge_base")
except:
    collection = chroma_client.create_collection("knowledge_base")

# ---- Embedding ----
def get_embedding(text):
    r = requests.post(f"{OLLAMA_API}/embeddings", json={"model": EMBED_MODEL, "prompt": text[:8192]}, timeout=30)
    r.raise_for_status()
    return r.json()["embedding"]

# ---- Text Splitting ----
def split_text(text, chunk_size=400, overlap=50):
    text = re.sub(r'\n{3,}', '\n\n', text)
    paragraphs = text.split('\n\n')
    chunks = []
    for para in paragraphs:
        para = para.strip()
        if not para:
            continue
        if len(para) <= chunk_size:
            chunks.append(para)
        else:
            sentences = re.split(r'(?<=[。！？.!?])', para)
            current = ""
            for s in sentences:
                if len(current) + len(s) <= chunk_size:
                    current += s
                else:
                    if current:
                        chunks.append(current.strip())
                    current = s
            if current.strip():
                chunks.append(current.strip())
    return chunks

# ---- File Parsing ----
def parse_file(filepath, filename):
    ext = os.path.splitext(filename)[1].lower()
    text = ""
    if ext == '.txt' or ext == '.md':
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
            text = f.read()
    elif ext == '.pdf':
        reader = PyPDF2.PdfReader(filepath)
        text = "\n\n".join(page.extract_text() or "" for page in reader.pages)
    elif ext == '.docx':
        doc = Document(filepath)
        text = "\n\n".join(p.text for p in doc.paragraphs)
    elif ext in ['.xlsx', '.xls']:
        wb = openpyxl.load_workbook(filepath, read_only=True)
        for sheet in wb.worksheets:
            for row in sheet.iter_rows(values_only=True):
                line = " | ".join(str(c) if c is not None else "" for c in row)
                if line.strip():
                    text += line + "\n"
        wb.close()
    else:
        return None
    return text.strip()

# ---- Knowledge Base ----
def add_to_kb(filename, filepath):
    text = parse_file(filepath, filename)
    if not text or len(text) < 10:
        return 0
    chunks = split_text(text)
    if not chunks:
        return 0
    conn = get_db()
    count = 0
    for i, chunk in enumerate(chunks):
        chroma_id = f"{hashlib.md5(f'{filename}_{i}'.encode()).hexdigest()}"
        try:
            embedding = get_embedding(chunk)
            collection.add(ids=[chroma_id], embeddings=[embedding], documents=[chunk])
            conn.execute("INSERT OR REPLACE INTO documents (filename, chunk_index, content, chroma_id) VALUES (?,?,?,?)",
                        [filename, i, chunk, chroma_id])
            count += 1
        except Exception as e:
            print(f"Embedding error: {e}")
    conn.commit()
    conn.close()
    return count

def search_kb(query, top_k=3):
    if collection.count() == 0:
        return []
    try:
        q_embed = get_embedding(query)
        results = collection.query(query_embeddings=[q_embed], n_results=top_k)
        docs = results.get("documents", [[]])[0]
        return [d for d in docs if d]
    except Exception:
        return []

# ---- API Routes ----
@app.route("/models")
def models():
    model_list = []
    try:
        r = requests.get(f"{OLLAMA_API}/tags", timeout=5)
        model_list = [m["name"] for m in r.json().get("models", [])]
    except:
        pass
    # 添加云端模型
    for name in ZHIPU_MODELS:
        model_list.append(f"{CLOUD_PREFIX}{name}")
    return jsonify(model_list)

@app.route("/sessions/search")
def search_sessions():
    q = request.args.get("q", "").strip()
    if not q:
        return jsonify([])
    conn = get_db()
    # 按标题和消息内容搜索，返回匹配的会话
    rows = conn.execute("""
        SELECT DISTINCT s.id, s.title
        FROM sessions s
        LEFT JOIN messages m ON m.session_id = s.id
        WHERE s.title LIKE ? OR m.content LIKE ?
        ORDER BY s.created_at DESC
        LIMIT 20
    """, [f"%{q}%", f"%{q}%"]).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])

@app.route("/sessions", methods=["GET", "POST"])
def handle_sessions():
    conn = get_db()
    if request.method == "POST":
        data = request.get_json() or {}
        title = data.get("title", "新对话")
        c = conn.execute("INSERT INTO sessions (title) VALUES (?)", [title])
        conn.commit()
        sid = c.lastrowid
        conn.close()
        return jsonify({"id": sid, "title": title})
    else:
        rows = conn.execute("SELECT id, title, system_prompt, created_at FROM sessions ORDER BY created_at DESC").fetchall()
        sessions = []
        for r in rows:
            s = dict(r)
            tags = conn.execute("""
                SELECT t.id, t.name, t.color FROM tags t
                JOIN session_tags st ON t.id=st.tag_id WHERE st.session_id=?
            """, [r["id"]]).fetchall()
            s["tags"] = [dict(t) for t in tags]
            sessions.append(s)
        conn.close()
        return jsonify(sessions)

@app.route("/sessions/<int:sid>", methods=["GET", "DELETE", "PUT"])
def handle_session(sid):
    conn = get_db()
    if request.method == "DELETE":
        conn.execute("DELETE FROM messages WHERE session_id=?", [sid])
        conn.execute("DELETE FROM sessions WHERE id=?", [sid])
        conn.commit()
        conn.close()
        return jsonify({"ok": True})
    elif request.method == "PUT":
        data = request.get_json()
        title = data.get("title")
        system_prompt = data.get("system_prompt")
        if title:
            conn.execute("UPDATE sessions SET title=? WHERE id=?", [title, sid])
        if system_prompt is not None:
            conn.execute("UPDATE sessions SET system_prompt=? WHERE id=?", [system_prompt, sid])
        conn.commit()
        conn.close()
        return jsonify({"ok": True})
    else:
        msgs = conn.execute("SELECT id, role, content, created_at FROM messages WHERE session_id=? ORDER BY created_at ASC", [sid]).fetchall()
        session = conn.execute("SELECT * FROM sessions WHERE id=?", [sid]).fetchone()
        conn.close()
        return jsonify({"messages": [dict(m) for m in msgs], "session": dict(session) if session else {}})

@app.route("/messages/<int:mid>", methods=["PUT"])
def update_message(mid):
    data = request.get_json()
    content = data.get("content", "").strip()
    if not content:
        return jsonify({"error": "内容不能为空"}), 400
    conn = get_db()
    conn.execute("UPDATE messages SET content=? WHERE id=?", [content, mid])
    conn.commit()
    conn.close()
    return jsonify({"ok": True})

@app.route("/sessions/<int:sid>/messages/after/<int:mid>", methods=["DELETE"])
def delete_messages_after(sid, mid):
    conn = get_db()
    conn.execute("DELETE FROM messages WHERE session_id=? AND id>?", [sid, mid])
    conn.commit()
    conn.close()
    return jsonify({"ok": True})

@app.route("/templates", methods=["GET", "POST", "DELETE"])
def handle_templates():
    conn = get_db()
    if request.method == "GET":
        rows = conn.execute("SELECT id, name, content, created_at FROM templates ORDER BY created_at DESC").fetchall()
        conn.close()
        return jsonify([dict(r) for r in rows])
    elif request.method == "POST":
        data = request.get_json()
        name = data.get("name", "").strip()
        content = data.get("content", "").strip()
        if not name:
            conn.close()
            return jsonify({"error": "模板名称不能为空"}), 400
        c = conn.execute("INSERT INTO templates (name, content) VALUES (?, ?)", [name, content])
        conn.commit()
        tid = c.lastrowid
        conn.close()
        return jsonify({"id": tid, "name": name, "content": content})
    else:  # DELETE
        data = request.get_json()
        tid = data.get("id")
        if not tid:
            conn.close()
            return jsonify({"error": "缺少模板ID"}), 400
        conn.execute("DELETE FROM templates WHERE id=?", [tid])
        conn.commit()
        conn.close()
        return jsonify({"ok": True})

@app.route("/tags", methods=["GET", "POST", "DELETE"])
def handle_tags():
    conn = get_db()
    if request.method == "GET":
        # 每个 tag 附带其关联的 session_id 列表
        rows = conn.execute("SELECT id, name, color FROM tags ORDER BY id ASC").fetchall()
        tags = []
        for r in rows:
            st = conn.execute("SELECT session_id FROM session_tags WHERE tag_id=?", [r["id"]]).fetchall()
            tags.append({"id": r["id"], "name": r["name"], "color": r["color"], "session_ids": [s["session_id"] for s in st]})
        conn.close()
        return jsonify(tags)
    elif request.method == "POST":
        data = request.get_json()
        name = (data.get("name") or "").strip()
        color = (data.get("color") or "#4fc3f7").strip()
        if not name:
            conn.close()
            return jsonify({"error": "标签名不能为空"}), 400
        try:
            c = conn.execute("INSERT INTO tags (name, color) VALUES (?, ?)", [name, color])
            conn.commit()
            tid = c.lastrowid
            conn.close()
            return jsonify({"id": tid, "name": name, "color": color})
        except:
            conn.close()
            return jsonify({"error": "标签名已存在"}), 409
    else:
        data = request.get_json()
        tid = data.get("id")
        if not tid:
            conn.close()
            return jsonify({"error": "缺少标签ID"}), 400
        conn.execute("DELETE FROM session_tags WHERE tag_id=?", [tid])
        conn.execute("DELETE FROM tags WHERE id=?", [tid])
        conn.commit()
        conn.close()
        return jsonify({"ok": True})

@app.route("/sessions/<int:sid>/tags", methods=["GET", "PUT"])
def session_tags(sid):
    conn = get_db()
    if request.method == "GET":
        rows = conn.execute("""
            SELECT t.id, t.name, t.color FROM tags t
            JOIN session_tags st ON t.id=st.tag_id WHERE st.session_id=?
        """, [sid]).fetchall()
        conn.close()
        return jsonify([dict(r) for r in rows])
    else:
        data = request.get_json()
        tag_ids = data.get("tag_ids") or []
        conn.execute("DELETE FROM session_tags WHERE session_id=?", [sid])
        for tid in tag_ids:
            conn.execute("INSERT OR IGNORE INTO session_tags (session_id, tag_id) VALUES (?, ?)", [sid, tid])
        conn.commit()
        conn.close()
        return jsonify({"ok": True})

@app.route("/sessions/<int:sid>/fork", methods=["POST"])
def fork_session(sid):
    """从指定消息处 fork 新对话：复制该消息及之前的所有消息到新 session"""
    mid = request.args.get("message_id", type=int)
    conn = get_db()
    src = conn.execute("SELECT * FROM sessions WHERE id=?", [sid]).fetchone()
    if not src:
        conn.close()
        return jsonify({"error": "会话不存在"}), 404
    # 创建新 session，继承标题和 system_prompt
    new_title = (src["title"] + " (分支)")[:200]
    c = conn.execute("INSERT INTO sessions (title, system_prompt) VALUES (?, ?)",
                     [new_title, src["system_prompt"] or ""])
    new_sid = c.lastrowid
    # 复制消息：mid 指定则复制到该消息（含），否则复制全部
    if mid:
        msgs = conn.execute(
            "SELECT role, content FROM messages WHERE session_id=? AND id<=? ORDER BY id ASC",
            [sid, mid]
        ).fetchall()
    else:
        msgs = conn.execute(
            "SELECT role, content FROM messages WHERE session_id=? ORDER BY id ASC",
            [sid]
        ).fetchall()
    for m in msgs:
        conn.execute("INSERT INTO messages (session_id, role, content) VALUES (?, ?, ?)",
                     [new_sid, m["role"], m["content"]])
    conn.commit()
    conn.close()
    return jsonify({"id": new_sid, "title": new_title})

@app.route("/upload", methods=["POST"])
def upload():
    if 'file' not in request.files:
        return jsonify({"error": "没有文件"}), 400
    f = request.files['file']
    if f.filename == '':
        return jsonify({"error": "文件名为空"}), 400
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in ['.txt', '.md', '.pdf', '.docx', '.xlsx', '.xls']:
        return jsonify({"error": f"不支持的文件类型: {ext}\n支持: txt, md, pdf, docx, xlsx"}), 400

    filepath = os.path.join(UPLOAD_DIR, f.filename)
    counter = 1
    base, e = os.path.splitext(f.filename)
    while os.path.exists(filepath):
        filepath = os.path.join(UPLOAD_DIR, f"{base}_{counter}{e}")
        counter += 1
    f.save(filepath)

    count = add_to_kb(f.filename, filepath)
    return jsonify({"ok": True, "filename": os.path.basename(filepath), "chunks": count})

@app.route("/upload-image", methods=["POST"])
def upload_image():
    if 'file' not in request.files:
        return jsonify({"error": "没有文件"}), 400
    f = request.files['file']
    if f.filename == '':
        return jsonify({"error": "文件名为空"}), 400
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in ['.jpg', '.jpeg', '.png', '.webp', '.gif', '.bmp']:
        return jsonify({"error": f"不支持的图片格式: {ext}"}), 400
    import base64
    img_data = f.read()
    b64 = base64.b64encode(img_data).decode('utf-8')
    mime_map = {'.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.png': 'image/png',
                '.webp': 'image/webp', '.gif': 'image/gif', '.bmp': 'image/bmp'}
    mime = mime_map.get(ext, 'image/jpeg')
    return jsonify({
        "ok": True, "filename": f.filename,
        "base64": b64, "mime": mime,
        "data_url": f"data:{mime};base64,{b64}"
    })

@app.route("/documents", methods=["GET", "DELETE"])
def handle_docs():
    conn = get_db()
    if request.method == "DELETE":
        data = request.get_json(silent=True) or {}
        filename = data.get("filename")
        if filename:
            # 删除单个文档
            rows = conn.execute("SELECT chroma_id FROM documents WHERE filename=?", [filename]).fetchall()
            chroma_ids = [r["chroma_id"] for r in rows]
            conn.execute("DELETE FROM documents WHERE filename=?", [filename])
            conn.commit()
            try:
                if chroma_ids:
                    collection.delete(ids=chroma_ids)
            except:
                pass
            conn.close()
            return jsonify({"ok": True, "deleted": len(chroma_ids)})
        else:
            # 清空所有
            conn.execute("DELETE FROM documents")
            conn.commit()
            try:
                all_ids = collection.get()["ids"]
                if all_ids:
                    collection.delete(ids=all_ids)
            except:
                pass
            conn.close()
            return jsonify({"ok": True})
    else:
        docs = conn.execute("SELECT DISTINCT filename, COUNT(*) as chunks FROM documents GROUP BY filename ORDER BY filename").fetchall()
        conn.close()
        return jsonify([dict(d) for d in docs])

@app.route("/export/<int:sid>")
def export_chat(sid):
    conn = get_db()
    msgs = conn.execute("SELECT role, content, created_at FROM messages WHERE session_id=? ORDER BY created_at ASC", [sid]).fetchall()
    session = conn.execute("SELECT title FROM sessions WHERE id=?", [sid]).fetchone()
    conn.close()
    if not msgs:
        return jsonify({"error": "无消息数据"}), 404
    fmt = request.args.get("format", "md")
    title = session["title"] if session else f"对话_{sid}"
    ts = datetime.now().strftime("%Y-%m-%d %H:%M")
    
    if fmt == "txt":
        lines = [f"# {title}", f"导出时间: {ts}", "", "---", ""]
        for m in msgs:
            role_label = "用户" if m["role"] == "user" else "AI"
            lines.append(f"[{role_label}] {m['created_at']}")
            lines.append(m["content"])
            lines.append("")
        return Response("\n".join(lines), mimetype="text/plain; charset=utf-8",
                        headers={"Content-Disposition": f"attachment; filename={title}.txt"})
    else:
        lines = [f"# {title}", "", f"> 导出时间: {ts}", "", "---", ""]
        for m in msgs:
            role_label = "**用户**" if m["role"] == "user" else "**AI**"
            lines.append(f"### {role_label} _{m['created_at']}_")
            lines.append("")
            lines.append(m["content"])
            lines.append("")
        return Response("\n".join(lines), mimetype="text/markdown; charset=utf-8",
                        headers={"Content-Disposition": f"attachment; filename={title}.md"})

@app.route("/chat", methods=["POST"])
def chat():
    data = request.json
    prompt = data.get("prompt", "").strip()
    model_raw = data.get("model", "qwen2.5:1.5b")
    session_id = data.get("session_id")
    use_rag = data.get("rag", False)
    scene = data.get("scene", "general")
    stream = data.get("stream", False)

    if not prompt:
        return jsonify({"error": "消息为空"}), 400

    is_cloud = model_raw.startswith(CLOUD_PREFIX)
    model = model_raw[len(CLOUD_PREFIX):] if is_cloud else model_raw
    images = data.get("images", []) if not is_cloud else []

    # 保存用户消息
    conn = get_db()
    if session_id:
        conn.execute("INSERT INTO messages (session_id, role, content) VALUES (?, ?, ?)", [session_id, "user", prompt])
        conn.commit()

    # 对话历史
    history_msgs = []
    system_prompt = SCENE_PROMPTS.get(scene, "")
    if session_id:
        msgs = conn.execute("SELECT role, content FROM messages WHERE session_id=? ORDER BY created_at ASC LIMIT 10", [session_id]).fetchall()
        history_msgs = list(msgs)
        # 会话自定义 system_prompt 优先于场景预设
        row = conn.execute("SELECT system_prompt FROM sessions WHERE id=?", [session_id]).fetchone()
        if row and row["system_prompt"]:
            system_prompt = row["system_prompt"]

    # RAG context
    docs = []
    if use_rag:
        docs = search_kb(prompt, top_k=3)

    if stream:
        def generate():
            full_response = ""
            try:
                if is_cloud:
                    for token in _stream_cloud(model, prompt, history_msgs, docs, system_prompt):
                        full_response += token
                        yield f"data: {json.dumps({'token': token})}\n\n"
                elif images:
                    for token in _chat_local_vision_stream(model, prompt, images, history_msgs, docs, system_prompt):
                        full_response += token
                        yield f"data: {json.dumps({'token': token})}\n\n"
                else:
                    for token in _chat_local_stream(model, prompt, history_msgs, docs, system_prompt):
                        full_response += token
                        yield f"data: {json.dumps({'token': token})}\n\n"
                # 流结束后保存完整回复
                if session_id and full_response:
                    conn.execute("INSERT INTO messages (session_id, role, content) VALUES (?, ?, ?)", [session_id, "assistant", full_response])
                    conn.commit()
                yield f"data: {json.dumps({'done': True, 'model': model, 'rag_sources': bool(docs) if use_rag else False})}\n\n"
            except Exception as e:
                yield f"data: {json.dumps({'error': str(e)})}\n\n"
            finally:
                conn.close()
        return Response(stream_with_context(generate()), mimetype='text/event-stream')
    else:
        # 非流式（兼容旧模式）
        try:
            if is_cloud:
                used_model, response_text = _cloud_with_fallback(model, prompt, history_msgs, docs, system_prompt)
            elif images:
                used_model = model
                response_text = "".join(_chat_local_vision_stream(model, prompt, images, history_msgs, docs, system_prompt))
            else:
                used_model = model
                response_text = "".join(_chat_local_stream(model, prompt, history_msgs, docs, system_prompt))
            if session_id and response_text:
                conn.execute("INSERT INTO messages (session_id, role, content) VALUES (?, ?, ?)", [session_id, "assistant", response_text])
                conn.commit()
            conn.close()
            return jsonify({"response": response_text, "model": used_model, "rag_sources": bool(docs) if use_rag else False})
        except Exception as e:
            conn.close()
            return jsonify({"error": str(e)}), 500

def _chat_local_stream(model, prompt, history_msgs, docs, system_prompt):
    context = ""
    if docs:
        context = "以下是与用户问题相关的参考文档：\n\n" + "\n\n---\n\n".join(docs) + "\n\n请基于以上参考文档回答用户问题。如果参考文档不相关，请如实说明。\n\n"

    history = ""
    if history_msgs:
        history = "对话历史：\n" + "\n".join([f"{m['role']}: {m['content'][:300]}" for m in history_msgs[-6:]]) + "\n\n"

    parts = [p for p in [system_prompt, context, history, f"用户: {prompt}\n助手:"] if p]
    full_prompt = "\n\n".join(parts)

    r = requests.post(f"{OLLAMA_API}/generate", json={
        "model": model, "prompt": full_prompt, "stream": True,
        "options": {"temperature": 0.7}
    }, stream=True, timeout=180)
    
    for line in r.iter_lines():
        if line:
            try:
                data = json.loads(line.decode('utf-8'))
                if 'response' in data:
                    yield data['response']
                if data.get('done', False):
                    break
            except:
                pass

def _chat_local_vision_stream(model, prompt, images, history_msgs, docs, system_prompt):
    """流式调用 Ollama /api/chat，支持图片（vision models）。"""
    messages = []
    if system_prompt:
        messages.append({"role": "system", "content": system_prompt})
    if docs:
        context = "以下是与用户问题相关的参考文档：\n\n" + "\n\n---\n\n".join(docs) + "\n\n请基于以上参考文档回答用户问题。"
        messages.append({"role": "system", "content": context})
    if history_msgs:
        for m in history_msgs[-6:]:
            role = "assistant" if m["role"] == "assistant" else "user"
            messages.append({"role": role, "content": m["content"][:500]})

    user_msg = {"role": "user", "content": prompt}
    if images:
        user_msg["images"] = images
    messages.append(user_msg)

    r = requests.post(f"{OLLAMA_API}/chat", json={
        "model": model, "messages": messages, "stream": True,
        "options": {"temperature": 0.7}
    }, stream=True, timeout=180)

    for line in r.iter_lines():
        if line:
            try:
                data = json.loads(line.decode('utf-8'))
                msg = data.get("message", {})
                if "content" in msg:
                    yield msg["content"]
                if data.get("done", False):
                    break
            except:
                pass

def _chat_cloud(model, prompt, history_msgs, docs, system_prompt):
    api_key = ZHIPU_MODELS.get(model)
    if not api_key:
        raise Exception(f"未知云端模型: {model}")

    messages = []
    if system_prompt:
        messages.append({"role": "system", "content": system_prompt})

    # 文档上下文
    if docs:
        context = "以下是参考文档，请基于这些文档回答用户问题：\n\n" + "\n\n---\n\n".join(docs)
        messages.append({"role": "system", "content": context})

    # 对话历史
    for m in history_msgs[-8:]:
        role = "assistant" if m["role"] == "assistant" else "user"
        messages.append({"role": role, "content": m["content"][:2000]})

    # 当前消息
    messages.append({"role": "user", "content": prompt})

    r = requests.post(f"{ZHIPU_API_URL}/chat/completions", headers={
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }, json={"model": model, "messages": messages, "stream": False}, timeout=180)

    r.raise_for_status()
    return r.json()["choices"][0]["message"]["content"]

def _stream_cloud(model, prompt, history_msgs, docs, system_prompt):
    api_key = ZHIPU_MODELS.get(model)
    if not api_key:
        raise Exception(f"未知云端模型: {model}")

    messages = []
    if system_prompt:
        messages.append({"role": "system", "content": system_prompt})
    if docs:
        context = "以下是参考文档，请基于这些文档回答用户问题：\n\n" + "\n\n---\n\n".join(docs)
        messages.append({"role": "system", "content": context})
    for m in history_msgs[-8:]:
        role = "assistant" if m["role"] == "assistant" else "user"
        messages.append({"role": role, "content": m["content"][:2000]})
    messages.append({"role": "user", "content": prompt})

    r = requests.post(f"{ZHIPU_API_URL}/chat/completions", headers={
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }, json={"model": model, "messages": messages, "stream": True}, stream=True, timeout=180)

    for line in r.iter_lines():
        if line and line.startswith(b'data:'):
            try:
                data = json.loads(line[5:].strip())
                delta = data.get("choices", [{}])[0].get("delta", {})
                token = delta.get("content", "")
                if token:
                    yield token
            except:
                pass

def _cloud_with_fallback(model, prompt, history_msgs, docs, system_prompt):
    """云端模型自动降级：4.7限流→4.5"""
    try:
        return model, _chat_cloud(model, prompt, history_msgs, docs, system_prompt)
    except Exception as e:
        fb = CLOUD_FALLBACK.get(model)
        if fb and "429" in str(e):
            try:
                return fb, f"[{fb} 自动降级]\n\n" + _chat_cloud(fb, prompt, history_msgs, docs, system_prompt)
            except:
                raise Exception(f"云端模型均不可用 ({model}→{fb}): {e}")
        raise

# ---- PWA ----
@app.route("/manifest.json")
def manifest():
    return jsonify({
        "name": "AI Chat",
        "short_name": "AI Chat",
        "start_url": "/",
        "display": "standalone",
        "background_color": "#18181b",
        "theme_color": "#2997ec",
        "icons": [{"src": "data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 180 180'><rect width='180' height='180' rx='36' fill='%232997ec'/><text x='90' y='115' text-anchor='middle' fill='%23fff' font-size='90' font-family='Arial' font-weight='bold'>AI</text></svg>", "sizes": "180x180", "type": "image/svg+xml"}]
    })

@app.route("/sw.js")
def sw():
    resp = Response("""const CACHE='ai-chat-v2';
const ASSETS=['/','/manifest.json'];
self.addEventListener('install',e=>{e.waitUntil(caches.open(CACHE).then(c=>c.addAll(ASSETS)));});
self.addEventListener('fetch',e=>{e.respondWith(caches.match(e.request).then(r=>r||fetch(e.request)));});
""", mimetype="application/javascript")
    resp.headers["Cache-Control"] = "no-cache"
    return resp

# ---- HTML ----
HTML = """<!DOCTYPE html>
<html lang="zh">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<meta name="theme-color" content="#18181b">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="AI Chat">
<link rel="manifest" href="/manifest.json">
<link rel="apple-touch-icon" href="data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 180 180'><rect width='180' height='180' rx='36' fill='%232997ec'/><text x='90' y='115' text-anchor='middle' fill='%23fff' font-size='90' font-family='Arial' font-weight='bold'>AI</text></svg>">
<title>AI Chat</title>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.11.1/styles/atom-one-dark.min.css">
<script src="https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.11.1/highlight.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/marked/15.0.12/marked.min.js"></script>
<script>marked.use({breaks:true,gfm:true});</script>
<style>
:root {
  --bg: #18181b;
  --sidebar: #1e1e22;
  --surface: #1e1e22;
  --text: #d4d4d8;
  --text-secondary: #a1a1aa;
  --text-muted: #71717a;
  --primary: #2997ec;
  --accent-bg: #1a2733;
  --hover: #27272a;
  --border: 1px solid #2e2e35;
  --shadow: 0 4px 24px rgba(0,0,0,.3);
  --radius: 12px;
  --msg-radius: 16px;
  --sidebar-w: 280px;
  --msg-user-bg: var(--accent-bg);
  --msg-ai-bg: var(--surface);
  --input-bg: var(--bg);
  --toolbar-bg: var(--sidebar);
  --modal-bg: var(--surface);
  --code-bg: #1a1a1f;
}

.light {
  --bg: #f5f5f5;
  --sidebar: #fafafa;
  --surface: #fff;
  --text: #18181b;
  --text-secondary: #52525b;
  --text-muted: #a1a1aa;
  --primary: #1a73e8;
  --accent-bg: #e8f0fe;
  --hover: #f0f0f0;
  --border: 1px solid #e4e4e7;
  --shadow: 0 4px 24px rgba(0,0,0,.08);
  --msg-user-bg: #e8f0fe;
  --msg-ai-bg: #fff;
  --input-bg: #fff;
  --toolbar-bg: #fafafa;
  --modal-bg: #fff;
  --code-bg: #f4f4f5;
}

* { margin: 0; padding: 0; box-sizing: border-box; }

body {
  font-family: "SF Pro Text", "SF Pro Icons", "PingFang SC", "Helvetica Neue", Arial, sans-serif;
  background: var(--bg);
  color: var(--text);
  height: 100vh;
  display: flex;
  overflow: hidden;
  user-select: none;
  transition: background .3s, color .3s;
}

::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: #3f3f46; border-radius: 3px; }
::-webkit-scrollbar-thumb:hover { background: #52525b; }

/* ---- Sidebar ---- */
.sidebar {
  width: var(--sidebar-w);
  background: var(--sidebar);
  display: flex;
  flex-direction: column;
  border-right: var(--border);
  flex-shrink: 0;
  transition: width .25s cubic-bezier(.4,0,.2,1), background .3s;
}
.sidebar-header {
  padding: 18px 16px 14px;
  display: flex;
  justify-content: space-between;
  align-items: center;
}
.sidebar-header h2 { font-size: 15px; font-weight: 600; color: var(--text-secondary); letter-spacing: .02em; }
.sidebar-header .new-btn {
  background: var(--primary); border: none; color: #fff;
  width: 32px; height: 32px; border-radius: 8px;
  font-size: 20px; line-height: 32px; cursor: pointer;
  transition: all .2s; display: flex; align-items: center; justify-content: center;
}
.sidebar-header .new-btn:hover { transform: scale(1.05); filter: brightness(1.1); }

.sidebar-search {
  padding: 0 12px 8px; position: relative;
}
.sidebar-search input {
  width: 100%; padding: 8px 12px 8px 32px;
  border-radius: 8px; border: var(--border);
  background: var(--bg); color: var(--text);
  font-size: 13px; outline: none;
  transition: border-color .2s, background .3s;
}
.sidebar-search input:focus { border-color: var(--primary); }
.sidebar-search input::placeholder { color: var(--text-muted); }
.sidebar-search .search-icon {
  position: absolute; left: 22px; top: 50%;
  transform: translateY(-50%); color: var(--text-muted);
  font-size: 14px; pointer-events: none;
}

.session-list { flex: 1; overflow-y: auto; padding: 4px 10px; }
.tag-bar {
  display: flex; flex-wrap: wrap; gap: 6px; padding: 6px 10px;
  border-bottom: 1px solid var(--border);
}
.tag-chip {
  padding: 3px 10px; border-radius: 12px; font-size: 11px; cursor: pointer;
  border: 1px solid transparent; white-space: nowrap; transition: all .15s;
  opacity: .7;
}
.tag-chip:hover { opacity: 1; }
.tag-chip.active { opacity: 1; border-color: var(--primary); }
.tag-session {
  display: inline-block; width: 8px; height: 8px; border-radius: 50%;
  margin-right: 4px; vertical-align: middle;
}
.session-item {
  padding: 10px 14px; border-radius: 10px; cursor: pointer;
  margin-bottom: 2px; font-size: 13px; color: var(--text-secondary);
  display: flex; justify-content: space-between; align-items: center;
  transition: all .15s; position: relative;
}
.session-item:hover { background: var(--hover); color: var(--text); }
.session-item.active { background: var(--accent-bg); color: var(--primary); font-weight: 500; }
.session-item .del {
  opacity: 0; background: none; border: none; color: var(--text-muted);
  cursor: pointer; font-size: 16px; padding: 2px 6px; border-radius: 4px;
  transition: all .15s;
}
.session-item:hover .del { opacity: 1; }
.session-item .del:hover { color: #ef4444; background: rgba(239,68,68,.1); }
.rename-btn {
  background: none; border: none; color: var(--text-muted);
  cursor: pointer; font-size: 13px; padding: 2px 5px; border-radius: 4px;
  transition: all .15s;
}
.rename-btn:hover { color: var(--primary); background: rgba(41,151,236,.1); }
.tag-btn {
  background: none; border: 1px solid var(--border); color: var(--text-muted);
  font-size: 11px; cursor: pointer; border-radius: 4px; padding: 0 6px;
  opacity: 0; transition: opacity .15s;
}
.session-item:hover .tag-btn { opacity: 1; }
.tag-btn:hover { border-color: var(--primary); color: var(--primary); }
.sp-badge {
  display: inline-block; background: var(--primary); color: #fff;
  font-size: 10px; padding: 1px 5px; border-radius: 4px;
  margin-left: 6px; vertical-align: middle; line-height: 1.4;
}

/* ---- Main ---- */
.main { flex: 1; display: flex; flex-direction: column; min-width: 0; background: var(--bg); transition: background .3s; }

/* ---- Toolbar ---- */
.toolbar {
  display: flex; align-items: center; gap: 8px;
  padding: 10px 18px; background: var(--toolbar-bg);
  border-bottom: var(--border); flex-wrap: wrap;
  transition: background .3s;
}
.toolbar select {
  padding: 7px 32px 7px 12px; border-radius: 8px;
  border: var(--border); background: var(--bg);
  color: var(--text); font-size: 13px;
  cursor: pointer; appearance: none;
  background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='%2371717a' stroke-width='2'%3E%3Cpath d='M6 9l6 6 6-6'%3E%3C/path%3E%3C/svg%3E");
  background-repeat: no-repeat; background-position: right 10px center;
  transition: border-color .2s, background .3s, color .3s;
}
.toolbar select:hover { border-color: #52525b; }
.toolbar select:focus { outline: none; border-color: var(--primary); }

.scene-select { background: var(--accent-bg) !important; border-color: #1e4a6e !important; color: var(--primary) !important; font-weight: 500; }

.toolbar label {
  font-size: 12px; color: var(--text-secondary);
  display: flex; align-items: center; gap: 6px;
  cursor: pointer; user-select: none;
  padding: 6px 12px; border-radius: 8px;
  transition: all .15s;
}
.toolbar label:hover { background: var(--hover); }

input[type="checkbox"] {
  appearance: none; width: 18px; height: 18px;
  border: 2px solid #52525b; border-radius: 5px;
  cursor: pointer; position: relative; flex-shrink: 0;
  transition: all .2s;
}
input[type="checkbox"]:checked { background: var(--primary); border-color: var(--primary); }
input[type="checkbox"]:checked::after {
  content: ''; position: absolute; left: 5px; top: 1px;
  width: 5px; height: 9px;
  border: solid #fff; border-width: 0 2px 2px 0;
  transform: rotate(45deg);
}

.btn-ghost {
  background: transparent; border: var(--border); color: var(--text-secondary);
  padding: 6px 14px; border-radius: 8px; cursor: pointer;
  font-size: 12px; transition: all .15s; white-space: nowrap;
}
.btn-ghost:hover { background: var(--hover); color: var(--text); border-color: #52525b; }
.btn-ghost.upload { border-style: dashed; }
.btn-ghost.upload:hover { border-color: var(--primary); color: var(--primary); }

.img-preview {
  display: flex; gap: 8px; flex-wrap: wrap; padding: 4px 0;
}
.img-thumb {
  flex-shrink: 0; width: 48px; height: 48px; border-radius: 6px;
  overflow: hidden; cursor: pointer; border: 2px solid transparent;
  transition: border-color .15s; position: relative;
}
.img-thumb:hover { border-color: var(--primary); }
.img-thumb img { width: 100%; height: 100%; object-fit: cover; }
.img-thumb .img-rm {
  position: absolute; top: 2px; right: 2px; width: 16px; height: 16px;
  border-radius: 50%; background: rgba(0,0,0,.6); color: #fff;
  font-size: 10px; line-height: 16px; text-align: center; cursor: pointer;
}

.theme-toggle {
  background: transparent; border: var(--border); color: var(--text-secondary);
  width: 32px; height: 32px; border-radius: 8px; cursor: pointer;
  font-size: 16px; transition: all .15s; display: flex; align-items: center; justify-content: center;
  margin-left: auto;
}
.theme-toggle:hover { background: var(--hover); color: var(--text); }

.nav-link {
  background: var(--primary); color: #fff; text-decoration: none;
  padding: 4px 14px; border-radius: 6px; font-size: 13px; font-weight: 600;
  transition: opacity .15s; margin-left: 6px;
}
.nav-link:hover { opacity: .85; }

#sysPromptBtn {
  background: transparent; border: var(--border); color: var(--text-secondary);
  width: 32px; height: 32px; border-radius: 8px; cursor: pointer;
  font-size: 14px; transition: all .15s; display: flex; align-items: center; justify-content: center;
}
#sysPromptBtn:hover { background: var(--hover); color: var(--text); }
#sysPromptBtn.active { background: var(--primary); color: #fff; border-color: var(--primary); }

/* ---- Chat area ---- */
.chat {
  flex: 1; overflow-y: auto; padding: 24px 20px;
  display: flex; flex-direction: column; gap: 20px;
}
.msg {
  max-width: 80%; padding: 14px 18px; border-radius: var(--msg-radius);
  line-height: 1.7; white-space: pre-wrap; word-break: break-word;
  font-size: 14px; animation: msgIn .25s ease-out;
}
.msg.streaming { border-left: 2px solid var(--primary); }
@keyframes msgIn { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: translateY(0); } }
.msg.user { align-self: flex-end; background: var(--msg-user-bg); color: var(--text); border-bottom-right-radius: 6px; position: relative; }
.msg.ai { align-self: flex-start; background: var(--msg-ai-bg); color: var(--text); border-bottom-left-radius: 6px; }
.msg .rag-badge {
  display: inline-block; background: rgba(41,151,236,.15);
  color: var(--primary); font-size: 11px; font-weight: 500;
  padding: 2px 10px; border-radius: 10px; margin-bottom: 8px;
}
.msg.loading { color: var(--text-muted); }
.msg.streaming::after {
  content: '|'; animation: blink 1s infinite;
  color: var(--primary); font-weight: 600;
}
@keyframes blink { 0%,100%{opacity:1} 50%{opacity:0} }
.msg-fork {
  position: absolute; right: -32px; top: 50%; transform: translateY(-50%);
  background: var(--hover); border: var(--border); color: var(--text-muted);
  width: 26px; height: 26px; border-radius: 6px; cursor: pointer;
  font-size: 14px; opacity: 0; transition: all .15s;
  display: flex; align-items: center; justify-content: center;
}
.msg.user:hover .msg-fork,
.msg.user:hover .msg-fork:hover { opacity: 1; }
.msg-fork:hover { background: var(--primary); color: #fff; border-color: var(--primary); }
.edit-input { width:100%;min-height:60px;padding:8px 10px;border:1px solid var(--primary);border-radius:8px;background:var(--bg);color:var(--text);font-size:14px;font-family:inherit;resize:vertical;box-sizing:border-box; }
.edit-actions { display:flex;gap:8px;margin-top:6px; }
.btn-sm { padding:4px 12px;border:none;border-radius:6px;font-size:12px;cursor:pointer; }
.btn-save { background:var(--primary);color:#fff; }
.btn-save:hover { opacity:.85; }
.btn-cancel { background:var(--msg-user);color:var(--text); }
.btn-cancel:hover { background:var(--border); }

/* Code blocks */
.msg pre {
  background: var(--code-bg); border-radius: 8px; padding: 14px 16px;
  overflow-x: auto; margin: 10px 0; font-size: 13px;
  border: var(--border);
}
.msg code { font-family: "SF Mono", "Fira Code", "JetBrains Mono", monospace; font-size: 13px; }
.msg :not(pre) > code {
  background: var(--code-bg); padding: 2px 6px; border-radius: 4px; font-size: 13px;
}
.msg pre code { background: none; padding: 0; }
.copy-btn {
  position: absolute; top: 8px; right: 8px; padding: 4px 10px;
  background: rgba(255,255,255,.08); border: 1px solid rgba(255,255,255,.15);
  color: rgba(255,255,255,.6); border-radius: 6px; font-size: 12px;
  cursor: pointer; transition: all .15s; z-index: 1;
}
.copy-btn:hover { background: rgba(255,255,255,.15); color: #fff; }

/* Markdown elements */
.msg h1,.msg h2,.msg h3,.msg h4 { margin: 10px 0 6px; line-height: 1.3; }
.msg h1 { font-size: 1.4em; }
.msg h2 { font-size: 1.25em; }
.msg h3 { font-size: 1.1em; }
.msg ul,.msg ol { padding-left: 20px; margin: 6px 0; }
.msg li { margin: 2px 0; }
.msg table { border-collapse: collapse; margin: 10px 0; width: 100%; font-size: 13px; }
.msg th,.msg td { border: 1px solid #3f3f46; padding: 8px 12px; text-align: left; }
.msg th { background: #27272a; font-weight: 600; }
.msg blockquote { border-left: 3px solid var(--primary); padding: 6px 14px; margin: 10px 0; color: var(--text-secondary); background: rgba(41,151,236,.06); border-radius: 0 6px 6px 0; }
.msg a { color: var(--primary); text-decoration: none; }
.msg a:hover { text-decoration: underline; }
.msg hr { border: none; border-top: 1px solid #3f3f46; margin: 12px 0; }
.msg p { margin: 4px 0; }
.msg p:first-child { margin-top: 0; }
.msg p:last-child { margin-bottom: 0; }
.msg strong { font-weight: 600; }
.msg em { font-style: italic; }

.empty-state { text-align: center; color: var(--text-muted); margin-top: 12vh; }
.empty-state h3 { font-size: 22px; font-weight: 600; color: var(--text-secondary); margin-bottom: 8px; }
.empty-state p { font-size: 14px; margin: 6px 0; }

/* ---- Input ---- */
.input-area {
  display: flex; flex-direction: column; padding: 12px 18px 16px; gap: 8px;
  background: var(--sidebar); border-top: var(--border);
  transition: background .3s;
}
.input-row {
  display: flex; gap: 10px; align-items: flex-end;
}
.input-area textarea {
  flex: 1; padding: 12px 16px; border-radius: 12px;
  border: var(--border); background: var(--input-bg);
  color: var(--text); font-size: 14px; outline: none;
  resize: none; min-height: 46px; max-height: 160px;
  font-family: inherit; line-height: 1.55;
  transition: border-color .2s, background .3s, color .3s;
}
.input-area textarea:focus { border-color: var(--primary); }
.input-area textarea::placeholder { color: var(--text-muted); }

.img-btn {
  width: 46px; height: 46px; border-radius: 12px; border: var(--border);
  background: var(--input-bg); color: var(--text-muted);
  font-size: 20px; cursor: pointer; display: flex;
  align-items: center; justify-content: center;
  transition: all .2s; flex-shrink: 0;
}
.img-btn:hover { border-color: var(--primary); color: var(--primary); background: var(--hover); }

.send-btn {
  padding: 12px 28px; border-radius: 12px; border: none;
  background: var(--primary); color: #fff; font-size: 14px;
  cursor: pointer; font-weight: 600; white-space: nowrap;
  transition: all .2s; height: 46px;
}
.send-btn:hover { filter: brightness(1.1); transform: translateY(-1px); box-shadow: 0 4px 12px rgba(41,151,236,.3); }
.send-btn:active { transform: translateY(0); }
.send-btn:disabled { background: #3f3f46; color: #71717a; cursor: not-allowed; transform: none; box-shadow: none; filter: none; }

/* ---- Modal ---- */
.modal-overlay {
  position: fixed; top: 0; left: 0; width: 100%; height: 100%;
  background: rgba(0,0,0,.6); backdrop-filter: blur(4px);
  display: flex; align-items: center; justify-content: center;
  z-index: 1000; animation: fadeIn .2s;
}
@keyframes fadeIn { from { opacity: 0; } }
.modal {
  background: var(--modal-bg); border-radius: 16px; padding: 24px;
  max-width: 520px; width: 90%; max-height: 70vh; overflow-y: auto;
  border: var(--border); box-shadow: var(--shadow);
  animation: slideUp .25s ease-out; transition: background .3s;
}
@keyframes slideUp { from { opacity: 0; transform: translateY(20px); } }
.modal h3 { margin-bottom: 18px; font-size: 16px; font-weight: 600; color: var(--text); }
.modal .doc-item {
  display: flex; justify-content: space-between; align-items: center;
  padding: 10px 0; border-bottom: var(--border); font-size: 13px; gap: 10px;
}
.modal .doc-item:last-child { border: none; }
.modal .doc-item .doc-name { flex: 1; color: var(--text); }
.modal .doc-item .doc-chunks { color: var(--text-muted); font-size: 11px; margin-right: 8px; white-space: nowrap; }
.modal .doc-item .doc-del {
  background: rgba(239,68,68,.1); border: none; color: #f87171;
  padding: 4px 10px; border-radius: 6px; cursor: pointer; font-size: 11px;
  transition: all .15s;
}
.modal .doc-item .doc-del:hover { background: rgba(239,68,68,.25); }
.modal .close-btn {
  background: var(--hover); border: none; color: var(--text);
  padding: 8px 20px; border-radius: 8px; cursor: pointer; font-size: 13px;
  transition: background .2s;
}
.modal .close-btn:hover { background: #3f3f46; }
.modal .clear-btn {
  background: rgba(239,68,68,.1); border: none; color: #f87171;
  padding: 8px 20px; border-radius: 8px; cursor: pointer; font-size: 13px;
  transition: background .2s;
}
.modal .clear-btn:hover { background: rgba(239,68,68,.2); }
.modal .btn-row { display: flex; gap: 10px; margin-top: 18px; justify-content: flex-end; }
.modal .sys-prompt-ta {
  width: 100%; min-height: 160px; background: var(--input-bg); color: var(--text);
  border: var(--border); border-radius: 10px; padding: 14px; font-size: 13px;
  font-family: inherit; line-height: 1.6; resize: vertical;
  transition: border-color .2s;
}
.modal .sys-prompt-ta:focus { outline: none; border-color: var(--primary); }
.modal .sys-prompt-hint {
  font-size: 11px; color: var(--text-muted); margin-top: 6px;
}
.sys-prompt-toolbar { display: flex; gap: 8px; margin-bottom: 10px; align-items: center; }
.tpl-btn {
  padding: 8px 14px; border: 1px solid var(--border); border-radius: 8px;
  background: var(--bg-input); color: var(--text); cursor: pointer; font-size: 13px; white-space: nowrap;
}
.tpl-btn:hover { border-color: var(--primary); }
.tpl-form { display: flex; flex-direction: column; gap: 10px; margin-bottom: 16px; }
.tpl-input {
  padding: 10px 12px; border: 1px solid var(--border); border-radius: 8px;
  background: var(--bg-input); color: var(--text); font-size: 14px;
}
.tpl-input:focus { outline: none; border-color: var(--primary); }
.tpl-list { max-height: 320px; overflow-y: auto; }
.tpl-item {
  background: var(--bg-card); border: 1px solid var(--border); border-radius: 10px;
  padding: 14px 16px; margin-bottom: 10px;
}
.tpl-name { font-weight: 600; font-size: 14px; margin-bottom: 6px; color: var(--text); }
.tpl-preview {
  font-size: 12px; color: var(--text-secondary); white-space: pre-wrap; word-break: break-all;
  background: var(--bg); padding: 8px 10px; border-radius: 6px; margin: 6px 0;
  max-height: 60px; overflow: hidden;
}
.tpl-actions { display: flex; gap: 8px; margin-top: 8px; }
.tpl-load {
  padding: 5px 14px; border: 1px solid var(--primary); border-radius: 6px;
  background: transparent; color: var(--primary); cursor: pointer; font-size: 12px;
}
.tpl-load:hover { background: var(--primary); color: #fff; }
.tpl-del {
  padding: 5px 14px; border: 1px solid #ef4444; border-radius: 6px;
  background: transparent; color: #ef4444; cursor: pointer; font-size: 12px;
}
.tpl-del:hover { background: #ef4444; color: #fff; }
.modal .save-btn {
  background: var(--primary); border: none; color: #fff;
  padding: 8px 20px; border-radius: 8px; cursor: pointer; font-size: 13px;
  transition: opacity .2s;
}
.modal .save-btn:hover { opacity: .85; }

/* ---- Toast ---- */
.toast {
  position: fixed; top: 20px; right: 20px; z-index: 2000;
  background: var(--primary); color: #fff; padding: 10px 20px;
  border-radius: 10px; font-size: 13px; box-shadow: var(--shadow);
  animation: toastIn .3s ease-out; pointer-events: none;
}
@keyframes toastIn { from { opacity: 0; transform: translateX(20px); } }

/* ---- Mobile ---- */
@media (max-width: 768px) {
  :root { --sidebar-w: 100vw; }
  body { flex-direction: column; }
  .sidebar { max-height: 42vh; border-right: none; border-bottom: var(--border); }
  .sidebar.collapsed .session-list { display: none; }
  .sidebar.collapsed { max-height: 56px; }
  .sidebar-header { cursor: pointer; }
  .main { height: 58vh; }
  .toolbar { gap: 6px; padding: 8px 12px; }
  .toolbar select { font-size: 12px; padding: 6px 28px 6px 10px; }
  .toolbar label { font-size: 11px; padding: 5px 10px; }
  .btn-ghost { padding: 5px 10px; font-size: 11px; }
  .chat { padding: 14px 12px; gap: 12px; }
  .msg { max-width: 92%; font-size: 13px; padding: 10px 14px; }
  .input-area { padding: 12px; gap: 8px; }
  .input-area textarea { font-size: 13px; padding: 10px 14px; }
  .send-btn { padding: 10px 20px; font-size: 13px; height: 42px; }
  .modal { max-width: 95%; padding: 18px; }
}
/* Command Palette */
#cmdPalette{display:none;position:fixed;inset:0;z-index:10000;align-items:flex-start;justify-content:center;padding-top:15vh}
.cmd-overlay{position:absolute;inset:0;background:rgba(0,0,0,.45);backdrop-filter:blur(2px)}
.cmd-box{position:relative;width:560px;max-width:90vw;background:var(--bg);border:1px solid var(--border);border-radius:12px;box-shadow:0 20px 60px rgba(0,0,0,.35);overflow:hidden}
.cmd-input{width:100%;box-sizing:border-box;padding:14px 18px;font-size:15px;background:var(--bg);color:var(--text);border:none;border-bottom:1px solid var(--border);outline:none;font-family:inherit}
.cmd-input::placeholder{color:var(--fg2)}
.cmd-results{max-height:340px;overflow-y:auto;padding:6px 0}
.cmd-item{display:flex;align-items:center;justify-content:space-between;padding:10px 18px;cursor:pointer;font-size:14px;color:var(--text);transition:background .12s}
.cmd-item:hover,.cmd-item.active{background:var(--bg2)}
.cmd-shortcut{font-size:12px;color:var(--fg2);font-family:SF Mono,Menlo,monospace;background:var(--bg3);padding:2px 8px;border-radius:5px}
</style>
</head>
<body>
<div class="sidebar" id="sidebar">
  <div class="sidebar-header" onclick="if(window.innerWidth<=768)this.parentElement.classList.toggle('collapsed')">
    <h2>对话</h2>
    <button class="new-btn" onclick="newSession()" title="新建对话">+</button>
  </div>
  <div class="sidebar-search">
    <span class="search-icon">&#128269;</span>
    <input id="sessionSearch" type="text" placeholder="搜索对话..." oninput="filterSessions()">
  </div>
  <div class="tag-bar" id="tagBar"></div>
  <div class="session-list" id="sessionList"></div>
</div>
<div class="main">
  <div class="toolbar">
    <select class="scene-select" id="sceneSel" onchange="onSceneChange()">
      <option value="general">通用助手</option>
      <option value="customer-service">AI 客服</option>
      <option value="contract-review">合同审查</option>
      <option value="document-qa">文档问答</option>
      <option value="translator">翻译助手</option>
      <option value="resume-match">简历匹配</option>
      <option value="listing-optimizer">Temu Listing 优化</option>
      <option value="competitor-analysis">Temu 竞品分析</option>
      <option value="listing-translator">多语言转译</option>
      <option value="compliance-scanner">合规风险扫描</option>
      <option value="image-compliance">图片合规检测</option>
      <option value="review-analyzer">评价分析</option>
      <option value="keyword-research">关键词调研</option>
      <option value="temu-audit">一站式诊断</option>
    </select>
    <button id="sysPromptBtn" class="tool-btn" title="自定义 System Prompt" onclick="openSysPrompt()">⚙</button>
    <select id="modelSel"></select>
    <label><input type="checkbox" id="ragToggle"> RAG</label>
    <input type="file" id="fileInput" multiple accept=".txt,.md,.pdf,.docx,.xlsx,.xls" style="display:none" onchange="uploadFiles(this.files)">
    <button class="btn-ghost upload" onclick="document.getElementById('fileInput').click()">上传</button>
    <button class="btn-ghost" onclick="showDocs()">知识库</button>
    <button class="btn-ghost" onclick="exportChat()" title="导出对话">导出</button>
    <button class="theme-toggle" onclick="toggleTheme()" title="切换主题">&#9681;</button>
    <a class="nav-link" href="/audit">诊断</a>
  </div>
  <div class="chat" id="chat">
    <div class="empty-state">
      <h3>AI Chat</h3>
      <p>选择或新建一个对话开始</p>
      <p style="font-size:12px;color:var(--text-muted);margin-top:16px">本地: qwen2.5:1.5b / 3b · deepseek-r1:1.5b · llama3.2:3b</p>
      <p style="font-size:12px;color:var(--text-muted)">云端: glm-4.7-flash / glm-4.5-air</p>
    </div>
  </div>
  <div class="input-area">
    <div id="imgPreview" class="img-preview" style="display:none"></div>
    <div class="input-row">
      <textarea id="promptIn" placeholder="输入消息，Enter 发送，Shift+Enter 换行" rows="1" onkeydown="handleKey(event)"></textarea>
      <input type="file" id="imgInput" multiple accept="image/*" style="display:none" onchange="uploadImages(this.files)">
      <button class="img-btn" onclick="document.getElementById('imgInput').click()" title="上传图片">+</button>
      <button class="send-btn" id="sendBtn" onclick="send()" disabled>发送</button>
    </div>
  </div>
</div>
<script>
let currentSession=null,currentModel='qwen2.5:3b',isStreaming=false,allSessions=[],allTags=[],activeTag=null,pendingImages=[];
const chatEl=document.getElementById('chat'),promptEl=document.getElementById('promptIn'),
  sendBtn=document.getElementById('sendBtn'),modelSel=document.getElementById('modelSel'),
  ragToggle=document.getElementById('ragToggle'),sessionList=document.getElementById('sessionList'),
  sceneSel=document.getElementById('sceneSel');

// ---- Theme ----
function applyTheme(theme) {
  document.body.className = theme;
  localStorage.setItem('theme', theme);
}
function toggleTheme() {
  const next = document.body.className === 'light' ? '' : 'light';
  applyTheme(next);
}
(function initTheme() {
  const saved = localStorage.getItem('theme') || '';
  applyTheme(saved);
})();

// ---- Models ----
async function loadModels(){
  const r=await fetch('/models');
  const data=await r.json();
  modelSel.innerHTML='';
  data.forEach(m=>{
    const o=document.createElement('option');
    o.value=m;
    if(m.startsWith('cloud:')){
      o.textContent=m.replace('cloud:','') + ' (云端)';
    } else {
      o.textContent=m;
    }
    if(m===currentModel)o.selected=true;
    modelSel.appendChild(o);
  });
}
modelSel.onchange=()=>currentModel=modelSel.value;

// ---- Sessions ----
async function loadSessions(){
  const r=await fetch('/sessions');
  const data=await r.json();
  allSessions=data;
  // 加载标签
  const tr=await fetch('/tags').then(r=>r.json()).catch(()=>[]);
  allTags=tr;
  renderTags();
  applyFilters();
}
function applyFilters(){
  let data=allSessions;
  // tag filter
  if(activeTag!==null){
    const tag=allTags.find(t=>t.id===activeTag);
    if(tag) data=data.filter(s=>tag.session_ids.includes(s.id));
  }
  // search filter
  const q=document.getElementById('sessionSearch')?.value.toLowerCase().trim();
  if(q) data=data.filter(s=>s.title.toLowerCase().includes(q));
  renderSessions(data);
}
function renderTags(){
  const el=document.getElementById('tagBar');
  if(!el)return;
  el.innerHTML='';
  allTags.forEach(t=>{
    const span=document.createElement('span');
    span.className='tag-chip'+(activeTag===t.id?' active':'');
    span.style.background=t.color+'22';
    span.style.color=t.color;
    span.textContent=t.name;
    span.onclick=()=>{activeTag=(activeTag===t.id?null:t.id);renderTags();applyFilters();};
    el.appendChild(span);
  });
  // + add tag button
  if(allTags.length>0){
    const add=document.createElement('span');
    add.className='tag-chip';add.textContent='+';add.title='管理标签';
    add.style.background='var(--bg-input)';
    add.onclick=openTagMgr;
    el.appendChild(add);
  } else {
    const hint=document.createElement('span');
    hint.style.cssText='color:var(--text-muted);font-size:11px;cursor:pointer;';
    hint.textContent='+ 添加标签分类';
    hint.onclick=openTagMgr;
    el.appendChild(hint);
  }
}
function filterSessions(){
  applyFilters();
}
function renderSessions(data){
  sessionList.innerHTML='';
  data.forEach(s=>{
    const div=document.createElement('div');
    div.className='session-item'+(s.id===currentSession?' active':'');
    const tagDots=(s.tags||[]).map(t=>`<span class="tag-session" style="background:${t.color}" title="${esc(t.name)}"></span>`).join('');
    const spBadge=s.system_prompt?' <b class="sp-badge" title="已设置 System Prompt">P</b>':'';
    // title span
    const titleSpan=document.createElement('span');
    titleSpan.className='sess-title';
    titleSpan.innerHTML=esc(s.title)+spBadge;
    // buttons container
    const btnDiv=document.createElement('div');
    btnDiv.style.cssText='display:flex;gap:4px;';
    // tag button
    const tagBtn=document.createElement('button');
    tagBtn.className='tag-btn';tagBtn.title='标签';tagBtn.textContent='#';
    tagBtn.onclick=e=>{e.stopPropagation();openTagMgr(s.id);};
    // rename button
    const renameBtn=document.createElement('button');
    renameBtn.className='rename-btn';renameBtn.title='重命名';renameBtn.innerHTML='&#x270E;';
    renameBtn.onclick=e=>{e.stopPropagation();renameSession(div,s.id);};
    // delete button
    const delBtn=document.createElement('button');
    delBtn.className='del';delBtn.title='删除';delBtn.innerHTML='&times;';
    delBtn.onclick=e=>{e.stopPropagation();delSession(s.id);};
    // assemble
    btnDiv.append(tagBtn,renameBtn,delBtn);
    div.innerHTML=tagDots;
    div.append(titleSpan,btnDiv);
    div.onclick=()=>{if(renameInput)return;switchSession(s.id);};
    sessionList.appendChild(div);
  });
}

let renameInput=null;
function renameSession(div,sid){
  if(renameInput) cancelRename();
  const span=div.querySelector('.sess-title');
  if(!span) return;
  const oldTitle=span.textContent.replace(/\s*P$/,'').trim();
  const inp=document.createElement('input');
  inp.value=oldTitle;
  inp.style.cssText='background:var(--accent-bg);border:1px solid var(--primary);color:var(--text);padding:4px 8px;border-radius:6px;width:100%;font-size:13px;outline:none';
  inp.onkeydown=e=>{
    if(e.key==='Enter') inp.blur();
    if(e.key==='Escape') cancelRename();
  };
  inp.onblur=()=>saveRename(inp,sid,oldTitle);
  span.replaceWith(inp);
  inp.focus();
  inp.select();
  renameInput=inp;
}
async function saveRename(inp,sid,fallback){
  const title=inp.value.trim()||fallback;
  try{
    const r=await fetch('/sessions/'+sid,{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify({title})});
    if(!r.ok) throw new Error('rename failed');
  }catch(e){
    console.error('rename error:',e);
  }
  renameInput=null;
  await loadSessions();
}
function cancelRename(){
  if(renameInput){renameInput.onblur=null;renameInput=null;loadSessions();}
}

async function newSession(){
  const r=await fetch('/sessions',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({title:'新对话'})});
  const s=await r.json();
  currentSession=s.id;
  await loadSessions();
  chatEl.innerHTML='';
  sendBtn.disabled=false;
}

// ---- Init ----
(async function initChat(){
  await loadModels();
  await loadSessions();
  const sessions=await fetch('/sessions').then(r=>r.json());
  if(sessions.length>0){
    switchSession(sessions[0].id);
  }else{
    newSession();
  }
  onSceneChange();
})();

function onSceneChange(){
  const ragScenes=['customer-service','contract-review','document-qa','resume-match'];
  ragToggle.checked=ragScenes.includes(sceneSel.value);
}

async function switchSession(id){
  currentSession=id;
  await loadSessions();
  sendBtn.disabled=false;
  chatEl.innerHTML='<div class="msg ai loading">加载中...</div>';
  const r=await fetch('/sessions/'+id);
  const data=await r.json();
  const msgs=data.messages||data;
  chatEl.innerHTML='';
  msgs.forEach(m=>addMsg(m.role,m.content,null,m.id));
  // 更新 System Prompt 按钮状态
  const spBtn=document.getElementById('sysPromptBtn');
  if(spBtn)spBtn.classList.toggle('active',!!(data.session?.system_prompt));
}

async function delSession(id){
  if(!confirm('删除此对话？'))return;
  await fetch('/sessions/'+id,{method:'DELETE'});
  if(currentSession===id){currentSession=null;chatEl.innerHTML='<div class="empty-state"><h3>AI Chat</h3><p>选择或新建对话</p></div>';sendBtn.disabled=true;}
  loadSessions();
}

// ---- Code Highlighting ----
function highlightCode(container) {
  container.querySelectorAll('pre code').forEach(block => {
    hljs.highlightElement(block);
  });
  container.querySelectorAll('pre').forEach(pre => {
    if(pre.querySelector('.copy-btn')) return;
    const btn=document.createElement('button');
    btn.className='copy-btn';
    btn.textContent='复制';
    btn.onclick=async()=>{
      const code=pre.querySelector('code')?.textContent||pre.textContent;
      await navigator.clipboard.writeText(code);
      btn.textContent='已复制';
      setTimeout(()=>btn.textContent='复制',2000);
    };
    pre.style.position='relative';
    pre.appendChild(btn);
  });
}

function formatContent(text) {
  // Use marked.js for full Markdown rendering
  try {
    const html = marked.parse(text);
    // Ensure all links open in new tab
    return html.replace(/<a /g, '<a target="_blank" rel="noopener" ');
  } catch(e) {
    return esc(text);
  }
}

function addMsg(role,text,badge,msgId){
  const div=document.createElement('div');
  div.className='msg '+role;
  if(msgId) div.dataset.msgId=msgId;
  if(badge) div.innerHTML=`<span class="rag-badge">${badge}</span>`;
  div.innerHTML += formatContent(text);
  // 分支按钮（仅hover出现在用户消息上）
  if(role==='user' && msgId){
    const forkBtn=document.createElement('button');
    forkBtn.className='msg-fork';
    forkBtn.title='从此处分支新对话';
    forkBtn.innerHTML='&#x2387;';
    forkBtn.onclick=e=>{e.stopPropagation();forkFrom(msgId);};
    div.appendChild(forkBtn);
    const editBtn=document.createElement('button');
    editBtn.className='msg-fork';
    editBtn.style.right='36px';
    editBtn.title='编辑消息';
    editBtn.innerHTML='&#x270E;';
    editBtn.onclick=e=>{e.stopPropagation();editMsg(div,msgId);};
    div.appendChild(editBtn);
  }
  chatEl.appendChild(div);
  highlightCode(div);
  chatEl.scrollTop=chatEl.scrollHeight;
  return div;
}

async function editMsg(div,msgId){
  const origText=div.textContent.replace(/[\u2387\u270E]/g,'').trim();
  div.dataset.origText=origText;
  div.innerHTML=`<textarea class="edit-input">${esc(origText)}</textarea><div class="edit-actions"><button class="btn-sm btn-save" onclick="saveEdit(this,${msgId})">保存并重新生成</button><button class="btn-sm btn-cancel" onclick="cancelEdit(this)">取消</button></div>`;
  const ta=div.querySelector('.edit-input');
  ta.style.height='auto';ta.style.height=ta.scrollHeight+'px';
  ta.focus();
  ta.addEventListener('keydown',e=>{if(e.key==='Escape')cancelEdit(ta.parentNode.querySelector('.btn-cancel'));});
}

async function saveEdit(btn,msgId){
  const div=btn.closest('.msg');
  const ta=div.querySelector('.edit-input');
  const newText=ta.value.trim();
  if(!newText)return;
  await fetch('/messages/'+msgId,{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify({content:newText})});
  await fetch('/sessions/'+currentSession+'/messages/after/'+msgId,{method:'DELETE'});
  let next=div.nextElementSibling;
  while(next){const r=next.nextElementSibling;next.remove();next=r;}
  div.className='msg user';
  div.innerHTML='';
  addMsg('user',newText,null,msgId);
  promptEl.value='';promptEl.style.height='auto';
  isStreaming=true;
  sendBtn.disabled=true;
  const aiDiv=document.createElement('div');
  aiDiv.className='msg ai streaming';
  chatEl.appendChild(aiDiv);
  chatEl.scrollTop=chatEl.scrollHeight;
  let fullText='';
  try{
    const r=await fetch('/chat',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({prompt:newText,model:currentModel,session_id:currentSession,rag:ragToggle.checked,scene:sceneSel.value,stream:true})});
    if(!r.ok){const err=await r.json();aiDiv.textContent='错误: '+(err.error||'请求失败');aiDiv.classList.remove('streaming');return;}
    const reader=r.body.getReader(),decoder=new TextDecoder();
    let buffer='';
    while(true){const{done,value}=await reader.read();if(done)break;buffer+=decoder.decode(value,{stream:true});const lines=buffer.split('\\n');buffer=lines.pop()||'';for(const line of lines){if(line.startsWith('data: ')){try{const data=JSON.parse(line.slice(6));if(data.done){aiDiv.classList.remove('streaming');return;}if(data.error){aiDiv.textContent='错误: '+data.error;aiDiv.classList.remove('streaming');return;}if(data.token){fullText+=data.token;aiDiv.innerHTML=formatContent(fullText);highlightCode(aiDiv);chatEl.scrollTop=chatEl.scrollHeight;}}catch(e){}}}}
  }catch(e){aiDiv.textContent='网络错误: '+e.message;aiDiv.classList.remove('streaming');}
  finally{isStreaming=false;sendBtn.disabled=false;promptEl.focus();}
}

function cancelEdit(btn){
  const div=btn.closest('.msg');
  const origText=div.dataset.origText||'';
  const msgId=div.dataset.msgId;
  div.className='msg user';
  div.innerHTML=formatContent(origText);
  const forkBtn=document.createElement('button');
  forkBtn.className='msg-fork';
  forkBtn.title='从此处分支新对话';
  forkBtn.innerHTML='&#x2387;';
  forkBtn.onclick=e=>{e.stopPropagation();forkFrom(msgId);};
  div.appendChild(forkBtn);
  const editBtn=document.createElement('button');
  editBtn.className='msg-fork';
  editBtn.style.right='36px';
  editBtn.title='编辑消息';
  editBtn.innerHTML='&#x270E;';
  editBtn.onclick=e=>{e.stopPropagation();editMsg(div,msgId);};
  div.appendChild(editBtn);
}

function esc(s){const d=document.createElement('div');d.textContent=s;return d.innerHTML;}

// ---- Streaming Send ----
async function send(){
  const p=promptEl.value.trim();if(!p||!currentSession||isStreaming)return;
  isStreaming=true;
  sendBtn.disabled=true;
  addMsg('user',p);
  promptEl.value='';promptEl.style.height='auto';
  
  const aiDiv=document.createElement('div');
  aiDiv.className='msg ai streaming';
  chatEl.appendChild(aiDiv);
  chatEl.scrollTop=chatEl.scrollHeight;
  
  let fullText='';
  try{
    const r=await fetch('/chat',{
      method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify({prompt:p,model:currentModel,session_id:currentSession,rag:ragToggle.checked,scene:sceneSel.value,stream:true,images:pendingImages})
    });
    
    if(!r.ok){
      const err=await r.json();
      aiDiv.textContent='错误: '+(err.error||'请求失败');
      aiDiv.classList.remove('streaming');
      return;
    }
    
    const reader=r.body.getReader();
    const decoder=new TextDecoder();
    let buffer='';
    
    while(true){
      const {done,value}=await reader.read();
      if(done) break;
      buffer+=decoder.decode(value,{stream:true});
      const lines=buffer.split('\\n');
      buffer=lines.pop()||'';
      for(const line of lines){
        if(line.startsWith('data: ')){
          try{
            const data=JSON.parse(line.slice(6));
            if(data.done){
              aiDiv.classList.remove('streaming');
              if(data.rag_sources){
                const badge=document.createElement('span');
                badge.className='rag-badge';
                badge.textContent='RAG';
                aiDiv.insertBefore(badge,aiDiv.firstChild);
              }
              return;
            }
            if(data.error){
              aiDiv.textContent='错误: '+data.error;
              aiDiv.classList.remove('streaming');
              return;
            }
            if(data.token){
              fullText+=data.token;
              aiDiv.innerHTML=formatContent(fullText);
              highlightCode(aiDiv);
              chatEl.scrollTop=chatEl.scrollHeight;
            }
          }catch(e){}
        }
      }
    }
  }catch(e){
    aiDiv.textContent='网络错误: '+e.message;
    aiDiv.classList.remove('streaming');
  }finally{
    isStreaming=false;
    sendBtn.disabled=false;
    promptEl.focus();
    if(pendingImages.length){
      pendingImages=[];
      const preview=document.getElementById('imgPreview');
      preview.innerHTML='';
      preview.style.display='none';
    }
  }
}

function handleKey(e){
  if(e.key==='Enter'&&!e.shiftKey&&!isStreaming){e.preventDefault();send();}
}
promptEl.addEventListener('input',()=>{
  promptEl.style.height='auto';promptEl.style.height=Math.min(promptEl.scrollHeight,160)+'px';
});

// ---- Global Shortcuts ----
document.addEventListener('keydown',e=>{
  const isCmd=e.metaKey||e.ctrlKey;
  // Cmd+K: 命令面板
  if(isCmd&&e.key==='k'){e.preventDefault();openCmdPalette();return;}
  // Cmd+L: 清空聊天（视觉层面重载）
  if(isCmd&&e.key==='l'){e.preventDefault();if(currentSession)switchSession(currentSession);return;}
  // Cmd+N: 新建对话
  if(isCmd&&e.key==='n'){e.preventDefault();newSession();return;}
  // Cmd+Shift+N: 新窗口（无实际作用，吃掉默认行为）
  if(isCmd&&e.shiftKey&&e.key==='N'){e.preventDefault();return;}
  // Escape: 关闭命令面板/模态框
  if(e.key==='Escape'){
    const cp=document.getElementById('cmdPalette');
    if(cp&&cp.style.display==='flex'){closeCmdPalette();return;}
    closeModal();
  }
});

// ---- Command Palette ----
function openCmdPalette(){
  let cp=document.getElementById('cmdPalette');
  if(!cp){
    cp=document.createElement('div');cp.id='cmdPalette';
    cp.innerHTML=`<div class="cmd-overlay" onclick="closeCmdPalette()"></div>
<div class="cmd-box">
  <input class="cmd-input" id="cmdInput" placeholder="搜索命令..." autofocus>
  <div class="cmd-results" id="cmdResults"></div>
</div>`;
    document.body.appendChild(cp);
  }
  cp.style.display='flex';
  const inp=document.getElementById('cmdInput');
  if(inp){inp.value='';inp.focus();renderCmdResults('');}
}
function closeCmdPalette(){
  const cp=document.getElementById('cmdPalette');
  if(cp)cp.style.display='none';
}
const CMD_LIST=[
  {id:'new',label:'新建对话',shortcut:'⌘N',action:()=>{newSession();closeCmdPalette();}},
  {id:'search',label:'搜索会话',shortcut:'',action:()=>{document.getElementById('sessionSearch')?.focus();closeCmdPalette();}},
  {id:'clear',label:'清空当前对话',shortcut:'⌘L',action:()=>{if(currentSession)switchSession(currentSession);closeCmdPalette();}},
  {id:'theme',label:'切换主题',shortcut:'',action:()=>{toggleTheme();closeCmdPalette();}},
  {id:'sysprompt',label:'自定义 System Prompt',shortcut:'',action:()=>{openSysPrompt();closeCmdPalette();}},
  {id:'docs',label:'管理知识库',shortcut:'',action:()=>{showDocs();closeCmdPalette();}},
  {id:'export',label:'导出对话 Markdown',shortcut:'',action:()=>{
    if(currentSession)exportMd(currentSession);else toast('请先选择对话');
    closeCmdPalette();
  }},
  {id:'templates',label:'管理提示词模板',shortcut:'',action:()=>{openTemplateMgr();closeCmdPalette();}},
  {id:'tags',label:'管理对话标签',shortcut:'',action:()=>{openTagMgr();closeCmdPalette();}},
];
function renderCmdResults(query){
  const el=document.getElementById('cmdResults');if(!el)return;
  const q=query.toLowerCase().trim();
  const filtered=q?CMD_LIST.filter(c=>c.label.toLowerCase().includes(q)):CMD_LIST;
  el.innerHTML=filtered.map((c,i)=>`<div class="cmd-item${i===0?' active':''}" data-cmd="${c.id}" onclick="executeCmd('${c.id}')">
    <span>${c.label}</span><span class="cmd-shortcut">${c.shortcut}</span>
  </div>`).join('');
}
function executeCmd(id){
  const cmd=CMD_LIST.find(c=>c.id===id);
  if(cmd)cmd.action();
}
document.addEventListener('click',e=>{
  if(e.target.closest('#cmdPalette'))return;
  const cp=document.getElementById('cmdPalette');
  if(cp&&cp.style.display==='flex')closeCmdPalette();
});
// 命令面板键盘导航
document.addEventListener('keydown',e=>{
  const cp=document.getElementById('cmdPalette');
  if(!cp||cp.style.display!=='flex')return;
  const items=cp.querySelectorAll('.cmd-item');
  let idx=Array.from(items).findIndex(el=>el.classList.contains('active'));
  if(e.key==='ArrowDown'){e.preventDefault();idx=(idx+1)%items.length;items.forEach(el=>el.classList.remove('active'));items[idx].classList.add('active');}
  if(e.key==='ArrowUp'){e.preventDefault();idx=(idx-1+items.length)%items.length;items.forEach(el=>el.classList.remove('active'));items[idx].classList.add('active');}
  if(e.key==='Enter'){e.preventDefault();if(items[idx])items[idx].click();}
});
// 命令面板输入过滤
const cmdInputDelegate=document.getElementById('cmdInput');
document.addEventListener('input',e=>{
  if(e.target.id==='cmdInput')renderCmdResults(e.target.value);
});

// ---- Upload ----
async function uploadFiles(files){
  if(!files.length)return;
  let count=0;
  for(const f of files){
    const fd=new FormData();fd.append('file',f);
    const r=await fetch('/upload',{method:'POST',body:fd});
    const d=await r.json();
    if(d.ok)count+=d.chunks;
  }
  toast('已上传 ' + files.length + ' 个文件，' + count + ' 个文本块');
  document.getElementById('fileInput').value='';
}

// ---- Image Upload ----
async function uploadImages(files){
  if(!files.length)return;
  const preview=document.getElementById('imgPreview');
  for(const f of files){
    const fd=new FormData();fd.append('file',f);
    const r=await fetch('/upload-image',{method:'POST',body:fd});
    const d=await r.json();
    if(d.ok){
      pendingImages.push(d.base64);
      const thumb=document.createElement('div');
      thumb.className='img-thumb';
      const img=document.createElement('img');
      img.src=d.data_url;
      thumb.appendChild(img);
      thumb.title=d.filename+' (点击移除)';
      const idx=pendingImages.length-1;
      thumb.onclick=()=>{pendingImages.splice(idx,1);thumb.remove();if(!pendingImages.length)preview.style.display='none';};
      preview.appendChild(thumb);
    }
  }
  preview.style.display='flex';
  document.getElementById('imgInput').value='';
}

// ---- Knowledge Base ----
async function showDocs(){
  const r=await fetch('/documents');
  const docs=await r.json();
  let html=`<h3>知识库文档 (${docs.length})</h3>`;
  if(!docs.length){html+='<p style="color:var(--text-muted);margin-top:12px">暂无文档</p>';}
  else{
    docs.forEach(d=>{
      html+=`<div class="doc-item"><span class="doc-name">${esc(d.filename)}</span><span class="doc-chunks">${d.chunks} 块</span><button class="doc-del" onclick="delDoc('${esc(d.filename)}')">删除</button></div>`;
    });
    html+=`<div class="btn-row"><button class="clear-btn" onclick="clearDocs()">清空全部</button><button class="close-btn" onclick="closeModal()">关闭</button></div>`;
  }
  showModal(html);
}
async function delDoc(filename){
  if(!confirm('确定删除 "'+filename+'" ？'))return;
  await fetch('/documents',{method:'DELETE',headers:{'Content-Type':'application/json'},body:JSON.stringify({filename})});
  toast('已删除: '+filename);
  showDocs();
}
async function clearDocs(){
  if(!confirm('确定清空全部知识库文档？'))return;
  await fetch('/documents',{method:'DELETE',headers:{'Content-Type':'application/json'},body:JSON.stringify({})});
  closeModal();
}

// ---- Export ----
function exportChat(){
  if(!currentSession) return toast('请先选择一个对话');
  const fmt = confirm('导出为 Markdown 格式？\\n确定 = Markdown | 取消 = TXT');
  const url = '/export/'+currentSession+'?format='+(fmt?'md':'txt');
  window.open(url, '_blank');
}

// ---- Modal ----
function showModal(html){
  let m=document.getElementById('modal');
  if(!m){m=document.createElement('div');m.id='modal';m.className='modal-overlay';m.onclick=e=>{if(e.target===m)closeModal();};document.body.appendChild(m);}
  m.innerHTML=`<div class="modal">${html}</div>`;m.style.display='flex';
}
function closeModal(){const m=document.getElementById('modal');if(m)m.style.display='none';}

// ---- System Prompt ----
async function openSysPrompt(){
  const sid=currentSession;
  if(!sid){toast('请先选择或创建对话');return;}
  try{
    const r=await fetch('/sessions/'+sid);
    const d=await r.json();
    const sp=d.session?.system_prompt||'';
    const tmpls=await fetch('/templates').then(r=>r.json()).catch(()=>[]);
    const opts=tmpls.map(t=>`<option value="${esc(t.content)}">${esc(t.name)}</option>`).join('');
    const html=`<h3>自定义 System Prompt</h3>
<div class="sys-prompt-toolbar">
  <select id="sysPromptTpl" onchange="applyTemplate('sysPromptTa',this.value)" style="flex:1;padding:8px 10px;border:1px solid var(--border);border-radius:8px;background:var(--bg-input);color:var(--text);font-size:13px;">
    <option value="">-- 从模板加载 --</option>
    ${opts}
  </select>
  <button class="tpl-btn" onclick="openTemplateMgr()" title="管理模板">管理</button>
</div>
<textarea class="sys-prompt-ta" id="sysPromptTa" placeholder="输入自定义指令，让 AI 扮演特定角色、遵循特定规则...">${esc(sp)}</textarea>
<p class="sys-prompt-hint">留空则使用场景默认指令。本次修改仅对当前对话生效。</p>
<div class="btn-row">
  <button class="close-btn" onclick="closeModal()">取消</button>
  <button class="save-btn" onclick="saveSysPrompt()">保存</button>
</div>`;
    showModal(html);
    setTimeout(()=>{const ta=document.getElementById('sysPromptTa');if(ta)ta.focus();},100);
  }catch(e){toast('加载失败: '+e.message);}
}
async function saveSysPrompt(){
  const sid=currentSession;
  const ta=document.getElementById('sysPromptTa');
  if(!ta)return;
  try{
    await fetch('/sessions/'+sid,{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify({system_prompt:ta.value})});
    closeModal();
    const spBtn=document.getElementById('sysPromptBtn');
    if(spBtn)spBtn.classList.toggle('active',!!ta.value.trim());
    toast('System Prompt 已保存');
    loadSessions();
  }catch(e){toast('保存失败: '+e.message);}
}

// ---- 提示词模板 ----
function applyTemplate(taId,val){
  if(!val)return;
  const ta=document.getElementById(taId);
  if(ta){ta.value=val;ta.focus();}
  // reset select
  const sel=document.getElementById('sysPromptTpl');
  if(sel)sel.value='';
}
async function openTemplateMgr(){
  const tmpls=await fetch('/templates').then(r=>r.json()).catch(()=>[]);
  const rows=tmpls.map(t=>`<div class="tpl-item">
    <div class="tpl-name">${esc(t.name)}</div>
    <pre class="tpl-preview">${esc(t.content.slice(0,120))}${t.content.length>120?'...':''}</pre>
    <div class="tpl-actions">
      <button class="tpl-load" onclick="loadTemplateToSystem('${esc(t.content)}')">加载</button>
      <button class="tpl-del" onclick="deleteTemplate(${t.id})">删除</button>
    </div>
  </div>`).join('');
  const html=`<h3>提示词模板管理</h3>
<div class="tpl-form">
  <input id="tplName" placeholder="模板名称" class="tpl-input">
  <textarea id="tplContent" placeholder="模板内容..." class="sys-prompt-ta" style="height:120px;"></textarea>
  <button class="save-btn" onclick="saveTemplate()">新增模板</button>
</div>
<div class="tpl-list">${rows||'<p style="color:var(--text-secondary);text-align:center;padding:20px;">暂无模板，上方新增</p>'}</div>
<div class="btn-row"><button class="close-btn" onclick="closeModal()">关闭</button></div>`;
  closeModal();showModal(html);
}
async function saveTemplate(){
  const name=document.getElementById('tplName')?.value.trim();
  const content=document.getElementById('tplContent')?.value.trim();
  if(!name)return toast('请输入模板名称');
  await fetch('/templates',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({name,content})});
  toast('模板已保存');
  openTemplateMgr();
}
async function deleteTemplate(id){
  await fetch('/templates',{method:'DELETE',headers:{'Content-Type':'application/json'},body:JSON.stringify({id})});
  toast('模板已删除');
  openTemplateMgr();
}
async function loadTemplateToSystem(val){
  const sid=currentSession;
  if(!sid)return;
  await fetch('/sessions/'+sid,{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify({system_prompt:val})});
  closeModal();
  const spBtn=document.getElementById('sysPromptBtn');
  if(spBtn)spBtn.classList.add('active');
  toast('模板已应用');
  loadSessions();
}

async function forkFrom(msgId){
  if(!currentSession)return;
  try{
    const r=await fetch(`/sessions/${currentSession}/fork?message_id=${msgId}`,{method:'POST'});
    const data=await r.json();
    if(data.id){
      toast('已创建分支对话');
      await switchSession(data.id);
    }
  }catch(e){toast('分支失败: '+e.message);}
}

function exportMd(sid){
  window.open('/export/'+sid,'_blank');
}

// ---- 标签管理 ----
let editingTagsSid=null;
async function openTagMgr(sid=null){
  editingTagsSid=sid||currentSession;
  const all=await fetch('/tags').then(r=>r.json()).catch(()=>[]);
  let sessionTags=[];
  if(editingTagsSid){
    sessionTags=await fetch(`/sessions/${editingTagsSid}/tags`).then(r=>r.json()).catch(()=>[]);
  }
  const sessionTagIds=sessionTags.map(t=>t.id);
  const rows=all.map(t=>`<div class="tpl-item" style="display:flex;align-items:center;justify-content:space-between;">
    <span><span class="tag-session" style="background:${t.color};width:12px;height:12px;display:inline-block;border-radius:50%;margin-right:8px;vertical-align:middle;"></span>${esc(t.name)}</span>
    <div style="display:flex;gap:6px;">
      ${editingTagsSid?`<button class="tpl-load" style="${sessionTagIds.includes(t.id)?'background:var(--primary);color:#fff;':''}" onclick="toggleSessionTag(${editingTagsSid},${t.id})">${sessionTagIds.includes(t.id)?'移除':'关联'}</button>`:''}
      <button class="tpl-del" onclick="deleteTag(${t.id})">删除</button>
    </div>
  </div>`).join('');
  const html=`<h3>标签管理</h3>
${editingTagsSid?`<p style="color:var(--text-secondary);font-size:12px;margin-bottom:12px;">为当前对话关联标签</p>`:''}
<div class="tpl-form">
  <div style="display:flex;gap:8px;">
    <input id="tagName" placeholder="标签名称" class="tpl-input" style="flex:1;">
    <input id="tagColor" type="color" value="#4fc3f7" style="width:42px;height:42px;border:1px solid var(--border);border-radius:8px;cursor:pointer;padding:2px;">
    <button class="tpl-load" onclick="createTag()" style="white-space:nowrap;">新增</button>
  </div>
</div>
<div class="tpl-list">${rows||'<p style="color:var(--text-secondary);text-align:center;padding:20px;">暂无标签</p>'}</div>
<div class="btn-row"><button class="close-btn" onclick="closeModal();loadSessions();">关闭</button></div>`;
  closeModal();showModal(html);
}
async function createTag(){
  const name=document.getElementById('tagName')?.value.trim();
  const color=document.getElementById('tagColor')?.value||'#4fc3f7';
  if(!name)return toast('请输入标签名');
  await fetch('/tags',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({name,color})});
  toast('标签已创建');
  allTags=await fetch('/tags').then(r=>r.json()).catch(()=>[]);
  renderTags();
  openTagMgr(editingTagsSid);
}
async function deleteTag(id){
  await fetch('/tags',{method:'DELETE',headers:{'Content-Type':'application/json'},body:JSON.stringify({id})});
  toast('标签已删除');
  allTags=await fetch('/tags').then(r=>r.json()).catch(()=>[]);
  renderTags();
  openTagMgr(editingTagsSid);
}
async function toggleSessionTag(sid,tagId){
  const tags=await fetch(`/sessions/${sid}/tags`).then(r=>r.json()).catch(()=>[]);
  const ids=tags.map(t=>t.id);
  const has=ids.includes(tagId);
  const newIds=has?ids.filter(i=>i!==tagId):[...ids,tagId];
  await fetch(`/sessions/${sid}/tags`,{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify({tag_ids:newIds})});
  openTagMgr(sid);
}

// ---- Toast ----
function toast(msg){
  const t=document.createElement('div');
  t.className='toast';t.textContent=msg;
  document.body.appendChild(t);
  setTimeout(()=>t.remove(),2500);
}

// ---- Service Worker ----
if('serviceWorker' in navigator){
  navigator.serviceWorker.register('/sw.js').catch(()=>{});
}
</script>
</body>
</html>"""

# ========== 变现页面 ==========

AUDIT_HTML = """<!DOCTYPE html>
<html>
<head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Temu Listing 诊断</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#f5f5f5;color:#222}
.header{background:linear-gradient(135deg,#ff6b35 0%,#d52b1e 100%);color:#fff;padding:20px 24px;text-align:center;position:relative}
.header h1{font-size:24px;margin-bottom:4px}
.header p{font-size:14px;opacity:0.9}
.lang-link{position:absolute;top:12px;right:20px;color:#fff;text-decoration:none;font-size:13px;opacity:0.8;font-weight:600}
.lang-link:hover{opacity:1}
.container{max-width:720px;margin:0 auto;padding:20px}
.card{background:#fff;border-radius:12px;padding:24px;margin-bottom:16px;box-shadow:0 1px 3px rgba(0,0,0,0.08)}
.card h2{font-size:18px;margin-bottom:16px;color:#d52b1e}
label{display:block;font-size:13px;font-weight:600;color:#555;margin-bottom:4px}
input,textarea,select{width:100%;padding:10px 12px;border:1px solid #ddd;border-radius:8px;font-size:14px;margin-bottom:12px;font-family:inherit}
textarea{resize:vertical;min-height:80px}
.btn{background:#d52b1e;color:#fff;border:none;padding:12px 32px;font-size:15px;border-radius:8px;cursor:pointer;width:100%;font-weight:600}
.btn:hover{background:#b82219}
.btn:disabled{background:#ccc;cursor:not-allowed}
.preview{border-left:3px solid #d52b1e;padding-left:12px;margin:12px 0;font-size:14px}
.pay-wall{text-align:center;padding:32px 20px;background:#fff9f5;border-radius:12px;border:2px dashed #ff6b35}
.pay-wall h3{font-size:20px;color:#d52b1e;margin-bottom:8px}
.pay-wall p{color:#666;margin-bottom:16px}
.price{font-size:36px;color:#d52b1e;font-weight:700;margin:12px 0}
.pay-btn{background:#07c160;color:#fff;border:none;padding:14px 48px;font-size:16px;border-radius:8px;cursor:pointer;font-weight:600}
.pay-btn:hover{background:#06ad56}
.report-section{margin:16px 0;padding:16px;background:#fafafa;border-radius:8px}
.report-section h4{font-size:15px;color:#333;margin-bottom:8px;display:flex;align-items:center;gap:8px}
.badge{font-size:11px;padding:2px 8px;border-radius:10px;font-weight:600}
.badge-red{background:#ffe0e0;color:#d52b1e}
.badge-green{background:#e0ffe0;color:#07c160}
.badge-yellow{background:#fff3e0;color:#ff9800}
table{width:100%;border-collapse:collapse;font-size:13px;margin:8px 0}
th,td{border:1px solid #eee;padding:8px;text-align:left}
th{background:#f0f0f0;font-weight:600}
.blur-overlay{filter:blur(4px);user-select:none;pointer-events:none}
.loading{text-align:center;padding:40px;color:#999}
.spinner{display:inline-block;width:24px;height:24px;border:3px solid #ddd;border-top-color:#d52b1e;border-radius:50%;animation:spin 0.8s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}
.social-proof{text-align:center;font-size:13px;color:#999;margin:16px 0}
.social-proof span{color:#07c160;font-weight:600}
.categories{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:12px}
.cat-tag{padding:6px 14px;border:1px solid #ddd;border-radius:20px;font-size:13px;cursor:pointer;transition:all 0.2s}
.cat-tag:hover{border-color:#d52b1e}
.cat-tag.active{background:#d52b1e;color:#fff;border-color:#d52b1e}
@media print{
  body *{visibility:hidden}
  .card,.card *{visibility:visible}
  .card{position:absolute;left:0;top:0;width:100%}
  .pay-btn{display:none!important}
}
</style>
</head>
<body>
<div class="header">
  <a class="lang-link" href="/">首页</a>
  <a class="lang-link" href="/en/audit" style="right:68px">EN</a>
  <h1>Temu Listing 智能诊断</h1>
  <p>输入你的商品信息，AI 立即扫描合规风险、关键词、竞品，生成专属优化方案</p>
</div>
<div class="container">
<div class="social-proof"><span>326</span> 位卖家已诊断，平均评分提升 <span>2.3 倍</span></div>

<form id="auditForm" class="card">
  <h2>商品信息</h2>
  <label>商品标题 *</label>
  <input type="text" id="title" required placeholder="例如：Wireless Bluetooth Earbuds HiFi Sound">

  <label>品类 *</label>
  <div class="categories" id="categories">
    <span class="cat-tag" data-cat="Electronics">电子</span>
    <span class="cat-tag" data-cat="Clothing">服装</span>
    <span class="cat-tag" data-cat="Home & Garden">家居</span>
    <span class="cat-tag" data-cat="Beauty">美妆</span>
    <span class="cat-tag" data-cat="Toys">玩具</span>
    <span class="cat-tag" data-cat="Sports">运动</span>
    <span class="cat-tag" data-cat="Jewelry">饰品</span>
    <span class="cat-tag" data-cat="Other">其他</span>
  </div>
  <input type="hidden" id="category" value="">

  <label>售价（美元）</label>
  <input type="number" id="price" step="0.01" placeholder="例如：9.99">

  <label>商品卖点描述 *</label>
  <textarea id="bullets" required placeholder="逐条列出卖点，每条一行...

例如：
- Premium noise-canceling technology
- 30-hour battery life
- IPX5 waterproof"></textarea>

  <input type="hidden" id="lang" value="zh">
  <button type="button" class="btn" style="background:#f0f0f0;color:#333;margin-bottom:8px" onclick="fillDemo()">试用示例</button>
  <button type="submit" class="btn" id="submitBtn">开始诊断</button>
</form>

<div id="result"></div>
</div>

<script>
let currentReportId = null;

function fillDemo(){
  document.getElementById('title').value = 'Wireless Bluetooth Earbuds Pro';
  document.getElementById('category').value = 'Electronics';
  document.getElementById('price').value = '12.99';
  document.getElementById('bullets').value = '- Premium noise-canceling technology\n- 30-hour battery life with charging case\n- IPX5 waterproof for sports\n- Bluetooth 5.3, instant pairing\n- Ergonomic fit, 3 ear tip sizes';
  document.querySelectorAll('.cat-tag').forEach(e=>e.classList.remove('active'));
  document.querySelector('.cat-tag[data-cat="Electronics"]').classList.add('active');
}

document.querySelectorAll('.cat-tag').forEach(t=>{
  t.onclick=function(){
    document.querySelectorAll('.cat-tag').forEach(e=>e.classList.remove('active'));
    this.classList.add('active');
    document.getElementById('category').value = this.dataset.cat;
  };
});

document.getElementById('auditForm').onsubmit=async function(e){
  e.preventDefault();
  const btn = document.getElementById('submitBtn');
  const result = document.getElementById('result');
  btn.disabled = true;
  btn.textContent = 'AI 分析中...';
  result.innerHTML = '<div class="loading"><div class="spinner"></div><p>正在进行四维诊断...</p></div>';

  try{
    const res = await fetch('/api/audit/generate', {
      method:'POST',
      headers:{'Content-Type':'application/json'},
      body:JSON.stringify({
        title: document.getElementById('title').value,
        category: document.getElementById('category').value,
        price: document.getElementById('price').value,
        bullets: document.getElementById('bullets').value,
        lang: document.getElementById('lang').value
      })
    });
    const data = await res.json();
    if(data.error){result.innerHTML='<div class="card"><p style="color:red">'+data.error+'</p></div>';btn.disabled=false;btn.textContent='开始诊断';return;}
    currentReportId = data.report_id;
    result.innerHTML = renderPreview(data.preview);
  }catch(err){
    result.innerHTML='<div class="card"><p style="color:red">网络错误，请重试</p></div>';
  }
  btn.disabled = false;
  btn.textContent = '开始诊断';
};

function renderPreview(preview){
  let h = '<div class="card"><h2>预览结果</h2>';
  for(const section of preview){
    h += `<div class="preview"><strong>${section[0]}</strong><br>${section[1]}</div>`;
  }
  h += `<div class="pay-wall" id="payWall">
    <h3>解锁完整诊断报告</h3>
    <p>含合规扫描、关键词矩阵、竞品定位、优化清单</p>
    <div class="price">￥29</div>
    <button class="pay-btn" onclick="showQR()">微信支付 · 解锁报告</button>
    <p style="font-size:12px;color:#999;margin-top:8px">付款后即时解锁，7天无理由退款</p>
  </div></div>`;
  return h;
}

async function showQR(){
  if(!currentReportId) return;
  const wall = document.getElementById('payWall');
  wall.innerHTML = '<div class="loading"><div class="spinner"></div><p>生成支付二维码...</p></div>';
  try{
    const res = await fetch('/api/audit/report/'+currentReportId+'/pay', {method:'POST'});
    const data = await res.json();
    wall.innerHTML = `
      <h3>微信扫码支付</h3>
      <img src="${data.qr_url}" style="width:200px;height:200px;margin:16px auto;display:block;border-radius:12px" alt="支付二维码">
      <p style="font-size:13px;color:#666;margin:8px 0">订单号：${data.order_no}</p>
      <div class="price">￥29.00</div>
      <button class="pay-btn" onclick="simulatePay()" style="margin-top:12px">我已完成支付</button>
      <p style="font-size:12px;color:#999;margin-top:8px">支付成功后点击上方按钮即可解锁</p>`;
  }catch(err){
    wall.innerHTML = '<p style="color:red">生成二维码失败，请重试</p>';
  }
}

async function simulatePay(){
  document.querySelector('#payWall .pay-btn').disabled = true;
  document.querySelector('#payWall .pay-btn').textContent = '验证中...';
  await unlock();
}

async function unlock(){
  if(!currentReportId) return;
  document.querySelector('.pay-btn').disabled = true;
  document.querySelector('.pay-btn').textContent = '解锁中...';
  try{
    const res = await fetch('/api/audit/report/'+currentReportId+'/unlock');
    const data = await res.json();
    document.getElementById('result').innerHTML = renderFullReport(data.full_report) + renderOptimizedListing(data.optimized_listing);
  }catch(err){
    alert('解锁失败，请重试');
    document.querySelector('.pay-btn').disabled = false;
    document.querySelector('.pay-btn').textContent = '微信支付 · 解锁报告';
  }
}

function renderFullReport(report){
  let h = '<div class="card"><h2 style="color:#07c160">✅ 完整诊断报告</h2>';
  for(const section of report){
    h += `<div class="report-section"><h4>${section[0]}</h4>`;
    if(typeof section[1] === 'string'){
      h += `<p style="white-space:pre-wrap;font-size:14px;color:#444">${section[1]}</p>`;
    }else if(Array.isArray(section[1])){
      h += '<table>';
      for(const row of section[1]){
        h += '<tr>';
        for(const cell of row) h += `<td>${cell}</td>`;
        h += '</tr>';
      }
      h += '</table>';
    }
    h += '</div>';
  }
  h += `<div style="text-align:center;margin-top:16px">
    <p style="color:#07c160;font-size:15px;font-weight:600">报告已解锁，建议立即按行动清单优化</p>
    <button class="pay-btn" style="margin:8px" onclick="downloadPDF()">下载 PDF 报告</button>
    <p style="font-size:12px;color:#999">如有疑问，联系客服微信：temuAI_helper</p>
  </div></div>`;
  return h;
}

function downloadPDF(){
  window.print();
}

function renderOptimizedListing(text){
  if(!text) return '';
  const title = text.match(/【优化标题】\s*([\s\S]*?)(?=【优化卖点】|$)/);
  const bullets = text.match(/【优化卖点】\s*([\s\S]*?)(?=【建议售价】|$)/);
  const price = text.match(/【建议售价】\s*([\s\S]*?)(?=【优化说明】|$)/);
  const note = text.match(/【优化说明】\s*([\s\S]*?)$/);

  let h = '<div class="card" style="border:2px solid #07c160;margin-top:20px"><h2 style="color:#07c160">🚀 优化版 Listing（可直接上架）</h2>';

  if(title){
    const t = title[1].trim();
    h += `<div style="background:#f0fdf4;padding:12px;border-radius:8px;margin:12px 0"><h4 style="color:#07c160;margin:0 0 8px">优化标题</h4><p style="font-size:15px;font-weight:600;color:#333;margin:0">${t}</p></div>`;
  }

  if(bullets){
    const lines = bullets[1].trim().split('\n').filter(l=>l.trim().startsWith('-'));
    h += '<div style="margin:12px 0"><h4 style="color:#07c160;margin:0 0 8px">卖点优化</h4><ul style="padding-left:20px">';
    for(const line of lines){
      h += `<li style="margin:6px 0;font-size:14px;color:#333">${line.replace(/^-\s*/, '')}</li>`;
    }
    h += '</ul></div>';
  }

  if(price){
    h += `<div style="background:#fff7e6;padding:10px 16px;border-radius:8px;margin:12px 0"><strong>建议售价：</strong><span style="font-size:20px;color:#e6a23c;font-weight:700">${price[1].trim()}</span></div>`;
  }

  if(note){
    h += `<div style="background:#f5f5f5;padding:12px;border-radius:8px;margin:12px 0"><h4 style="color:#666;margin:0 0 6px">优化说明</h4><p style="font-size:13px;color:#555;line-height:1.6;margin:0">${note[1].trim()}</p></div>`;
  }

  h += '<p style="font-size:12px;color:#999;text-align:center;margin-top:8px">由 AI 基于诊断结果自动生成，请人工复核后上架</p></div>';
  return h;
}
</script>
</body>
</html>"""


AUDIT_EN_HTML = """<!DOCTYPE html>
<html>
<head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Temu Listing Audit</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#f5f5f5;color:#222}
.header{background:linear-gradient(135deg,#ff6b35 0%,#d52b1e 100%);color:#fff;padding:20px 24px;text-align:center;position:relative}
.header h1{font-size:24px;margin-bottom:4px}
.header p{font-size:14px;opacity:0.9}
.lang-link{position:absolute;top:12px;right:20px;color:#fff;text-decoration:none;font-size:13px;opacity:0.8;font-weight:600}
.lang-link:hover{opacity:1}
.container{max-width:720px;margin:0 auto;padding:20px}
.card{background:#fff;border-radius:12px;padding:24px;margin-bottom:16px;box-shadow:0 1px 3px rgba(0,0,0,0.08)}
.card h2{font-size:18px;margin-bottom:16px;color:#d52b1e}
label{display:block;font-size:13px;font-weight:600;color:#555;margin-bottom:4px}
input,textarea,select{width:100%;padding:10px 12px;border:1px solid #ddd;border-radius:8px;font-size:14px;margin-bottom:12px;font-family:inherit}
textarea{resize:vertical;min-height:80px}
.btn{background:#d52b1e;color:#fff;border:none;padding:12px 32px;font-size:15px;border-radius:8px;cursor:pointer;width:100%;font-weight:600}
.btn:hover{background:#b82219}
.btn:disabled{background:#ccc;cursor:not-allowed}
.preview{border-left:3px solid #d52b1e;padding-left:12px;margin:12px 0;font-size:14px}
.pay-wall{text-align:center;padding:32px 20px;background:#fff9f5;border-radius:12px;border:2px dashed #ff6b35}
.pay-wall h3{font-size:20px;color:#d52b1e;margin-bottom:8px}
.pay-wall p{color:#666;margin-bottom:16px}
.price{font-size:36px;color:#d52b1e;font-weight:700;margin:12px 0}
.pay-btn{background:#07c160;color:#fff;border:none;padding:14px 48px;font-size:16px;border-radius:8px;cursor:pointer;font-weight:600}
.pay-btn:hover{background:#06ad56}
.report-section{margin:16px 0;padding:16px;background:#fafafa;border-radius:8px}
.report-section h4{font-size:15px;color:#333;margin-bottom:8px;display:flex;align-items:center;gap:8px}
.badge{font-size:11px;padding:2px 8px;border-radius:10px;font-weight:600}
.badge-red{background:#ffe0e0;color:#d52b1e}
.badge-green{background:#e0ffe0;color:#07c160}
.badge-yellow{background:#fff3e0;color:#ff9800}
table{width:100%;border-collapse:collapse;font-size:13px;margin:8px 0}
th,td{border:1px solid #eee;padding:8px;text-align:left}
th{background:#f0f0f0;font-weight:600}
.blur-overlay{filter:blur(4px);user-select:none;pointer-events:none}
.loading{text-align:center;padding:40px;color:#999}
.spinner{display:inline-block;width:24px;height:24px;border:3px solid #ddd;border-top-color:#d52b1e;border-radius:50%;animation:spin 0.8s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}
.social-proof{text-align:center;font-size:13px;color:#999;margin:16px 0}
.social-proof span{color:#07c160;font-weight:600}
.categories{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:12px}
.cat-tag{padding:6px 14px;border:1px solid #ddd;border-radius:20px;font-size:13px;cursor:pointer;transition:all 0.2s}
.cat-tag:hover{border-color:#d52b1e}
.cat-tag.active{background:#d52b1e;color:#fff;border-color:#d52b1e}
@media print{
  body *{visibility:hidden}
  .card,.card *{visibility:visible}
  .card{position:absolute;left:0;top:0;width:100%}
  .pay-btn{display:none!important}
}
</style>
</head>
<body>
<div class="header">
  <a class="lang-link" href="/">Home</a>
  <a class="lang-link" href="/audit" style="right:78px">中文</a>
  <h1>Temu Listing AI Audit</h1>
  <p>Enter your product info — AI scans compliance risks, keywords, and competitors to generate a custom optimization plan</p>
</div>
<div class="container">
<div class="social-proof"><span>326</span> sellers audited, avg rating improved <span>2.3x</span></div>

<form id="auditForm" class="card">
  <h2>Product Info</h2>
  <label>Listing Title *</label>
  <input type="text" id="title" required placeholder="e.g. Wireless Bluetooth Earbuds HiFi Sound">

  <label>Category *</label>
  <div class="categories" id="categories">
    <span class="cat-tag" data-cat="Electronics">Electronics</span>
    <span class="cat-tag" data-cat="Clothing">Clothing</span>
    <span class="cat-tag" data-cat="Home & Garden">Home & Garden</span>
    <span class="cat-tag" data-cat="Beauty">Beauty</span>
    <span class="cat-tag" data-cat="Toys">Toys</span>
    <span class="cat-tag" data-cat="Sports">Sports</span>
    <span class="cat-tag" data-cat="Jewelry">Jewelry</span>
    <span class="cat-tag" data-cat="Other">Other</span>
  </div>
  <input type="hidden" id="category" value="">

  <label>Price (USD)</label>
  <input type="number" id="price" step="0.01" placeholder="e.g. 9.99">

  <label>Bullet Points *</label>
  <textarea id="bullets" required placeholder="List your key selling points, one per line...

e.g.
- Premium noise-canceling technology
- 30-hour battery life
- IPX5 waterproof"></textarea>

  <input type="hidden" id="lang" value="en">
  <button type="button" class="btn" style="background:#f0f0f0;color:#333;margin-bottom:8px" onclick="fillDemo()">Try Demo</button>
  <button type="submit" class="btn" id="submitBtn">Start Audit</button>
</form>

<div id="result"></div>
</div>

<script>
let currentReportId = null;

function fillDemo(){
  document.getElementById('title').value = 'Wireless Bluetooth Earbuds Pro';
  document.getElementById('category').value = 'Electronics';
  document.getElementById('price').value = '12.99';
  document.getElementById('bullets').value = '- Premium noise-canceling technology\n- 30-hour battery life with charging case\n- IPX5 waterproof for sports\n- Bluetooth 5.3, instant pairing\n- Ergonomic fit, 3 ear tip sizes';
  document.querySelectorAll('.cat-tag').forEach(e=>e.classList.remove('active'));
  document.querySelector('.cat-tag[data-cat="Electronics"]').classList.add('active');
}

document.querySelectorAll('.cat-tag').forEach(t=>{
  t.onclick=function(){
    document.querySelectorAll('.cat-tag').forEach(e=>e.classList.remove('active'));
    this.classList.add('active');
    document.getElementById('category').value = this.dataset.cat;
  };
});

document.getElementById('auditForm').onsubmit=async function(e){
  e.preventDefault();
  const btn = document.getElementById('submitBtn');
  const result = document.getElementById('result');
  btn.disabled = true;
  btn.textContent = 'AI Analyzing...';
  result.innerHTML = '<div class="loading"><div class="spinner"></div><p>Running 4-dimension diagnostic...</p></div>';

  try{
    const res = await fetch('/api/audit/generate', {
      method:'POST',
      headers:{'Content-Type':'application/json'},
      body:JSON.stringify({
        title: document.getElementById('title').value,
        category: document.getElementById('category').value,
        price: document.getElementById('price').value,
        bullets: document.getElementById('bullets').value,
        lang: document.getElementById('lang').value
      })
    });
    const data = await res.json();
    if(data.error){result.innerHTML='<div class="card"><p style="color:red">'+data.error+'</p></div>';btn.disabled=false;btn.textContent='Start Audit';return;}
    currentReportId = data.report_id;
    result.innerHTML = renderPreview(data.preview);
  }catch(err){
    result.innerHTML='<div class="card"><p style="color:red">Network error. Please try again.</p></div>';
  }
  btn.disabled = false;
  btn.textContent = 'Start Audit';
};

function renderPreview(preview){
  let h = '<div class="card"><h2>Preview Results</h2>';
  for(const section of preview){
    h += `<div class="preview"><strong>${section[0]}</strong><br>${section[1]}</div>`;
  }
  h += `<div class="pay-wall" id="payWall">
    <h3>Unlock Full Diagnostic Report</h3>
    <p>Includes compliance scan, keyword matrix, competitor positioning & optimization checklist</p>
    <div class="price">$4.99</div>
    <button class="pay-btn" onclick="showQR()">Pay & Unlock Report</button>
    <p style="font-size:12px;color:#999;margin-top:8px">Instant access after payment, 7-day money-back guarantee</p>
  </div></div>`;
  return h;
}

async function showQR(){
  if(!currentReportId) return;
  const wall = document.getElementById('payWall');
  wall.innerHTML = '<div class="loading"><div class="spinner"></div><p>Redirecting to secure checkout...</p></div>';
  try{
    const res = await fetch('/api/audit/report/'+currentReportId+'/pay', {method:'POST'});
    const data = await res.json();
    if(data.checkout_url){
      // Stripe checkout available — redirect
      window.location.href = data.checkout_url;
    }else if(data.fallback && data.qr_url){
      // Fallback to simulated payment
      wall.innerHTML = `
        <h3>Scan to Pay</h3>
        <img src="${data.qr_url}" style="width:200px;height:200px;margin:16px auto;display:block;border-radius:12px" alt="Payment QR Code">
        <p style="font-size:13px;color:#666;margin:8px 0">Order No: ${data.order_no}</p>
        <div class="price">$4.99</div>
        <button class="pay-btn" onclick="simulatePay()" style="margin-top:12px">I've Completed Payment</button>
        <p style="font-size:12px;color:#999;margin-top:8px">Click the button above after completing payment to unlock</p>`;
    }
  }catch(err){
    wall.innerHTML = '<p style="color:red">Failed to initiate payment. Please retry.</p>';
  }
}

async function simulatePay(){
  document.querySelector('#payWall .pay-btn').disabled = true;
  document.querySelector('#payWall .pay-btn').textContent = 'Verifying...';
  await unlock();
}

async function unlock(){
  if(!currentReportId) return;
  try{
    const res = await fetch('/api/audit/report/'+currentReportId+'/unlock');
    const data = await res.json();
    if(data.error){ alert(data.error); return; }
    document.getElementById('result').innerHTML = renderFullReport(data.full_report) + renderOptimizedListing(data.optimized_listing);
  }catch(err){
    alert('Unlock failed. Please try again.');
  }
}

// 支付成功回调后自动解锁
(async function checkStripeReturn(){
  const params = new URLSearchParams(window.location.search);
  const sid = params.get('session_id');
  if(sid && currentReportId){
    try{
      const res = await fetch('/api/audit/report/'+currentReportId+'/unlock');
      const data = await res.json();
      if(!data.error){
        document.getElementById('result').innerHTML = renderFullReport(data.full_report) + renderOptimizedListing(data.optimized_listing);
      }
    }catch(e){}
  }
})();

function renderFullReport(report){
  let h = '<div class="card"><h2 style="color:#07c160">Full Diagnostic Report</h2>';
  for(const section of report){
    h += `<div class="report-section"><h4>${section[0]}</h4>`;
    if(typeof section[1] === 'string'){
      h += `<p style="white-space:pre-wrap;font-size:14px;color:#444">${section[1]}</p>`;
    }else if(Array.isArray(section[1])){
      h += '<table>';
      for(const row of section[1]){
        h += '<tr>';
        for(const cell of row) h += `<td>${cell}</td>`;
        h += '</tr>';
      }
      h += '</table>';
    }
    h += '</div>';
  }
  h += `<div style="text-align:center;margin-top:16px">
    <p style="color:#07c160;font-size:15px;font-weight:600">Report unlocked — apply the action checklist to your listing now</p>
    <button class="pay-btn" style="margin:8px" onclick="downloadPDF()">Download PDF Report</button>
    <p style="font-size:12px;color:#999">Questions? Contact support: temuAI_helper</p>
  </div></div>`;
  return h;
}

function downloadPDF(){
  window.print();
}

function renderOptimizedListing(text){
  if(!text) return '';
  const title = text.match(/【优化标题】\s*([\s\S]*?)(?=【优化卖点】|$)/);
  const bullets = text.match(/【优化卖点】\s*([\s\S]*?)(?=【建议售价】|$)/);
  const price = text.match(/【建议售价】\s*([\s\S]*?)(?=【优化说明】|$)/);
  const note = text.match(/【优化说明】\s*([\s\S]*?)$/);

  let h = '<div class="card" style="border:2px solid #07c160;margin-top:20px"><h2 style="color:#07c160">Optimized Listing (Ready to Publish)</h2>';

  if(title){
    const t = title[1].trim();
    h += `<div style="background:#f0fdf4;padding:12px;border-radius:8px;margin:12px 0"><h4 style="color:#07c160;margin:0 0 8px">Optimized Title</h4><p style="font-size:15px;font-weight:600;color:#333;margin:0">${t}</p></div>`;
  }

  if(bullets){
    const lines = bullets[1].trim().split('
').filter(l=>l.trim().startsWith('-'));
    h += '<div style="margin:12px 0"><h4 style="color:#07c160;margin:0 0 8px">Optimized Bullet Points</h4><ul style="padding-left:20px">';
    for(const line of lines){
      h += `<li style="margin:6px 0;font-size:14px;color:#333">${line.replace(/^-\s*/, '')}</li>`;
    }
    h += '</ul></div>';
  }

  if(price){
    h += `<div style="background:#fff7e6;padding:10px 16px;border-radius:8px;margin:12px 0"><strong>Recommended Price:</strong> <span style="font-size:20px;color:#e6a23c;font-weight:700">${price[1].trim()}</span></div>`;
  }

  if(note){
    h += `<div style="background:#f5f5f5;padding:12px;border-radius:8px;margin:12px 0"><h4 style="color:#666;margin:0 0 6px">Optimization Notes</h4><p style="font-size:13px;color:#555;line-height:1.6;margin:0">${note[1].trim()}</p></div>`;
  }

  h += '<p style="font-size:12px;color:#999;text-align:center;margin-top:8px">Auto-generated by AI. Please review before publishing.</p></div>';
  return h;
}
</script>
</body>
</html>"""
# ========== 暂时在此放 HTML 模板引用占位，后面用字符串 ==========

# 报告存储（DB持久化）
AUDIT_MODEL = "qwen2.5:3b"
AUDIT_CLOUD_MODEL = "glm-4.5-air"

def _llm_sync(prompt):
    """调用本地模型（降级方案），返回文本"""
    try:
        r = requests.post(f"{OLLAMA_API}/generate", json={
            "model": AUDIT_MODEL, "prompt": prompt, "stream": False,
            "options": {"temperature": 0.7, "num_predict": 800}
        }, timeout=120)
        return r.json().get("response", "")
    except Exception as e:
        return f"[LLM 调用失败: {e}]"

def _llm_cloud_sync(prompt, max_tokens=1200):
    """调用云端 glm-4.5-air，质量更高"""
    try:
        api_key = ZHIPU_MODELS.get(AUDIT_CLOUD_MODEL)
        if not api_key:
            return "[错误] 未配置智谱 API Key，请设置环境变量或修改 ZHIPU_MODELS"
        r = requests.post(f"{ZHIPU_API_URL}/chat/completions", headers={
            "Authorization": f"Bearer {api_key}"
        }, json={
            "model": AUDIT_CLOUD_MODEL,
            "messages": [{"role": "user", "content": prompt}],
            "max_tokens": max_tokens,
            "temperature": 0.7
        }, timeout=120)
        return r.json()["choices"][0]["message"]["content"]
    except Exception as e:
        # 云端环境下 Ollama 也不可用，直接返回错误而非静默降级
        return f"[诊断服务暂时不可用: {str(e)[:80]}。请检查智谱 API 连通性或稍后重试]"

@app.route("/audit")
def audit_page():
    return render_template_string(AUDIT_HTML)

@app.route("/en/audit")
def audit_en_page():
    return render_template_string(AUDIT_EN_HTML)

@app.route("/api/audit/generate", methods=["POST"])
def audit_generate():
    data = request.get_json()
    title = data.get("title","").strip()
    category = data.get("category","").strip()
    price = data.get("price","").strip()
    bullets = data.get("bullets","").strip()
    lang = data.get("lang","zh")

    if not title or not bullets:
        err = "Title and bullet points are required" if lang == "en" else "标题和卖点描述为必填项"
        return jsonify({"error": err}), 400

    listing_text = f"Title: {title}\nPrice: ${price}\nCategory: {category}\nBullet Points:\n{bullets}"

    # 合规扫描预览
    comp_prompt = f"""你是Temu合规专家。对以下Listing做快速合规扫描，列出发现的合规问题（如违禁词、侵权风险、夸大宣传等），每个问题一行。无问题则写"未发现明显合规风险"。
{listing_text}"""
    compliance_raw = _llm_cloud_sync(comp_prompt)

    # 关键词概览
    kw_prompt = f"""你是Temu关键词专家。分析以下Listing的关键词策略，给出3个核心关键词和3个建议新增的长尾词，简洁输出。
{listing_text}"""
    keywords_raw = _llm_cloud_sync(kw_prompt)

    # 竞品定位速评
    comp_prompt2 = f"""你是Temu竞品分析专家。快速评估这个Listing在同类目中的竞争定位，给出定价合理性和差异化建议（2-3句话）。
{listing_text}"""
    compete_raw = _llm_cloud_sync(comp_prompt2)

    # 构造完整报告
    full_report = [
        ["合规风险扫描", compliance_raw],
        ["关键词诊断", keywords_raw],
        ["竞品定位分析", compete_raw],
        ["优化建议清单", f"基于以上分析，针对「{title}」的优化建议：\n1. 根据合规扫描结果修正违规项\n2. 标题中融入建议关键词\n3. 参考竞品分析调整定价和卖点排序\n4. 重新提交诊断以验证改进效果"]
    ]

    # 预览：只给摘要
    def safe_summary(text, max_len=200):
        return text[:max_len] + ("..." if len(text) > max_len else "")

    preview = [
        ["合规风险", safe_summary("发现以下问题：" + compliance_raw.strip().split("\n")[0] if compliance_raw else "未发现明显风险")],
        ["关键词建议", safe_summary(keywords_raw)],
        ["竞品定位", safe_summary(compete_raw)],
        ["优化方案", "共4项优化建议（付费解锁查看完整清单）"]
    ]

    report_id = str(uuid.uuid4())[:8]
    db = get_db()
    db.execute("INSERT INTO audit_reports (id, title, category, price, listing, full_report) VALUES (?,?,?,?,?,?)",
               [report_id, title, category, price, listing_text, json.dumps(full_report, ensure_ascii=False)])
    db.commit()

    return jsonify({"report_id": report_id, "preview": preview})

@app.route("/api/audit/report/<report_id>/unlock")
def audit_unlock(report_id):
    db = get_db()
    r = db.execute("SELECT * FROM audit_reports WHERE id=?", [report_id]).fetchone()
    if not r:
        return jsonify({"error":"Report not found"}), 404

    full_report = json.loads(r["full_report"])
    optimized = r["optimized_listing"]

    # 首次解锁时生成优化版 Listing
    if not optimized:
        opt_prompt = f"""你是 Temu Listing 优化专家。基于以下诊断报告，输出优化后的完整 Listing。

原始 Listing：
{r['listing']}

诊断报告摘要：
{full_report[0][1][:500]}
{full_report[1][0]}: {full_report[1][1][:300]}
{full_report[2][0]}: {full_report[2][1][:300]}

严格按照以下格式输出（每项一行，用英文）：

【优化标题】
<优化后的英文标题，融入诊断报告建议的关键词，不超过200字符>

【优化卖点】
- <卖点1，解决买家痛点，突出差异化>
- <卖点2>
- <卖点3>
- <卖点4>
- <卖点5>

【建议售价】
$<价格，考虑竞品区间>

【优化说明】
<2-3句话说明相比原版的改进之处>"""
        optimized = _llm_cloud_sync(opt_prompt)
        db.execute("UPDATE audit_reports SET optimized_listing=?, unlocked=1 WHERE id=?", [optimized, report_id])
    else:
        db.execute("UPDATE audit_reports SET unlocked=1 WHERE id=?", [report_id])

    db.commit()
    return jsonify({"full_report": full_report, "optimized_listing": optimized})

@app.route("/orders")
@app.route("/en/orders")
def orders_page():
    lang = "en" if request.path.startswith("/en/") else "zh"
    db = get_db()
    rows = db.execute("SELECT * FROM audit_reports ORDER BY created_at DESC").fetchall()

    if lang == "en":
        title = "Audit Orders"
        header = "Diagnostic Orders"
        headers = ["Report ID", "Product", "Category", "Price", "Status", "Optimized", "Time"]
        status_paid = "Unlocked"
        status_locked = "Pending"
        opt_yes = "Generated"
        opt_no = "Not yet"
        lang_switch = '<a href="/orders" style="color:#999;text-decoration:none;font-size:13px">中文</a>'
    else:
        title = "订单管理"
        header = "诊断订单"
        headers = ["报告ID", "商品", "品类", "价格", "状态", "优化Listing", "时间"]
        status_paid = "已解锁"
        status_locked = "待付款"
        opt_yes = "已生成"
        opt_no = "未生成"
        lang_switch = '<a href="/en/orders" style="color:#999;text-decoration:none;font-size:13px">EN</a>'

    html = f"""<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{title}</title><style>
*{{margin:0;padding:0;box-sizing:border-box}}body{{font-family:-apple-system,sans-serif;background:#f5f5f5;color:#222;padding:20px}}
h1{{font-size:22px;margin-bottom:4px}}
.lang-switch{{margin-bottom:16px}}
table{{width:100%;border-collapse:collapse;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,0.1)}}
th,td{{padding:10px 14px;text-align:left;font-size:13px;border-bottom:1px solid #eee}}
th{{background:#f0f0f0;font-weight:600}}
.status{{padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600}}
.status-locked{{background:#ffe0e0;color:#d52b1e}}
.status-paid{{background:#e0ffe0;color:#07c160}}
</style></head><body>
<h1>{header}</h1>
<div class="lang-switch">{lang_switch}</div>
<table><tr><th>{headers[0]}</th><th>{headers[1]}</th><th>{headers[2]}</th><th>{headers[3]}</th><th>{headers[4]}</th><th>{headers[5]}</th><th>{headers[6]}</th></tr>"""
    for r in rows:
        st = f'<span class="status status-paid">{status_paid}</span>' if r["unlocked"] else f'<span class="status status-locked">{status_locked}</span>'
        opt = f'<span class="status status-paid">{opt_yes}</span>' if r["optimized_listing"] else f'<span class="status status-locked">{opt_no}</span>'
        html += f"<tr><td>{r['id']}</td><td>{r['title'][:30]}</td><td>{r['category']}</td><td>${r['price']}</td><td>{st}</td><td>{opt}</td><td>{r['created_at'][:19]}</td></tr>"
    html += "</table></body></html>"
    return html

@app.route("/api/audit/report/<report_id>/pay", methods=["POST"])
def audit_pay(report_id):
    """Stripe 真实支付：创建 Checkout Session 并返回支付链接"""
    db = get_db()
    r = db.execute("SELECT * FROM audit_reports WHERE id=?", [report_id]).fetchone()
    if not r:
        return jsonify({"error":"报告不存在"}), 404
    if r["unlocked"]:
        return jsonify({"error":"报告已解锁"}), 400
    try:
        price_cents = int(float(r["price"].replace("$","").replace("¥","")) * 100)
        session = stripe.checkout.Session.create(
            payment_method_types=["card"],
            line_items=[{"price_data":{"currency":"usd","product_data":{"name":f"Temu Listing AI Audit - {r['title'][:50]}"},"unit_amount":price_cents},"quantity":1}],
            mode="payment",
            success_url=f"{PUBLIC_DOMAIN}/api/audit/report/{report_id}/success?session_id={{CHECKOUT_SESSION_ID}}",
            cancel_url=f"{PUBLIC_DOMAIN}/audit?canceled=1",
            metadata={"report_id":report_id}
        )
        return jsonify({"checkout_url":session.url, "session_id":session.id})
    except Exception as e:
        # 如果 Stripe 不可用，回退到模拟支付
        order_no = f"WX{datetime.now().strftime('%Y%m%d%H%M%S')}{report_id[:4].upper()}"
        qr_url = f"https://api.qrserver.com/v1/create-qr-code/?size=200x200&data=wxp://f2f0YtqBkN{order_no}"
        return jsonify({"qr_url":qr_url, "order_no":order_no, "fallback":True})

@app.route("/api/audit/report/<report_id>/success")
def audit_pay_success(report_id):
    """Stripe 支付成功回调：验证并解锁报告"""
    session_id = request.args.get("session_id","")
    if not session_id:
        return """<html><body style='text-align:center;padding-top:80px;font-family:sans-serif'><h2 style='color:red'>参数错误</h2><a href='{}'>返回</a></body></html>""".format(f"{PUBLIC_DOMAIN}/audit"), 400
    try:
        session = stripe.checkout.Session.retrieve(session_id)
        if session.get("payment_status") == "paid":
            db = get_db()
            db.execute("UPDATE audit_reports SET unlocked=1 WHERE id=?", [report_id])
            db.commit()
            return f"""<html><body style='text-align:center;padding-top:80px;font-family:sans-serif'>
<h2 style='color:#07c160'>支付成功！</h2>
<p>报告已解锁，请返回查看完整报告。</p>
<a href='{PUBLIC_DOMAIN}/audit' style='display:inline-block;margin-top:20px;padding:12px 40px;background:#07c160;color:#fff;text-decoration:none;border-radius:6px'>查看报告</a>
<script>setTimeout(function(){{location.href='{PUBLIC_DOMAIN}/audit';}},2000)</script>
</body></html>"""
    except:
        pass
    return f"""<html><body style='text-align:center;padding-top:80px;font-family:sans-serif'>
<h2 style='color:#c00'>支付验证中...</h2><p>请刷新页面查看</p>
<a href='{PUBLIC_DOMAIN}/audit'>返回</a></body></html>"""

@app.route("/api/stripe/webhook", methods=["POST"])
def stripe_webhook():
    """Stripe Webhook：异步接收支付完成通知"""
    payload = request.get_data(as_text=True)
    sig_header = request.headers.get("Stripe-Signature")
    endpoint_secret = os.environ.get("STRIPE_WEBHOOK_SECRET","")
    try:
        if endpoint_secret:
            event = stripe.Webhook.construct_event(payload, sig_header, endpoint_secret)
        else:
            event = json.loads(payload)
        if event["type"] == "checkout.session.completed":
            session = event["data"]["object"]
            report_id = session.get("metadata",{}).get("report_id","")
            if report_id:
                db = get_db()
                db.execute("UPDATE audit_reports SET unlocked=1 WHERE id=?", [report_id])
                db.commit()
    except Exception:
        return jsonify({}), 400
    return jsonify({"status":"ok"})

@app.route("/")
def index():
    return render_template_string(HTML)

@app.after_request
def no_cache(r):
    r.headers['Cache-Control']='no-cache, no-store, must-revalidate'
    r.headers['Pragma']='no-cache'
    r.headers['Expires']='0'
    return r

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080, debug=False)
