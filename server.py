#!/usr/bin/env python3
import html
import io
import json
import os
import re
import shutil
import tempfile
import urllib.parse
import base64
import secrets
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

import pdfplumber
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

ROOT = Path(__file__).resolve().parent
OUTPUTS = ROOT / "public"
PORT = int(os.environ.get("PORT") or os.environ.get("RESUME_AGENT_PORT", "8766"))
APP_PASSWORD = os.environ.get("APP_PASSWORD", "")
APP_USERNAME = os.environ.get("APP_USERNAME", "admin")

LEXICON = [
    "AI", "Agent", "LLM", "RAG", "Prompt", "数据分析", "增长", "商业化", "B端",
    "SaaS", "产品策略", "用户研究", "A/B测试", "SQL", "Python", "机器学习",
    "深度学习", "模型评估", "推荐系统", "搜索", "支付", "风控", "CRM",
    "低代码", "自动化", "GTM", "运营", "战略", "跨团队", "项目管理",
    "0到1", "指标体系", "留存", "转化", "收入", "成本", "效率", "合规",
    "沟通", "组织协调", "抗压", "出差", "加班", "客户", "交付", "管理",
]

STOPWORDS = set(
    "岗位 职责 任职 要求 优先 以及 相关 负责 进行 具备 熟悉 能够 以上 工作 公司 团队 业务 产品 项目 经验 能力 包括 通过 对于 使用 一个 一份 我们 你将 需要".split()
)


def tokenize(text):
    found = []
    lower = text.lower()
    for word in LEXICON:
        if word.lower() in lower:
            found.append(word)
    words = re.findall(r"[A-Za-z][A-Za-z+#.]{1,28}|[\u4e00-\u9fa5]{2,8}", text)
    for word in words:
        if word not in STOPWORDS and len(word) > 1:
            found.append(word)
    counts = {}
    for word in found:
        counts[word] = counts.get(word, 0) + 1
    return [word for word, _ in sorted(counts.items(), key=lambda item: (-item[1], len(item[0]), item[0]))[:18]]


def normalize(text):
    return re.sub(r"\n{3,}", "\n\n", text.replace("\r", "").replace("\t", " ")).strip()


def extract_docx(path):
    doc = Document(path)
    lines = []
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if text:
            lines.append(text)
    for table in doc.tables:
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells if cell.text.strip()]
            if cells:
                lines.append(" | ".join(cells))
    return "\n".join(lines)


def extract_pdf(path):
    lines = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if text.strip():
                lines.append(text.strip())
    return "\n\n".join(lines)


def extract_text(path, filename):
    suffix = Path(filename).suffix.lower()
    if suffix == ".docx":
        return extract_docx(path)
    if suffix == ".pdf":
        return extract_pdf(path)
    if suffix == ".txt":
        return Path(path).read_text(encoding="utf-8", errors="ignore")
    raise ValueError("仅支持 .docx、.pdf、.txt。旧版 .doc 请先另存为 .docx。")


def score_line(line, keywords):
    lower = line.lower()
    score = sum(4 for kw in keywords if kw.lower() in lower)
    if re.search(r"[0-9%万亿kK+]", line):
        score += 3
    if re.search(r"负责|主导|搭建|推动|优化|提升|增长|降低|设计|上线|交付|管理|协作|沟通", line):
        score += 2
    return score


def select_lines(resume_text, jd, requirements, intensity):
    keywords = tokenize(jd + "\n" + requirements)
    lines = [line.strip(" -*\u2022") for line in normalize(resume_text).splitlines() if len(line.strip()) > 6]
    scored = sorted(((line, score_line(line, keywords)) for line in lines), key=lambda item: item[1], reverse=True)
    count = {"conservative": 7, "balanced": 9, "aggressive": 12}.get(intensity, 9)
    picked = [line for line, score in scored if score > 0][:count] or lines[: min(count, len(lines))]
    return keywords, picked


def build_evaluation(resume_text, jd, requirements, role, company, company_type, keywords):
    all_text = resume_text + "\n" + jd + "\n" + requirements
    matched = [kw for kw in keywords if kw.lower() in resume_text.lower()]
    skill_words = "、".join((matched or keywords)[:6]) or "岗位相关技能"
    role_name = role or "目标岗位"
    company_name = company or "目标公司"
    has_numbers = bool(re.search(r"[0-9%万亿kK+]", resume_text))
    quant = "简历中已有一定量化结果，可继续补充业务指标、规模、转化或效率提升数据" if has_numbers else "建议补充可量化成果，例如规模、转化率、成本、周期、收入或效率变化"
    mobility = "可在面试中进一步确认加班、出差和高强度节奏适应度"
    if re.search(r"出差|加班|抗压|高压|吃苦|驻场|客户现场", all_text):
        mobility = "材料中已出现高强度协作、客户现场或节奏适应相关线索，具备进一步验证价值"
    return [
        f"系统学习与知识结构：围绕 {company_type} 与 {role_name} 所需能力，已体现对 {skill_words} 等知识模块的接触和应用基础。",
        f"技能掌握与实践经验：过往经历可迁移到 {company_name} 的 {role_name}，重点优势在任务拆解、结果交付、跨团队协同和岗位关键词相关实践。",
        f"岗位胜任度：与 JD 匹配的经历应优先呈现为“目标 - 动作 - 结果 - 复盘”，适合投递 {role_name} 及相近的业务/产品/项目/运营协同类岗位。",
        f"性格与沟通协作：从材料看具备主动推进、沟通表达、组织协调和问题闭环意识；建议在面试中准备 1-2 个冲突协调或复杂协作案例。",
        f"抗压与适应度：{quant}；{mobility}，避免在简历中作无法核验的绝对化承诺。",
    ]


def adapt_bullets(lines, keywords, company_type, tone):
    tone_prefix = {
        "direct": "主导",
        "executive": "统筹",
        "builder": "落地",
    }.get(tone, "主导")
    if "早期创业" in company_type:
        tone_prefix = "从0到1推进"
    bullets = []
    for line in lines:
        clean = re.sub(r"^[-*\u2022]\s*", "", line).rstrip("。；;")
        matched = [kw for kw in keywords if kw.lower() in clean.lower()][:3]
        suffix = f"，对应岗位关注的{'、'.join(matched)}" if matched else "，强化目标岗位所需的业务判断与执行闭环"
        if re.match(r"^(主导|负责|推动|搭建|设计|优化|统筹|落地|从0到1|管理)", clean):
            bullets.append(f"{clean}{suffix}")
        else:
            bullets.append(f"{tone_prefix}{clean}{suffix}")
    return bullets


def add_run(paragraph, text, bold=False, size=10.5, color=None):
    run = paragraph.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = "Microsoft YaHei"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    if color:
        run.font.color.rgb = RGBColor(*color)
    return run


def insert_before(paragraph, text="", style=None):
    new_p = paragraph.insert_paragraph_before(text)
    if style:
        new_p.style = style
    return new_p


def insert_frontmatter(doc, evaluation, bullets, role, company):
    first = doc.paragraphs[0] if doc.paragraphs else doc.add_paragraph()
    title = insert_before(first)
    title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    add_run(title, "综合评价", bold=True, size=13, color=(23, 78, 166))
    for item in evaluation:
        p = insert_before(first)
        p.style = doc.styles["Normal"]
        add_run(p, item, size=9.5)
    p = insert_before(first)
    add_run(p, "", size=4)
    focus = insert_before(first)
    add_run(focus, f"岗位适配重点：{company or '目标公司'}｜{role or '目标岗位'}", bold=True, size=11, color=(23, 78, 166))
    for item in bullets:
        p = insert_before(first)
        p.style = doc.styles["Normal"]
        add_run(p, item, size=9.5)
    spacer = insert_before(first)
    add_run(spacer, "", size=4)


def create_new_docx(out_path, source_text, evaluation, bullets, role, company):
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Pt(42)
    section.bottom_margin = Pt(42)
    section.left_margin = Pt(50)
    section.right_margin = Pt(50)
    title = doc.add_paragraph()
    add_run(title, f"{company or '目标公司'}｜{role or '目标岗位'}｜适配版简历", bold=True, size=15, color=(23, 78, 166))
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()
    h = doc.add_paragraph()
    add_run(h, "综合评价", bold=True, size=12, color=(23, 78, 166))
    for item in evaluation:
        p = doc.add_paragraph(style=None)
        p.paragraph_format.left_indent = Pt(12)
        add_run(p, item, size=10)
    h = doc.add_paragraph()
    add_run(h, "岗位适配重点", bold=True, size=12, color=(23, 78, 166))
    for item in bullets:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Pt(12)
        add_run(p, item, size=10)
    h = doc.add_paragraph()
    add_run(h, "原简历文本（便于核对）", bold=True, size=12, color=(23, 78, 166))
    for line in [line for line in source_text.splitlines() if line.strip()][:80]:
        p = doc.add_paragraph()
        add_run(p, line.strip(), size=9.5)
    doc.save(out_path)


def generate_docx(upload_path, filename, payload):
    source_text = extract_text(upload_path, filename)
    keywords, lines = select_lines(source_text, payload["jd"], payload["requirements"], payload["intensity"])
    evaluation = build_evaluation(
        source_text,
        payload["jd"],
        payload["requirements"],
        payload["role"],
        payload["company"],
        payload["companyType"],
        keywords,
    )
    bullets = adapt_bullets(lines, keywords, payload["companyType"], payload["tone"])
    safe_company = re.sub(r"[\\/:*?\"<>|]+", "-", payload["company"] or "target-company")
    safe_role = re.sub(r"[\\/:*?\"<>|]+", "-", payload["role"] or "target-role")
    out_path = Path(tempfile.gettempdir()) / f"{safe_company}-{safe_role}-resume.docx"
    if filename.lower().endswith(".docx"):
        shutil.copyfile(upload_path, out_path)
        doc = Document(out_path)
        insert_frontmatter(doc, evaluation, bullets, payload["role"], payload["company"])
        doc.save(out_path)
    else:
        create_new_docx(out_path, source_text, evaluation, bullets, payload["role"], payload["company"])
    preview = "\n".join(["综合评价", *[f"- {item}" for item in evaluation], "", "岗位适配重点", *[f"- {item}" for item in bullets]])
    return out_path, preview, keywords


class Handler(BaseHTTPRequestHandler):
    def is_authorized(self):
        if not APP_PASSWORD:
            return True
        auth = self.headers.get("Authorization", "")
        if not auth.startswith("Basic "):
            return False
        try:
            decoded = base64.b64decode(auth.removeprefix("Basic ").strip()).decode("utf-8")
            username, password = decoded.split(":", 1)
        except Exception:
            return False
        return secrets.compare_digest(username, APP_USERNAME) and secrets.compare_digest(password, APP_PASSWORD)

    def require_auth(self):
        if self.is_authorized():
            return True
        self.send_response(401)
        self.send_header("WWW-Authenticate", 'Basic realm="Resume Agent"')
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.end_headers()
        self.wfile.write("需要输入访问账号和密码。".encode("utf-8"))
        return False

    def do_GET(self):
        if self.path == "/healthz":
            self.send_response(200)
            self.send_header("Content-Type", "text/plain; charset=utf-8")
            self.end_headers()
            self.wfile.write(b"ok")
            return
        if not self.require_auth():
            return
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        if path == "/":
            path = "/resume-agent.html"
        target = (OUTPUTS / path.lstrip("/")).resolve()
        if not str(target).startswith(str(OUTPUTS.resolve())) or not target.exists():
            self.send_error(404)
            return
        content_type = "text/html; charset=utf-8" if target.suffix == ".html" else "application/octet-stream"
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.end_headers()
        self.wfile.write(target.read_bytes())

    def do_POST(self):
        if not self.require_auth():
            return
        if self.path != "/api/generate":
            self.send_error(404)
            return
        upload_path = None
        out_path = None
        try:
            import cgi

            form = cgi.FieldStorage(
                fp=self.rfile,
                headers=self.headers,
                environ={
                    "REQUEST_METHOD": "POST",
                    "CONTENT_TYPE": self.headers.get("Content-Type"),
                },
            )
            file_item = form["resumeFile"]
            filename = Path(file_item.filename or "resume").name
            with tempfile.NamedTemporaryFile(delete=False, suffix=Path(filename).suffix) as tmp:
                tmp.write(file_item.file.read())
                upload_path = tmp.name
            payload = {
                "company": form.getfirst("company", ""),
                "role": form.getfirst("role", ""),
                "companyType": form.getfirst("companyType", "AI 原生公司"),
                "intensity": form.getfirst("intensity", "balanced"),
                "tone": form.getfirst("tone", "direct"),
                "jd": form.getfirst("jd", ""),
                "requirements": form.getfirst("requirements", ""),
            }
            out_path, preview, keywords = generate_docx(upload_path, filename, payload)
            data = out_path.read_bytes()
            encoded_preview = html.escape(preview)
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            self.send_header("Content-Disposition", f"attachment; filename={urllib.parse.quote(out_path.name)}")
            self.send_header("X-Preview", urllib.parse.quote(encoded_preview))
            self.send_header("X-Keywords", urllib.parse.quote(json.dumps(keywords, ensure_ascii=False)))
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
        except Exception as exc:
            message = str(exc).encode("utf-8")
            self.send_response(400)
            self.send_header("Content-Type", "text/plain; charset=utf-8")
            self.send_header("Content-Length", str(len(message)))
            self.end_headers()
            self.wfile.write(message)
        finally:
            for path in (upload_path, out_path):
                if path:
                    try:
                        Path(path).unlink(missing_ok=True)
                    except Exception:
                        pass


if __name__ == "__main__":
    OUTPUTS.mkdir(exist_ok=True)
    server = ThreadingHTTPServer(("", PORT), Handler)
    print(f"Resume Agent running on port {PORT}", flush=True)
    server.serve_forever()
