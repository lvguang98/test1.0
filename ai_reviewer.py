#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AI 智能审核模块 - 工伤案件数据完整性、法规匹配度、文书质量审查
"""
import json
import re
import urllib.error
import urllib.request
from PyQt5.QtCore import QThread, pyqtSignal


class AIReviewError(Exception):
    """AI 审查异常"""
    pass


# 默认模型名（可在运行时通过 UI 配置修改）
# 常见模型：deepseek-chat, gpt-4o-mini, claude-sonnet-4-20250514
DEFAULT_MODEL = "deepseek-chat"

# API 超时时间（秒）
API_TIMEOUT = 60

# 审查温度（低温度保证审查结果一致性）
REVIEW_TEMPERATURE = 0.3

# 最大输出 token
MAX_TOKENS = 2000

# 7 条法规定义
REGULATIONS = [
    "第十四条第一款第一项 - 在工作时间和工作场所内，因工作原因受到事故伤害的（普通工伤案件）",
    "第十四条第一款第二项 - 工作时间前后在工作场所内，从事与工作有关的预备性或者收尾性工作受到事故伤害的（预备收尾案件）",
    "第十四条第一款第三项 - 在工作时间和工作场所内，因履行工作职责受到暴力等意外伤害的（暴力伤害案件）",
    "第十四条第一款第四项 - 患职业病的（患职业病案件）",
    "第十四条第一款第五项 - 因工外出期间，由于工作原因受到伤害或者发生事故下落不明的（因工外出案件）",
    "第十四条第一款第六项 - 在上下班途中，受到非本人主要责任的交通事故或者城市轨道交通、客运轮渡、火车事故伤害的（上下班时案件）",
    "第十五条第一款第一项 - 在工作时间和工作岗位，突发疾病死亡或者在48小时之内经抢救无效死亡的（工作时因病亡故案件）",
]

# 每条法规的审查要素（法律构成要件）
REGULATION_CRITERIA = {
    "普通工伤案件": "核心审查要素（第十四条第一项）：\n"
        "1. 是否在工作时间内受伤？\n2. 是否在工作场所内受伤？\n3. 是否因工作原因受伤？\n"
        "4. 三者必须同时具备。请据此审查笔录是否充分证明了这三个要素。",
    "预备收尾案件": "核心审查要素（第十四条第二项）：\n"
        "1. 受伤时间是否在正常工作时间的「前后」？（如上班前准备、下班后收拾）\n"
        "2. 受伤时从事的是否为「预备性或收尾性」工作？（如准备工具、打扫现场）\n"
        "3. 该工作是否与本职工作有关？\n"
        "4. 注意：本项不要求在工作时间「内」，而是时间「前后」。请据此审查。",
    "暴力伤害案件": "核心审查要素（第十四条第三项）：\n"
        "1. 是否在工作时间和工作场所内？\n"
        "2. 是否因「履行工作职责」受到暴力伤害？\n"
        "3. 暴力来源是谁？是否与工作职责有直接关联？\n"
        "4. 是否已报案或有其他证据？请据此审查。",
    "患职业病案件": "核心审查要素（第十四条第四项）：\n"
        "1. 是否经有资质的职业病诊断机构确诊？\n"
        "2. 是否有职业病诊断证明书？\n"
        "3. 所患疾病是否在《职业病分类和目录》内？\n"
        "4. 工作环境是否存在导致该职业病的危害因素？请据此审查。",
    "因工外出案件": "核心审查要素（第十四条第五项）：\n"
        "1. 是否属于「因工外出」期间？（单位指派或工作需要的出差）\n"
        "2. 是否因工作原因受到伤害或发生事故？\n"
        "3. 外出路线、时间、目的与工作是否有直接关系？\n"
        "4. 如是下落不明，是否有相关证明？请据此审查。",
    "上下班时案件": "核心审查要素（第十四条第六项）：\n"
        "1. 是否在「上下班途中」？（合理时间、合理路线）\n"
        "2. 是否受到交通事故或城市轨道交通、客运轮渡、火车事故伤害？\n"
        "3. 本人是否承担「主要责任」？（非本人主要责任才能认定）\n"
        "4. 是否有交警责任认定书？路线是否必经之路？请据此审查。",
    "工作时因病亡故案件": "核心审查要素（第十五条第一项）：\n"
        "1. 是否在「工作时间」和「工作岗位」上突发疾病？\n"
        "2. 是否在「48小时之内」经抢救无效死亡？\n"
        "3. 突发疾病与死亡之间是否有连续抢救记录？\n"
        "4. 是否有医疗机构的死亡证明和抢救记录？请据此审查。",
}

def _get_regulation_criteria(regulation_text):
    """根据条例文本获取对应的审查要素"""
    for key, criteria in REGULATION_CRITERIA.items():
        if key in regulation_text:
            return criteria
    return "请根据选定条例的法律构成要件进行审查。"


class AIReviewer:
    """AI 审查引擎 — 封装 API 通信和 prompt 逻辑"""

    def __init__(self, api_url, api_key):
        self.api_url = api_url.rstrip("/")
        self.api_key = api_key
        # 自动检测 API 类型：URL 含 /anthropic 则为 Anthropic 兼容接口
        self.api_type = "anthropic" if "/anthropic" in self.api_url else "openai"
        if self.api_type == "anthropic":
            # Anthropic 接口一般在 /anthropic 后面不需要额外路径
            # 例如 https://api.deepseek.com/anthropic → POST /v1/messages
            self.endpoint = f"{self.api_url}/v1/messages"
        else:
            self.endpoint = f"{self.api_url}/v1/chat/completions"

    def _call_api(self, system_prompt, user_prompt):
        """调用 AI API（自动适配 OpenAI / Anthropic 格式），返回响应文本"""
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }
        if self.api_type == "anthropic":
            headers["anthropic-version"] = "2023-06-01"
            payload = json.dumps({
                "model": DEFAULT_MODEL,
                "system": system_prompt,
                "messages": [
                    {"role": "user", "content": user_prompt},
                ],
                "temperature": REVIEW_TEMPERATURE,
                "max_tokens": MAX_TOKENS,
            }).encode("utf-8")
        else:
            payload = json.dumps({
                "model": DEFAULT_MODEL,
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt},
                ],
                "temperature": REVIEW_TEMPERATURE,
                "max_tokens": MAX_TOKENS,
            }).encode("utf-8")

        req = urllib.request.Request(
            self.endpoint, data=payload, headers=headers, method="POST")
        try:
            with urllib.request.urlopen(req, timeout=API_TIMEOUT) as resp:
                body = json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            detail = ""
            try:
                detail = e.read().decode("utf-8")[:300]
            except Exception:
                pass
            raise AIReviewError(f"API返回错误 ({e.code}): {detail}")
        except urllib.error.URLError as e:
            raise AIReviewError(f"无法连接到API服务，请检查API地址: {e.reason}")
        except (OSError, TimeoutError) as e:
            raise AIReviewError(f"审核请求超时或网络异常: {e}")

        # 解析响应：Anthropic 和 OpenAI 格式不同
        try:
            if self.api_type == "anthropic":
                # Anthropic: {"content": [{"type":"text","text":"..."}], ...}
                content = body.get("content", [])
                text_blocks = [b["text"] for b in content if b.get("type") == "text"]
                return "\n".join(text_blocks)
            else:
                # OpenAI: {"choices": [{"message": {"content": "..."}}], ...}
                return body["choices"][0]["message"]["content"]
        except (KeyError, json.JSONDecodeError, IndexError) as e:
            raise AIReviewError(f"API返回格式异常: {e}")

    def _parse_json_response(self, text):
        """从 AI 响应中提取 JSON 字典，失败则用原始文本作为 summary"""
        # 尝试直接解析
        try:
            return json.loads(text)
        except json.JSONDecodeError:
            pass
        # 尝试从 markdown 代码块中提取
        m = re.search(r"```(?:json)?\s*\n?(.*?)\n?```", text, re.DOTALL)
        if m:
            try:
                return json.loads(m.group(1))
            except json.JSONDecodeError:
                pass
        # 尝试从文本中提取最外层花括号范围
        brace_start = text.find("{")
        brace_end = text.rfind("}")
        if brace_start >= 0 and brace_end > brace_start:
            try:
                return json.loads(text[brace_start : brace_end + 1])
            except json.JSONDecodeError:
                pass
        # 降级：原文本作为摘要
        return {"summary": text, "issues": [], "suggestions": [], "score": None}

    # ─── 三个审查方向 ─────────────────────────────────────────────────

    def review_data_completeness(self, form_data, case_data=None, doc_text=None):
        """数据完整性检查（含全文审查，按选定条例要素审查）"""
        regulation = form_data.get("条例", "")
        criteria = _get_regulation_criteria(regulation)
        system_prompt = (
            "你是一名工伤案件审核专家。根据表单数据和笔录全文，检查数据完整性和一致性。\n\n"
            "注意：年龄、性别等可通过身份证推算的信息，无需比对和报告。\n"
            "只检查以下内容：\n"
            "1. 必填项是否缺失（姓名、身份证号、用人单位等）\n"
            "2. 身份证号格式是否正确（18位）\n"
            "3. 笔录中描述的事实与表单信息是否有明确矛盾\n\n"
            f"【选定条例的法定审查要素】\n{criteria}\n\n"
            "根据以上法定要素，重点审查笔录中是否充分证明了各项要件。\n"
            "对每个发现的问题，提供可直接用于笔录的追问和可能的回答。\n"
            "请以JSON格式返回：\n"
            '{"summary": "总体评价（一句话）", '
            '"issues": [{"description": "问题摘要", "question": "问：追问？", "answer": "答：可能回答。"}], '
            '"completeness_score": 85}'
        )
        user_prompt = self._build_full_review_prompt(form_data, case_data, doc_text)
        raw = self._call_api(system_prompt, user_prompt)
        result = self._parse_json_response(raw)
        structured_issues = self._normalize_issues(result.get("issues", []))
        return {
            "title": "数据完整性检查",
            "result": self._build_html_result(
                "数据完整性检查",
                result.get("summary", ""),
                result.get("completeness_score"),
                structured_issues,
                result.get("suggestions", []),
            ),
            "issues": structured_issues,
            "score": result.get("completeness_score"),
        }

    def review_regulation_match(self, form_data, case_data=None, doc_text=None):
        """法规匹配度分析（含全文审查，按选定条例要素审查）"""
        regulation = form_data.get("条例", "")
        criteria = _get_regulation_criteria(regulation)
        regulations_text = "\n".join(REGULATIONS)
        system_prompt = (
            "你是一名工伤认定法律专家，精通《工伤保险条例》。"
            "根据案件事实（特别是笔录全文），判断用户选择的条例是否最合适。\n\n"
            f"可用条例列表：\n{regulations_text}\n\n"
            f"用户选定条例及法定审查要素：\n{criteria}\n\n"
            "请严格根据选定条例的法定构成要件，逐项审查笔录内容是否充分证明各要件。\n"
            "如果笔录事实不符合选定条例的要素，推荐最匹配的条例。\n"
            "对每个发现的问题，提供一条可用于笔录核实的追问和可能的回答。\n"
            "请以JSON格式返回：\n"
            '{"summary": "匹配度评价", "selected_regulation": "用户选择的条例", '
            '"recommended_regulation": "推荐的条例（一致则写\\"一致\\"）", '
            '"match_level": "完全匹配|基本匹配|部分匹配|不匹配", '
            '"analysis": "详细分析（200字以内）", '
            '"issues": [{"description": "问题摘要", "question": "问：追问？", "answer": "答：可能回答。"}], '
            '"match_score": 90}'
        )
        user_prompt = self._build_full_review_prompt(form_data, case_data, doc_text)
        raw = self._call_api(system_prompt, user_prompt)
        result = self._parse_json_response(raw)
        structured_issues = self._normalize_issues(result.get("issues", []))
        return {
            "title": "条例匹配度分析",
            "result": self._build_html_result(
                "条例匹配度分析",
                result.get("summary", ""),
                result.get("match_score"),
                structured_issues,
                result.get("suggestions", []),
                extra_info={
                    "选定条例": result.get("selected_regulation", ""),
                    "推荐条例": result.get("recommended_regulation", ""),
                    "匹配程度": result.get("match_level", ""),
                    "分析": result.get("analysis", ""),
                },
            ),
            "issues": structured_issues,
            "score": result.get("match_score"),
        }

    def review_document_quality(self, doc_text, form_data=None):
        """文书质量审查"""
        system_prompt = (
            "你是一名工伤案件文书审核专家。请审查工伤认定谈话笔录的质量。\n\n"
            "审查要点：\n"
            "1. 事实陈述是否清晰完整（时间、地点、人物、经过、结果）\n"
            "2. 逻辑是否自洽（陈述与基本信息是否一致）\n"
            "3. 用词是否专业规范\n"
            "4. 受伤经过、就医情况、医疗结论是否相互印证\n"
            "5. 是否存在明显遗漏或矛盾\n\n"
            "对每个发现的问题，提供一条可用于补充笔录的追问和可能的回答。\n"
            "请以JSON格式返回：\n"
            '{"summary": "总体评价", '
            '"issues": [{"description": "问题摘要", "question": "问：追问？", "answer": "答：可能回答。"}], '
            '"quality_score": 80}'
        )
        user_prompt = f"请审查以下谈话笔录文书内容：\n\n{doc_text[:4000]}"
        if form_data:
            user_prompt = (
                f"案件基本信息：\n"
                f"  受伤职工：{form_data.get('受伤职工', '')}\n"
                f"  用人单位：{form_data.get('用人单位', '')}\n"
                f"  案件类型：{form_data.get('案件类型', '')}\n"
                f"  选用条例：{form_data.get('条例', '')}\n\n"
                f"笔录内容：\n{doc_text[:3500]}"
            )
        raw = self._call_api(system_prompt, user_prompt)
        result = self._parse_json_response(raw)
        structured_issues = self._normalize_issues(result.get("issues", []))
        return {
            "title": "文书质量审查",
            "result": self._build_html_result(
                "文书质量审查",
                result.get("summary", ""),
                result.get("quality_score"),
                structured_issues,
                result.get("suggestions", []),
            ),
            "issues": structured_issues,
            "score": result.get("quality_score"),
        }

    # ─── Issues 标准化 ───────────────────────────────────────────────

    def _normalize_issues(self, raw_issues):
        """将 AI 返回的各种 issues 格式统一为 [{description, question, answer}]"""
        result = []
        for item in raw_issues:
            if isinstance(item, dict):
                result.append({
                    "description": item.get("description", item.get("desc", "")),
                    "question": item.get("question", item.get("q", "问：请补充说明。")),
                    "answer": item.get("answer", item.get("a", "答：...")),
                })
            elif isinstance(item, str):
                result.append({
                    "description": item,
                    "question": f"问：关于「{item}」，请补充说明。",
                    "answer": "答：...",
                })
        return result

    # ─── 提示词构建 ───────────────────────────────────────────────────

    def _build_full_review_prompt(self, form_data, case_data=None, doc_text=None):
        """构建包含笔录全文的审查提示"""
        parts = [self._build_data_prompt(form_data, case_data)]
        if doc_text:
            parts.append("")
            parts.append("【笔录全文（请对照审查）】")
            parts.append(doc_text)
        return "\n".join(parts)

    def _build_data_prompt(self, form_data, case_data=None):
        """构建标准化的表单数据提示"""
        person_type = form_data.get("人员类型", "本人")
        lines = [
            "请审核以下工伤案件数据：",
            "",
            "【案件基本信息】",
            f"案本号：{form_data.get('案本号', '（未生成）')}",
            f"受伤职工：{form_data.get('受伤职工', '')}",
            f"申请人：{form_data.get('申请人', '')}",
            f"案件类型：{form_data.get('案件类型', '')}",
            f"拟用条例：{form_data.get('条例', '')}",
            "",
            f"【人员信息 - {person_type}】",
            f"姓名：{form_data.get(f'{person_type}姓名', '')}",
            f"性别：{form_data.get(f'{person_type}性别', '')}",
            f"身份证号：{form_data.get(f'{person_type}身份证号', '')}",
            f"身份证地址：{form_data.get(f'{person_type}身份证地址', '')}",
            f"现住址：{form_data.get(f'{person_type}现住址', '')}",
            f"电话：{form_data.get(f'{person_type}电话', '')}",
            f"岗位：{form_data.get(f'{person_type}岗位', '')}",
            "",
            "【单位信息】",
            f"用人单位：{form_data.get('用人单位', '')}",
            f"用工单位：{form_data.get('用工单位', '')}",
            f"工作场所：{form_data.get('工作场所', '')}",
            "",
            "【操作信息】",
            f"操作员：{form_data.get('操作员', '')}",
            f"当前日期：{form_data.get('当前日期', '')}",
        ]

        # 如果有已有案件的 person_info，一并提供
        if case_data and isinstance(case_data, dict):
            person_info = case_data.get("person_info", {})
            if person_info:
                lines.append("")
                lines.append("【已记录的受伤信息（来自已有笔录）】")
                for key, label in [
                    ("受伤经过", "受伤经过"),
                    ("就医情况", "就医情况"),
                    ("医疗结论", "医疗结论"),
                    ("自我介绍", "自我介绍"),
                ]:
                    val = person_info.get(key, "")
                    if val:
                        lines.append(f"{label}：{val}")

        return "\n".join(lines)

    # ─── 文书 AI 润色 ──────────────────────────────────────────────

    def ai_polish_approval(self, self_intro, injury_desc, medical_desc, person_name):
        """对审批表中的三段文字进行法律文书规范化润色"""
        system_prompt = (
            "你是一名工伤案件法律文书审核员。请对以下三段文字进行规范化润色，"
            "使其符合法律文书的要求。\n\n"
            "润色规范：\n"
            f"1. 全文使用第三人称「{person_name}」替代「我」「我们」\n"
            "2. 语言正式、客观、精炼，去除口语化表达\n"
            "3. 保持事实准确性，不添加原文没有的信息\n"
            "4. 时间、地点、人物、经过要清晰完整\n"
            "5. 自我介绍去掉「我是」开头，直接陈述身份\n\n"
            "严格以JSON格式返回：\n"
            '{"自我介绍":"","受伤经过":"","就医情况":""}'
        )
        user_prompt = (
            f"【自我介绍】\n{self_intro}\n\n"
            f"【受伤经过】\n{injury_desc}\n\n"
            f"【就医情况】\n{medical_desc}"
        )
        raw = self._call_api(system_prompt, user_prompt)
        return self._parse_json_response(raw)

    def ai_extract_notice_info(self, approval_text):
        """从案件审批表 AI 提取告知书/通知书所需字段"""
        system_prompt = (
            "你是一名工伤案件文书审核员。审查案件审批表内容，检查信息是否一致，"
            "提取告知书/通知书所需字段。\n"
            "严格以JSON格式返回：\n"
            '{"受伤职工":"","用人单位":"","受伤时间":"","医疗结论":"",'
            '"综合意见":"","申请时间":"","受理时间":""}'
        )
        user_prompt = f"案件审批表内容：\n{approval_text}"
        raw = self._call_api(system_prompt, user_prompt)
        return self._parse_json_response(raw)

    # ─── HTML 构建 ────────────────────────────────────────────────────

    def _build_html_result(self, title, summary, score=None, issues=None,
                           suggestions=None, extra_info=None):
        """将审查结果转换为 HTML 供 QTextBrowser 显示"""
        parts = [f'<h3 style="color:#2c3e50;">{title}</h3>']

        if score is not None:
            if score >= 80:
                color = "#27ae60"
            elif score >= 60:
                color = "#e67e22"
            else:
                color = "#e74c3c"
            parts.append(
                f'<p><b>评分：</b>'
                f'<span style="color:{color};font-size:18px;font-weight:bold;">'
                f'{score}/100</span></p>'
            )

        if summary:
            parts.append(f'<p><b>评价：</b>{summary}</p>')

        if extra_info:
            parts.append('<hr style="border:1px solid #eee;">')
            for k, v in extra_info.items():
                if v:
                    parts.append(f'<p><b>{k}：</b>{v}</p>')

        if issues:
            count = len(issues)
            parts.append(
                f'<p style="color:#c0392b;"><b>发现 {count} 个问题</b>'
                f'（详见下方可勾选列表，勾选后可加入笔录）</p>')

        if suggestions:
            parts.append('<h4 style="color:#2980b9;">改进建议：</h4><ul>')
            for item in suggestions:
                parts.append(f'<li style="margin:4px 0;">{item}</li>')
            parts.append("</ul>")

        return "".join(parts)

    # ─── 全流程审查 ───────────────────────────────────────────────────

class ReviewWorker(QThread):
    """后台审查线程 — 在子线程中调用 AI，通过信号通知主线程"""

    progress = pyqtSignal(str)
    step_completed = pyqtSignal(str, str)
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, api_url, api_key, form_data, case_data=None,
                 doc_text=None):
        super().__init__()
        self.api_url = api_url
        self.api_key = api_key
        self.form_data = form_data
        self.case_data = case_data
        self.doc_text = doc_text

    def run(self):
        reviewer = AIReviewer(self.api_url, self.api_key)
        empty = {"title": "", "result": "", "issues": [], "score": None}

        # 数据完整性
        self.progress.emit("正在检查数据完整性...")
        try:
            r1 = reviewer.review_data_completeness(
                self.form_data, self.case_data, self.doc_text)
        except Exception as e:
            r1 = {"title": "数据完整性检查",
                  "result": f'<p style="color:#e74c3c;">审查失败：{e}</p>',
                  "issues": [], "score": None}
        try:
            self.step_completed.emit("completeness", json.dumps(r1, ensure_ascii=False))
        except Exception:
            self.step_completed.emit("completeness", json.dumps(empty, ensure_ascii=False))

        # 法规匹配
        self.progress.emit("正在分析条例匹配度...")
        try:
            r2 = reviewer.review_regulation_match(
                self.form_data, self.case_data, self.doc_text)
        except Exception as e:
            r2 = {"title": "条例匹配度分析",
                  "result": f'<p style="color:#e74c3c;">审查失败：{e}</p>',
                  "issues": [], "score": None}
        try:
            self.step_completed.emit("regulation", json.dumps(r2, ensure_ascii=False))
        except Exception:
            self.step_completed.emit("regulation", json.dumps(empty, ensure_ascii=False))

        # 文书质量
        if self.doc_text:
            self.progress.emit("正在审查文书质量...")
            try:
                r3 = reviewer.review_document_quality(self.doc_text, self.form_data)
            except Exception as e:
                r3 = {"title": "文书质量审查",
                      "result": f'<p style="color:#e74c3c;">审查失败：{e}</p>',
                      "issues": [], "score": None}
        else:
            r3 = {"title": "文书质量审查",
                  "result": '<p style="color:#888;">暂无笔录文书。请先生成笔录。</p>',
                  "issues": [], "score": None}
        try:
            self.step_completed.emit("document_quality", json.dumps(r3, ensure_ascii=False))
        except Exception:
            self.step_completed.emit("document_quality", json.dumps(empty, ensure_ascii=False))

        # 汇总
        scores = []
        all_issues = []
        for r in [r1, r2, r3]:
            if r.get("score") is not None:
                scores.append(r["score"])
            all_issues.extend(r.get("issues", []))
        overall = round(sum(scores) / len(scores)) if scores else None

        full_result = {
            "completeness": r1, "regulation": r2, "document_quality": r3,
            "overall_score": overall, "total_issues": all_issues,
        }
        try:
            self.finished.emit(json.dumps(full_result, ensure_ascii=False))
        except Exception:
            pass
