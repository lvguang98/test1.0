# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 运行命令

```bash
# 激活虚拟环境
venv310\Scripts\activate

# 安装依赖（阿里云镜像）
venv310\Scripts\pip install -r requirements.txt -i https://mirrors.aliyun.com/pypi/simple/

# 运行主程序
python main.py
```

## 架构概览

这是一个 **工伤案件管理系统** — PyQt5 桌面应用，用于生成工伤认定相关的 Word 文书。

### 核心文件

| 文件 | 职责 |
|------|------|
| `main.py` (91KB, ~2170行) | 全部业务逻辑，单文件 `MainWindow` 类 |
| `main_window.ui` | Qt Designer 生成的 UI 布局文件 |
| `config_manager.py` | 配置持久化（QSettings + base64 编码 API key） |
| `templates/` | 21 个 docx 模板文件 |
| `*.xlsx` (3个) | 用人单位/用工单位/工作场所的下拉数据源 |

### 数据存储

- **案件索引**: `cases_index.json` — 所有案件的结构化元数据（案本号、人员信息、审批时间等）
- **下拉选项**: 3 个 Excel 文件 (`用人单位名称汇总.xlsx` 等)，运行时可通过 ComboBox 新增/删除
- **案件文件**: `{年份}/{案本号}/` 目录下存放生成的 docx 文书
- 配置通过 `QSettings` 存储在系统默认位置（操作员名、API URL/Key）

### 案本号规则

格式：`{PREFIX}-{姓名}-{序号:03d}`

| 案件类型 | 前缀 |
|----------|------|
| 普通工伤（单位申请） | GS |
| 个人申请 | GR |
| 工亡（单位申请） | GSW |
| 个人申请工亡 | GRW |

### 人员类型与模板

三种人员角色（`radio_self` / `radio_witness` / `radio_legal_entity`），每种对应不同模板：
- **本人**: 模板名 `本人谈话笔录（{案件类型}）.docx`
- **证人**: 模板名 `证人谈话笔录（{案件类型}）.docx`
- **法人**: 模板名 `法人谈话笔录（{案件类型}）.docx`

模板使用 **docxtpl** 库渲染（Jinja2 风格 `{{变量}}` 占位符）。

### 关键数据流

1. 用户填写表单 → `collect_form_data()` 返回字典
2. `on_generate_record()` 按人员类型分流到 `handle_person_case()` / `handle_witness_case()` / `handle_legal_case()`
3. 生成案本号 → 创建年份文件夹 → 更新 `cases_index.json`
4. `generate_transcript_unified()` 渲染 docx 模板 → 保存 → 启动 Word 打开
5. 本人笔录关闭后，后台线程 `extract_person_info_from_doc()` 从文档中提取「受伤经过/就医情况/医疗结论」回写到索引

### 身份证读卡器

通过 `ctypes` 调用 `sdtapi.dll`（SDT 身份证读卡器），流程：`StartFindIDCard → SelectIDCard → ReadBaseMsg → ClosePort`。

### 注意事项

- `main_window.ui` 中的 widget objectName 必须与 `main.py` 中的 `self.xxx` 属性名一致（PyQt `loadUi` 自动绑定）
- 案件数据在 `cases_index.json` 和 docx 文件中各存一份，修改时需两边同步
- `update_case_index()` 是合并更新（保留旧数据，用新数据覆盖），不会清空已有字段
- 只设计了 `venv310`（Python 3.10），系统需预装 Python 3.10
