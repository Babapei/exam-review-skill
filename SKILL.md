---
name: exam-review-sop
description: Use when the user is preparing for a course final exam and needs to organize study materials from course slides/syllabus into a comprehensive review package. Triggers on keywords like "复习", "期末", "整理资料", "exam prep", "study guide", "课程" plus file uploads of PPTs/PDFs/syllabus.
---

# Exam Review Material Generation SOP

> A standardized workflow for transforming course materials (PPTs, syllabus, past exams) into a comprehensive, exam-ready study package.
> Compatible with: OpenCode, Claude Code, Claude Desktop, GitHub Copilot (via instructions).

## Overview

This skill guides an AI assistant through a 6-phase process to:
1. Parse exam requirements and course materials
2. Extract and catalog all knowledge points
3. Generate standardized review documents
4. Verify completeness against exam syllabus
5. Export in Markdown format (with merged version for Notion import)
6. **(Optional)** Generate exam-approach guide organized by question type with mastery task chains

**Target output**: A dynamically-configured review package (5-11 documents) adapted to the specific exam type and course nature, ready for import to Notion or hand-copying to cheat sheet.

---

## Adaptive Configuration (动态适配)

The AI MUST first detect exam type and course nature from the syllabus, then adjust the entire workflow. Do NOT blindly generate the same 7 documents for every course.

### Exam Type Detection (from syllabus)

| Detected Type | Key Indicators | Output Adjustment |
|--------------|----------------|-------------------|
| **Essay/问答** (like this NLP exam) | "问答题", "简答", "论述", "分析", "比较" | Full Q&A predictions, comparison tables, calculation workshop, architecture diagrams |
| **Multiple Choice/选择** | "单选", "多选", "选择", "判断" | Flashcard-style facts, distractor analysis, concept traps, quick-reference tables |
| **Programming/编程** | "编程", "代码", "算法实现", "实验" | Code templates, algorithm pseudocode, test cases, debugging checklists |
| **Calculation/计算** | "计算", "数值", "推导", "证明" | Formula sheets with derivations, step-by-step templates, unit/dimension checks |
| **Mixed/混合** | Multiple types listed | Combine relevant document types from above |

### Phase 6 Activation Trigger (new in v1.1)

Phase 6 (Exam-Approach Guide) should be activated when **ALL** of the following are true:

1. The exam has **multiple distinct question types** with clear structural patterns (e.g., "简答 + 算法设计 + 计算" or "填空 + 选择 + 大题")
2. Past exams or syllabus indicate **consistent question templates** that repeat across years (e.g., "Sort always asks: given an array, write the process + pseudocode + complexity")
3. The course involves **procedural/algorithmic knowledge** where "how to answer" can be broken into step-by-step task chains (not just concept recall)

Typical courses that trigger Phase 6:
- CS Algorithm courses (算法设计与分析, Data Structures)
- Engineering calculation courses (信号与系统, 电路分析)
- Mixed-format STEM exams with problem-solving patterns

Courses that typically do NOT trigger Phase 6:
- Pure essay/memorization exams (历史, 文学)
- Pure multiple-choice exams without consistent patterns
- Exams where all questions are unique and unpredictable

### Course Type Detection (from content)

| Course Nature | Typical Content | Emphasis Shift |
|--------------|-----------------|----------------|
| **STEM/理工科** (CS, Math, Physics, Engineering) | Formulas, algorithms, architectures, derivations | Heavy on calculation workshop, formula cheat sheet, architecture diagrams |
| **Liberal Arts/文科** (History, Literature, Philosophy) | Timelines, figures, concepts, arguments | Heavy on timeline, concept maps, argument structures, quote banks |
| **Medical/医学** | Diseases, symptoms, treatments, anatomy | Heavy on comparison tables (disease differentiation), procedure flows, mnemonic devices |
| **Business/商科** (Economics, Management, Finance) | Models, cases, frameworks, calculations | Heavy on framework summaries, case analysis templates, formula + concept mix |
| **Language/语言** (Linguistics, Translation) | Rules, exceptions, structures, theories | Heavy on rule summaries, exception lists, structure diagrams |

### Dynamic Document Selection

Based on detected type, select from this menu:

**Base Documents** (always generate — 任何考试都有):
1. `README.md` — Study plan + file index
2. `期末复习指南.md` — Comprehensive chapter-by-chapter guide
3. `A4救命纸.md` — Cheat sheet (emergency condensed reference, always useful even if exam is closed-book for quick memorization)

**Essay-Optimized Additions** (问答题为主时追加):
4. `考试问答预测.md` — 30-50 predicted Q&A with model answers
5. `计算题精讲.md` — Calculation workshop with worked examples
6. `核心对比表.md` — Comparison tables for analysis questions
7. `发展时间线.md` — Timeline/flowchart memory aid
8. `完整性评估报告.md` — Coverage gap analysis

**Choice-Optimized Additions** (replace Q&A with):
3. `知识点速记卡.md` — Flashcard-style fact sheets
4. `易混辨析.md` — Concepts that are easily confused (distractor analysis)
5. `重点概念表.md` — Key terms with one-sentence definitions

**Programming-Optimized Additions**:
3. `代码模板.md` — Common algorithm templates and patterns
4. `实验步骤.md` — Lab procedure checklists
5. `调试清单.md` — Common bugs and fixes
6. `测试用例.md` — Typical input/output examples

**Calculation-Optimized Additions**:
3. `公式推导.md` — Full derivations of key formulas
4. `量纲检查.md` — Unit and dimension verification checklist
5. `常见陷阱.md` — Common calculation errors

---

## Phase 0: Input Collection

### Required Inputs
The user should provide:
- [ ] **Exam Syllabus** (分值分布, 题型, 开/闭卷, 考试时间)
- [ ] **All Course PPTs/PDFs** (ideally named with chapter numbers: `1-xxx.pptx`, `2-xxx.pptx`)
- [ ] **Past Exams** (if available — for question pattern analysis)

### Naming Convention
Prompt user to rename files before processing:
```
0-概述.pptx
1-文本表示.pptx
2-RNN.pptx
3-序列到序列.pptx
4-注意力.pptx
5-Transformer.pptx
6-GPT-BERT.pptx
7-LLM新技术.pptx
```

---

## Phase 1: Material Extraction (材料提取)

### 1.1 Syllabus Parsing
Extract from exam syllabus document:
- Total score, exam format (open/closed/semi-open), question types
- Chapter weight distribution (分值表)
- Cognitive level requirements per chapter (记忆/理解/分析/应用/综合)
- Time allocation recommendations

### 1.2 Slide Text Extraction
For each PPT file, extract per-slide text content including:
- Slide title and body text
- Table content (preserve structure)
- Notes/comments sections (often contain hidden explanations)
- Deep XML parsing for grouped shapes and hidden text nodes

**Tool**: Use Python `python-pptx` with XML deep traversal:
```python
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

# Surface extraction
for slide in prs.slides:
    for shape in slide.shapes:
        if hasattr(shape, 'text_frame'):
            print(shape.text_frame.text)

# Deep XML extraction for hidden content
from lxml import etree
for slide in prs.slides:
    tree = etree.fromstring(slide._element.xml)
    # Extract all text nodes, notes, grouped shapes
```

### 1.3 Image/Formula Catalog
- Count slides with images vs text-only
- Extract embedded images from formula-heavy slides
- Note which slides contain diagrams that need to be described in text form

### 1.4 Chapter Mapping
Create index mapping each PPT to syllabus chapter:
```
PPT File          → Chapter    → Weight → Skills Required
1-文本表示.pptx    → Ch2       → 10%    → 理解
2-RNN.pptx         → Ch3       → 15%    → 理解+分析
...
```

---

## Phase 2: Knowledge Curation (知识整理)

### 2.1 Content Classification
For each chapter, categorize extracted content:

| Category | Description | Exam Relevance |
|----------|-------------|----------------|
| **Core Concepts** | Definitions, key terms, milestones | High (记忆题) |
| **Formulas** | Mathematical formulas with variables explained | High (计算题) |
| **Architecture** | Model structures, data flow diagrams | High (分析题) |
| **Algorithms** | Step-by-step procedures (BPTT, RLHF, etc.) | High (应用题) |
| **Comparisons** | X vs Y tables, pros/cons | High (分析题) |
| **Numerical Examples** | Worked examples from slides | High (计算题) |
| **History/Timeline** | Development milestones | Medium (记忆题) |
| **Applications** | Use cases, downstream tasks | Medium (理解题) |

### 2.2 Formula Standardization (CRITICAL — LaTeX Wrapping Required)

All mathematical formulas MUST be wrapped in LaTeX math delimiters `$...$` or `$$...$$` so they display correctly in any math-aware markdown viewer (including Notion, Typora, VS Code with extensions).

**Rule**: Every formula — whether inline or display — must use `$...$` (inline) or `$$...$$` (display block).

**WRONG** (plain text — renders as literal characters):
```
h_{t-1}, C_t = f_t ⊙ C_{t-1}, √d_k, σ²
```

**CORRECT** (LaTeX wrapped — renders as proper math):
```markdown
$h_{t-1}$, $C_t = f_t \odot C_{t-1}$, $\sqrt{d_k}$, $\sigma^2$

$$
\text{Attention}(Q,K,V) = \text{softmax}\left(\frac{QK^T}{\sqrt{d_k}}\right)V
$$
```

**Common Conversions**:
| Plain Text | LaTeX |
|-----------|-------|
| `h_{t-1}` | `$h_{t-1}$` |
| `C̃_t` | `$\tilde{C}_t$` |
| `f_t ⊙ C_{t-1}` | `$f_t \odot C_{t-1}$` |
| `√d_k` | `$\sqrt{d_k}$` |
| `σ²` | `$\sigma^2$` |
| `Σ_j` | `$\sum_j$` or `$\Sigma_j$` |
| `α̂_ij` | `$\hat{\alpha}_{ij}$` |
| `∂E/∂W` | `$\frac{\partial E}{\partial W}$` |
| `→` | `$\rightarrow$` |
| `[h_{t-1}, x_t]` | `$[h_{t-1}, x_t]$` |

**Multi-line formulas** (use `$$...$$` with `aligned`):
```markdown
$$
\begin{aligned}
f_t &= \sigma(W_f \cdot [h_{t-1}, x_t] + b_f) \\
i_t &= \sigma(W_i \cdot [h_{t-1}, x_t] + b_i) \\
\tilde{C}_t &= \tanh(W_C \cdot [h_{t-1}, x_t] + b_C)
\end{aligned}
$$
```

**Variable definitions** (outside math mode):
```markdown
Where:
- $X$: input embedding matrix $[\text{seq\_len} \times d_{\text{model}}]$
- $W^Q, W^K, W^V$: learned weight matrices $[d_{\text{model}} \times d_k]$
- $d_k$: dimension per head ($d_{\text{model}} / \text{num\_heads}$)
```

**Verification**: After generating .md files, the AI must verify that every formula is wrapped in `$...$` or `$$...$$`. Any bare formula like `h_{t-1}` without delimiters is a bug and must be fixed.

### 2.3 Diagram-to-Text Conversion
For architecture diagrams (Transformer, LSTM, Encoder-Decoder):
- Create ASCII/text-based structural descriptions
- Annotate data flow direction
- Label all components (Q/K/V source, Mask positions, residual connections)

---

## Phase 3: Document Generation (动态生成文档)

> **CRITICAL**: Refer to the "Adaptive Configuration" section above. Generate ONLY the document types relevant to the detected exam type and course nature. Do not generate all documents blindly.

### Base Documents (Always Generate)

#### Document 1: README & Study Plan (README.md)
**Purpose**: Single double-sided A4 paper for exam (if semi-open)
**Constraints**: Must be hand-copyable; prioritize formulas and comparisons
**Structure**:
```markdown
# Course Final Exam A4 Cheat Sheet

## Side 1: Core Formulas & Architecture
- [ ] Most important formula (highest weight chapter)
- [ ] Second most important formula
- [ ] Architecture diagram (text-based)
- [ ] Key algorithm steps
- [ ] Calculation procedure templates

## Side 2: Comparison Tables & Keywords
- [ ] Top 5 comparison tables (X vs Y)
- [ ] Key terminology definitions
- [ ] Common "Why" questions and answers
- [ ] Timeline/milestones (if memory-heavy)
- [ ] Exam-specific tips
```
**Rule**: Only include content that fits on 2 sides of A4 with normal handwriting. Each item ≤ 3 lines.

### Document 2: Comprehensive Study Guide (期末复习指南.md)
**Purpose**: Complete chapter-by-chapter review
**Structure**:
```markdown
# Course Final Exam Review Guide

## Chapter N: Title (Weight% — Skill Level)
### 1. Core Concepts
- Definition list with key terms

### 2. Key Formulas (⚠️ Calculation Focus)
- Formula with variable explanations

### 3. Architecture/Algorithm
- Step-by-step description
- Text-based diagram

### 4. Comparison Tables
| Dimension | Method A | Method B |

### 5. Common Questions
- Predicted exam questions with model answers

### 6. Calculation Examples (if applicable)
- Worked numerical example from slides
```
**Rule**: Match chapter order to syllabus; emphasize high-weight chapters.

### Document 3: Q&A Predictions (考试问答预测.md)
**Purpose**: Practice with predicted exam questions
**Structure**:
```markdown
## Chapter X: Title (Weight%)

### Q1: [Memory question]
**Answer**: Concise, complete

### Q2: [Understanding question]
**Answer**: Explanation with examples

### Q3: [Analysis/Comparison question]
**Answer**: Structured comparison with evidence

### Q4: [Application/Calculation question]
**Answer**: Formula → Steps → Numerical result
```
**Rule**: Generate 5-8 questions per chapter; mix cognitive levels to match syllabus.

### Document 4: Calculation Workshop (计算题精讲.md)
**Purpose**: Step-by-step worked examples for all calculation types
**Structure**:
```markdown
## Type N: [Calculation Name]

### Example N-1
**Given**: [Specific numbers]
**Find**: [What's being asked]

**Solution**:
Step 1: [Apply formula]
Step 2: [Substitute values]
Step 3: [Compute result]
Step 4: [Interpret/verify]

**Answer**: [Final value with units]

### Example N-2: Conceptual
**Question**: Why do we [operation]?
**Answer**: [Explanation of mathematical/ML rationale]
```
**Rule**: Include at least one complete numerical example for each formula in A4 sheet.

### Document 5: Comparison Tables (核心对比表.md)
**Purpose**: Quick reference for "compare/contrast" analysis questions
**Structure**:
```markdown
# Core Comparison Tables

## 1. [Model A] vs [Model B] (⚠️ High Probability)
| Dimension | A | B |
| Name | | |
| Year/Institution | | |
| Architecture | | |
| Key Innovation | | |
| Strengths | | |
| Weaknesses | | |
| Use Cases | | |
```
**Rule**: Create tables for every "vs" mentioned in course; bold rows that are exam favorites.

### Document 6: Timeline/Flowchart (发展时间线.md 或 流程图.md)
**Purpose**: Memory aid for chronological or procedural content
**Structure**:
- Timeline: Year → Event → Significance
- Flowchart: Step → Sub-step → Key Detail

### Document 7: Study Plan README (README.md)
**Purpose**: Navigation, time allocation, learning strategy. When Phase 6 is active, this document MUST include BOTH study approaches so the user can choose.
**Structure**:
```markdown
# Review Package Guide

## File Inventory
| File | Purpose | When to Read |

## Study Plan (choose one approach, or do both in sequence)

### Approach A: Knowledge-First (按章节，适合时间充裕)
| Round | Duration | Focus | Files |
| Round 1 (Skeleton) | 2h | Structure + Formulas | Review Guide + A4 Sheet |
| Round 2 (Q&A) | 3h | Predicted questions | Q&A Predictions |
| Round 3 (Calculations) | 2h | Practice problems | Calculation Workshop |
| Round 4 (A4 Copy) | 1h | Memorize cheat sheet | A4 Sheet |

### Approach B: Exam-Question-First (按题型，适合时间紧张) [only when Phase 6 is active]
Detailed task chains per question type — see **01-全题型模版与解答指南.md**.
| Batch | Question Types | Score | Estimated Time |
| Batch 1 | Master Method + Recursion Tree + Insertion Sort | ~30 | 3h |
| Batch 2 | LCS + Matrix Chain | ~20-25 | 3h |
| ... | ... | ... | ... |

## Chapter Priority (by exam weight)
1. Chapter X (20%) — Start here
2. Chapter Y (20%) — Most formulas
...
```

---

## Phase 4: Quality Verification (质量检查)

### 4.1 Coverage Checklist
For each chapter in syllabus:
- [ ] All memory-level points covered in Review Guide
- [ ] All understanding-level concepts explained with examples
- [ ] All analysis-level comparisons have dedicated tables
- [ ] All application-level procedures have step-by-step breakdowns
- [ ] All formulas have variable definitions
- [ ] All diagrams have text-based descriptions
- [ ] At least one calculation example per formula-heavy chapter

### 4.2 Completeness Report
Generate ` completeness_assessment_report.md ` with:
```markdown
# Review Material Completeness Report

## Coverage Summary
| Chapter | Weight | Coverage | Missing | Priority |

## Key Gaps Identified
1. [Gap description] → [Action needed]

## Recommended Additions
- [ ] Item to add
```

### 4.3 Formula Verification
- Verify each formula against original PPT images
- Confirm LaTeX notation renders correctly
- Check numerical examples produce correct results

---

## Phase 5: Multi-Format Export (多格式输出)

### Markdown (Source — Primary Output)
- Keep all `.md` files in `review/` directory
- Maintain clean directory structure:
  ```
  course-folder/
  ├── source/          ← Original PPTs + syllabus
  ├── review/          ← Generated documents (.md)
  └── temp/            ← Extraction artifacts (deletable)
  ```

### Notion Import
- Merge all `.md` files into single `course_review_for_notion.md`
- Notion → Import → Markdown & CSV
- Each H1 becomes a toggleable section

---

## Phase 6: Exam-Approach Guide Generation (Optional — 题型导向复习指南)

> **Activation**: See "Phase 6 Activation Trigger" in Adaptive Configuration above. Only generate when the exam has identifiable question-type patterns.
> **Prerequisite**: All Phase 3 knowledge documents must be complete first. This phase depends on them.

### Purpose

Transform the chapter-organized knowledge documents into an **exam-question-first** study approach. Instead of "read Chapter 1, then Chapter 2", the user gets: "To master Master Method questions: Step 1 read §1.4 → Step 2 do Example 1-1 → Step 3 verify with 2024 Past Paper A-1".

This is especially valuable when:
- The user has limited time and wants to directly practice "what will be on the paper"
- The exam has stable question templates that repeat across years
- The user finds chapter-by-chapter reading too slow and unfocused

### 6.1 Question-Type Pattern Extraction

Analyze ALL available materials (syllabus, past exams, teacher hints, extracted course content) to identify every distinct question type. For each type, determine:

| Attribute | Source | Example |
|-----------|--------|---------|
| **What it looks like on the paper** | Past exams + syllabus | "给定递归式 T(n)=aT(n/b)+f(n)，求解" |
| **Sub-questions always asked** | Past exam patterns | "(1) 写出a,b,n^{log_b a} (2) 比较f(n) (3) 结论" |
| **Score weight** | Syllabus + past exams | 10 points |
| **Frequency/certainty** | Teacher hints + past data | 🔴必考 / 🟠高频 / 🟡可能 / 🟢边缘 |
| **Required skills** | Cognitive level mapping | 记忆 / 理解 / 计算 / 综合 |

### 6.2 Document Generation: Question-Type Template Guide (01-全题型模版与解答指南.md)

**Structure for each question type**:
```markdown
## 📝 题型N：[Question Type Name]（分值）

### 🔴/🟠/🟡 Priority Level

### 卷子上的样子
[Complete, realistic exam question as it appears on paper — with actual data, tables, diagrams]

### 你怎么答（答题标准步骤）
[Step-by-step: exactly what to write, in what order, with what format]

### 答题检查清单
- [ ] Required element 1
- [ ] Required element 2

### 例题训练
| 例题 | 数据 | 在哪个文件练 |
|------|------|------------|
| 例N-1 | [specific data] | 计算题精讲 第X-Y行 |
| 真题 | [past paper data] | 2024大题复习题 第X-Y行 |

### 基础知识来源
| 要理解的概念 | 详细讲解位置 |
|------------|------------|
| Concept A | 期末复习指南 §X.Y |
| Pseudocode | 伪代码模板集 第X-Y行 |
| Quick Ref | A4救命纸 第X-Y行 |
```

**CRITICAL rules**:
1. Every "where to find" reference MUST include exact file name + section/chapter/line range (not just "see X file")
2. Priority levels MUST align with teacher hints and past exam data — do not guess
3. The "卷子上的样子" section must be a COMPLETE, realistic question, not just "给你一个数组，让你排序" — include actual numbers, actual tables, actual diagrams
4. Each question type must link to ALL relevant knowledge sources (guide + pseudocode + calculation workshop + Q&A + comparison tables + A4 sheet)
5. Include a "快速索引" section at the end mapping "I want to practice X" → exact file + line range

### 6.3 Mastery Task Chains (掌握路径)

For each question type, generate a step-by-step mastery chain. This is DIFFERENT from the answer template — it tells the user how to LEARN to answer this type, not how to answer it on the exam.

**Format**:
```markdown
### 🔴 题型X：[Name]（Score）

**最终要求**：[What the user should be able to do independently]

**掌握路径：**

| 步骤 | 做什么 | 具体位置 |
|:---:|--------|---------|
| 1 | [Understand the core concept] | [File §Section, line range] |
| 2 | [Follow a worked example step by step] | [File, example number, line range] |
| 3 | [Independent practice with different data] | [File, exercise number, line range] |
| 4 | [Verify with past paper question] | [File, past paper ID, line range] |

**⚠️ 陷阱/扣分点**：[Common mistakes specific to this question type]
```

**Rules for task chains**:
- Each step must reference EXACT locations, not general file names
- Steps must be ordered: understand → follow-along → independent-practice → verify
- Include common traps/扣分点 specific to that question type
- At the end, provide a "自检标准" (self-check criteria): what does "mastered" look like for this type?

### 6.4 Batch Execution Order (执行顺序建议)

After all task chains are generated, organize them into execution batches ordered by score weight:

```markdown
| 批次 | 题型 | 分值 | 预计耗时 |
|:---:|------|:---:|:---:|
| 第1批 | [Highest score types] | ~X分 | Xh |
| 第2批 | [Next priority] | ~X分 | Xh |
...
| 考前 | A4纸最后一遍 + 填空/选择再过一遍 | — | 0.5h |
```

### 6.5 Update README with Dual-Plan Navigation

When Phase 6 is complete, update `README.md` to present BOTH study approaches:

```markdown
## Study Plan (choose one)

### Approach A: Knowledge-First（按章节，适合时间充裕）
[Existing chapter/timeline schedule]

### Approach B: Exam-Question-First（按题型，适合时间紧张）
详细任务链见 **01-全题型模版与解答指南.md**。
[Batch execution order table from 6.4]
```

Also update the file inventory to include the new `01-全题型模版与解答指南.md` as a ⭐ priority file.

### 6.6 Quality Check for Phase 6

- [ ] Every question type has a realistic "卷子上的样子" with actual data
- [ ] Every question type links to all relevant knowledge sources with exact locations
- [ ] Task chain steps are ordered: understand → follow → practice → verify
- [ ] Common traps/扣分点 are identified for each type
- [ ] Self-check criteria are concrete and testable (not vague like "understood")
- [ ] Batch execution order respects score weight priority
- [ ] README.md updated with dual-plan navigation
- [ ] No hallucinated question types — all types are backed by syllabus/past exams/teacher hints

---

## Cross-Platform Installation Guide

### OpenCode
```bash
# Option 1: Project-level (recommended)
mkdir -p .opencode/skills/exam-review-sop
cp SKILL.md .opencode/skills/exam-review-sop/SKILL.md
# Restart opencode

# Option 2: Global
mkdir -p ~/.config/opencode/skills/exam-review-sop
cp SKILL.md ~/.config/opencode/skills/exam-review-sop/SKILL.md
```

### Claude Code
```bash
mkdir -p ~/.claude/skills/exam-review-sop
cp SKILL.md ~/.claude/skills/exam-review-sop/SKILL.md
# Claude Code auto-detects skills in ~/.claude/skills/
```

### Claude Desktop (Web)
Use **Projects** feature:
1. Create new Project
2. In Project Instructions, paste the full SKILL.md content
3. Upload course materials to Project files

### GitHub Copilot / Codex
Create `.github/copilot-instructions.md`:
```markdown
# Copilot Instructions: Exam Review Generator

When the user uploads course materials (PPTs, syllabus) and mentions exam preparation,
follow this workflow:

1. Parse syllabus for chapter weights and question types
2. Extract text from all uploaded PPTs using python-pptx
3. Organize content by chapter and cognitive level
4. Generate documents: A4 Cheat Sheet, Study Guide, Q&A, Calculations, Comparisons, Timeline, README
5. Verify coverage against syllabus
6. (Optional) If exam has clear question-type patterns, generate exam-approach guide (Phase 6)
7. Output in Markdown format
```

---

## Prompt Templates for Users

### Quick Start Prompt
"帮我整理期末复习资料。这是考试大纲 [upload]，这是所有课件 [upload]。考试是半开卷，可以带一张A4纸。"

### Advanced Prompt
"请按 exam-review-sop 流程执行：
1. 解析大纲分值分布
2. 提取所有PPT文字和图片
3. 生成7份复习文档
4. 检查覆盖率并补充遗漏
5. 输出为 Markdown"

### Follow-up Prompts
- "对比现有复习资料和PPT，检查有没有遗漏"
- "把A4纸内容再精简一些，确保两面能写完"
- "合并成一个文件方便导入Notion"
- "帮我按照题型整理一份复习指南，我想直接按卷子上的题型来复习"

---

## Quality Principles

1. **Syllabus-Driven**: Always prioritize chapters by exam weight, not by PPT page count
2. **Formula-First**: Any formula in PPT must appear in A4 sheet + Study Guide + Calculations
3. **Text-Complete**: If PPT has images/diagrams, create text descriptions so the AI can verify coverage without image reading
4. **Minimal Intrusion**: Don't add content not in course materials (unless standard textbook supplement for formula notation)
5. **Calculation Verified**: Every numerical example must be recalculated to confirm correctness
6. **Source Preserved**: Keep original PPTs and extracted text for reference; never overwrite source materials

---

## Troubleshooting

| Problem | Solution |
|---------|----------|
| PPT has many images, text extraction incomplete | Use XML deep parsing; extract embedded images and describe them |
| Formula images not readable by AI | Supplement with standard textbook notation based on semantic text descriptions |
| User has no exam outline | Ask for course schedule/announcements; infer weights from lecture hours or PPT slide counts |
| Syllabus is image-only (scanned PDF) | Use OCR or ask user to type key information (weights, question types) |
| Content too much for A4 | Prioritize: formulas > comparisons > algorithms > definitions |

---

*Version: 1.1*
*Compatible with: OpenCode ≥ 0.x, Claude Code, Claude Desktop, GitHub Copilot*
*License: MIT (or user-specified)*
