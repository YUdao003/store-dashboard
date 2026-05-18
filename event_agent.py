#!/usr/bin/env python3
"""活动公司多岗位 Agent 系统 — 四岗并行，汇总成完整活动方案"""

import asyncio
import argparse
import sys
from datetime import datetime

import anthropic

MODEL = "claude-sonnet-4-6"

# ─────────────────────────── 系统提示词 ───────────────────────────

PLANNING_SYSTEM = """你是一名资深活动策划市场总监，拥有10年以上大型活动策划经验。
你负责从市场角度为活动制定完整的策划方案，包括：
- 活动概念与主题创意
- 市场定位与目标受众分析
- 营销传播策略（线上+线下渠道组合）
- 预期KPI与效果评估指标（量化）
- 预算大纲与资源配置建议

输出要求：
- 使用清晰的Markdown二级/三级标题结构
- 数据和建议要具体可执行，避免空话
- KPI必须给出具体数字范围
- 语言专业、简洁有力"""

DESIGN_SYSTEM = """你是一名专业活动视觉设计总监，擅长品牌视觉体系搭建与活动物料设计。
你负责为活动制定完整的设计方案，包括：
- 视觉识别系统（VI）方向与风格定义
- 主视觉设计概念与色彩方案（给出具体色值）
- 所有活动物料清单及规格说明（名称/尺寸/数量/工艺/交付时间）
- 设计排期与交付节点
- 设计规范要点（字体、排版、禁止用法）

输出要求：
- 物料清单使用表格格式（| 物料名 | 规格 | 数量 | 用途 | 交付节点 |）
- 色彩方案给出 HEX 色值
- 字体推荐具体到字体名称
- 便于与客户及供应商直接沟通"""

EXECUTION_SYSTEM = """你是一名经验丰富的活动执行总监，专注于大型活动的落地执行管理。
你负责制定详细的执行清单，供真实执行团队人员参照操作，包括：
- 按时间节点划分的执行清单（T-30天/T-7天/T-1天/活动当天/T+1天）
- 每项任务的具体操作步骤、负责角色、完成标准
- 现场布置与设备清单
- 应急预案与风险处理
- 供应商管理与对接要点

严格要求：
1. 所有任务必须使用 - [ ] 格式（可勾选的复选框）
2. 每条任务末尾注明负责角色，格式：【角色名】
3. 时间节点用二级标题（## T-30天）分组
4. 内容足够详细，执行人员无需额外培训即可按清单操作
5. 活动当天需细化到小时级时间节点"""

CUSTOMER_SERVICE_SYSTEM = """你是一名专业的活动客服跟进总监，擅长全流程客户关系管理。
你负责制定活动的客服跟进方案，包括：
- 活动前：报名咨询响应SOP、确认通知、提醒机制
- 活动中：现场客服支持流程、突发问题处理SOP
- 活动后：满意度调研、反馈收集、关系维护节点
- 常见FAQ问题库（含标准回复话术，可直接使用）
- 客服响应时间标准与问题升级机制
- 后续跟进时间线与沟通模板

输出要求：
- FAQ格式：**Q: 问题** / **A: 回复话术**
- 沟通模板要可直接复制发送（含称呼、正文、落款）
- 流程用有序列表，时间节点用标题分组
- 话术语气专业、温暖、高效"""

SYNTHESIS_SYSTEM = """你是活动公司的CEO，负责审阅各部门方案并向客户提交执行摘要。
请根据四个部门的完整方案，撰写一份约200字的执行摘要，要求：
- 第一句点明活动核心价值主张
- 各岗位最关键行动要点各一句（策划/设计/执行/客服）
- 指出1-2个需要跨部门协调的关键节点
- 语言精炼、有高度，适合向客户汇报
- 不要使用列表，用段落式流畅表达"""


# ─────────────────────────── Agent 函数 ───────────────────────────

async def run_planning_agent(client: anthropic.AsyncAnthropic, brief: str) -> str:
    try:
        response = await client.messages.create(
            model=MODEL,
            max_tokens=4096,
            system=[{
                "type": "text",
                "text": PLANNING_SYSTEM,
                "cache_control": {"type": "ephemeral"},
            }],
            messages=[{
                "role": "user",
                "content": f"请根据以下活动简报，制定完整的策划市场方案：\n\n{brief}",
            }],
        )
        text = next((b.text for b in response.content if b.type == "text"), "（无输出）")
        print("✓ 策划市场岗 完成", flush=True)
        return f"## 策划市场方案\n\n{text}"
    except anthropic.APIError as e:
        print(f"✗ 策划市场岗 失败: {e}", flush=True)
        return f"## 策划市场方案\n\n> ⚠️ Agent 调用失败：{e}"


async def run_design_agent(client: anthropic.AsyncAnthropic, brief: str) -> str:
    try:
        response = await client.messages.create(
            model=MODEL,
            max_tokens=4096,
            system=[{
                "type": "text",
                "text": DESIGN_SYSTEM,
                "cache_control": {"type": "ephemeral"},
            }],
            messages=[{
                "role": "user",
                "content": f"请根据以下活动简报，制定完整的视觉设计方案：\n\n{brief}",
            }],
        )
        text = next((b.text for b in response.content if b.type == "text"), "（无输出）")
        print("✓ 设计岗 完成", flush=True)
        return f"## 视觉设计方案\n\n{text}"
    except anthropic.APIError as e:
        print(f"✗ 设计岗 失败: {e}", flush=True)
        return f"## 视觉设计方案\n\n> ⚠️ Agent 调用失败：{e}"


async def run_execution_agent(client: anthropic.AsyncAnthropic, brief: str) -> str:
    try:
        response = await client.messages.create(
            model=MODEL,
            max_tokens=6000,
            system=[{
                "type": "text",
                "text": EXECUTION_SYSTEM,
                "cache_control": {"type": "ephemeral"},
            }],
            messages=[{
                "role": "user",
                "content": (
                    f"请根据以下活动简报，制定详细执行清单。\n"
                    f"按 T-30天/T-7天/T-1天/活动当天（细化到小时）/T+1天 分阶段，"
                    f"每条任务使用 - [ ] 格式并注明【负责角色】：\n\n{brief}"
                ),
            }],
        )
        text = next((b.text for b in response.content if b.type == "text"), "（无输出）")
        print("✓ 执行岗 完成", flush=True)
        return f"## 执行方案\n\n{text}"
    except anthropic.APIError as e:
        print(f"✗ 执行岗 失败: {e}", flush=True)
        return f"## 执行方案\n\n> ⚠️ Agent 调用失败：{e}"


async def run_customer_service_agent(client: anthropic.AsyncAnthropic, brief: str) -> str:
    try:
        response = await client.messages.create(
            model=MODEL,
            max_tokens=4096,
            system=[{
                "type": "text",
                "text": CUSTOMER_SERVICE_SYSTEM,
                "cache_control": {"type": "ephemeral"},
            }],
            messages=[{
                "role": "user",
                "content": f"请根据以下活动简报，制定完整的客服跟进方案：\n\n{brief}",
            }],
        )
        text = next((b.text for b in response.content if b.type == "text"), "（无输出）")
        print("✓ 客服跟进岗 完成", flush=True)
        return f"## 客服跟进方案\n\n{text}"
    except anthropic.APIError as e:
        print(f"✗ 客服跟进岗 失败: {e}", flush=True)
        return f"## 客服跟进方案\n\n> ⚠️ Agent 调用失败：{e}"


async def run_synthesis(
    client: anthropic.AsyncAnthropic,
    brief: str,
    sections: list[str],
) -> str:
    combined = "\n\n---\n\n".join(sections)
    try:
        response = await client.messages.create(
            model=MODEL,
            max_tokens=1024,
            system=[{
                "type": "text",
                "text": SYNTHESIS_SYSTEM,
                "cache_control": {"type": "ephemeral"},
            }],
            messages=[{
                "role": "user",
                "content": (
                    f"活动简报：\n{brief}\n\n"
                    f"四个部门方案如下：\n\n{combined}\n\n"
                    f"请撰写约200字的执行摘要。"
                ),
            }],
        )
        text = next((b.text for b in response.content if b.type == "text"), "（无输出）")
        print("✓ 执行摘要 完成", flush=True)
        return text
    except anthropic.APIError as e:
        print(f"✗ 执行摘要 失败: {e}", flush=True)
        return f"> ⚠️ 摘要生成失败：{e}"


# ─────────────────────────── 主流程 ───────────────────────────

async def run(brief: str) -> None:
    client = anthropic.AsyncAnthropic()

    print("\n🚀 四个岗位 Agent 并行启动中...\n", flush=True)

    planning, design, execution, cs = await asyncio.gather(
        run_planning_agent(client, brief),
        run_design_agent(client, brief),
        run_execution_agent(client, brief),
        run_customer_service_agent(client, brief),
    )

    print("\n📝 生成执行摘要...\n", flush=True)

    sections = [planning, design, execution, cs]
    summary = await run_synthesis(client, brief, sections)

    now = datetime.now()
    timestamp = now.strftime("%Y%m%d_%H%M%S")
    date_str = now.strftime("%Y年%m月%d日 %H:%M")

    report_parts = [
        f"# 活动策划全案报告\n\n> 生成时间：{date_str}\n",
        f"## 执行摘要\n\n{summary}\n",
        f"## 活动简报\n\n> {brief}\n",
        planning,
        design,
        execution,
        cs,
    ]
    report = "\n\n---\n\n".join(report_parts)

    filename = f"event_report_{timestamp}.md"
    with open(filename, "w", encoding="utf-8") as f:
        f.write(report)

    print(f"\n{'=' * 60}")
    print(report)
    print(f"{'=' * 60}\n")
    print(f"✅ 报告已保存至：{filename}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="活动公司多岗位 Agent 系统",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="示例：\n  python event_agent.py --brief \"2024年品牌发布会，500人规模\"",
    )
    parser.add_argument("--brief", type=str, help="活动简报（直接传入，省略则交互输入）")
    args = parser.parse_args()

    if args.brief:
        brief = args.brief.strip()
    else:
        print("请输入活动简报（完成后按 Enter 两次确认）：")
        lines: list[str] = []
        try:
            while True:
                line = input()
                if not line and lines:
                    break
                lines.append(line)
        except EOFError:
            pass
        brief = "\n".join(lines).strip()

    if not brief:
        print("错误：活动简报不能为空", file=sys.stderr)
        sys.exit(1)

    asyncio.run(run(brief))


if __name__ == "__main__":
    main()
