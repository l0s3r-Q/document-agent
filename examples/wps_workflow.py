"""
wps-office 融合端到端示例：document-agent 离线生成 + wps-office 在线编辑/导出

场景：生成一份安全生产检查通知 → 质量检查 → 转 PDF → 在 WPS 中打开

依赖：
- document-toolkit MCP（本仓库 mcp/document-toolkit/server.py）— 离线生成
- wps-office MCP（外部，见 docs/wps-office-integration.md）— 在线编辑
- 运行前：两个 MCP 均已注册到 Reasonix / Claude Code 等工具

本脚本用 MCP 协议直连两个 server 演示完整链路；
在 AI 工具中，等价操作为按顺序调用各 MCP 工具。
"""

import json
import os
import subprocess
import sys
import time

# ============ 配置 ============
DOC_TOOLKIT = r"C:\Users\36078\skills\mcp\document-toolkit\server.py"
WPS_MCP = r"C:\Users\36078\skills\mcp\wps-office\dist\index.js"
OUT_DIR = os.path.join(os.path.expanduser("~"), "Desktop", "文档")
DOCX_PATH = os.path.join(OUT_DIR, "安全生产检查通知.docx")
PDF_PATH = os.path.join(OUT_DIR, "安全生产检查通知.pdf")


def call_mcp(server_cmd, server_args, tool, arguments, env=None):
    """通过 MCP stdio 协议调用工具（最小客户端实现）。env 可传额外环境变量。"""
    full_env = dict(os.environ)
    if env:
        full_env.update(env)
    proc = subprocess.Popen(
        [server_cmd, *server_args],
        stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        text=True, encoding="utf-8", env=full_env,
    )

    def send(obj):
        proc.stdin.write(json.dumps(obj) + "\n")
        proc.stdin.flush()

    def recv(expected_id):
        while True:
            line = proc.stdout.readline()
            if not line:
                return None
            try:
                msg = json.loads(line)
            except Exception:
                continue
            if msg.get("id") == expected_id:
                return msg

    send({"jsonrpc": "2.0", "id": 1, "method": "initialize",
          "params": {"protocolVersion": "2024-11-05", "capabilities": {},
                     "clientInfo": {"name": "wps-workflow-demo", "version": "1.0"}}})
    recv(1)
    send({"jsonrpc": "2.0", "id": 2, "method": "tools/call",
          "params": {"name": tool, "arguments": arguments}})
    msg = recv(2)
    proc.terminate()
    if msg is None:
        return {"error": "no response"}
    if "error" in msg:
        return {"error": msg["error"]}
    content = msg.get("result", {}).get("content", [])
    text = "".join(c.get("text", "") for c in content if c.get("type") == "text")
    is_err = msg.get("result", {}).get("isError")
    return {"ok": not is_err, "text": text}


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    print("=" * 60)
    print("wps-office 融合示例：生成 → 质量检查 → 转 PDF → 打开")
    print("=" * 60)

    # 1. document-toolkit: 生成 docx（离线，无需 WPS）
    spec = {
        "doc_type": "notice",
        "title": "关于开展安全生产检查的通知",
        "sections": [
            {"type": "paragraph",
             "text": "各部门、各车间：为落实安全生产主体责任，防范各类安全事故，现决定开展全厂安全生产专项检查，现将有关事项通知如下。"},
            {"type": "heading1", "text": "一、检查时间"},
            {"type": "paragraph", "text": "2026年8月10日至8月20日。"},
            {"type": "heading1", "text": "二、检查内容"},
            {"type": "list", "items": ["消防设施设备完好情况", "用电用气安全", "特种设备运行状态", "安全教育培训记录"]},
            {"type": "heading1", "text": "三、工作要求"},
            {"type": "paragraph", "text": "各部门要高度重视，对照检查内容逐项自查，发现隐患立即整改。检查结果于8月21日前书面报送安全办。"},
            {"type": "paragraph", "text": "特此通知。"},
        ],
    }
    print(f"\n[1/4] document-toolkit.build_docx → {DOCX_PATH}")
    r = call_mcp("python", [DOC_TOOLKIT], "build_docx",
                 {"spec_json": json.dumps(spec, ensure_ascii=False), "output_path": DOCX_PATH})
    print("    ", r.get("text", r)[:120])

    # 2. document-toolkit: 质量检查
    print(f"\n[2/4] document-toolkit.quality_check")
    r = call_mcp("python", [DOC_TOOLKIT], "quality_check", {"path": DOCX_PATH})
    print("    ", r.get("text", r)[:200])

    # 3. document-toolkit: 转 PDF（Word COM → WPS COM → LibreOffice 降级链）
    print(f"\n[3/4] document-toolkit.convert_to_pdf → {PDF_PATH}")
    r = call_mcp("python", [DOC_TOOLKIT], "convert_to_pdf", {"docx_path": DOCX_PATH, "output_path": PDF_PATH})
    print("    ", r.get("text", r)[:120])

    # 4. wps-office: 在 WPS 中打开（在线预览/编辑）
    # 说明：Reasonix/Claude Code 等工具已注册 wps-office MCP 时，直接调用
    #       `wps_word_open_document(filePath=...)` 即可，无需独立进程。
    #       独立进程仅用于演示，且需 WPS 运行 + Word 组件已打开。
    print(f"\n[4/4] wps-office.wps_word_open_document（Reasonix 已注册时直接调用 MCP 工具）")
    if os.path.exists(r"C:\Users\36078\AppData\Roaming\reasonix\config.toml"):
        print("    提示: 在 AI 工具中调用 wps-office 的 `wps_word_open_document`，参数 filePath。")
        print("    前提: WPS 已运行且 Word 组件已打开（本脚本独立进程会与 Reasonix 的 58891 端口冲突，故跳过实际调用）。")
    else:
        r = call_mcp("node", [WPS_MCP], "wps_word_open_document", {"filePath": DOCX_PATH},
                     env={"WPS_USE_POLL": "1"})
        print("    ", r.get("text", r)[:150])

    print("\n完成。生成的文档：")
    for p in (DOCX_PATH, PDF_PATH):
        if os.path.exists(p):
            print(f"  [OK] {p} ({os.path.getsize(p)} bytes)")
        else:
            print(f"  [WARN] {p} not generated")

    print("\n后续可继续（按需取消注释）：")
    print("  # wps-office: 在 WPS 中美化已打开的文档")
    print("  # call_mcp('node', [WPS_MCP], 'wps_word_set_font', {'font_name': '微软雅黑', 'font_size': 12, 'range': 'all'})")
    print("  # wps-office: 导出为 PDF（在线方式）")
    print("  # call_mcp('node', [WPS_MCP], 'wps_convert_to_pdf', {'outputPath': PDF_PATH})")


if __name__ == "__main__":
    main()
