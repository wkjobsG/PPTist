import asyncio
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List

from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware

# 导入PPT模块
from ppt import (
    PPTOutlineRequest,
    PPTContentRequest,
    AIWritingRequest,
    PPTSaveRequest,
    build_outline_chain,
    build_ppt_content_chain,
    build_ai_writing_chain,
    save_ppt_json,
    STATIC_DIR
)

# 创建FastAPI应用（用于兼容原有接口）
app = FastAPI(title="AI PPT MCP Server", version="1.0.0")

# 添加CORS中间件
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # 允许所有来源，生产环境建议指定具体域名
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.mount("/files", StaticFiles(directory=str(STATIC_DIR)), name="files")

# 原有FastAPI接口（用于兼容性）
@app.post("/tools/aippt_outline")
async def generate_ppt_outline_stream(request: PPTOutlineRequest):
    chain = build_outline_chain(request.model)

    async def token_stream():
        try:
            print("[调试] /tools/aippt_outline 调用参数：", {
                "require": request.content,
                "language": request.language,                           
                "model": request.model
            })
            async for chunk in chain.astream({
                "require": request.content,
                "language": request.language
            }):
                yield chunk
        except Exception as e:
            import traceback
            print("[AIPPT Outline Streaming Error]", e)
            traceback.print_exc()
            yield f"\n[ERROR]: {str(e)}\n"

    return StreamingResponse(token_stream(), media_type="text/event-stream")

@app.post("/tools/aippt")
async def generate_ppt_content_stream(request: PPTContentRequest):
    chain = build_ppt_content_chain(request.model)

    full_output: str = await chain.ainvoke({
        "language": request.language,
        "content": request.content
    })

    lines = [line.strip() for line in full_output.strip().split("\n") if line.strip()]

    async def stream_pages():
        for line in lines:
            yield line + "\n\n"
            await asyncio.sleep(0.5)

    return StreamingResponse(stream_pages(), media_type="text/event-stream")

@app.post("/tools/ai_writing")
async def generate_ai_writing_stream(request: AIWritingRequest):
    chain = build_ai_writing_chain(request.model)

    full_output: str = await chain.ainvoke({
        "command": request.command,
        "content": request.content
    })

    lines = [line.strip() for line in full_output.strip().split("\n") if line.strip()]

    async def stream_pages():
        for line in lines:
            yield line + "\n\n"
            await asyncio.sleep(0.5)

    return StreamingResponse(stream_pages(), media_type="text/event-stream")

# API文档和状态接口
@app.get("/")
async def root():
    return {
        "name": "AI PPT MCP Server",
        "version": "1.0.0",
        "mcp_tools": [
            "generate_ppt_outline",
            "generate_ppt_content",
            "optimize_ppt_content",
            "save_ppt_to_file",
            "generate_complete_ppt"
        ],
        "legacy_endpoints": [
            "/tools/aippt_outline",
            "/tools/aippt",
            "/tools/ai_writing"
        ]
    }

@app.get("/health")
async def health_check():
    return {"status": "healthy", "timestamp": datetime.now().isoformat()}

# 启动函数
def run_fastapi_server():
    """运行FastAPI服务器"""
    import uvicorn
    print("启动 AI PPT FastAPI Server...")
    print("兼容API地址: http://127.0.0.1:5000")
    print("API文档: http://127.0.0.1:5000/docs")
    print("健康检查: http://127.0.0.1:5000/health")
    
    # 启动FastAPI服务器用于兼容接口
    uvicorn.run("main:app", host="0.0.0.0", port=5000, reload=False)

# 主程序入口
if __name__ == "__main__":
    # 可以选择运行FastAPI服务器或MCP服务器
    import sys
    
    if len(sys.argv) > 1 and sys.argv[1] == "mcp":
        # 运行MCP服务器
        from mcp_starter import run_mcp_server
        run_mcp_server()
    else:
        # 默认运行FastAPI服务器
        run_fastapi_server()