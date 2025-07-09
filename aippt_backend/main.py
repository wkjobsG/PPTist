from fastapi import FastAPI
from fastapi.responses import StreamingResponse
from pydantic import BaseModel, Field
import asyncio
import os
from dotenv import load_dotenv

from langchain.prompts import PromptTemplate
from langchain_core.output_parsers import StrOutputParser
from langchain_openai import ChatOpenAI

# 加载.env环境变量
load_dotenv()

app = FastAPI()

# PPT大纲生成Prompt
def get_outline_prompt():
    template = """你是用户的PPT大纲生成助手，请根据下列主题生成章节结构。

输出格式为：
# PPT标题（只有一个）
## 章的名字
### 节的名字
- 内容1
- 内容2
- 内容3
### 节的名字
- xxxxx
- xxxxx
- xxxxx
### 节的名字
- xxxxx
- xxxxx
- xxxxx

这是生成要求：{require}\n
这是生成的语言要求：{language}
"""
    return PromptTemplate.from_template(template)

def build_outline_chain(model_name: str = None):
    api_key = os.environ.get("OPENAI_API_KEY")
    base_url = os.environ.get("OPENAI_BASE_URL")
    default_model = os.environ.get("AIPPT_MODEL", "gpt-4o")
    model = model_name or default_model
    llm = ChatOpenAI(
        temperature=0.7,
        model=model,
        openai_api_key=api_key,
        base_url=base_url
    )
    return get_outline_prompt() | llm | StrOutputParser()

# PPT内容生成Prompt
def get_ppt_content_prompt():
    ppt_content_template = """
你是一个专业的PPT内容生成助手，请根据给定的大纲内容和原始要求，生成完整的PPT页面内容结构。

页面类型包括：
- 封面页："cover"
- 目录页："contents"
- 内容页："content"
- 过渡页："transition"
- 结束页："end"

输出格式要求如下：
- 每一页为一个独立 JSON 对象
- 每个 JSON 对象写在**同一行**
- 页面之间用两个换行符分隔
- 不要添加任何注释或解释说明
- 如果用户要求根据提供的材料生成ppt，尽量把用户提供的材料内容充分利用

示例格式（注意每个 JSON 占一行）：

{{"type": "cover", "data": {{ "title": "接口相关内容介绍", "text": "了解接口定义、设计与实现要点" }}}}

{{"type": "contents", "data": {{ "items": ["接口定义概述", "接口分类详情", "接口设计原则"] }}}}

{{"type": "transition", "data": {{ "title": "接口定义", "text": "开始介绍接口的基本含义" }}}}

{{"type": "content", "data": {{ "title": "接口定义", "items": [ {{ "title": "基本概念", "text": "接口是系统中模块通信的协议" }}, {{ "title": "作用", "text": "促进模块解耦，提高系统灵活性" }} ] }}}}

{{"type": "end"}}

请根据以下信息生成 PPT 内容：

语言：{language}
大纲内容和原始要求：{content}
"""
    return PromptTemplate.from_template(ppt_content_template)

def build_ppt_content_chain(model_name: str = None):
    api_key = os.environ.get("OPENAI_API_KEY")
    base_url = os.environ.get("OPENAI_BASE_URL")
    default_model = os.environ.get("AIPPT_MODEL", "gpt-4o")
    model = model_name or default_model
    llm = ChatOpenAI(
        temperature=0.7,
        model=model,
        openai_api_key=api_key,
        base_url=base_url
    )
    return get_ppt_content_prompt() | llm | StrOutputParser()

# PPT内容优化Prompt
def get_ai_writing_prompt():
    ai_writing_template = """
你是一个专业的PPT内容优化助手，请根据给定的优化指令和原始内容，生成优化后的内容。

优化指令：{command}
原始内容：{content}
"""
    return PromptTemplate.from_template(ai_writing_template)

def build_ai_writing_chain(model_name: str = None):
    api_key = os.environ.get("OPENAI_API_KEY")
    base_url = os.environ.get("OPENAI_BASE_URL")
    default_model = os.environ.get("AIPPT_MODEL", "gpt-4o")
    model = model_name or default_model
    llm = ChatOpenAI(
        temperature=0.7,
        model=model,
        openai_api_key=api_key,
        base_url=base_url
    )
    return get_ai_writing_prompt() | llm | StrOutputParser()

class PPTOutlineRequest(BaseModel):
    model: str = Field(None, description="使用的模型名称，例如 gpt-4o 或 gpt-4o-mini")
    language: str = Field(..., description="生成内容的语言，例如 中文、English")
    content: str = Field(..., description="生成的要求")
    stream: bool = True

class PPTContentRequest(BaseModel):
    model: str = Field(None, description="使用的模型名称，例如 gpt-4o 或 gpt-4o-mini")
    language: str = Field(..., description="生成内容的语言，例如 中文、English")
    content: str = Field(..., description="大纲内容")
    stream: bool = True
class AIWritingRequest(BaseModel):
    model: str = Field(None, description="使用的模型名称，例如 gpt-4o 或 gpt-4o-mini")
    content: str = Field(None, description="需要AI优化的内容")
    command: str = Field(..., description="AI优化指令")
    stream: bool = True

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
                "require": request.content,  # 用 content 字段传递到 prompt
                "language": request.language
            }):
                yield chunk
        except Exception as e:
            import traceback
            print("[AIPPT Outline Streaming Error]", e)
            traceback.print_exc()
            # 返回错误信息到前端
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
async def generate_ppt_content_stream(request: AIWritingRequest):
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

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="127.0.0.1", port=5000, reload=True)