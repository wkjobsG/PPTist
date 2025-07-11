import json
import os
import uuid
import asyncio
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

from pydantic import BaseModel, Field
from dotenv import load_dotenv

from langchain.prompts import PromptTemplate
from langchain_core.output_parsers import StrOutputParser
from langchain_openai import ChatOpenAI

# 加载.env环境变量
load_dotenv()

# 静态文件目录
STATIC_DIR = Path("ppt_files")
STATIC_DIR.mkdir(exist_ok=True)

# Pydantic模型定义
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

class PPTSaveRequest(BaseModel):
    title: str = Field(..., description="PPT标题")
    slides: List[Dict[str, Any]] = Field(..., description="PPT幻灯片数据")
    template: str = Field("template_7", description="PPT模板")

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

{{"type": "content", "data": {{ "title": "接口定义", "items": [ {{ "title": "基本概念", "text": "接口是系统中模块通信的协议" ,"image": "图片url", "table": "表格HTML"}}, {{ "title": "作用", "text": "促进模块解耦，提高系统灵活性" }} ] }}}}

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

# 辅助函数
def save_ppt_json(title: str, slides: List[Dict], template: str = "template_7") -> str:
    """保存PPT JSON并返回访问URL"""
    ppt_id = str(uuid.uuid4())
    filename = f"ppt_{ppt_id}.json"
    filepath = STATIC_DIR / filename
    
    ppt_data = {
        "slides": slides,
        "template": template,
        "title": f"ppt_aippt_{title}",
        "type": "aippt_slides",
        "created_at": datetime.now().isoformat(),
        "id": ppt_id
    }
    
    with open(filepath, 'w', encoding='utf-8') as f:
        json.dump(ppt_data, f, ensure_ascii=False, indent=2)
    
    # 返回JSON文件的HTTP访问地址
    base_url = os.environ.get("SERVER_BASE_URL", "http://127.0.0.1:5000")
    json_url = f"{base_url}/files/{filename}"
    return json_url

def parse_ppt_content_to_slides(content: str) -> List[Dict]:
    """解析PPT内容为slides格式"""
    slides = []
    lines = [line.strip() for line in content.strip().split("\n") if line.strip()]
    
    for line in lines:
        try:
            slide_data = json.loads(line)
            slides.append(slide_data)
        except json.JSONDecodeError:
            # 忽略解析失败的行
            continue
    
    return slides