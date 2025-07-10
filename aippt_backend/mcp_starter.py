import asyncio
from typing import Any, Dict, List
from mcp.server import FastMCP

# 导入PPT模块
from ppt import (
    build_outline_chain,
    build_ppt_content_chain,
    build_ai_writing_chain,
    parse_ppt_content_to_slides,
    save_ppt_json
)

# 创建FastMCP实例
mcp = FastMCP("AI PPT Generator 🚀", port=5001, host="0.0.0.0")

# MCP工具定义
@mcp.tool()
async def generate_ppt_outline(
    content: str,
    language: str = "中文",
    model: str = None
) -> str:
    """
    生成PPT大纲
    
    Args:
        content: PPT主题或要求描述
        language: 生成内容的语言，默认为中文
        model: 使用的AI模型名称，例如gpt-4o
    
    Returns:
        生成的PPT大纲文本
    """
    chain = build_outline_chain(model)
    try:
        result = await chain.ainvoke({
            "require": content,
            "language": language
        })
        return result
    except Exception as e:
        return f"生成大纲时出错: {str(e)}"

@mcp.tool()
async def generate_ppt_content(
    content: str,
    language: str = "中文",
    model: str = None
) -> Dict[str, Any]:
    """
    根据大纲生成完整PPT内容
    
    Args:
        content: PPT大纲内容
        language: 生成内容的语言，默认为中文
        model: 使用的AI模型名称，例如gpt-4o
    
    Returns:
        包含slides数组的字典
    """
    chain = build_ppt_content_chain(model)
    try:
        result = await chain.ainvoke({
            "language": language,
            "content": content
        })
        slides = parse_ppt_content_to_slides(result)
        return {
            "slides": slides,
            "success": True
        }
    except Exception as e:
        return {
            "slides": [],
            "raw_content": "",
            "success": False,
            "error": str(e)
        }

@mcp.tool()
async def optimize_ppt_content(
    content: str,
    command: str,
    model: str = None
) -> str:
    """
    优化PPT内容
    
    Args:
        content: 需要优化的内容
        command: 优化指令，例如"使内容更简洁"
        model: 使用的AI模型名称，例如gpt-4o
    
    Returns:
        优化后的内容
    """
    chain = build_ai_writing_chain(model)
    try:
        result = await chain.ainvoke({
            "command": command,
            "content": content
        })
        return result
    except Exception as e:
        return f"优化内容时出错: {str(e)}"

@mcp.tool()
def save_ppt_to_file(
    title: str,
    slides: List[Dict[str, Any]],
    template: str = "template_7"
) -> Dict[str, Any]:
    """
    保存PPT到文件并返回查看链接
    
    Args:
        title: PPT标题
        slides: PPT幻灯片数据数组
        template: PPT模板名称，默认为template_7
    
    Returns:
        包含保存结果和查看链接的字典
    """
    try:
        json_url = save_ppt_json(title, slides, template)
        view_url = f"http://127.0.0.1:5173/?aippt_config_url={json_url}"
        
        return {
            "success": True,
            "json_url": json_url,
            "view_url": view_url,
            "message": "PPT保存成功"
        }
    except Exception as e:
        return {
            "success": False,
            "error": f"保存PPT失败: {str(e)}"
        }

@mcp.tool()
async def generate_complete_ppt(
    content: str,
    title: str = "AI生成PPT",
    language: str = "中文",
    template: str = "template_7",
    model: str = None
) -> Dict[str, Any]:
    """
    一键生成完整PPT（包含大纲生成、内容生成和保存）
    
    Args:
        content: PPT主题或要求描述
        title: PPT标题
        language: 生成内容的语言，默认为中文
        template: PPT模板名称，默认为template_7
        model: 使用的AI模型名称，例如gpt-4o
    
    Returns:
        包含完整PPT生成结果的字典
    """
    try:
        # 1. 生成大纲
        outline_chain = build_outline_chain(model)
        outline = await outline_chain.ainvoke({
            "require": content,
            "language": language
        })
        
        # 2. 生成PPT内容
        content_chain = build_ppt_content_chain(model)
        ppt_content = await content_chain.ainvoke({
            "language": language,
            "content": f"原始要求: {content}\n\n大纲内容:\n{outline}"
        })
        
        # 3. 解析并保存PPT
        slides = parse_ppt_content_to_slides(ppt_content)
        json_url = save_ppt_json(title, slides, template)
        view_url = f"http://127.0.0.1:5173/?aippt_config_url={json_url}"
        
        return {
            "success": True,
            "outline": outline,
            "slides": slides,
            "json_url": json_url,
            "view_url": view_url,
            "message": "完整PPT生成成功"
        }
    except Exception as e:
        return {
            "success": False,
            "error": f"生成完整PPT时出错: {str(e)}"
        }

# 运行MCP服务器的函数
def run_mcp_server():
    """运行MCP服务器"""
    print("启动 MCP Server...")
    print("MCP Server 地址: http://127.0.0.1:5001")
    mcp.run(transport="sse")