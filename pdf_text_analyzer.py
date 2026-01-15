"""
PDF文本结构分析工具
演示PDF中文本的存储方式和提取方法
"""
import fitz  # PyMuPDF
from PyPDF2 import PdfReader
import json


def analyze_pdf_text_structure_pymupdf(pdf_path):
    """
    使用PyMuPDF分析PDF文本结构
    """
    print("=" * 60)
    print("使用 PyMuPDF (fitz) 分析PDF文本结构")
    print("=" * 60)
    
    doc = fitz.open(pdf_path)
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        print(f"\n--- 第 {page_num + 1} 页 ---")
        
        # 获取文本块（blocks）- 这些是文本的视觉分组
        blocks = page.get_text("dict")["blocks"]
        
        text_blocks = []
        for block_idx, block in enumerate(blocks):
            if "lines" in block:  # 文本块
                block_text = ""
                for line in block["lines"]:
                    for span in line["spans"]:
                        block_text += span["text"]
                
                # 获取块的位置信息（类似文本框）
                bbox = block["bbox"]  # [x0, y0, x1, y1]
                
                text_blocks.append({
                    "block_index": block_idx,
                    "text": block_text.strip(),
                    "bbox": bbox,
                    "font": line["spans"][0]["font"] if line["spans"] else "unknown",
                    "size": line["spans"][0]["size"] if line["spans"] else 0
                })
                
                print(f"\n文本块 {block_idx + 1}:")
                print(f"  位置: ({bbox[0]:.1f}, {bbox[1]:.1f}) - ({bbox[2]:.1f}, {bbox[3]:.1f})")
                print(f"  字体: {line['spans'][0]['font'] if line['spans'] else 'unknown'}")
                print(f"  大小: {line['spans'][0]['size'] if line['spans'] else 0}")
                print(f"  内容: {block_text[:100]}..." if len(block_text) > 100 else f"  内容: {block_text}")
        
        print(f"\n本页共有 {len(text_blocks)} 个文本块")
    
    doc.close()


def analyze_pdf_text_structure_pypdf2(pdf_path):
    """
    使用PyPDF2分析PDF文本结构
    """
    print("\n" + "=" * 60)
    print("使用 PyPDF2 分析PDF文本结构")
    print("=" * 60)
    
    reader = PdfReader(pdf_path)
    
    for page_num, page in enumerate(reader.pages):
        print(f"\n--- 第 {page_num + 1} 页 ---")
        
        # PyPDF2提取文本（简单方式）
        text = page.extract_text()
        
        print(f"提取的文本（前200字符）:")
        print(text[:200] + "..." if len(text) > 200 else text)
        
        # 尝试提取文本对象（如果支持）
        if hasattr(page, 'get_contents'):
            try:
                contents = page.get_contents()
                print(f"\n内容对象数量: {len(contents) if isinstance(contents, list) else 1}")
            except:
                pass


def extract_text_with_positions(pdf_path):
    """
    提取文本及其位置信息（模拟文本框）
    """
    print("\n" + "=" * 60)
    print("提取文本及其位置信息（模拟文本框）")
    print("=" * 60)
    
    doc = fitz.open(pdf_path)
    
    all_text_boxes = []
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        blocks = page.get_text("dict")["blocks"]
        
        for block_idx, block in enumerate(blocks):
            if "lines" in block:
                # 收集块中的所有文本
                full_text = ""
                for line in block["lines"]:
                    for span in line["spans"]:
                        full_text += span["text"]
                
                if full_text.strip():
                    # 获取边界框（类似文本框）
                    bbox = block["bbox"]
                    
                    text_box = {
                        "page": page_num + 1,
                        "box_index": block_idx,
                        "text": full_text.strip(),
                        "x0": bbox[0],
                        "y0": bbox[1],
                        "x1": bbox[2],
                        "y1": bbox[3],
                        "width": bbox[2] - bbox[0],
                        "height": bbox[3] - bbox[1]
                    }
                    
                    all_text_boxes.append(text_box)
                    
                    print(f"\n页面 {page_num + 1}, 文本框 {block_idx + 1}:")
                    print(f"  位置: ({bbox[0]:.1f}, {bbox[1]:.1f})")
                    print(f"  大小: {bbox[2] - bbox[0]:.1f} x {bbox[3] - bbox[1]:.1f}")
                    print(f"  内容: {full_text[:80]}..." if len(full_text) > 80 else f"  内容: {full_text}")
    
    doc.close()
    
    return all_text_boxes


def analyze_paragraph_structure(text_boxes):
    """
    尝试从文本框中识别段落结构
    """
    print("\n" + "=" * 60)
    print("尝试识别段落结构")
    print("=" * 60)
    
    # 简单的段落识别：基于垂直间距和字体大小
    paragraphs = []
    current_paragraph = []
    
    for i, box in enumerate(text_boxes):
        if i == 0:
            current_paragraph.append(box)
        else:
            prev_box = text_boxes[i - 1]
            # 如果垂直间距较大，可能是新段落
            vertical_gap = box["y0"] - prev_box["y1"]
            
            if vertical_gap > 10:  # 阈值，可根据实际情况调整
                if current_paragraph:
                    paragraphs.append(current_paragraph)
                current_paragraph = [box]
            else:
                current_paragraph.append(box)
    
    if current_paragraph:
        paragraphs.append(current_paragraph)
    
    print(f"\n识别出 {len(paragraphs)} 个可能的段落:")
    for para_idx, para in enumerate(paragraphs):
        para_text = " ".join([box["text"] for box in para])
        print(f"\n段落 {para_idx + 1} ({len(para)} 个文本框):")
        print(f"  {para_text[:150]}..." if len(para_text) > 150 else f"  {para_text}")


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 2:
        print("使用方法: python pdf_text_analyzer.py <PDF文件路径>")
        sys.exit(1)
    
    pdf_path = sys.argv[1]
    
    try:
        # 方法1: 使用PyMuPDF分析
        analyze_pdf_text_structure_pymupdf(pdf_path)
        
        # 方法2: 使用PyPDF2分析
        analyze_pdf_text_structure_pypdf2(pdf_path)
        
        # 方法3: 提取文本位置信息（模拟文本框）
        text_boxes = extract_text_with_positions(pdf_path)
        
        # 方法4: 尝试识别段落结构
        if text_boxes:
            analyze_paragraph_structure(text_boxes)
        
    except Exception as e:
        print(f"错误: {e}")
        import traceback
        traceback.print_exc()
