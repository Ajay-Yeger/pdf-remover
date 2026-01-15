import sys
import os
import json
import requests
import base64
from PyQt5.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QLabel,
    QFileDialog,
    QTextEdit,
    QProgressBar,
    QMessageBox,
    QLineEdit,
    QDateEdit,
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QDate
from PyPDF2 import PdfReader, PdfWriter
import fitz  # PyMuPDF
from datetime import datetime
import re
import credit_score_visualizer

def get_resource_path(relative_path):
    """
    获取资源文件的绝对路径
    支持PyInstaller打包后的环境
    """
    try:
        # PyInstaller打包后的临时目录
        base_path = sys._MEIPASS
    except Exception:
        # 开发环境，使用脚本所在目录
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

def get_base_dir():
    """
    获取程序基础目录（用于保存配置文件等）
    打包后返回exe所在目录，开发环境返回脚本所在目录
    """
    if getattr(sys, 'frozen', False):
        # 打包后的exe环境
        return os.path.dirname(sys.executable)
    else:
        # 开发环境
        return os.path.dirname(os.path.abspath(__file__))

# 配置文件路径（保存在程序目录下）
CONFIG_FILE = os.path.join(get_base_dir(), "pdf_processor_config.json")


class PDFProcessorThread(QThread):
    """处理PDF的线程类"""
    progress = pyqtSignal(int)  # 进度信号
    status = pyqtSignal(str)     # 状态信号
    finished = pyqtSignal()     # 完成信号
    
    def __init__(self, pdf_files, output_dir, image_output_dir, update_date=None, employee_id=None, employee_name=None, region_code=None, parent=None):
        super().__init__()
        self.pdf_files = pdf_files
        self.output_dir = output_dir
        self.image_output_dir = image_output_dir
        self.update_date = update_date if update_date else datetime.now()  # 使用传入的日期，默认为当前日期
        self.employee_id = employee_id  # 工号
        self.employee_name = employee_name  # 姓名
        self.region_code = region_code  # 地区编码
        self.parent = parent  # 保存父窗口引用，用于访问huawei_token

        
    def run(self):
        total_files = len(self.pdf_files)
        
        for index, pdf_path in enumerate(self.pdf_files):
            try:
                self.status.emit(f"正在处理: {os.path.basename(pdf_path)}")
                
                # 读取PDF
                reader = PdfReader(pdf_path)
                total_pages = len(reader.pages)
                
                # 检查页数
                if total_pages <= 2:
                    self.status.emit(f"跳过 {os.path.basename(pdf_path)}: 页数不足（只有{total_pages}页）")
                    self.progress.emit(int((index + 1) / total_files * 100))
                    continue
                
                # 创建新的PDF写入器
                writer = PdfWriter()
                
                # 添加除第一页和最后一页外的所有页面
                for page_num in range(1, total_pages - 1):
                    writer.add_page(reader.pages[page_num])
                
                # 生成输出文件名
                base_name = os.path.splitext(os.path.basename(pdf_path))[0]
                # 添加工号-姓名前缀
                if self.employee_id and self.employee_name:
                    prefix = f"{self.employee_id}-{self.employee_name}_"
                else:
                    prefix = ""
                # 使用os.path.join确保路径正确，并规范化路径分隔符
                base_output_path = os.path.normpath(os.path.join(self.output_dir, f"{prefix}{base_name}_processed.pdf"))

                # 如果已存在同名文件，则自动添加索引后缀：_1, _2, ...
                output_path = base_output_path
                if os.path.exists(output_path):
                    idx = 1
                    while True:
                        candidate = os.path.normpath(
                            os.path.join(self.output_dir, f"{prefix}{base_name}_processed_{idx}.pdf")
                        )
                        if not os.path.exists(candidate):
                            output_path = candidate
                            break
                        idx += 1
                
                # 保存处理后的PDF
                with open(output_path, 'wb') as output_file:
                    writer.write(output_file)
                
                self.status.emit(
                    f"✓ PDF处理完成: {os.path.basename(pdf_path)} -> {os.path.basename(output_path)}"
                )

                # 在删除后的第1页的 "1.2 工商信息" 上方插入 "BOSS来单指数评估"
                # try:
                #     add_subtitle_above_text_in_page1(output_path, target_text="1.2 工商信息", subtitle="1.1 BOSS来单指数评估", font_size=12)
                #     self.status.emit("  - 已在第1页的1.2 工商信息上方添加二级标题")
                # except Exception as e:
                #     self.status.emit(f"  - 在第1页添加二级标题时出错: {e}")

                # 替换第1页中以"1.1 企查分"开头的文本块为"1.1 BOSS来单指数评估"
                try:
                    replace_text_starting_with(output_path, target_prefix="1.1 企查分", new_text="BOSS来单指数评估", font_size=12)
                    self.status.emit("  - 已替换第1页的1.1 企查分为1.1 BOSS来单指数评估")
                except Exception as e:
                    self.status.emit(f"  - 替换文本时出错: {e}")

                # 删除（覆盖）以"联系电话"开头的文本块
                try:
                    remove_tel_blocks_from_pdf(output_path, prefix="联系电话")
                    self.status.emit("  - 已移除以联系电话开头的文本块（如存在）")
                except Exception as e:
                    self.status.emit(f"  - 移除联系电话文本块时出错: {e}")

                # 删除包含"企查查"或"企查分"的文本块
                try:
                    remove_keyword_blocks_from_pdf(output_path, keywords=["企查查", "企查分"])
                    self.status.emit("  - 已删除包含企查查或企查分的文本块（如存在）")
                except Exception as e:
                    self.status.emit(f"  - 删除企查查/企查分文本块时出错: {e}")

                # 替换左上角 logo 为 newlogo.png
                try:
                    logo_path = get_resource_path("newlogo.png")
                    if os.path.exists(logo_path):
                        replace_top_left_logo(output_path, logo_path)
                        self.status.emit("  - 已替换左上角 logo 为 newlogo.png（如存在）")
                    else:
                        self.status.emit(f"  - 未找到 newlogo.png，跳过 logo 替换（查找路径: {logo_path}）")
                except Exception as e:
                    self.status.emit(f"  - 替换左上角 logo 时出错: {e}")

                # 在右上角添加 newlogo2.jpeg
                try:
                    top_right_logo_path = get_resource_path("newlogo2.jpeg")
                    if os.path.exists(top_right_logo_path):
                        add_top_right_logo(output_path, top_right_logo_path)
                        self.status.emit("  - 已在右上角添加 newlogo2.jpeg")
                    else:
                        self.status.emit(f"  - 未找到 newlogo2.jpeg，跳过右上角 logo 添加（查找路径: {top_right_logo_path}）")
                except Exception as e:
                    self.status.emit(f"  - 添加右上角 logo 时出错: {e}")

                # 在每页页眉右侧添加文档编码
                try:
                    if self.region_code:
                        add_header_document_code(output_path, self.region_code)
                        self.status.emit(f"  - 已在每页页眉添加文档编码: {self.region_code}-XXXXXX")
                    else:
                        self.status.emit("  - 未设置地区编码，跳过页眉文档编码添加")
                except Exception as e:
                    self.status.emit(f"  - 添加页眉文档编码时出错: {e}")

                # 在 "1 基本信息" 下面添加二级标题
                try:
                    add_subtitle_after_text(output_path, target_text="1 基本信息", subtitle="1.1 BOSS来单指数评估", font_size=12)
                    self.status.emit("  - 已在1 基本信息下方添加二级标题")
                except Exception as e:
                    self.status.emit(f"  - 添加二级标题时出错: {e}")
                
                # 提取图片（即使未选择图片输出目录也执行，用于OCR和替换page2_img2）
                self.extract_images(output_path, base_name)
                
                # 更新进度
                self.progress.emit(int((index + 1) / total_files * 100))
                
            except Exception as e:
                self.status.emit(f"✗ 错误 {os.path.basename(pdf_path)}: {str(e)}")
                import traceback
                print(f"错误详情: {traceback.format_exc()}")
                # 即使出错也要更新进度
                self.progress.emit(int((index + 1) / total_files * 100))
        
        self.finished.emit()
    
    def extract_images(self, pdf_path, base_name):
        """从PDF中提取所有图片（用于OCR和替换page2_img2）
        
        如果未选择图片输出目录（self.image_output_dir 为空），
        则使用程序目录下的临时子目录保存中间图片文件。
        """
        doc = None  # 初始化doc变量，确保在finally中能正确关闭
        try:
            # 确保图片输出目录存在
            if self.image_output_dir:
                pdf_image_dir = os.path.join(self.image_output_dir, base_name)
            else:
                # 未选择图片输出目录时，使用程序基础目录下的 images_temp 目录
                base_dir = get_base_dir()
                pdf_image_dir = os.path.join(base_dir, "images_temp", base_name)
            os.makedirs(pdf_image_dir, exist_ok=True)
            
            # 使用PyMuPDF打开PDF
            doc = fitz.open(pdf_path)
            img_count = 0
            image_info_list = []
            
            # 仅提取 page2_img2（即第2页的第2张图片，1-based）
            target_page_index = 1   # 第2页，0-based 索引为1
            target_img_index = 1    # 第2张图片，enumerate 从0开始

            for page_index in range(len(doc)):
                page = doc[page_index]

                # 只在目标页上查找图片，其它页跳过
                if page_index != target_page_index:
                    continue

                image_list = page.get_images(full=True)
                
                for img_index, img in enumerate(image_list):
                    # 只处理目标图片，其它图片跳过
                    if img_index != target_img_index:
                        continue

                    xref = img[0]
                    base_image = doc.extract_image(xref)
                    image_bytes = base_image["image"]
                    image_ext = base_image["ext"]
                    
                    img_count += 1
                    # 生成图片文件名：PDF名_页码_图片索引.扩展名（保持原命名规则）
                    img_filename = f"{base_name}_page{page_index+1}_img{img_index+1}.{image_ext}"
                    img_path = os.path.normpath(os.path.join(pdf_image_dir, img_filename))
                    
                    # 保存图片
                    with open(img_path, "wb") as f:
                        f.write(image_bytes)
                    
                    # 记录图片信息
                    image_info = {
                        'index': img_count,
                        'pdf_name': base_name,
                        'page': page_index + 1,
                        'img_in_page': img_index + 1,
                        'filename': img_filename,
                        'path': img_path,
                        'ext': image_ext
                    }
                    image_info_list.append(image_info)
                    
                    # 在控制台输出图片下标信息
                    print(f"[图片下标: {img_count}] PDF: {base_name}, 页码: {page_index+1}, "
                          f"页面内图片索引: {img_index+1}, 文件名: {img_filename}, "
                          f"扩展名: {image_ext}")
                    
                    # 调用华为云OCR API识别图片文字
                    if self.parent and hasattr(self.parent, 'huawei_token') and self.parent.huawei_token:
                        self.status.emit("  - 正在调用华为云OCR API识别图片文字...")
                        config = load_config()
                        project_id = config.get("huawei_project_id", "")
                        region = config.get("huawei_project", "cn-north-4")
                        
                        if project_id:
                            ocr_result = call_huawei_ocr_api(
                                image_bytes, 
                                self.parent.huawei_token, 
                                project_id,
                                region
                            )
                            if ocr_result:
                                # 提取识别结果
                                words_block_list = ocr_result.get("result", {}).get("words_block_list", [])
                                if words_block_list:
                                    recognized_text = "\n".join([block.get("words", "") for block in words_block_list])
                                    self.status.emit(f"  - OCR识别成功，识别到 {len(words_block_list)} 个文字块")
                                    print(f"\n=== OCR识别结果 ===")
                                    print(f"识别到的文字块数量: {len(words_block_list)}")
                                    print(f"识别内容:\n{recognized_text}\n")
                                    
                                    # 提取倒数第三个文字块作为信用分
                                    credit_score = None
                                    if len(words_block_list) >= 3:
                                        third_last_block = words_block_list[-3]
                                        third_last_text = third_last_block.get("words", "")
                                        print(f"倒数第三个文字块: {third_last_text}")
                                        
                                        # 尝试从文字中提取数字
                                        numbers = re.findall(r'\d+', third_last_text)
                                        if numbers:
                                            credit_score = int(numbers[0])  # 取第一个数字
                                            print(f"提取的信用分: {credit_score}")
                                            self.status.emit(f"  - 提取到信用分: {credit_score}")
                                            
                                            # 创建信用分可视化图片并替换PDF中的图片
                                            tmp_path = None  # 用于跟踪临时文件，确保异常时清理
                                            try:
                                                self.status.emit(f"  - 正在创建信用分可视化图片（分数: {credit_score}）...")
                                                # 获取图片位置信息（必须在doc关闭前获取）
                                                rects = page.get_image_rects(xref)
                                                if rects:
                                                    original_rect = rects[0]
                                                    
                                                    # 获取页面尺寸
                                                    page_rect = page.rect
                                                    page_width = page_rect.width
                                                    page_height = page_rect.height
                                                    
                                                    # 计算放大后的尺寸（原尺寸的3倍）
                                                    original_width = original_rect.x1 - original_rect.x0
                                                    original_height = original_rect.y1 - original_rect.y0
                                                    scale_factor = 3.3  # 放大3倍
                                                    enlarged_width = original_width * scale_factor
                                                    enlarged_height = original_height * scale_factor
                                                    
                                                    # 计算原图片的中心点（用于保持y坐标）
                                                    center_y = (original_rect.y0 + original_rect.y1) / 2
                                                    
                                                    # 向下移动的距离（可以根据需要调整）
                                                    vertical_offset = 15  # 向下移动30像素
                                                    center_y = center_y + vertical_offset
                                                    
                                                    # 在页面上水平居中：x坐标 = (页面宽度 - 图片宽度) / 2
                                                    center_x = page_width / 2
                                                    
                                                    # 以页面中心为x坐标，原图片中心为y坐标（向下偏移），计算放大后的矩形
                                                    target_img_rect = fitz.Rect(
                                                        center_x - enlarged_width / 2,  # 左上角x = 页面中心x - 宽度/2
                                                        center_y - enlarged_height / 2,  # 左上角y = 原图片中心y（已下移）- 高度/2
                                                        center_x + enlarged_width / 2,   # 右下角x = 页面中心x + 宽度/2
                                                        center_y + enlarged_height / 2    # 右下角y = 原图片中心y（已下移）+ 高度/2
                                                    )
                                                    
                                                    # 打印调试信息
                                                    print(f"页面尺寸: {page_width:.1f} x {page_height:.1f}")
                                                    print(f"原图片位置: ({original_rect.x0:.1f}, {original_rect.y0:.1f}) - ({original_rect.x1:.1f}, {original_rect.y1:.1f})")
                                                    print(f"原图片尺寸: {original_width:.1f} x {original_height:.1f}")
                                                    print(f"放大后尺寸: {enlarged_width:.1f} x {enlarged_height:.1f}")
                                                    print(f"页面中心x: {center_x:.1f}, 原图片中心y: {center_y:.1f}")
                                                    print(f"新图片位置: ({target_img_rect.x0:.1f}, {target_img_rect.y0:.1f}) - ({target_img_rect.x1:.1f}, {target_img_rect.y1:.1f})")
                                                    
                                                    # 创建临时图片文件
                                                    temp_image_path = os.path.normpath(os.path.join(pdf_image_dir, f"{base_name}_credit_score_temp.png"))
                                                    
                                                    # 使用传入的日期（从GUI选择）
                                                    update_date = self.update_date
                                                    
                                                    # 调用创建可视化图片的函数
                                                    credit_score_visualizer.create_credit_score_visualization(
                                                        score=credit_score,
                                                        update_date=update_date,
                                                        output_path=temp_image_path
                                                    )
                                                    
                                                    self.status.emit("  - 信用分可视化图片创建成功，正在替换PDF中的图片...")
                                                    
                                                    # 替换PDF中的图片
                                                    # 先删除原图片
                                                    try:
                                                        page.delete_image(xref)
                                                    except:
                                                        pass
                                                    
                                                    # 在原位置插入新图片（放大一倍）
                                                    page.insert_image(target_img_rect, filename=temp_image_path, keep_proportion=True)
                                                    
                                                    # 保存PDF（使用临时文件方式）
                                                    # 注意：必须关闭文档后才能用os.replace()替换文件，否则可能因文件锁定而失败
                                                    tmp_path = pdf_path + ".credit_score.tmp"
                                                    doc.save(tmp_path, deflate=True)
                                                    doc.close()  # 关闭文档，确保文件可以被os.replace()替换
                                                    doc = None  # 标记doc已关闭
                                                    
                                                    try:
                                                        os.replace(tmp_path, pdf_path)
                                                        tmp_path = None  # 替换成功，清除临时文件标记
                                                    except Exception as replace_error:
                                                        # 如果替换失败，尝试清理临时文件
                                                        if tmp_path and os.path.exists(tmp_path):
                                                            try:
                                                                os.remove(tmp_path)
                                                            except:
                                                                pass
                                                        raise replace_error
                                                    
                                                    # 重新打开PDF，因为函数末尾需要统一关闭
                                                    doc = fitz.open(pdf_path)
                                                    page = doc[target_page_index]
                                                    
                                                    self.status.emit(f"  - ✓ 已成功替换PDF中的page2_img2为信用分可视化图片")
                                                    print(f"✓ 已成功替换PDF中的page2_img2为信用分可视化图片（分数: {credit_score}）")
                                                else:
                                                    self.status.emit("  - 警告: 无法获取图片位置，跳过替换")
                                            except Exception as e:
                                                # 如果doc已关闭，需要重新打开
                                                if doc is None:
                                                    try:
                                                        doc = fitz.open(pdf_path)
                                                    except:
                                                        pass
                                                else:
                                                    # 尝试访问文档属性来检查是否已关闭，如果已关闭则重新打开
                                                    try:
                                                        _ = len(doc)  # 尝试访问文档属性
                                                    except:
                                                        # 文档已关闭，重新打开
                                                        try:
                                                            doc = fitz.open(pdf_path)
                                                        except:
                                                            pass
                                                
                                                # 清理临时文件
                                                if tmp_path and os.path.exists(tmp_path):
                                                    try:
                                                        os.remove(tmp_path)
                                                    except:
                                                        pass
                                                
                                                self.status.emit(f"  - ✗ 替换图片时出错: {str(e)}")
                                                print(f"✗ 替换图片时出错: {str(e)}")
                                                import traceback
                                                print(traceback.format_exc())
                                        else:
                                            print(f"警告: 倒数第三个文字块中未找到数字: {third_last_text}")
                                            self.status.emit(f"  - 警告: 未能在倒数第三个文字块中找到数字")
                                    else:
                                        print(f"警告: 文字块数量不足3个，无法提取倒数第三个")
                                        self.status.emit("  - 警告: 文字块数量不足，无法提取信用分")
                                else:
                                    self.status.emit("  - OCR识别成功，但未识别到文字")
                                    print("OCR识别成功，但未识别到文字")
                            else:
                                self.status.emit("  - OCR识别失败")
                        else:
                            self.status.emit("  - 未配置项目ID，跳过OCR识别")
                    else:
                        self.status.emit("  - Token未获取，跳过OCR识别")
            
            # 更新状态
            self.status.emit(f"✓ 提取图片完成: {base_name} -> 共 {img_count} 张图片")
            print(f"\n=== {base_name} 图片提取完成 ===")
            print(f"共提取 {img_count} 张图片")
            print(f"保存目录: {pdf_image_dir}\n")
            
        except Exception as e:
            error_msg = f"✗ 提取图片错误 {base_name}: {str(e)}"
            self.status.emit(error_msg)
            print(error_msg)
            import traceback
            print(traceback.format_exc())
        finally:
            # 确保doc总是被关闭
            if doc is not None:
                try:
                    doc.close()
                except:
                    pass


def remove_tel_blocks_from_pdf(pdf_path: str, prefix: str = "联系电话"):
    """
    从 PDF 中删除（通过白色覆盖）以指定前缀开头的文本块。
    使用 PyMuPDF 的文本块信息，找到以 prefix 开头的块并画白色矩形覆盖。
    为避免“save to original must be incremental”错误，采用保存到临时文件再替换原文件的方式。
    """
    if not os.path.exists(pdf_path):
        print(f"错误: PDF文件不存在: {pdf_path}")
        return

    doc = fitz.open(pdf_path)
    changed = False

    for page_index, page in enumerate(doc):
        # get_text("blocks") 返回的每个元素通常为:
        # (x0, y0, x1, y1, text, block_no, ...)，其中 text 为该块的全部文本
        blocks = page.get_text("blocks")
        for b in blocks:
            if len(b) < 5:
                continue
            x0, y0, x1, y1, text = b[0], b[1], b[2], b[3], b[4]
            if isinstance(text, str) and text.strip().startswith(prefix):
                rect = fitz.Rect(x0, y0, x1, y1)
                # 在该区域画白色填充矩形，覆盖原有文本
                page.draw_rect(rect, color=(1, 1, 1), fill=(1, 1, 1), width=0)
                changed = True
                print(f"覆盖页面 {page_index + 1} 中文本块: '{text.strip()[:50]}'...")

    if changed:
        # 不能直接覆盖保存到原文件（PyMuPDF 要求 incremental 模式），
        # 这里采用保存到临时文件再替换原文件的方式。
        tmp_path = pdf_path + ".tmp"
        doc.save(tmp_path, deflate=True)
        doc.close()

        os.replace(tmp_path, pdf_path)
        print(f"已从PDF中删除以 '{prefix}' 开头的文本块: {pdf_path}")
    else:
        doc.close()
        print(f"未在PDF中找到以 '{prefix}' 开头的文本块: {pdf_path}")


def remove_keyword_blocks_from_pdf(pdf_path: str, keywords: list):
    """
    从 PDF 中删除（通过白色覆盖）包含指定关键词的文本块。
    使用 PyMuPDF 的文本块信息，找到包含关键词的块并画白色矩形覆盖。
    为避免"save to original must be incremental"错误，采用保存到临时文件再替换原文件的方式。
    
    参数:
        pdf_path: PDF文件路径
        keywords: 关键词列表，文本块中包含任一关键词即会被删除
    """
    if not os.path.exists(pdf_path):
        print(f"错误: PDF文件不存在: {pdf_path}")
        return

    doc = fitz.open(pdf_path)
    changed = False
    removed_count = 0

    for page_index, page in enumerate(doc):
        # get_text("blocks") 返回的每个元素通常为:
        # (x0, y0, x1, y1, text, block_no, ...)，其中 text 为该块的全部文本
        blocks = page.get_text("blocks")
        for b in blocks:
            if len(b) < 5:
                continue
            x0, y0, x1, y1, text = b[0], b[1], b[2], b[3], b[4]
            if isinstance(text, str):
                # 第1页（索引为0）不覆盖包含"企查分"的文本块
                if page_index == 0 and "企查分" in text:
                    print(f"跳过页面 {page_index + 1} 中包含 '企查分' 的文本块（第1页不覆盖）: '{text.strip()[:50]}'...")
                    continue  # 跳过这个文本块，继续处理下一个文本块
                
                # 检查文本中是否包含任一关键词
                for keyword in keywords:
                    if keyword in text:
                        rect = fitz.Rect(x0, y0, x1, y1)
                        # 在该区域画白色填充矩形，覆盖原有文本
                        page.draw_rect(rect, color=(1, 1, 1), fill=(1, 1, 1), width=0)
                        changed = True
                        removed_count += 1
                        print(f"覆盖页面 {page_index + 1} 中包含 '{keyword}' 的文本块: '{text.strip()[:50]}'...")
                        break  # 找到一个关键词就覆盖，避免重复处理

    if changed:
        # 不能直接覆盖保存到原文件（PyMuPDF 要求 incremental 模式），
        # 这里采用保存到临时文件再替换原文件的方式。
        tmp_path = pdf_path + ".keyword.tmp"
        doc.save(tmp_path, deflate=True)
        doc.close()

        os.replace(tmp_path, pdf_path)
        print(f"已从PDF中删除包含关键词 {keywords} 的文本块，共 {removed_count} 个: {pdf_path}")
    else:
        doc.close()
        print(f"未在PDF中找到包含关键词 {keywords} 的文本块: {pdf_path}")


def add_subtitle_after_text(pdf_path: str, target_text: str, subtitle: str, font_size: float = 12, spacing: float = 5):
    """
    在指定文本块下方添加二级标题
    
    参数:
        pdf_path: PDF文件路径
        target_text: 目标文本（要查找的文本块）
        subtitle: 要添加的二级标题文本
        font_size: 字体大小，默认12
        spacing: 与目标文本的间距，默认5
    """
    if not os.path.exists(pdf_path):
        print(f"错误: PDF文件不存在: {pdf_path}")
        return

    doc = fitz.open(pdf_path)
    changed = False
    
    # 尝试加载字体文件（优先从项目目录读取 HYQiHeiClassic-70S.ttf）
    font_name = None
    try:
        # 字体文件路径（优先项目目录中的 HYQiHeiClassic-70S.ttf）
        font_path = get_resource_path("HYQiHeiClassic-70S.ttf")
        
        if os.path.exists(font_path):
            # 将字体插入到第一页（只需要插入一次）
            if len(doc) > 0:
                font_xref = doc[0].insert_font(fontname="HYQiHeiClassic", fontfile=font_path)
                font_name = "HYQiHeiClassic"  # 使用字体名称字符串，与insert_font中的名称保持一致
                print(f"成功加载字体: {font_path}, 字体名称: {font_name}, xref: {font_xref}")
        else:
            print(f"警告: 未找到字体文件 HYQiHeiClassic-70S.ttf，将使用默认字体")
            print(f"提示: 请将字体文件 HYQiHeiClassic-70S.ttf 放在项目目录: {project_dir}")
            font_name = "china-s"  # 回退到内置中文字体
    except Exception as e:
        print(f"加载字体时出错: {e}，将使用默认字体")
        font_name = "china-s"  # 回退到内置中文字体

    for page_index, page in enumerate(doc):
        # 只处理第2页（索引为1），跳过其他页面
        if page_index != 1:
            continue
        
        # 如果使用了自定义字体，需要在该页面上插入字体
        if font_name and font_name != "china-s":
            try:
                font_path = get_resource_path("HYQiHeiClassic-70S.ttf")
                if os.path.exists(font_path):
                    page.insert_font(fontname="HYQiHeiClassic", fontfile=font_path)
            except Exception as e:
                print(f"在页面 {page_index + 1} 插入字体时出错: {e}")
        
        # 获取文本块
        blocks = page.get_text("blocks")
        for b in blocks:
            if len(b) < 5:
                continue
            x0, y0, x1, y1, text = b[0], b[1], b[2], b[3], b[4]
            
            # 查找包含目标文本的块
            if isinstance(text, str) and target_text in text:
                # 计算插入位置：在目标文本块下方
                insert_x = x0  # 保持左对齐
                insert_y = y1 + spacing  # 在文本块下方，加上间距
                
                # 插入文本
                try:
                    # 使用 insert_textbox 方法插入文本（支持中文和长文本）
                    # 计算文本框的宽度（使用页面宽度或原文本块的宽度）
                    page_rect = page.rect
                    textbox_width = min(page_rect.width - insert_x - 10, (x1 - x0) * 2)  # 至少留10像素边距
                    
                    # 创建文本框矩形
                    textbox_rect = fitz.Rect(insert_x, insert_y, insert_x + textbox_width, insert_y + font_size * 2)
                    
                    # 使用 insert_textbox 插入文本（自动换行，支持中文）
                    # 参数：矩形区域, 文本内容, fontsize, fontname, color, align
                    rc = page.insert_textbox(
                        textbox_rect,
                        subtitle,
                        fontsize=font_size,
                        fontname=font_name if font_name else "china-s",  # 使用加载的字体名称
                        color=(0, 0, 0),  # 黑色
                        align=0  # 左对齐
                    )
                    
                    if rc >= 0:  # 成功插入
                        changed = True
                        print(f"页面 {page_index + 1} 在 '{target_text}' 下方添加了 '{subtitle}'，位置: ({insert_x:.1f}, {insert_y:.1f})")
                    else:
                        print(f"警告: 文本可能超出文本框范围，返回码: {rc}")
                        # 如果失败，尝试使用更大的文本框
                        textbox_rect = fitz.Rect(insert_x, insert_y, page_rect.width - 10, insert_y + font_size * 3)
                        rc = page.insert_textbox(
                            textbox_rect,
                            subtitle,
                            fontsize=font_size,
                            fontname=font_name if font_name else "china-s",
                            color=(0, 0, 0),
                            align=0
                        )
                        if rc >= 0:
                            changed = True
                            print(f"页面 {page_index + 1} 在 '{target_text}' 下方添加了 '{subtitle}'（使用扩展文本框）")
                    
                    break  # 只处理第一个匹配的文本块
                except Exception as e:
                    print(f"插入文本时出错: {e}")
                    import traceback
                    print(traceback.format_exc())

    if changed:
        # 保存PDF（使用临时文件方式）
        tmp_path = pdf_path + ".subtitle.tmp"
        doc.save(tmp_path, deflate=True)
        doc.close()
        os.replace(tmp_path, pdf_path)
        print(f"已添加二级标题并保存到原文件: {pdf_path}")
    else:
        doc.close()
        print(f"未在PDF中找到文本 '{target_text}'，未添加二级标题: {pdf_path}")


def add_subtitle_above_text_in_page1(pdf_path: str, target_text: str, subtitle: str, font_size: float = 12, spacing: float = 5):
    """
    在删除页面后的第1页（索引0）指定文本块上方添加二级标题，使用 HYQiHeiClassic-55S.ttf 字体
    
    参数:
        pdf_path: PDF文件路径
        target_text: 目标文本（要查找的文本块）
        subtitle: 要添加的二级标题文本
        font_size: 字体大小，默认12
        spacing: 与目标文本的间距，默认5
    """
    if not os.path.exists(pdf_path):
        print(f"错误: PDF文件不存在: {pdf_path}")
        return

    doc = fitz.open(pdf_path)
    changed = False
    
    # 尝试加载字体文件（优先从项目目录读取 HYQiHeiClassic-55S.ttf）
    font_name = None
    try:
        # 字体文件路径（优先项目目录中的 HYQiHeiClassic-55S.ttf）
        font_path = get_resource_path("HYQiHeiClassic-55S.ttf")
        
        if os.path.exists(font_path):
            # 将字体插入到第一页（只需要插入一次）
            if len(doc) > 0:
                font_xref = doc[0].insert_font(fontname="HYQiHeiClassic55S", fontfile=font_path)
                font_name = "HYQiHeiClassic55S"  # 使用字体名称字符串，与insert_font中的名称保持一致
                print(f"成功加载字体: {font_path}, 字体名称: {font_name}, xref: {font_xref}")
        else:
            print(f"警告: 未找到字体文件 HYQiHeiClassic-55S.ttf，将使用默认字体")
            print(f"提示: 请将字体文件 HYQiHeiClassic-55S.ttf 放在项目目录: {project_dir}")
            font_name = "china-s"  # 回退到内置中文字体
    except Exception as e:
        print(f"加载字体时出错: {e}，将使用默认字体")
        font_name = "china-s"  # 回退到内置中文字体

    # 只处理第1页（索引为0），这是删除第一页和最后一页后的第1页
    if len(doc) > 0:
        page = doc[0]
        page_index = 0

        # 如果使用了自定义字体，需要在该页面上插入字体
        if font_name and font_name != "china-s":
            try:
                font_path = get_resource_path("HYQiHeiClassic-55S.ttf")
                if os.path.exists(font_path):
                    page.insert_font(fontname="HYQiHeiClassic55S", fontfile=font_path)
            except Exception as e:
                print(f"在页面 {page_index + 1} 插入字体时出错: {e}")
        
        # 获取文本块
        blocks = page.get_text("blocks")
        found_target = False
        for b in blocks:
            if len(b) < 5:
                continue
            x0, y0, x1, y1, text = b[0], b[1], b[2], b[3], b[4]

            # 查找包含目标文本的块
            if isinstance(text, str) and target_text in text:
                found_target = True
                print(f"[调试] 找到目标文本块: '{text.strip()}'")
                # 计算插入位置：在目标文本块上方
                insert_x = x0  # 保持左对齐
                page_rect = page.rect
                
                # 计算文本框高度（估算，留一些余量）
                estimated_height = font_size * 1.5
                
                # 在目标文本块上方插入，y坐标 = 目标文本块的顶部 - 间距 - 文本框高度
                insert_y = y0 - spacing - estimated_height
                
                # 确保不会超出页面顶部边界
                if insert_y < 0:
                    insert_y = max(0, y0 - spacing)
                    estimated_height = y0 - spacing - insert_y
                
                # 插入文本
                try:
                    # 使用 insert_textbox 方法插入文本（支持中文和长文本）
                    # 计算文本框的宽度（使用页面宽度或原文本块的宽度）
                    textbox_width = min(page_rect.width - insert_x - 10, (x1 - x0) * 2)  # 至少留10像素边距
                    
                    # 创建文本框矩形
                    textbox_rect = fitz.Rect(insert_x, insert_y, insert_x + textbox_width, insert_y + estimated_height)
                    
                    # 使用 insert_textbox 插入文本（自动换行，支持中文）
                    # 参数：矩形区域, 文本内容, fontsize, fontname, color, align
                    rc = page.insert_textbox(
                        textbox_rect,
                        subtitle,
                        fontsize=font_size,
                        fontname=font_name if font_name else "china-s",  # 使用加载的字体名称
                        color=(0, 0, 0),  # 黑色
                        align=0  # 左对齐
                    )
                    
                    if rc >= 0:  # 成功插入
                        changed = True
                        print(f"第1页（页面 {page_index + 1}）在 '{target_text}' 上方添加了 '{subtitle}'，位置: ({insert_x:.1f}, {insert_y:.1f})")
                    else:
                        print(f"警告: 文本可能超出文本框范围，返回码: {rc}")
                        # 如果失败，尝试使用更大的文本框
                        textbox_rect = fitz.Rect(insert_x, insert_y, page_rect.width - 10, insert_y + font_size * 3)
                        rc = page.insert_textbox(
                            textbox_rect,
                            subtitle,
                            fontsize=font_size,
                            fontname=font_name if font_name else "china-s",
                            color=(0, 0, 0),
                            align=0
                        )
                        if rc >= 0:
                            changed = True
                            print(f"第1页（页面 {page_index + 1}）在 '{target_text}' 上方添加了 '{subtitle}'（使用扩展文本框）")
                    
                    break  # 只处理第一个匹配的文本块
                except Exception as e:
                    print(f"插入文本时出错: {e}")
                    import traceback
                    print(traceback.format_exc())
        
        if not found_target:
            print(f"[调试] 警告: 在第1页未找到包含 '{target_text}' 的文本块")

    if changed:
        # 保存PDF（使用临时文件方式）
        tmp_path = pdf_path + ".page1_subtitle.tmp"
        doc.save(tmp_path, deflate=True)
        doc.close()
        os.replace(tmp_path, pdf_path)
        print(f"已在第1页添加二级标题并保存到原文件: {pdf_path}")
    else:
        doc.close()
        print(f"未在第1页找到文本 '{target_text}'，未添加二级标题: {pdf_path}")


def replace_text_starting_with(pdf_path: str, target_prefix: str, new_text: str, font_size: float = 12):
    """
    在删除页面后的第1页（索引0）替换以指定前缀开头的文本块，使用 HYQiHeiClassic-55S.ttf 字体
    
    参数:
        pdf_path: PDF文件路径
        target_prefix: 目标文本前缀（要查找的文本块以此开头）
        new_text: 要替换的新文本
        font_size: 字体大小，默认12
    """
    if not os.path.exists(pdf_path):
        print(f"错误: PDF文件不存在: {pdf_path}")
        return

    doc = fitz.open(pdf_path)
    changed = False
    
    # 尝试加载字体文件（优先从项目目录读取 HYQiHeiClassic-55S.ttf）
    font_name = None
    try:
        # 字体文件路径（优先项目目录中的 HYQiHeiClassic-55S.ttf）
        font_path = get_resource_path("HYQiHeiClassic-55S.ttf")
        
        if os.path.exists(font_path):
            # 将字体插入到第一页（只需要插入一次）
            if len(doc) > 0:
                font_xref = doc[0].insert_font(fontname="HYQiHeiClassic55S", fontfile=font_path)
                font_name = "HYQiHeiClassic55S"  # 使用字体名称字符串，与insert_font中的名称保持一致
                print(f"成功加载字体: {font_path}, 字体名称: {font_name}, xref: {font_xref}")
        else:
            print(f"警告: 未找到字体文件 HYQiHeiClassic-55S.ttf，将使用默认字体")
            print(f"提示: 请将字体文件 HYQiHeiClassic-55S.ttf 放在项目目录: {project_dir}")
            font_name = "china-s"  # 回退到内置中文字体
    except Exception as e:
        print(f"加载字体时出错: {e}，将使用默认字体")
        font_name = "china-s"  # 回退到内置中文字体

    # 只处理第1页（索引为0），这是删除第一页和最后一页后的第1页
    if len(doc) > 0:
        page = doc[0]
        page_index = 0
        
        # 如果使用了自定义字体，需要在该页面上插入字体
        if font_name and font_name != "china-s":
            try:
                font_path = get_resource_path("HYQiHeiClassic-55S.ttf")
                if os.path.exists(font_path):
                    page.insert_font(fontname="HYQiHeiClassic55S", fontfile=font_path)
            except Exception as e:
                print(f"在页面 {page_index + 1} 插入字体时出错: {e}")
        
        # 获取文本块
        blocks = page.get_text("blocks")
        found_target = False
        
        for b in blocks:
            if len(b) < 5:
                continue
            x0, y0, x1, y1, text = b[0], b[1], b[2], b[3], b[4]
            x0 = x0+21
            x1 = x1-400
            y0 = y0+2
            y1 = y1+2

            # 查找以目标前缀开头的文本块
            if isinstance(text, str) and text.strip().startswith(target_prefix):
                found_target = True

                
                # 计算原文本块的位置和尺寸
                rect = fitz.Rect(x0, y0, x1, y1)
                
                # 步骤1: 用白色矩形覆盖原文本块
                expanded_rect = fitz.Rect(
                    max(0, x0 - 2), 
                    max(0, y0 - 2), 
                    x1 + 2, 
                    y1 + 2
                )
                page.draw_rect(expanded_rect, color=(1, 1, 1), fill=(1, 1, 1), width=0)

                
                # 步骤2: 在原位置插入新文本
                try:
                    # 计算文本插入位置（使用原文本块的基线位置）
                    # insert_text 的 y 坐标是基线位置，需要加上字体大小
                    insert_point = (x0, y0 + font_size)
                    
                    # 使用 insert_text 插入新文本
                    page.insert_text(
                        point=insert_point,
                        text=new_text,
                        fontsize=font_size,
                        fontname=font_name if font_name else "china-s",
                        color=(0, 0, 0)  # 黑色
                    )
                    
                    changed = True
                    print(f"第1页（页面 {page_index + 1}）已替换文本 '{text.strip()[:30]}...' 为 '{new_text}'，位置: ({x0:.1f}, {y0:.1f})")
                    break  # 只处理第一个匹配的文本块
                except Exception as e:
                    print(f"插入文本时出错: {e}")
                    import traceback
                    print(traceback.format_exc())
        
        if not found_target:
            print(f"[调试] 警告: 在第1页未找到以 '{target_prefix}' 开头的文本块")

    if changed:
        # 保存PDF（使用临时文件方式）
        tmp_path = pdf_path + ".replace_text.tmp"
        doc.save(tmp_path, deflate=True)
        doc.close()
        os.replace(tmp_path, pdf_path)
        print(f"已替换文本并保存到原文件: {pdf_path}")
    else:
        doc.close()
        print(f"未在第1页找到以 '{target_prefix}' 开头的文本块，未替换文本: {pdf_path}")


def replace_top_left_logo(pdf_path: str, logo_path: str, max_x: float = 100, max_y: float = 100):
    """
    将每页左上角区域内的图片替换为指定的 logo 图片（newlogo.png）。
    逻辑：
      1. 查找每页中所有图片的绘制位置（Rect）
      2. 选出左上角区域内的 Rect（x0 < max_x 且 y0 < max_y）
      3. 先用白色矩形覆盖原logo区域（删除原logo）
      4. 然后在同一区域插入新的 logo 图片
    为避免 “save to original must be incremental” 错误，同样采用临时文件再替换的方式。
    """
    if not os.path.exists(pdf_path):
        print(f"错误: PDF文件不存在: {pdf_path}")
        return
    if not os.path.exists(logo_path):
        print(f"错误: logo 文件不存在: {logo_path}")
        return

    doc = fitz.open(pdf_path)
    changed = False

    for page_index, page in enumerate(doc):
        imgs = page.get_images(full=True)
        if not imgs:
            continue

        for img in imgs:
            xref = img[0]
            rects = page.get_image_rects(xref)
            for rect in rects:
                # 左上角区域判定
                if rect.x0 < max_x and rect.y0 < max_y:
                    # 步骤1: 先尝试删除图片对象（如果可能）
                    try:
                        page.delete_image(xref)
                    except:
                        pass  # 如果删除失败，继续用覆盖方式
                    
                    # 步骤2: 用白色矩形完全覆盖原logo区域（确保删除原logo）
                    # 稍微扩大覆盖范围，确保完全覆盖
                    expanded_rect = fitz.Rect(
                        max(0, rect.x0 - 2), 
                        max(0, rect.y0 - 2), 
                        rect.x1 + 2, 
                        rect.y1 + 2
                    )
                    page.draw_rect(expanded_rect, color=(1, 1, 1), fill=(1, 1, 1), width=0)
                    
                    # 步骤3: 计算放大10%后的区域（保持左上角位置不变）
                    original_width = rect.x1 - rect.x0
                    original_height = rect.y1 - rect.y0
                    enlarged_width = original_width * 1.2  # 放大20%
                    enlarged_height = original_height * 1.2  # 放大20%
                    
                    # 创建放大后的矩形（左上角位置不变，右下角扩展）
                    enlarged_rect = fitz.Rect(
                        rect.x0,
                        rect.y0,
                        rect.x0 + enlarged_width,
                        rect.y0 + enlarged_height
                    )
                    
                    # 在放大后的区域插入新的 logo 图片
                    page.insert_image(enlarged_rect, filename=logo_path, keep_proportion=True)
                    changed = True
                    print(
                        f"页面 {page_index + 1} 左上角 logo 已删除并替换为 {os.path.basename(logo_path)}（放大10%），"
                        f"原区域: ({rect.x0:.1f}, {rect.y0:.1f}) - ({rect.x1:.1f}, {rect.y1:.1f}), "
                        f"新区域: ({enlarged_rect.x0:.1f}, {enlarged_rect.y0:.1f}) - ({enlarged_rect.x1:.1f}, {enlarged_rect.y1:.1f})"
                    )
                    # 一个 rect 只需要替换一次
                    break

    if changed:
        tmp_path = pdf_path + ".logo.tmp"
        doc.save(tmp_path, deflate=True)
        doc.close()
        os.replace(tmp_path, pdf_path)
        print(f"已完成左上角 logo 替换并保存到原文件: {pdf_path}")
    else:
        doc.close()
        print(f"未在 PDF 中检测到左上角图片（x < {max_x}, y < {max_y}），未进行 logo 替换: {pdf_path}")


def load_config():
    """加载配置文件"""
    default_config = {
        "output_dir": "",
        "image_output_dir": "",
        "huawei_username": "tckeke123cck",
        "huawei_domain": "hid_46npa4c6rmavfjz",
        "huawei_password": "cc961121",
        "huawei_project": "cn-southwest-2",
        "huawei_project_id":"17bd5a8f587e44718ce0a9981d1893ff",
        "employee_id": "",
        "employee_name": "",
        "region_code": ""
    }
    
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                # 确保所有必需的字段都存在，如果不存在则使用默认值
                for key, default_value in default_config.items():
                    if key not in config:
                        config[key] = default_value
                # 如果配置文件缺少字段，自动保存更新后的配置
                if any(key not in config for key in default_config.keys()):
                    try:
                        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                            json.dump(config, f, ensure_ascii=False, indent=2)
                    except:
                        pass  # 如果保存失败，继续使用内存中的配置
                return config
        except Exception as e:
            print(f"加载配置文件失败: {e}")
            return default_config
    else:
        # 如果配置文件不存在，创建默认配置文件
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(default_config, f, ensure_ascii=False, indent=2)
        except:
            pass
        return default_config


def save_config(output_dir="", image_output_dir="", huawei_username="", huawei_domain="", huawei_password="", huawei_project=""):
    """保存配置文件"""
    # 如果只传了output_dir和image_output_dir，尝试保留现有的华为云配置
    if huawei_username == "" and huawei_domain == "" and huawei_password == "":
        existing_config = load_config()
        huawei_username = existing_config.get("huawei_username", "")
        huawei_domain = existing_config.get("huawei_domain", "")
        huawei_password = existing_config.get("huawei_password", "")
        huawei_project = existing_config.get("huawei_project", "cn-north-4")
    
    config = {
        "output_dir": output_dir,
        "image_output_dir": image_output_dir,
        "huawei_username": huawei_username,
        "huawei_domain": huawei_domain,
        "huawei_password": huawei_password,
        "huawei_project": huawei_project
    }
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"保存配置文件失败: {e}")


def save_login_info(employee_id, employee_name, region_code):
    """保存登录信息到配置文件"""
    config = load_config()
    config["employee_id"] = employee_id
    config["employee_name"] = employee_name
    config["region_code"] = region_code
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"保存登录信息失败: {e}")


def clear_login_info():
    """清除登录信息"""
    config = load_config()
    config["employee_id"] = ""
    config["employee_name"] = ""
    config["region_code"] = ""
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"清除登录信息失败: {e}")


def get_huawei_token(username, domain, password, project_name="cn-north-4"):
    """
    获取华为云Token
    
    参数:
        username: IAM用户名
        domain: 账号名
        password: 密码
        project_name: 项目名称，默认为cn-north-4
    
    返回:
        token: 如果成功返回token字符串，失败返回None
    """
    print(username)
    print(domain)
    print(password)
    print(project_name)
    url = f"https://iam.{project_name}.myhuaweicloud.com/v3/auth/tokens"
    
    payload = json.dumps({
        "auth": {
            "identity": {
                "methods": ["password"],
                "password": {
                    "user": {
                        "name": username,
                        "password": password,
                        "domain": {
                            "name": domain
                        }
                    }
                }
            },
            "scope": {
                "project": {
                    "name": project_name
                }
            }
        }
    })
    
    headers = {
        'Content-Type': 'application/json'
    }
    
    try:
        response = requests.post(url, headers=headers, data=payload, timeout=10)
        if response.status_code == 201:
            token = response.headers.get("X-Subject-Token")
            if token:
                print(f"✓ 成功获取华为云Token（有效期24小时）")
                return token
            else:
                print(f"✗ 获取Token失败: 响应头中未找到X-Subject-Token")
                return None
        else:
            print(f"✗ 获取Token失败: HTTP {response.status_code}, {response.text}")
            return None
    except Exception as e:
        print(f"✗ 获取Token时发生错误: {e}")
        return None


def call_huawei_ocr_api(image_bytes, token, project_id, region="cn-north-4"):
    """
    调用华为云OCR API进行文字识别
    
    参数:
        image_bytes: 图片的字节数据
        token: 华为云Token
        project_id: 项目ID
        region: 区域名称，默认为cn-north-4
    
    返回:
        result: 如果成功返回识别结果字典，失败返回None
    """
    # 将图片转换为base64编码
    image_base64 = base64.b64encode(image_bytes).decode('utf-8')
    
    # 构建请求URL
    endpoint = f"ocr.{region}.myhuaweicloud.com"
    url = f"https://{endpoint}/v2/{project_id}/ocr/general-text"
    
    # 构建请求体
    payload = json.dumps({
        "image": image_base64,
        "quick_mode": False,
        "detect_direction": False
    })
    
    # 构建请求头
    headers = {
        'X-Auth-Token': token,
        'Content-Type': 'application/json'
    }
    
    try:
        response = requests.post(url, headers=headers, data=payload, timeout=30)
        if response.status_code == 200:
            result = response.json()
            print(f"✓ OCR识别成功")
            return result
        else:
            print(f"✗ OCR识别失败: HTTP {response.status_code}, {response.text}")
            return None
    except Exception as e:
        print(f"✗ OCR识别时发生错误: {e}")
        import traceback
        print(traceback.format_exc())
        return None


def add_top_right_logo(pdf_path: str, logo_path: str, margin_x: float = 10, margin_y: float = 0, logo_width: float = 80, logo_height: float = 80):
    """
    在每页右上角添加指定的 logo 图片（newlogo2.jpeg）。
    逻辑：
      1. 获取每页的页面尺寸
      2. 计算右上角位置（页面宽度 - margin_x - logo_width, margin_y）
      3. 在该位置插入新的 logo 图片
    为避免 "save to original must be incremental" 错误，采用临时文件再替换的方式。
    
    参数:
        pdf_path: PDF文件路径
        logo_path: logo图片路径
        margin_x: 距离右边缘的边距（默认20）
        margin_y: 距离上边缘的边距（默认20）
        logo_width: logo宽度（默认80）
        logo_height: logo高度（默认80）
    """
    if not os.path.exists(pdf_path):
        print(f"错误: PDF文件不存在: {pdf_path}")
        return
    if not os.path.exists(logo_path):
        print(f"错误: logo 文件不存在: {logo_path}")
        return

    doc = fitz.open(pdf_path)
    changed = False

    for page_index, page in enumerate(doc):
        # 获取页面尺寸
        page_rect = page.rect
        page_width = page_rect.width
        page_height = page_rect.height
        
        # 计算右上角位置（logo缩小10%）
        scaled_logo_width = logo_width * 0.8  # 缩小10%
        scaled_logo_height = logo_height * 0.8  # 缩小10%
        
        # x0: 页面宽度 - 右边距 - 缩小后的logo宽度
        # y0: 上边距
        # x1: 页面宽度 - 右边距
        # y1: 上边距 + 缩小后的logo高度
        x0 = page_width - margin_x - scaled_logo_width
        y0 = margin_y
        x1 = page_width - margin_x
        y1 = margin_y + scaled_logo_height
        
        # 创建插入位置矩形
        logo_rect = fitz.Rect(x0, y0, x1, y1)
        
        # 在右上角插入logo
        page.insert_image(logo_rect, filename=logo_path, keep_proportion=True)
        changed = True
        print(
            f"页面 {page_index + 1} 右上角已添加 {os.path.basename(logo_path)}（缩小10%），位置: "
            f"({x0:.1f}, {y0:.1f}) - ({x1:.1f}, {y1:.1f}), 尺寸: {scaled_logo_width:.1f} x {scaled_logo_height:.1f}"
        )

    if changed:
        tmp_path = pdf_path + ".topright.tmp"
        doc.save(tmp_path, deflate=True)
        doc.close()
        os.replace(tmp_path, pdf_path)
        print(f"已完成右上角 logo 添加并保存到原文件: {pdf_path}")
    else:
        doc.close()
        print(f"未成功添加右上角 logo: {pdf_path}")


def add_header_document_code(pdf_path: str, region_code: str):
    """
    在PDF每页的页眉右侧位置添加文档编码
    
    参数:
        pdf_path: PDF文件路径
        region_code: 地区编码
    """
    import random
    
    # 生成随机六位整数
    random_int = random.randint(100000, 999999)
    document_code = f"{region_code}-{random_int}"
    
    doc = None
    try:
        doc = fitz.open(pdf_path)
        changed = False
        
        # 遍历每一页
        for page_index in range(len(doc)):
            page = doc[page_index]
            page_rect = page.rect
            
            # 页眉位置：偏右侧，距离顶部约10-15px，距离右边缘约20-30px
            # 字号设置为8（小字号）
            font_size = 8
            margin_top = 20  # 距离顶部12px
            margin_right = 75  # 距离右边缘25px
            
            # 计算文本位置（右对齐）
            # 文本框右边界距离右边缘margin_right
            textbox_right = page_rect.width - margin_right
            textbox_top = margin_top
            # 估算文本框宽度（足够宽以容纳文本）
            estimated_text_width = len(document_code) * font_size * 0.7
            textbox_left = max(0, textbox_right - estimated_text_width * 2)  # 给足够的宽度
            textbox_bottom = textbox_top + font_size * 2  # 给足够的高度
            
            # 使用浅灰色 (0.7, 0.7, 0.7)
            text_color = (0.7, 0.7, 0.7)
            
            try:
                # 创建文本框矩形（右对齐）
                textbox_rect = fitz.Rect(
                    textbox_left,
                    textbox_top,
                    textbox_right,
                    textbox_bottom
                )
                
                # 使用insert_textbox插入文本，右对齐（align=2）
                rc = page.insert_textbox(
                    textbox_rect,
                    document_code,
                    fontsize=font_size,
                    fontname="china-s",  # 使用中文字体支持
                    color=text_color,
                    align=2  # 2表示右对齐
                )
                
                if rc >= 0:
                    changed = True
                    print(f"页面 {page_index + 1} 已添加文档编码: {document_code}")
                else:
                    print(f"警告: 页面 {page_index + 1} 添加文档编码失败，返回码: {rc}")
                    
            except Exception as e:
                print(f"页面 {page_index + 1} 添加文档编码时出错: {e}")
        
        if changed:
            tmp_path = pdf_path + ".header.tmp"
            doc.save(tmp_path, deflate=True)
            doc.close()
            os.replace(tmp_path, pdf_path)
            print(f"已完成页眉文档编码添加并保存到原文件: {pdf_path}")
        else:
            doc.close()
            print(f"未成功添加页眉文档编码: {pdf_path}")
            
    except Exception as e:
        if doc is not None:
            try:
                doc.close()
            except:
                pass
        print(f"添加页眉文档编码时出错: {e}")
        import traceback
        print(traceback.format_exc())


class LoginDialog(QDialog):
    """登录对话框"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.employee_id = None
        self.employee_name = None
        self.region_code = None
        self.init_ui()
        self.load_saved_info()  # 加载保存的登录信息
    
    def init_ui(self):
        self.setWindowTitle("登录")
        self.setGeometry(300, 300, 350, 200)
        self.setModal(True)  # 设置为模态对话框
        
        layout = QVBoxLayout()
        self.setLayout(layout)
        
        # 标题
        title_label = QLabel("请输入工号、姓名和地区编码")
        title_label.setStyleSheet("font-size: 16px; font-weight: bold; margin: 10px;")
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)
        
        # 工号输入
        id_layout = QHBoxLayout()
        id_label = QLabel("工号:")
        id_label.setMinimumWidth(80)
        self.id_input = QLineEdit()
        self.id_input.setPlaceholderText("请输入工号")
        id_layout.addWidget(id_label)
        id_layout.addWidget(self.id_input)
        layout.addLayout(id_layout)
        
        # 姓名输入
        name_layout = QHBoxLayout()
        name_label = QLabel("姓名:")
        name_label.setMinimumWidth(80)
        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("请输入姓名")
        name_layout.addWidget(name_label)
        name_layout.addWidget(self.name_input)
        layout.addLayout(name_layout)
        
        # 地区编码输入
        region_layout = QHBoxLayout()
        region_label = QLabel("地区编码:")
        region_label.setMinimumWidth(80)
        self.region_input = QLineEdit()
        self.region_input.setPlaceholderText("请输入地区编码")
        region_layout.addWidget(region_label)
        region_layout.addWidget(self.region_input)
        layout.addLayout(region_layout)
        
        # 按钮
        button_layout = QHBoxLayout()
        login_btn = QPushButton("登录")
        login_btn.clicked.connect(self.login)
        login_btn.setStyleSheet("padding: 8px; font-size: 12px; background-color: #4CAF50; color: white;")
        button_layout.addWidget(login_btn)
        layout.addLayout(button_layout)
    
    def load_saved_info(self):
        """加载保存的登录信息"""
        config = load_config()
        saved_id = config.get("employee_id", "")
        saved_name = config.get("employee_name", "")
        saved_region = config.get("region_code", "")
        
        if saved_id:
            self.id_input.setText(saved_id)
        if saved_name:
            self.name_input.setText(saved_name)
        if saved_region:
            self.region_input.setText(saved_region)
    
    def login(self):
        """处理登录"""
        employee_id = self.id_input.text().strip()
        employee_name = self.name_input.text().strip()
        region_code = self.region_input.text().strip()
        
        if not employee_id:
            QMessageBox.warning(self, "警告", "请输入工号！")
            return
        
        if not employee_name:
            QMessageBox.warning(self, "警告", "请输入姓名！")
            return
        
        if not region_code:
            QMessageBox.warning(self, "警告", "请输入地区编码！")
            return
        
        self.employee_id = employee_id
        self.employee_name = employee_name
        self.region_code = region_code
        
        # 保存登录信息到配置文件
        save_login_info(employee_id, employee_name, region_code)
        
        self.accept()  # 接受对话框，返回 QDialog.Accepted
    
    def get_login_info(self):
        """获取登录信息"""
        return self.employee_id, self.employee_name, self.region_code


class PDFPageRemoverGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.pdf_files = []
        
        # 显示登录对话框
        login_dialog = LoginDialog(self)
        result = login_dialog.exec_()
        
        # 获取登录信息
        self.employee_id, self.employee_name, self.region_code = login_dialog.get_login_info()
        
        # 如果用户取消登录或未输入信息，关闭应用
        if result != QDialog.Accepted or not self.employee_id or not self.employee_name or not self.region_code:
            sys.exit(0)
        
        # 从配置文件加载保存的目录和华为云配置
        config = load_config()
        self.output_dir = config.get("output_dir", "")
        self.image_output_dir = config.get("image_output_dir", "")
        self.huawei_token = None
        
        self.init_ui()
        
        # 应用启动时尝试获取Token
        self.get_token_on_startup()
        
    def init_ui(self):
        # 窗口标题显示当前登录用户和地区编码
        user_info = f"{self.employee_id}-{self.employee_name}" if self.employee_id and self.employee_name else ""
        region_info = f"[{self.region_code}]" if self.region_code else ""
        self.setWindowTitle(f"PDF页面修改工具[{user_info}] {region_info}")
        self.setGeometry(100, 100, 800, 600)
        
        # 主窗口部件
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        
        # 主布局
        layout = QVBoxLayout()
        main_widget.setLayout(layout)
        
        # 标题区域（包含标题和退出登录按钮）
        title_layout = QHBoxLayout()
        title_label = QLabel("PDF页面修改工具")
        title_label.setStyleSheet("font-size: 20px; font-weight: bold; margin: 10px;")
        title_label.setAlignment(Qt.AlignCenter)
        title_layout.addWidget(title_label)
        
        # 退出登录按钮
        logout_btn = QPushButton("退出登录")
        logout_btn.clicked.connect(self.logout)
        logout_btn.setStyleSheet("padding: 6px; font-size: 11px; background-color: #f44336; color: white;")
        title_layout.addWidget(logout_btn)
        
        layout.addLayout(title_layout)
        
        # PDF文件选择区域
        pdf_layout = QVBoxLayout()
        pdf_label = QLabel("选择的PDF文件:")
        pdf_label.setStyleSheet("font-weight: bold;")
        pdf_layout.addWidget(pdf_label)
        
        # 文件列表显示
        self.file_list = QTextEdit()
        self.file_list.setReadOnly(True)
        self.file_list.setMaximumHeight(150)
        self.file_list.setPlaceholderText("未选择文件")
        pdf_layout.addWidget(self.file_list)
        
        # 选择PDF按钮
        select_pdf_btn = QPushButton("选择PDF文件（按住shift多选）")
        select_pdf_btn.clicked.connect(self.select_pdf_files)
        select_pdf_btn.setStyleSheet("padding: 8px; font-size: 12px;")
        pdf_layout.addWidget(select_pdf_btn)
        
        layout.addLayout(pdf_layout)
        
        # 输出目录选择区域
        output_layout = QVBoxLayout()
        output_label = QLabel("输出目录:")
        output_label.setStyleSheet("font-weight: bold;")
        output_layout.addWidget(output_label)
        
        self.output_path_label = QLabel("未选择输出目录")
        self.output_path_label.setStyleSheet("padding: 5px; background-color: #f0f0f0; border: 1px solid #ccc;")
        # 如果有保存的输出目录，显示它
        if self.output_dir:
            self.output_path_label.setText(self.output_dir)
        output_layout.addWidget(self.output_path_label)
        
        # 选择输出目录按钮
        select_output_btn = QPushButton("选择输出目录")
        select_output_btn.clicked.connect(self.select_output_directory)
        select_output_btn.setStyleSheet("padding: 8px; font-size: 12px;")
        output_layout.addWidget(select_output_btn)
        
        layout.addLayout(output_layout)
        
        # 图片输出目录选择区域
        image_output_layout = QVBoxLayout()
        image_output_label = QLabel("图片输出目录（必选）:")
        image_output_label.setStyleSheet("font-weight: bold;")
        image_output_layout.addWidget(image_output_label)
        
        self.image_output_path_label = QLabel("未选择图片输出目录（将产生额外磁盘垃圾）")
        self.image_output_path_label.setStyleSheet("padding: 5px; background-color: #f0f0f0; border: 1px solid #ccc;")
        # 如果有保存的图片输出目录，显示它
        if self.image_output_dir:
            self.image_output_path_label.setText(self.image_output_dir)
        image_output_layout.addWidget(self.image_output_path_label)
        
        # 选择图片输出目录按钮
        select_image_output_btn = QPushButton("选择图片输出目录（必选）")
        select_image_output_btn.clicked.connect(self.select_image_output_directory)
        select_image_output_btn.setStyleSheet("padding: 8px; font-size: 12px;")
        image_output_layout.addWidget(select_image_output_btn)
        
        layout.addLayout(image_output_layout)
        
        # 日期选择区域
        date_layout = QHBoxLayout()
        date_label = QLabel("更新日期:")
        date_label.setStyleSheet("font-weight: bold;")
        date_layout.addWidget(date_label)
        
        # 日期选择组件，默认当前日期
        self.date_edit = QDateEdit()
        self.date_edit.setDate(QDate.currentDate())  # 设置为当前日期
        self.date_edit.setCalendarPopup(True)  # 启用日历弹出
        self.date_edit.setDisplayFormat("yyyy-MM-dd")  # 设置日期格式
        self.date_edit.setStyleSheet("padding: 5px;")
        date_layout.addWidget(self.date_edit)
        
        date_layout.addStretch()  # 添加弹性空间，使日期选择器靠左
        layout.addLayout(date_layout)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)
        
        # 状态显示区域
        status_label = QLabel("处理状态:")
        status_label.setStyleSheet("font-weight: bold;")
        layout.addWidget(status_label)
        
        self.status_text = QTextEdit()
        self.status_text.setReadOnly(True)
        self.status_text.setPlaceholderText("等待开始处理...")
        layout.addWidget(self.status_text)
        
        # 处理按钮
        process_btn = QPushButton("开始处理")
        process_btn.clicked.connect(self.start_processing)
        process_btn.setStyleSheet("padding: 10px; font-size: 14px; font-weight: bold; background-color: #4CAF50; color: white;")
        layout.addWidget(process_btn)
        
        # 清空按钮
        clear_btn = QPushButton("清空列表")
        clear_btn.clicked.connect(self.clear_files)
        clear_btn.setStyleSheet("padding: 8px; font-size: 12px;")
        layout.addWidget(clear_btn)
        
    def select_pdf_files(self):
        """选择PDF文件（支持多选）"""
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "选择PDF文件",
            "",
            "PDF文件 (*.pdf);;所有文件 (*.*)"
        )
        
        if files:
            self.pdf_files = files
            file_names = [os.path.basename(f) for f in files]
            self.file_list.setPlainText("\n".join(file_names))
            self.status_text.append(f"已选择 {len(files)} 个PDF文件")
    
    def select_output_directory(self):
        """选择输出目录"""
        # 如果有保存的目录，从该目录开始选择
        start_dir = self.output_dir if self.output_dir and os.path.exists(self.output_dir) else ""
        
        directory = QFileDialog.getExistingDirectory(
            self,
            "选择输出目录",
            start_dir
        )
        
        if directory:
            # 规范化路径，确保使用正确的路径分隔符
            self.output_dir = os.path.normpath(directory)
            self.output_path_label.setText(self.output_dir)
            self.status_text.append(f"PDF输出目录: {self.output_dir}")
            # 保存配置
            save_config(self.output_dir, self.image_output_dir)
    
    def select_image_output_directory(self):
        """选择图片输出目录"""
        # 如果有保存的目录，从该目录开始选择
        start_dir = self.image_output_dir if self.image_output_dir and os.path.exists(self.image_output_dir) else ""
        
        directory = QFileDialog.getExistingDirectory(
            self,
            "选择图片输出目录",
            start_dir
        )
        
        if directory:
            # 规范化路径，确保使用正确的路径分隔符
            self.image_output_dir = os.path.normpath(directory)
            self.image_output_path_label.setText(self.image_output_dir)
            self.status_text.append(f"图片输出目录: {self.image_output_dir}")
            # 保存配置
            save_config(self.output_dir, self.image_output_dir)
    
    def clear_files(self):
        """清空文件列表"""
        self.pdf_files = []
        self.file_list.clear()
        self.status_text.append("已清空文件列表")
    
    def logout(self):
        """退出登录"""
        # 检查是否有正在进行的处理
        if hasattr(self, 'processor_thread') and self.processor_thread.isRunning():
            reply = QMessageBox.question(
                self, 
                "确认退出", 
                "处理正在进行中，确定要退出登录吗？",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            if reply == QMessageBox.No:
                return
            # 如果用户确认退出，停止处理线程
            self.processor_thread.terminate()
            self.processor_thread.wait()
        
        # 确认退出登录
        reply = QMessageBox.question(
            self,
            "确认退出登录",
            "确定要退出登录吗？退出后需要重新登录才能使用。",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            # 清除登录信息
            clear_login_info()
            # 关闭应用
            QApplication.quit()
    
    def start_processing(self):
        """开始处理PDF文件"""
        # 检查是否已有线程在运行
        if hasattr(self, 'processor_thread') and self.processor_thread.isRunning():
            QMessageBox.warning(self, "警告", "处理正在进行中，请等待完成！")
            return
        
        if not self.pdf_files:
            QMessageBox.warning(self, "警告", "请先选择PDF文件！")
            return
        
        if not self.output_dir:
            QMessageBox.warning(self, "警告", "请先选择输出目录！")
            return
        
        # 规范化并检查输出目录是否存在
        self.output_dir = os.path.normpath(self.output_dir)
        if not os.path.exists(self.output_dir):
            try:
                os.makedirs(self.output_dir, exist_ok=True)
            except Exception as e:
                QMessageBox.critical(self, "错误", f"无法创建输出目录: {str(e)}")
                return
        
        # 规范化图片输出目录
        if self.image_output_dir:
            self.image_output_dir = os.path.normpath(self.image_output_dir)
            if not os.path.exists(self.image_output_dir):
                try:
                    os.makedirs(self.image_output_dir, exist_ok=True)
                except Exception as e:
                    QMessageBox.critical(self, "错误", f"无法创建图片输出目录: {str(e)}")
                    return
        
        # 重置进度条
        self.progress_bar.setValue(0)
        self.status_text.clear()
        self.status_text.append("开始处理...")
        
        # 如果存在旧的线程，先断开连接（防止信号重复连接）
        if hasattr(self, 'processor_thread'):
            try:
                self.processor_thread.progress.disconnect()
                self.processor_thread.status.disconnect()
                self.processor_thread.finished.disconnect()
            except:
                pass  # 如果连接不存在，忽略错误
        
        # 获取选择的日期并转换为datetime对象
        selected_date = self.date_edit.date()
        update_date = datetime(selected_date.year(), selected_date.month(), selected_date.day())
        
        # 创建并启动处理线程
        self.processor_thread = PDFProcessorThread(
            self.pdf_files, 
            self.output_dir, 
            self.image_output_dir,
            update_date,  # 传递选择的日期
            self.employee_id,  # 传递工号
            self.employee_name,  # 传递姓名
            self.region_code,  # 传递地区编码
            self  # 传递父窗口引用，用于访问huawei_token
        )
        self.processor_thread.progress.connect(self.update_progress)
        self.processor_thread.status.connect(self.update_status)
        self.processor_thread.finished.connect(self.processing_finished)
        self.processor_thread.start()
    
    def update_progress(self, value):
        """更新进度条"""
        self.progress_bar.setValue(value)
    
    def update_status(self, message):
        """更新状态信息"""
        self.status_text.append(message)
        # 自动滚动到底部
        self.status_text.verticalScrollBar().setValue(
            self.status_text.verticalScrollBar().maximum()
        )
    
    def processing_finished(self):
        """处理完成"""
        self.status_text.append("\n✓ 所有文件处理完成！")
        QMessageBox.information(self, "完成", "所有PDF文件处理完成！")
    
    def get_token_on_startup(self):
        """应用启动时获取Token"""
        config = load_config()
        username = config.get("huawei_username", "")
        domain = config.get("huawei_domain", "")
        password = config.get("huawei_password", "")
        project = config.get("huawei_project", "cn-north-4")
        
        if username and domain and password:
            self.status_text.append("正在获取华为云Token...")
            self.status_text.repaint()  # 立即更新界面
            
            token = get_huawei_token(username, domain, password, project)
            if token:
                self.huawei_token = token
                self.status_text.append("✓ 华为云Token获取成功（有效期24小时）")
            else:
                self.status_text.append("✗ 华为云Token获取失败，请检查配置")
        else:
            self.status_text.append("提示: 未配置华为云账号信息，Token获取已跳过")
            self.status_text.append("如需使用OCR功能，请在配置文件中添加华为云账号信息")


def main():
    app = QApplication(sys.argv)
    window = PDFPageRemoverGUI()
    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()

