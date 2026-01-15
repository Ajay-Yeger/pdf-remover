import matplotlib
# 使用 Agg 后端（非 GUI 后端，可以在后台线程中使用，适合生成图片文件）
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.patches import Wedge, Circle, FancyBboxPatch
from matplotlib.colors import LinearSegmentedColormap
import numpy as np
from datetime import datetime

# 设置matplotlib支持中文
matplotlib.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'Arial Unicode MS', 'DejaVu Sans']
matplotlib.rcParams['axes.unicode_minus'] = False


def create_credit_score_visualization(score=1206, update_date=None, output_path='credit_score.png'):
    """
    创建信用评分可视化图表
    
    参数:
        score: 信用分数 (0-2000)
        update_date: 更新日期，可以是datetime对象或字符串，格式为"YYYY年MM月DD日"
        output_path: 输出PNG文件路径
    """
    # 设置图形大小和DPI
    fig, ax = plt.subplots(figsize=(18, 7), dpi=150)
    ax.set_xlim(0, 18)
    ax.set_ylim(0, 7)
    ax.axis('off')
    
    # 设置背景为透明
    fig.patch.set_facecolor('none')
    ax.set_facecolor('none')
    
    # 如果没有提供日期，使用当前日期
    if update_date is None:
        update_date = datetime.now()
    
    # 处理日期格式
    if isinstance(update_date, str):
        date_str = update_date
    else:
        date_str = update_date.strftime('%Y年%m月%d日')
    
    # 确保分数在有效范围内
    score = max(0, min(2000, score))
    
    # ========== 左侧圆形仪表盘 ==========
    center_x, center_y = 3, 3.5
    radius_outer = 1.3
    radius_inner = 1.1
    ring_width = radius_outer - radius_inner
    
    # 黄色部分在底部左侧，大约占15-20%的范围
    # 从大约225度开始到270度（底部左侧区域）
    yellow_start_angle = 225  # 左下角
    yellow_end_angle = 270     # 正下方
    yellow_span = yellow_end_angle - yellow_start_angle
    
    # 绘制黄色部分（底部左侧）
    yellow_wedge = Wedge((center_x, center_y), radius_outer, yellow_start_angle, yellow_end_angle, 
                        width=ring_width, facecolor='#FFD700', edgecolor='none', zorder=1)
    ax.add_patch(yellow_wedge)
    
    # 绘制绿色部分（其余大部分）
    green_wedge1 = Wedge((center_x, center_y), radius_outer, yellow_end_angle, yellow_start_angle + 360, 
                        width=ring_width, facecolor='#0E9643', edgecolor='none', zorder=1)
    ax.add_patch(green_wedge1)
    
    # 绘制白色内圆
    inner_circle = Circle((center_x, center_y), radius_inner, 
                         facecolor='white', edgecolor='none', zorder=2)
    ax.add_patch(inner_circle)
    
    # 绘制分数文本（大号绿色）
    ax.text(center_x, center_y, str(int(score)), 
           fontsize=60, fontweight='bold', color='#0E9643',
           ha='center', va='center', zorder=3)
    
    # 绘制更新时间文本（在圆环底部中心）
    ax.text(center_x, center_y - radius_inner - 0.4, f'上次更新:{date_str}', 
           fontsize=13, color='#333333', ha='center', va='top', zorder=3)
    
    # ========== 右侧水平渐变条 ==========
    bar_x_start, bar_x_end = 5, 17.5  # 长度从10增加到12.5（增加1/4）
    bar_y = 3.3  # 向下移动0.2，使与圆环中部对齐
    bar_height = 0.4
    
    # 创建渐变颜色映射（黄色到绿色）
    colors = ['#FFD700', '#FFEB3B', '#CDDC39', '#8BC34A', '#4CAF50', '#388E3C']
    n_bins = 200  # 增加分段数以获得更平滑的渐变
    cmap = LinearSegmentedColormap.from_list('credit_gradient', colors, N=n_bins)
    
    # 绘制渐变条（带圆角效果）
    bar_width = bar_x_end - bar_x_start
    corner_radius = bar_height / 2  # 圆角半径等于高度的一半，形成完全圆角（pill形状）
    
    # 使用imshow绘制渐变，然后用圆角路径裁剪
    # 创建渐变数据（水平方向）
    gradient_data = np.linspace(0, 1, n_bins).reshape(1, -1)
    # 扩展为多行以获得更好的渲染效果
    gradient_data = np.repeat(gradient_data, 50, axis=0)
    
    # 使用imshow绘制渐变
    extent = [bar_x_start, bar_x_end, bar_y - bar_height/2, bar_y + bar_height/2]
    im = ax.imshow(gradient_data, aspect='auto', extent=extent, 
                   cmap=cmap, interpolation='bilinear', zorder=1, origin='lower')
    
    # 创建圆角矩形路径用于裁剪
    rounded_rect = FancyBboxPatch((bar_x_start, bar_y - bar_height/2), bar_width, bar_height,
                                  boxstyle=f"round,pad=0,rounding_size={corner_radius}",
                                  transform=ax.transData)
    
    # 设置裁剪路径 - 使用圆角矩形的路径和变换
    im.set_clip_path(rounded_rect)
    
    # 计算分数在条上的位置
    score_position = bar_x_start + (bar_x_end - bar_x_start) * (score / 2000)
    
    # 绘制分数标记线（垂直线）
    marker_line_y_bottom = bar_y - bar_height/2 - 0.7  # 向下延伸更多
    circle_radius = 0.3
    
    # 绘制垂直线（从进度条到圆环上边缘）
    ax.plot([score_position, score_position], [bar_y - bar_height/2, marker_line_y_bottom + circle_radius],
           color='#0E9643', linewidth=4, zorder=2)
    
    # 绘制垂直线（从圆环下边缘到圆环中心下方）
    ax.plot([score_position, score_position], [marker_line_y_bottom - circle_radius, marker_line_y_bottom],
           color='#0E9643', linewidth=4, zorder=2)
    
    # 绘制白色填充圆来遮挡垂直线在圆环内的部分
    white_circle = Circle((score_position, marker_line_y_bottom), circle_radius,
                         facecolor='white', edgecolor='none', zorder=3)
    ax.add_patch(white_circle)
    
    # 绘制分数标记圆环（只有边框，无填充）
    score_circle = Circle((score_position, marker_line_y_bottom), circle_radius,
                         facecolor='none', edgecolor='#0E9643', linewidth=4, zorder=4)
    ax.add_patch(score_circle)
    
    # 在圆环中显示分数（绿色文字）
    ax.text(score_position, marker_line_y_bottom, str(int(score)),
           fontsize=16, fontweight='bold', color='#0E9643',
           ha='center', va='center', zorder=5)
    
    # ========== 标签文本 ==========
    # 所有标签都在进度条上方，文字标签在数字标签上方
    label_y_text = bar_y + bar_height/2 + 0.5     # 文字标签位置（最上方）
    label_y_numbers = bar_y + bar_height/2 + 0.3  # 数字标签位置（文字下方）
    
    # 左侧标签：信用表现一般 (0)
    ax.text(bar_x_start+0.1, label_y_text, '信用表现一般',
           fontsize=13, color='#333333', ha='left', va='bottom', zorder=3)
    ax.text(bar_x_start+0.1, label_y_numbers-0.15, '0',
           fontsize=13, color='#333333', ha='left', va='bottom', zorder=3)
    
    # 中间标签：基础分段 (600-700)
    mid_x = (bar_x_start + bar_x_end) / 2
    ax.text(mid_x, label_y_text, '基础分段',
           fontsize=13, color='#333333', ha='center', va='bottom', zorder=3)
    ax.text(mid_x, label_y_numbers-0.15, '600-700',
           fontsize=13, color='#333333', ha='center', va='bottom', zorder=3)
    
    # 右侧标签：信用表现卓越 (2000)
    ax.text(bar_x_end-0.1, label_y_text, '信用表现卓越',
           fontsize=13, color='#333333', ha='right', va='bottom', zorder=3)
    ax.text(bar_x_end-0.1, label_y_numbers-0.15, '2000',
           fontsize=13, color='#333333', ha='right', va='bottom', zorder=3)
    
    # 保存图片（透明背景）
    plt.tight_layout()
    plt.savefig(output_path, dpi=150, bbox_inches='tight', transparent=True, facecolor='none', edgecolor='none')
    plt.close()
    
    print(f"信用评分可视化已保存到: {output_path}")


if __name__ == '__main__':
    # 示例使用
    # 使用默认值
    #create_credit_score_visualization()
    
    # 自定义分数和日期
    create_credit_score_visualization(
        score=1500,
        update_date='2025年12月15日',
        output_path='credit_score_custom.png'
    )
    
    # # 使用datetime对象
    # create_credit_score_visualization(
    #     score=800,
    #     update_date=datetime(2025, 12, 20),
    #     output_path='credit_score_datetime.png'
    # )

