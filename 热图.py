import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import numpy as np
import os
import matplotlib as mpl
from matplotlib.font_manager import FontProperties, fontManager
import shutil

def setup_chinese_support():
    """设置中文支持功能"""
    try:
        # 检查并清除matplotlib缓存目录
        cache_dir = mpl.get_cachedir()
        print(f"Matplotlib缓存目录: {cache_dir}")
        if os.path.exists(cache_dir):
            print("清除matplotlib缓存...")
            shutil.rmtree(cache_dir)
    except Exception as e:
        print(f"清除缓存出错: {e}")

    # 手动添加SimHei字体（适用于Windows）
    if os.name == 'nt':  # Windows系统
        system_font_dir = os.path.join(os.environ['WINDIR'], 'Fonts')
        simhei_path = os.path.join(system_font_dir, 'simhei.ttf')
        if os.path.exists(simhei_path):
            # 添加到字体管理器
            fontManager.addfont(simhei_path)
            print(f"已添加SimHei字体: {simhei_path}")
    
    # 尝试不同的中文字体
    chinese_fonts = ['SimHei', 'Microsoft YaHei', 'KaiTi', 'STSong', 'STHeiti', 'STFangsong', 'FangSong', 'SimSun']
    
    # 查找可用字体
    system_fonts = set([f.name for f in fontManager.ttflist])
    print("系统可用字体:")
    for font in sorted(system_fonts)[:20]:  # 只显示前20个
        print(f" - {font}")
    
    # 选择第一个可用的中文字体
    selected_font = None
    for font in chinese_fonts:
        if any(f for f in system_fonts if font.lower() in f.lower()):
            selected_font = font
            print(f"使用中文字体: {selected_font}")
            break
    
    if selected_font:
        # 设置图表字体
        plt.rcParams['font.sans-serif'] = [selected_font]
        plt.rcParams['axes.unicode_minus'] = False
        
        # 找到字体文件路径
        font_path = None
        for f in fontManager.ttflist:
            if selected_font.lower() in f.name.lower():
                font_path = f.fname
                break
        
        if font_path:
            print(f"字体路径: {font_path}")
            font_prop = FontProperties(fname=font_path)
        else:
            print("警告: 无法找到字体文件路径!")
            font_prop = FontProperties(family=selected_font)
    else:
        print("警告: 未找到任何中文字体!")
        font_prop = FontProperties()
    
    return font_prop

def plot_correlation_heatmap(data, title, output_file, font_prop, max_columns=15):
    """
    为给定的数据绘制相关性热力图
    
    参数:
        data: 包含数据的DataFrame
        title: 图表标题
        output_file: 输出文件名
        font_prop: 字体属性
        max_columns: 最大列数（避免太多列导致图像无法显示）
    """
    # 只选择数值类型的列
    numeric_data = data.select_dtypes(include=[np.number])
    
    # 如果列太多，只显示强相关关系（|相关性|>0.5)
    if len(numeric_data.columns) > max_columns:
        print(f"列数太多({len(numeric_data.columns)})，只显示强相关关系（|相关性|>0.5)")
        corr_matrix = numeric_data.corr()
        
        # 创建一个布尔矩阵，标记强相关关系
        strong_corr = np.abs(corr_matrix) > 0.5
        # 找出至少有一个强相关关系大于0.5的列
        columns_to_keep = corr_matrix.columns[strong_corr.sum(axis=1) > 1].tolist()
        
        # 确保保留至少一些列
        if len(columns_to_keep) < 5:
            # 如果太少，保留相关性最高的前10列
            all_columns = corr_matrix.abs().sum().sort_values(ascending=False).index.tolist()
            columns_to_keep = all_columns[:min(10, len(all_columns))]
        
        numeric_data = numeric_data[columns_to_keep]
        print(f"保留列: {', '.join(columns_to_keep)}")
    
    # 计算相关性矩阵
    corr_matrix = numeric_data.corr()
    
    # 创建更大的图形
    plt.figure(figsize=(8, 10))
    
    # 绘制热力图
    ax = sns.heatmap(
        corr_matrix,
        annot=True,
        fmt=".2f",
        cmap='coolwarm',
        center=0,
        square=True,
        linewidths=.5,
        annot_kws={'size': 8, 'fontproperties': font_prop, 'horizontalalignment': 'center'}
    )
    
    # 设置中文标题
    ax.set_title(title, 
                fontsize=18, 
                pad=20,
                fontproperties=font_prop)
    
    # 旋转X轴标签
    ax.set_xticklabels(
        ax.get_xticklabels(), 
        rotation=45, 
        ha='right',
        fontproperties=font_prop,
        fontsize=11
    )
    
    # 设置Y轴标签
    ax.set_yticklabels(
        ax.get_yticklabels(), 
        rotation=0,
        ha='right',
        fontproperties=font_prop,
        fontsize=11
    )
    
    # 添加边框
    for _, spine in ax.spines.items():
        spine.set_visible(True)
        spine.set_linewidth(1)
    
    # 调整布局防止标签被裁剪
    plt.tight_layout()
    
    # 保存图形
    print(f"保存文件到: {output_file}")
    plt.savefig(output_file, dpi=300, bbox_inches='tight')
    
    # 显示图形
    plt.show()
    plt.close()

# 设置中文支持
font_prop = setup_chinese_support()

# 读取Excel文件
file_path = r'D:\1793672313\Desktop\下载 (1)\河湖源草地土壤理化性质.xlsx'

# 读取所有工作表
xls = pd.ExcelFile(file_path)
sheet_names = xls.sheet_names
print(f"工作表列表: {', '.join(sheet_names)}")

# 处理每个工作表
for i, sheet_name in enumerate(sheet_names):
    # 跳过系统生成的工作表
    if sheet_name.lower().startswith('sheet') or sheet_name.lower().startswith('工作表'):
        print(f"跳过系统生成的工作表: {sheet_name}")
        continue
    
    print(f"\n处理工作表: {sheet_name}")
    
    # 读取指定工作表
    data = pd.read_excel(file_path, sheet_name=sheet_name)
    
    # 打印所有列名确保中文正确读取
    print("数据列名:")
    for col in data.columns:
        print(col)
    
    # 为不同工作表设置不同的图表标题和文件名
    if "0-10cm" in sheet_name or "0-20cm" in sheet_name or "表一" in sheet_name:
        title = f'河湖源草地土壤理化性质相关性分析 ({sheet_name})'
        output_file = f'表一_{sheet_name}_相关性热图.png'
    else:
        title = f'河湖源草地土壤理化性质相关性分析 ({sheet_name})'
        output_file = f'表二_{sheet_name}_相关性热图.png'
    
    # 绘制热力图
    plot_correlation_heatmap(data, title, output_file, font_prop, max_columns=15)
    
    print(f"已完成工作表: {sheet_name}")

print("\n所有工作表处理完成!")
