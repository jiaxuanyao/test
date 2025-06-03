import pandas as pd
import matplotlib.pyplot as plt
import matplotlib as mpl
import numpy as np
import os
from matplotlib.font_manager import FontProperties
from matplotlib.patches import Patch
from datetime import datetime

# --- 中文支持设置 ---
def set_chinese_font(font_size=12):
    """设置中文字体支持"""
    try:
        # 获取系统可用字体
        fm = mpl.font_manager.FontManager()
        available_fonts = [f.name for f in fm.ttflist]
        
        # 优选字体列表
        pref_fonts = ['Microsoft YaHei', 'SimHei', 'KaiTi', 'STSong', 'NSimSun', 
                      'FangSong', 'SimSun', 'SimHei', 'Microsoft JhengHei']
        
        # 找到第一个可用中文字体
        for font in pref_fonts:
            if any(font in f for f in available_fonts):
                mpl.rcParams['font.family'] = 'sans-serif'
                mpl.rcParams['font.sans-serif'] = [font]
                mpl.rcParams['axes.unicode_minus'] = False
                print(f"使用中文字体: {font}")
                break
        
        # 设置全局字体大小
        plt.rcParams.update({
            'font.size': font_size,
            'axes.titlesize': font_size,
            'axes.labelsize': font_size,
            'xtick.labelsize': font_size - 2,
            'ytick.labelsize': font_size - 2,
            'legend.fontsize': font_size - 2,
            'figure.titlesize': font_size
        })
    except Exception as e:
        print(f"设置中文字体时出错: {e}")

# 初始化中文字体
set_chinese_font(10)

# --- 土壤质地分类函数 ---
def soil_texture_classification(sand, silt, clay):
    """根据砂粒、粉粒和黏粒含量分类土壤质地"""
    total = sand + silt + clay
    if abs(total - 100) > 1:  # 允许1%的容差
        sand = (sand / total) * 100
        silt = (silt / total) * 100
        clay = (clay / total) * 100
        
    if sand >= 85:
        if clay < 10 and silt < 15: return '砂土'
        return '壤质砂土'
    elif sand >= 70:
        if clay < 15 and silt < 15: return '砂土'
        if clay < 20 and silt < 30: return '砂质壤土'
    elif sand >= 52:
        if silt < 28 and clay < 15: return '壤质砂土'
        if silt < 30 and clay < 25: return '砂质壤土'
    elif sand >= 43:
        if silt < 50 and clay < 15: return '砂质壤土'
        if silt < 53 and clay < 20 and sand < 52: return '壤土'
    elif silt >= 50:
        if clay < 27: return '粉砂质壤土'
        if clay < 35: return '粉砂质黏壤土'
        return '粉砂质黏土'
    elif clay < 20:
        if silt < 30: return '壤质砂土'
        if silt < 50: return '壤土'
        return '粉砂质壤土'
    elif clay < 35:
        if sand < 45: return '黏壤土'
        return '砂质黏壤土'
    else:
        if sand > 45: return '砂质黏土'
        if silt > 40: return '粉砂质黏土'
        return '黏土'

# --- 主程序 ---
def main():
    # 文件路径
    file_path = r'D:\1793672313\Desktop\下载 (1)\河湖源草地土壤理化性质.xlsx'
    
    if not os.path.exists(file_path):
        print(f"错误: 文件不存在 - {file_path}")
        return
    
    try:
        # 读取Excel文件
        excel_file = pd.ExcelFile(file_path)
        
        # 选择要处理的工作表 (0-10cm和10-20cm)
        target_sheets = []
        for sheet_name in excel_file.sheet_names:
            if '0-10' in sheet_name:
                target_sheets.append(('0-10cm', sheet_name))
            elif '0-20' in sheet_name or '10-20' in sheet_name:
                target_sheets.append(('10-20cm', sheet_name))
            
            if len(target_sheets) >= 2:
                break
        
        if len(target_sheets) < 2:
            print("警告: 未找到足够的工作表")
            target_sheets = [(name, name) for name in excel_file.sheet_names[:2]]
        
        print("将处理的工作表:")
        for depth, name in target_sheets:
            print(f"{depth}: {name}")
        
        # 创建图形 - 增加垂直空间容纳下移的标签和注释
        plt.figure(figsize=(14, 8), dpi=120, facecolor='#f7f7f7')  # 增加高度
        ax = plt.subplot(111)
        
        # 设置土壤质地颜色方案
        soil_colors = {
            '砂土': '#4e79a7', 
            '壤质砂土': '#f28e2b', 
            '砂质壤土': '#e15759', 
            '壤土': '#76b7b2',
            '粉砂质壤土': '#59a14f', 
            '粉砂质黏壤土': '#edc948', 
            '黏壤土': '#b07aa1',
            '砂质黏壤土': '#ff9da7', 
            '粉砂': '#9c755f', 
            '粉砂质黏土': '#bab0ac',
            '黏土': '#d37295', 
            '砂质黏土': '#b3a2d0',
            '未分類': '#cccccc'
        }
        
        # 存储数据
        soil_data = {}
        all_soil_types = set()
        
        # 处理每个工作表
        for i, (depth, sheet_name) in enumerate(target_sheets):
            try:
                df = excel_file.parse(sheet_name)
                
                # 查找包含粒级的列
                sand_col, silt_col, clay_col = None, None, None
                for col in df.columns:
                    if '砂' in col: sand_col = col
                    elif '粉' in col or '泥' in col: silt_col = col
                    elif '粘' in col or '黏' in col: clay_col = col
                
                if not all([sand_col, silt_col, clay_col]):
                    print(f"警告: 工作表 {sheet_name} 缺少颗粒含量列")
                    continue
                
                # 应用土壤分类
                df['土壤质地'] = df.apply(lambda row: soil_texture_classification(
                    row[sand_col], row[silt_col], row[clay_col]), axis=1
                )
                
                # 统计土壤类型
                count_data = df['土壤质地'].value_counts().to_dict()
                soil_data[depth] = count_data
                
                # 收集所有土壤类型
                all_soil_types.update(count_data.keys())
                
                print(f"\n{depth} 土壤质地统计:")
                for soil, count in count_data.items():
                    print(f"  {soil}: {count}个样本")
            
            except KeyError as e:
                print(f"处理工作表 {sheet_name} 时出错: {str(e)}")
        
        # 确保所有土壤类型都包括
        soil_types = sorted(all_soil_types, reverse=True)
        
        # 设置柱状图参数
        n_types = len(soil_types)
        bar_width = 0.35
        index = np.arange(n_types)
        
        colors = [soil_colors.get(soil_type, '#777777') for soil_type in soil_types]
        
        # 绘制柱状图
        for i, depth in enumerate(soil_data.keys()):
            counts = [soil_data[depth].get(soil_type, 0) for soil_type in soil_types]
            
            # 计算柱子位置
            offset = (i - 0.5) * bar_width
            positions = index + offset
            
            # 创建柱状图 - 移除了label参数
            bars = ax.bar(
                positions, counts, bar_width,
                color=colors,
                edgecolor='white',
                linewidth=1,
                zorder=2
            )
            
            # 添加数据标签 - 显示每个柱子的数值
            for bar in bars:
                height = bar.get_height()
                if height > 0:
                    ax.text(
                        bar.get_x() + bar.get_width()/2 + bar_width/10,  # 微调避免重叠
                        height + 0.2,
                        int(height),
                        ha='center',
                        va='bottom',
                        fontsize=9,
                        color='#333333'
                    )
        
        # 添加分隔线 - 增强可读性
        for x in np.arange(0.5, n_types, 1):
            ax.axvline(x=x-0.5, color='lightgrey', linestyle=':', alpha=0.5)
        
        # 设置坐标轴
        ax.set_title('不同深度土壤质地分布比较', fontsize=13, pad=20, color='#333333')
        ax.set_ylabel('样本数量', fontsize=11, labelpad=8)
        ax.set_xlabel('土壤质地类型', fontsize=11, labelpad=15)  # 增加标签间距
        
        # ************** 修改：下移X轴标签并添加注释 **************
        
        # 设置X轴刻度和标签 - 留更多空间给注释
        ax.set_xticks(index)
        
        # 调整标签位置并添加注释
        label_line_height = -max(counts)*0.15  # 标签行的高度
        annotation_line_height = -max(counts)*0.25  # 注释行的高度
        
        # 添加土壤类型标签（X轴）
        for i, soil_type in enumerate(soil_types):
            ax.text(
                i, label_line_height, 
                soil_type, 
                ha='center', 
                va='top', 
                fontsize=10, 
                color='#333333',
                weight='bold'
            )
            
            # 添加注释信息 - 这里使用示例注释，实际应替换为您的注释
            # 示例注释：0-10cm深度数量和10-20cm深度数量
            count_0_10 = soil_data['0-10cm'].get(soil_type, 0)
            count_10_20 = soil_data['10-20cm'].get(soil_type, 0)
            
            comment = f"0-10cm: {count_0_10}份 | 10-20cm: {count_10_20}份"
            
            ax.text(
                i, annotation_line_height, 
                comment, 
                ha='center', 
                va='top', 
                fontsize=8, 
                color='#666666',
                alpha=0.8
            )
        
        # 隐藏原X轴刻度标签
        ax.set_xticklabels([''] * n_types)
        
        # ************** 修改结束 **************
        
        # 设置Y轴范围 - 考虑下移的标签空间
        max_count = max(
            max([soil_data[d].get(s, 0) for s in soil_types])
            for d in soil_data.keys()
        )
        
        # 底部额外空间为最大值的30%
        ax.set_ylim(-max_count * 0.3, max_count * 1.25)
        
        # 添加网格线
        ax.grid(True, axis='y', linestyle='--', alpha=0.5, color='#dddddd', zorder=0)
        
        # 美化坐标轴
        ax.spines[['top', 'right']].set_visible(False)
        ax.spines[['left', 'bottom']].set_color('#cccccc')
        ax.spines['bottom'].set_position(('data', -max_count * 0.3))  # 底部坐标轴下移
        ax.tick_params(axis='y', length=0)
        
        # 创建土壤质地图例
        soil_leg_handles = [
            Patch(facecolor=soil_colors.get(st, '#777777'), edgecolor='#666666', label=st)
            for st in soil_types
        ]
        
        # 在右上角添加图例
        soil_legend = ax.legend(
            handles=soil_leg_handles,
            title='土壤质地类别',
            loc='upper right',
            frameon=True,
            edgecolor='#999999',
            facecolor='white',
            title_fontproperties=FontProperties(weight='bold', size=10),
            fontsize=9
        )
        
        # 在图表右下角添加深度说明
        ax.text(
            0.98, 0.01, 
            f"■ {list(soil_data.keys())[0]}   ■ {list(soil_data.keys())[1]}",
            transform=ax.transAxes,
            ha='right',
            va='bottom',
            fontsize=10,
            bbox=dict(facecolor='white', alpha=0.7, edgecolor='lightgrey', boxstyle='round,pad=0.5')
        )
        
        # 添加数据源信息
        plt.figtext(
            0.5, 0.02,
            f'数据来源: {os.path.basename(file_path)}, 日期: {datetime.now().strftime("%Y-%m-%d")}',
            ha='center', 
            fontsize=9, 
            color='#777777'
        )
        
        # 调整布局 - 增加底部空间
        plt.subplots_adjust(bottom=0.15)  # 增加底部空间容纳标签注释
        
        # 保存图像
        output_path = os.path.join(os.path.expanduser('~'), 'Desktop', '土壤质地深度比较图.png')
        plt.savefig(output_path, dpi=150, bbox_inches='tight')
        print(f"\n图表已保存至: {output_path}")
        
        plt.show()
    
    except Exception as e:
        print(f"处理文件时发生错误: {str(e)}")

# 运行主程序
if __name__ == "__main__":
    main()
