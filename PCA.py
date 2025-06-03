import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.decomposition import PCA
from sklearn.preprocessing import StandardScaler
from matplotlib.font_manager import FontProperties, findfont, FontManager
import os
import matplotlib as mpl
import traceback

# --- 改进的中文字体设置 ---
def set_chinese_font(font_size=12):
    """更可靠的设置中文字体支持"""
    try:
        # 获取系统中字体列表
        fm = FontManager()
        font_names = [f.name for f in fm.ttflist]
        
        # 优选字体列表（按优先级排序）
        pref_fonts = ['Microsoft YaHei', 'SimHei', 'KaiTi', 'STSong', 
                      'NSimSun', 'FangSong', 'SimSun', 'Microsoft JhengHei', 
                      'WenQuanYi Micro Hei', 'Arial Unicode MS']
        
        # 尝试找到第一个可用中文字体
        chosen_font = None
        for font in pref_fonts:
            try:
                # 尝试查找字体，如果找到则使用
                findfont(font)
                chosen_font = font
                print(f"找到可用中文字体: {chosen_font}")
                break
            except:
                continue
        
        if chosen_font:
            # 设置为matplotlib默认字体
            plt.rcParams['font.family'] = 'sans-serif'
            plt.rcParams['font.sans-serif'] = [chosen_font]
            plt.rcParams['axes.unicode_minus'] = False
            
            # 创建字体属性对象供后续显示中文使用
            font_prop = FontProperties(fname=findfont(chosen_font))
        else:
            print("警告: 未找到合适的系统字体，使用matplotlib默认字体")
            font_prop = None
        
        # 设置全局字体大小
        plt.rcParams.update({
            'font.size': font_size,
            'axes.titlesize': font_size,
            'axes.labelsize': font_size,
            'xtick.labelsize': font_size - 2,
            'ytick.labelsize': font_size - 2,
            'legend.fontsize': font_size - 2,
            'figure.titlesize': font_size + 2
        })
        
        return font_prop
        
    except Exception as e:
        print(f"设置中文字体时出错: {e}")
        return None

# 初始化中文字体（获取字体属性对象）
font_prop = set_chinese_font(10)

# === 数据预处理函数 ===
def load_and_preprocess_data(file_path):
    """
    加载和预处理土壤理化性质数据
    参数:
        file_path: Excel文件路径
    返回:
        合并后的DataFrame，包含深度信息
    """
    # 1. 读取Excel文件
    excel_file = pd.ExcelFile(file_path)
    
    # 2. 检查存在的工作表
    print("工作簿中的工作表:", excel_file.sheet_names)
    
    # 3. 识别目标工作表
    target_sheets = []
    for sheet_name in excel_file.sheet_names:
        if '0-10' in sheet_name or '0_10' in sheet_name:
            target_sheets.append(('0-10cm', sheet_name))
        elif ('0-20' in sheet_name or '10-20' in sheet_name or 
              '0_20' in sheet_name or '10_20' in sheet_name):
            target_sheets.append(('10-20cm', sheet_name))
    
    # 如果没找到命名规范的sheet，取前两个sheet
    if not target_sheets:
        print("警告: 未识别到标准命名的工作表，使用前两个工作表")
        target_sheets = [(f"深度{i+1}", name) for i, name in enumerate(excel_file.sheet_names[:2])]
    else:
        print("处理的工作表:", target_sheets)
    
    # 4. 读取并合并数据
    dfs = []
    for depth, sheet_name in target_sheets:
        try:
            df = excel_file.parse(sheet_name)
            print(f"\n工作表 '{sheet_name}' 的列名: {df.columns.tolist()}")
            
            # 检查必要列是否存在
            required_columns = ['容重', '含水率', 'pH', '总氮', '有机质', '总磷', '总钾', '砂粒', '粉粒', '黏粒']
            
            # 尝试匹配中文列名
            col_mapping = {}
            for col in df.columns:
                # 将英文字符替换为UTF-8格式以避免编码问题
                col_clean = str(col).strip().replace('\u200b', '').replace(' ', '')
                
                for req_col in required_columns:
                    if req_col in col_clean:
                        col_mapping[col] = req_col
            
            # 重命名列
            if col_mapping:
                df = df.rename(columns=col_mapping)
                print(f"列重命名映射: {col_mapping}")
            
            # 添加深度信息
            df['深度'] = depth
            
            # 只保留需要的列
            cols_to_keep = required_columns + ['深度']
            available_cols = [col for col in cols_to_keep if col in df.columns]
            
            if len(available_cols) < 5:  # 如果缺少太多列
                print(f"警告: 工作表 '{sheet_name}' 缺少必要的列，跳过")
                continue
                
            # 移除完全空白的行
            df.dropna(how='all', inplace=True)
            
            # 将非数值数据转换为数值
            for col in available_cols:
                if col != '深度':
                    # 处理多种格式的数值（含逗号、百分号等）
                    if df[col].dtype == 'object':
                        df[col] = df[col].astype(str).str.replace(',', '').str.replace('%', '')
                    df[col] = pd.to_numeric(df[col], errors='coerce')
            
            # 添加到列表
            dfs.append(df[available_cols])
            print(f"工作表 '{sheet_name}' 有效样本数: {len(df)}")
            
        except Exception as e:
            print(f"处理工作表 '{sheet_name}' 时出错: {str(e)}")
            continue
    
    # 如果没有成功加载任何数据
    if not dfs:
        raise ValueError("未能加载任何有效数据")
    
    # 5. 合并所有数据
    combined_df = pd.concat(dfs, ignore_index=True)
    
    # 6. 检查缺失值
    print("\n缺失值统计:")
    print(combined_df.isnull().sum())
    
    # 7. 处理缺失值 - 使用列中位数填充
    numeric_cols = combined_df.select_dtypes(include=np.number).columns
    for col in numeric_cols:
        if col != '深度':  # '深度' 是分类列
            median_val = combined_df[col].median()
            combined_df[col].fillna(median_val, inplace=True)
    
    # 8. 查看合并后的数据基本信息
    print(f"\n合并后的数据形状: {combined_df.shape}")
    print(combined_df.info())
    print("\n数据描述性统计:")
    print(combined_df.describe())
    
    return combined_df

# === 主成分分析函数 ===
def perform_pca_and_plot(df, font_prop=None):
    """
    执行主成分分析并创建可视化图表
    参数:
        df: 包含土壤理化性质数据的DataFrame
        font_prop: 字体属性对象(用于中文显示)
    """
    # 1. 数据标准化 - PCA前必须标准化数据
    numeric_cols = ['容重', '含水率', 'pH', '总氮', '有机质', '总磷', '总钾', '砂粒', '粉粒', '黏粒']
    
    # 检查哪些列实际存在
    existing_cols = [col for col in numeric_cols if col in df.columns]
    print(f"用于PCA的变量: {existing_cols}")
    
    # 提取数值数据
    X = df[existing_cols]
    
    # 标准化处理
    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(X)
    
    # 2. 执行主成分分析
    print("\n执行主成分分析(PCA)...")
    pca = PCA()
    principal_components = pca.fit_transform(X_scaled)
    
    # 3. 打印PCA结果
    # 解释方差比
    explained_variance = pca.explained_variance_ratio_
    cumulative_variance = np.cumsum(explained_variance)
    
    print("\n主成分分析结果:")
    print(f"总方差解释比例: {np.sum(explained_variance):.4f}")
    print("\n每个主成分解释的方差比例:")
    for i, var in enumerate(explained_variance):
        print(f"PC{i+1}: {var:.4f} ({var * 100:.1f}%)")
    
    print("\n累计解释方差比例:")
    for i, cum_var in enumerate(cumulative_variance):
        print(f"前{i+1}个主成分: {cum_var:.4f} ({cum_var * 100:.1f}%)")
    
    # 4. 创建DataFrame存储主成分得分和原始数据
    pc_df = pd.DataFrame(
        data=principal_components[:, :3],  # 取前三个主成分
        columns=[f'PC{i+1}' for i in range(3)]
    )
    pc_df['深度'] = df['深度'].values
    
    # 5. 创建可视化图表
    plt.figure(figsize=(16, 10), dpi=100)
    plt.suptitle('草地土壤理化性质主成分分析', fontsize=16, y=0.95, fontproperties=font_prop)
    
    # 使用seaborn设置样式
    sns.set_style("whitegrid")
    sns.set_palette("Set2")
    
    # === 子图1: 主成分的方差解释比例 ===
    ax1 = plt.subplot(2, 2, 1)
    ax1.bar(
        range(1, len(explained_variance) + 1),
        explained_variance,
        alpha=0.7,
        color='skyblue',
        label='各主成分解释方差'
    )
    
    ax1.plot(
        range(1, len(cumulative_variance) + 1),
        cumulative_variance,
        'o-',
        color='orange',
        label='累计解释方差'
    )
    
    ax1.set_xticks(range(1, len(explained_variance) + 1))
    ax1.set_xlabel('主成分编号', fontproperties=font_prop)
    ax1.set_ylabel('解释方差比例', fontproperties=font_prop)
    ax1.set_title('主成分方差解释比例', fontproperties=font_prop)
    ax1.legend(prop=font_prop)
    ax1.grid(True, linestyle='--', alpha=0.6)
    
    # 确保数值标签使用中文字体
    for label in ax1.get_xticklabels() + ax1.get_yticklabels():
        label.set_fontproperties(font_prop)
    
    # 添加百分比标签
    for i, (var, cum_var) in enumerate(zip(explained_variance, cumulative_variance)):
        ax1.text(i+1, var+0.02, f'{var*100:.1f}%', ha='center', fontproperties=font_prop)
        ax1.text(i+1, cum_var+0.02, f'{cum_var*100:.1f}%', ha='center', color='orange', fontproperties=font_prop)
    
    # === 子图2: 前两个主成分的散点图 ===
    ax2 = plt.subplot(2, 2, 2)
    
    # 获取深度类别列表
    depth_categories = df['深度'].unique()
    
    # 创建散点图（显式传递图例信息）
    scatter = sns.scatterplot(
        x='PC1', 
        y='PC2', 
        data=pc_df, 
        hue='深度',
        palette='viridis',
        s=80,
        ax=ax2,
        legend="full"  # 确保图例显示
    )
    
    # 计算数据点分布的边界（动态设置坐标轴范围）
    pc1_min, pc1_max = pc_df['PC1'].min(), pc_df['PC1'].max()
    pc2_min, pc2_max = pc_df['PC2'].min(), pc_df['PC2'].max()
    
    pc1_range = pc1_max - pc1_min
    pc2_range = pc2_max - pc2_min
    
    ax2.set_xlim(pc1_min - pc1_range*0.1, pc1_max + pc1_range*0.1)
    ax2.set_ylim(pc2_min - pc2_range*0.1, pc2_max + pc2_range*0.1)
    
    ax2.set_xlabel(f'PC1 (解释方差: {explained_variance[0]*100:.1f}%)', fontproperties=font_prop)
    ax2.set_ylabel(f'PC2 (解释方差: {explained_variance[1]*100:.1f}%)', fontproperties=font_prop)
    ax2.set_title('PC1 vs PC2 - 按深度着色', fontproperties=font_prop)
    ax2.grid(True, linestyle='--', alpha=0.6)
    
    # 显式配置图例
    handles, labels = ax2.get_legend_handles_labels()
    
    # 创建新的图例对象（确保中文显示）
    new_legend = ax2.legend(
        handles=handles, 
        labels=[f"深度: {label}" for label in labels],  # 添加说明文本
        title='土壤深度', 
        title_fontproperties=font_prop,
        prop=font_prop,
        loc='best',  # 自动寻找最佳位置
        frameon=True,
        framealpha=0.8
    )
    
    # 确保图例的文本字体正确
    for text in new_legend.get_texts():
        text.set_fontproperties(font_prop)
    
    # 设置刻度标签的字体
    for label in ax2.get_xticklabels() + ax2.get_yticklabels():
        label.set_fontproperties(font_prop)
    
    # === 子图3: 变量在主成分上的载荷 ===
    # 计算载荷矩阵
    loadings = pca.components_.T * np.sqrt(pca.explained_variance_)
    
    ax3 = plt.subplot(2, 2, 3)
    
    # 绘制所有变量的载荷方向线
    for i, var in enumerate(existing_cols):
        ax3.arrow(
            0, 0,  # 原点
            loadings[i, 0],  # PC1方向
            loadings[i, 1],  # PC2方向
            head_width=0.05,
            head_length=0.1,
            fc='gray',
            ec='gray',
            alpha=0.5
        )
        
        # 在箭头末端添加变量标签（使用中文字体）
        ax3.text(
            loadings[i, 0] * 1.15,
            loadings[i, 1] * 1.15,
            var,
            color='black',
            ha='center',
            va='center',
            fontsize=10,
            bbox=dict(facecolor='white', alpha=0.7, edgecolor='none', boxstyle='round,pad=0.2'),
            fontproperties=font_prop
        )
    
    # 设置坐标轴范围
    max_loading = np.max(np.abs(loadings[:, :2])) * 1.3
    ax3.set_xlim(-max_loading, max_loading)
    ax3.set_ylim(-max_loading, max_loading)
    
    # 添加参考线
    ax3.axhline(0, color='gray', linestyle='--', alpha=0.5)
    ax3.axvline(0, color='gray', linestyle='--', alpha=0.5)
    
    # 添加圆和椭圆
    circle = plt.Circle((0, 0), 1, fill=False, color='blue', alpha=0.3)
    ax3.add_patch(circle)
    
    # 标签和标题（使用中文字体）
    ax3.set_xlabel('PC1载荷', fontproperties=font_prop)
    ax3.set_ylabel('PC2载荷', fontproperties=font_prop)
    ax3.set_title('变量载荷图', fontproperties=font_prop)
    
    # 设置刻度标签的字体
    for label in ax3.get_xticklabels() + ax3.get_yticklabels():
        label.set_fontproperties(font_prop)
    
    ax3.grid(True, linestyle='--', alpha=0.6)
    
    # 添加变量贡献信息（使用中文字体）
    ax3.text(
        0.98, 0.02,
        '箭头方向表示变量与主成分的关系\n箭头长度表示贡献强度',
        transform=ax3.transAxes,
        ha='right',
        va='bottom',
        fontsize=9,
        bbox=dict(facecolor='white', alpha=0.7, edgecolor='gray'),
        fontproperties=font_prop
    )
    
    # === 子图4: 所有变量在样本中的分布 ===
    ax4 = plt.subplot(2, 2, 4)
    
    # 选择最重要的几个变量展示
    n_show = min(5, len(existing_cols))  # 最多显示5个变量
    # 按在PC1上的载荷绝对值排序
    top_vars_idx = np.abs(loadings[:, 0]).argsort()[-n_show:]
    top_vars = [existing_cols[i] for i in top_vars_idx]
    
    # 使用箱线图+散点图展示这些变量在不同深度的分布
    melted_df = pd.melt(
        df[top_vars + ['深度']],
        id_vars=['深度'],
        var_name='变量',
        value_name='值'
    )
    
    # 创建箱线图+散点图
    bp = sns.boxplot(
        x='变量', 
        y='值', 
        data=melted_df, 
        hue='深度', 
        showmeans=True,  # 显示均值
        meanprops={'marker':'D', 'markerfacecolor':'white', 'markeredgecolor':'black', 'markersize': 6},
        ax=ax4
    )
    
    # 添加均值标签
    for i, variable in enumerate(top_vars):
        for j, depth in enumerate(pc_df['深度'].unique()):
            mean_val = melted_df[(melted_df['变量'] == variable) & 
                                (melted_df['深度'] == depth)]['值'].mean()
            ax4.text(i - 0.2 + j*0.4, mean_val + 0.02, f'{mean_val:.2f}', 
                    ha='center', va='bottom', fontsize=8, fontproperties=font_prop)
    
    # 调整图形
    ax4.set_xticks(range(len(top_vars)))
    ax4.set_xticklabels(top_vars, rotation=0, fontproperties=font_prop)
    ax4.set_xlabel('', fontproperties=font_prop)
    ax4.set_ylabel('值', fontproperties=font_prop)
    ax4.set_title(f'关键变量在不同深度的分布', fontproperties=font_prop)
    
    # 设置y轴刻度标签和图例的字体
    for label in ax4.get_yticklabels():
        label.set_fontproperties(font_prop)
    
    leg = ax4.legend(prop=font_prop)
    for text in leg.get_texts():
        text.set_fontproperties(font_prop)
    
    ax4.grid(True, linestyle='--', alpha=0.6)
    
    # === 调整布局并保存 ===
    plt.tight_layout(rect=[0, 0, 1, 0.95])
    
    # 添加整体标题（使用中文字体）
    plt.figtext(0.5, 0.98, '草地土壤理化性质主成分分析', 
               ha='center', va='top', fontsize=16, weight='bold', fontproperties=font_prop)
    
    # 添加脚注（使用中文字体）
    #plt.figtext(0.5, 0.03, f'数据来源: {os.path.basename(file_path)} | 分析日期: {pd.Timestamp.now().date()}',
               #ha='center', fontsize=9, fontproperties=font_prop)
    
    # 保存图像（确保使用正确的背景色和DPI）
    output_dir = os.path.join(os.path.expanduser('~'), 'Desktop')
    output_path = os.path.join(output_dir, '土壤理化性质_PCA分析图.png')
    plt.savefig(output_path, dpi=150, bbox_inches='tight', facecolor='white')
    print(f"\n图表已保存至: {output_path}")
    
    plt.show()
    
    return pc_df, explained_variance, loadings

# === 主程序 ===
if __name__ == "__main__":
    # === 文件路径设置 ===
    # 请修改为你的实际文件路径
    file_path = r'D:\1793672313\Desktop\下载 (1)\河湖源草地土壤理化性质.xlsx'
    
    if not os.path.exists(file_path):
        print(f"错误: 文件不存在 - {file_path}")
    else:
        print(f"处理文件: {file_path}")
        
        try:
            # === 加载和预处理数据 ===
            df = load_and_preprocess_data(file_path)
            
            # === 执行PCA分析和可视化 ===
            # 传入字体属性对象确保中文正常显示
            pc_df, explained_variance, loadings = perform_pca_and_plot(df, font_prop)
            
            # === 显示主要结果 ===
            print("\n=== PCA分析主要结论 ===")
            print(f"1. 前两个主成分(PC1 + PC2)解释了{sum(explained_variance[:2])*100:.1f}%的总变异")
            print(f"2. 第一主成分(PC1)解释{explained_variance[0]*100:.1f}%的变异，主要受以下变量影响:")
            # 找到对PC1影响最大的3个变量
            pc1_loadings = loadings[:, 0]
            top_idx = np.abs(pc1_loadings).argsort()[::-1]  # 降序排列
            
            # 获取实际的列名
            col_names = df.columns
            
            print(f"   - 正相关: {col_names[top_idx[0]]} (载荷: {pc1_loadings[top_idx[0]]:.2f})")
            print(f"   - 负相关: {col_names[top_idx[-1]]} (载荷: {pc1_loadings[top_idx[-1]]:.2f})")
            
            print(f"3. 不同深度的土壤在PCA图上的分布差异:")
            # 分组计算中心点差异
            grouped = pc_df.groupby('深度').agg({'PC1': 'mean', 'PC2': 'mean'})
            print(grouped)
            
            if grouped.shape[0] > 1:
                # 计算组间距离
                diff = grouped.iloc[0] - grouped.iloc[1]
                print(f"   0-10cm和10-20cm组的PC1平均分差异: {diff['PC1']:.2f}")
                print(f"   0-10cm和10-20cm组的PC2平均分差异: {diff['PC2']:.2f}")
                
                # 解读差异
                if abs(diff['PC1']) > abs(diff['PC2']):
                    print("   深度差异主要反映在PC1方向上")
                else:
                    print("   深度差异主要反映在PC2方向上")
            
            # === 保存结果到Excel ===
            # 创建包含所有结果的DataFrame
            full_df = df.copy()
            full_df[['PC1', 'PC2', 'PC3']] = pc_df[['PC1', 'PC2', 'PC3']]
            
            # 保存到桌面
            output_excel = os.path.join(os.path.expanduser('~'), 'Desktop', '土壤理化性质_PCA分析结果.xlsx')
            full_df.to_excel(output_excel, index=False)
            print(f"分析结果已保存至Excel文件: {output_excel}")
            
        except Exception as e:
            print(f"处理过程中出错: {str(e)}")
            traceback.print_exc()
