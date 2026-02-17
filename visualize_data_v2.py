import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
import platform
import matplotlib.font_manager as fm

def set_chinese_font():
    # Try to find a Chinese font
    font_path = None
    
    # Common Windows Chinese fonts
    windows_fonts = [
        r'C:\Windows\Fonts\simhei.ttf',
        r'C:\Windows\Fonts\msyh.ttc',
        r'C:\Windows\Fonts\simsun.ttc'
    ]
    
    for f in windows_fonts:
        if os.path.exists(f):
            font_path = f
            break
            
    if font_path:
        prop = fm.FontProperties(fname=font_path)
        plt.rcParams['font.family'] = prop.get_name()
    else:
        # Fallback
        plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'SimSun', 'Arial Unicode MS']
    
    plt.rcParams['axes.unicode_minus'] = False # Solve negative sign display issue

def load_data():
    filename = 'analysis_result.xlsx'
    print(f"Loading data from {filename}...")
    try:
        # Read all sheets
        df_students = pd.read_excel(filename, sheet_name='Student_Comparison')
        df_class = pd.read_excel(filename, sheet_name='Class_Summary')
        df_subject = pd.read_excel(filename, sheet_name='Subject_Summary')
        return df_students, df_class, df_subject
    except Exception as e:
        print(f"Error loading data: {e}")
        return None, None, None

def plot_total_score_distribution(df, output_dir):
    plt.figure(figsize=(10, 6))
    
    # Drop NaN
    data_monthly = df['Total_Score_Monthly'].dropna()
    data_midterm = df['Total_Score_Midterm'].dropna()
    
    sns.histplot(data_monthly, color='skyblue', label='第一次月考', kde=True, alpha=0.5)
    sns.histplot(data_midterm, color='orange', label='期中考试', kde=True, alpha=0.5)
    
    plt.title('第一次月考 vs 期中考试 总分分布对比')
    plt.xlabel('总分')
    plt.ylabel('人数')
    plt.legend()
    plt.grid(axis='y', alpha=0.3)
    
    plt.savefig(os.path.join(output_dir, 'total_score_distribution.png'))
    plt.close()
    print("Saved total_score_distribution.png")

def plot_subject_comparison(df_subject, output_dir):
    if df_subject is None or df_subject.empty:
        return

    # Bar chart for subject averages
    plt.figure(figsize=(12, 6))
    
    # Melt
    df_melted = df_subject.melt(id_vars=['Subject'], value_vars=['Avg_Score_Monthly', 'Avg_Score_Midterm'], 
                                var_name='Exam', value_name='Average Score')
    
    # Rename for legend
    df_melted['Exam'] = df_melted['Exam'].replace({
        'Avg_Score_Monthly': '第一次月考',
        'Avg_Score_Midterm': '期中考试'
    })
    
    # Enforce order: Monthly first, then Midterm
    sns.barplot(data=df_melted, x='Subject', y='Average Score', hue='Exam', 
                hue_order=['第一次月考', '期中考试'], palette=['skyblue', 'orange'])
                
    plt.title('各学科平均分对比 (月考 vs 期中)')
    plt.ylabel('平均分')
    plt.xlabel('学科')
    plt.grid(axis='y', alpha=0.3)
    plt.savefig(os.path.join(output_dir, 'subject_comparison_bar.png'))
    plt.close()
    print("Saved subject_comparison_bar.png")
    
    # Plot Delta
    plt.figure(figsize=(10, 5))
    # Use hue=Subject to avoid warning, set legend=False
    sns.barplot(data=df_subject, x='Subject', y='Delta', hue='Subject', legend=False, palette='vlag')
    plt.axhline(0, color='black', linewidth=0.8)
    plt.title('学科分数变化 (期中 - 月考)')
    plt.ylabel('分数变化 (正=进步, 负=退步)')
    plt.xlabel('学科')
    plt.grid(axis='y', alpha=0.3)
    plt.savefig(os.path.join(output_dir, 'subject_delta.png'))
    plt.close()
    print("Saved subject_delta.png")

def plot_class_performance(df_students, output_dir):
    if df_students is None or 'Class_Midterm' not in df_students.columns:
        print("Class column not found, skipping class plots.")
        return

    # Boxplot
    plt.figure(figsize=(14, 7))
    sorted_classes = sorted(df_students['Class_Midterm'].dropna().unique())
    
    sns.boxplot(data=df_students, x='Class_Midterm', y='Total_Score_Midterm', order=sorted_classes, hue='Class_Midterm', legend=False, palette="Set3")
    plt.title('期中考试各班级总分分布')
    plt.xlabel('班级')
    plt.ylabel('总分')
    plt.grid(axis='y', alpha=0.3)
    plt.savefig(os.path.join(output_dir, 'class_score_boxplot.png'))
    plt.close()
    print("Saved class_score_boxplot.png")

def plot_rank_changes(df_students, output_dir):
    if df_students is None: return
    
    if 'Total_School_Rank_Monthly' not in df_students.columns or 'Improvement_School_Rank' not in df_students.columns:
        print("Rank columns missing, skipping rank plot.")
        return
        
    plt.figure(figsize=(10, 8))
    sns.scatterplot(data=df_students, x='Total_School_Rank_Monthly', y='Improvement_School_Rank', alpha=0.6)
    plt.axhline(0, color='red', linestyle='--', linewidth=1)
    plt.title('排名变化分析')
    plt.xlabel('第一次月考排名 (数值越小越靠前)')
    plt.ylabel('排名进步量 (正数=进步, 负数=退步)')
    plt.grid(True, alpha=0.3)
                 
    plt.savefig(os.path.join(output_dir, 'rank_change_scatter.png'))
    plt.close()
    print("Saved rank_change_scatter.png")

def main():
    output_dir = 'charts'
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    set_chinese_font()
    
    df_students, df_class, df_subject = load_data()
    
    if df_students is not None:
        plot_total_score_distribution(df_students, output_dir)
        plot_class_performance(df_students, output_dir)
        plot_rank_changes(df_students, output_dir)
        
    if df_subject is not None:
        plot_subject_comparison(df_subject, output_dir)
        
    print("Visualization complete!")

if __name__ == "__main__":
    main()
