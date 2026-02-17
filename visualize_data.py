import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
import platform

def set_chinese_font():
    system_name = platform.system()
    if system_name == 'Windows':
        # Common Chinese fonts on Windows
        fonts = ['SimHei', 'Microsoft YaHei', 'SimSun']
    elif system_name == 'Darwin': # macOS
        fonts = ['Arial Unicode MS', 'PingFang SC', 'Heiti SC']
    else: # Linux
        fonts = ['WenQuanYi Micro Hei', 'Droid Sans Fallback']
        
    for font in fonts:
        try:
            plt.rcParams['font.sans-serif'] = [font]
            # Verify if font is actually available by trying to render a plot
            # But just setting it is usually enough if it exists.
            # Let's also set minus sign to False to handle negative numbers correctly
            plt.rcParams['axes.unicode_minus'] = False
            break
        except Exception:
            continue

def load_data():
    filename = 'analysis_result.xlsx'
    print(f"Loading data from {filename}...")
    try:
        xls = pd.ExcelFile(filename)
        df_students = pd.read_excel(filename, sheet_name='Student_Comparison')
        df_class = pd.read_excel(filename, sheet_name='Class_Summary')
        df_subject = pd.read_excel(filename, sheet_name='Subject_Summary')
        return df_students, df_class, df_subject
    except Exception as e:
        print(f"Error loading data: {e}")
        return None, None, None

def plot_total_score_distribution(df, output_dir):
    plt.figure(figsize=(10, 6))
    sns.histplot(df['Total_Score_Monthly'], color='skyblue', label='第一次月考', kde=True, alpha=0.5)
    sns.histplot(df['Total_Score_Midterm'], color='orange', label='期中考试', kde=True, alpha=0.5)
    plt.title('第一次月考 vs 期中考试 总分分布对比')
    plt.xlabel('总分')
    plt.ylabel('人数')
    plt.legend()
    plt.savefig(os.path.join(output_dir, 'total_score_distribution.png'))
    plt.close()
    print("Saved total_score_distribution.png")

def plot_subject_comparison(df_subject, output_dir):
    # Bar chart for subject averages
    plt.figure(figsize=(12, 6))
    
    # Melt the dataframe for seaborn
    df_melted = df_subject.melt(id_vars=['Subject'], value_vars=['Avg_Score_Monthly', 'Avg_Score_Midterm'], 
                                var_name='Exam', value_name='Average Score')
    
    # Rename for legend
    df_melted['Exam'] = df_melted['Exam'].replace({
        'Avg_Score_Monthly': '第一次月考',
        'Avg_Score_Midterm': '期中考试'
    })
    
    sns.barplot(data=df_melted, x='Subject', y='Average Score', hue='Exam')
    plt.title('各学科平均分对比')
    plt.ylabel('平均分')
    plt.xlabel('学科')
    plt.savefig(os.path.join(output_dir, 'subject_comparison_bar.png'))
    plt.close()
    print("Saved subject_comparison_bar.png")
    
    # Plot Delta
    plt.figure(figsize=(10, 5))
    sns.barplot(data=df_subject, x='Subject', y='Delta', palette='vlag')
    plt.axhline(0, color='black', linewidth=0.8)
    plt.title('各学科平均分变化 (期中 - 月考)')
    plt.ylabel('分数变化')
    plt.xlabel('学科')
    plt.savefig(os.path.join(output_dir, 'subject_delta.png'))
    plt.close()
    print("Saved subject_delta.png")

def plot_class_performance(df_students, output_dir):
    if 'Class_Midterm' not in df_students.columns:
        print("Class column not found, skipping class plots.")
        return

    # Boxplot of Total Scores by Class
    plt.figure(figsize=(14, 7))
    # Sort classes if possible
    sorted_classes = sorted(df_students['Class_Midterm'].unique())
    sns.boxplot(data=df_students, x='Class_Midterm', y='Total_Score_Midterm', order=sorted_classes, palette="Set3")
    plt.title('期中考试各班级总分分布 (箱线图)')
    plt.xlabel('班级')
    plt.ylabel('总分')
    plt.savefig(os.path.join(output_dir, 'class_score_boxplot.png'))
    plt.close()
    print("Saved class_score_boxplot.png")
    
    # Class Average Score Comparison (Monthly vs Midterm)
    # We can use the Class Summary df for this, but boxplot needs raw data
    
def plot_rank_changes(df_students, output_dir):
    if 'Total_School_Rank_Monthly' not in df_students.columns or 'Improvement_School_Rank' not in df_students.columns:
        print("Rank columns missing, skipping rank plot.")
        return
        
    plt.figure(figsize=(10, 8))
    sns.scatterplot(data=df_students, x='Total_School_Rank_Monthly', y='Improvement_School_Rank', alpha=0.6)
    plt.axhline(0, color='red', linestyle='--', linewidth=1)
    plt.title('月考排名 vs 排名变化量')
    plt.xlabel('第一次月考排名 (越小越好)')
    plt.ylabel('排名变化 (正数表示进步)')
    
    # Add annotations for top improvers (optional)
    # top_improvers = df_students.nlargest(5, 'Improvement_School_Rank')
    # for i, row in top_improvers.iterrows():
    #     plt.text(row['Total_School_Rank_Monthly'], row['Improvement_School_Rank'], 
    #              str(row['Name_Midterm']), fontsize=9)
                 
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
