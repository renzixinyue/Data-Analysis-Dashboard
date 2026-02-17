import pandas as pd
import numpy as np

def load_data(filepath, exam_suffix):
    print(f"Loading {filepath}...")
    try:
        df = pd.read_excel(filepath, header=[1, 2])
    except Exception as e:
        print(f"Error reading {filepath}: {e}")
        return None
    
    # Flatten columns
    new_columns = []
    last_subject = None
    
    for i in range(len(df.columns)):
        col = df.columns[i]
        subject = str(col[0]).strip()
        metric = str(col[1]).strip()
        
        if "Unnamed" in subject or subject == "nan":
            subject = last_subject
        else:
            last_subject = subject
            
        if "Unnamed" in metric or metric == "nan":
            metric = ""
            
        if subject and metric:
            new_columns.append(f"{subject}_{metric}")
        elif subject:
            new_columns.append(subject)
        else:
            new_columns.append(metric) 
            
    df.columns = new_columns
    
    # Clean up column names and map to standard names
    col_map = {}
    
    for col in df.columns:
        new_name = col
        if "姓名" in col: new_name = "Name"
        elif "学号" in col: new_name = "StudentID"
        elif "考号" in col: new_name = "ExamID"
        elif "班级" in col and "排名" not in col: new_name = "Class"
        elif "总分" in col and "排名" not in col: new_name = "Total_Score"
        elif "总分" in col and "联考排名" in col: new_name = "Total_Joint_Rank"
        elif "总分" in col and "学校排名" in col: new_name = "Total_School_Rank"
        elif "总分" in col and "班级排名" in col: new_name = "Total_Class_Rank"
        
        col_map[col] = new_name
    
    df = df.rename(columns=col_map)
    
    # Remove rows where Name or StudentID is missing
    df = df.dropna(subset=['Name', 'StudentID'])
    
    # Convert numeric columns
    for col in df.columns:
        if "Score" in col or "Rank" in col or "分数" in col or "排名" in col:
            df[col] = pd.to_numeric(df[col], errors='coerce')
            
    # Add suffix to all columns except join keys (StudentID)
    # We WILL suffix Name and Class to distinguish them
    cols_to_rename = {col: f"{col}_{exam_suffix}" for col in df.columns if col != 'StudentID'}
    df = df.rename(columns=cols_to_rename)
    
    return df

def analyze():
    # Load data
    df_monthly = load_data('diyiciyuekao.xlsx', 'Monthly')
    df_midterm = load_data('qizhognchengji.xlsx', 'Midterm')
    
    if df_monthly is None or df_midterm is None:
        return

    # Merge data on StudentID
    print("Merging data...")
    merged_df = pd.merge(df_monthly, df_midterm, on='StudentID', how='inner')
    
    # Calculate Deltas
    # Total Score
    if 'Total_Score_Midterm' in merged_df.columns and 'Total_Score_Monthly' in merged_df.columns:
        merged_df['Delta_Total_Score'] = merged_df['Total_Score_Midterm'] - merged_df['Total_Score_Monthly']
    
    # School Rank Improvement (Monthly - Midterm)
    if 'Total_School_Rank_Monthly' in merged_df.columns and 'Total_School_Rank_Midterm' in merged_df.columns:
        merged_df['Improvement_School_Rank'] = merged_df['Total_School_Rank_Monthly'] - merged_df['Total_School_Rank_Midterm']
    
    if 'Total_Class_Rank_Monthly' in merged_df.columns and 'Total_Class_Rank_Midterm' in merged_df.columns:
        merged_df['Improvement_Class_Rank'] = merged_df['Total_Class_Rank_Monthly'] - merged_df['Total_Class_Rank_Midterm']
    
    # Subject Analysis
    subjects = ['语文', '数学', '英语', '生物', '道德与法治', '历史', '地理']
    
    for sub in subjects:
        # Note: In load_data, we kept original subject names like "语文_分数"
        # So they became "语文_分数_Monthly"
        score_col_monthly = f"{sub}_分数_Monthly"
        score_col_midterm = f"{sub}_分数_Midterm"
        
        if score_col_monthly in merged_df.columns and score_col_midterm in merged_df.columns:
            merged_df[f'Delta_{sub}'] = merged_df[score_col_midterm] - merged_df[score_col_monthly]

    # Reorder columns
    # We want Name_Midterm to be the main Name column
    basic_cols = ['StudentID', 'Name_Midterm', 'Class_Midterm', 'Name_Monthly', 'Class_Monthly']
    score_cols = ['Total_Score_Monthly', 'Total_Score_Midterm', 'Delta_Total_Score', 
                  'Total_School_Rank_Monthly', 'Total_School_Rank_Midterm', 'Improvement_School_Rank',
                  'Total_Class_Rank_Monthly', 'Total_Class_Rank_Midterm', 'Improvement_Class_Rank']
    
    # Filter valid columns
    final_cols = [c for c in basic_cols if c in merged_df.columns] + \
                 [c for c in score_cols if c in merged_df.columns] + \
                 [c for c in merged_df.columns if c not in basic_cols and c not in score_cols]
                 
    merged_df = merged_df[final_cols]
    
    # Class Level Analysis
    print("Performing Class Analysis...")
    if 'Class_Midterm' in merged_df.columns:
        class_group = merged_df.groupby('Class_Midterm')
        
        agg_dict = {}
        if 'Total_Score_Monthly' in merged_df.columns: agg_dict['Total_Score_Monthly'] = 'mean'
        if 'Total_Score_Midterm' in merged_df.columns: agg_dict['Total_Score_Midterm'] = 'mean'
        if 'Delta_Total_Score' in merged_df.columns: agg_dict['Delta_Total_Score'] = 'mean'
        if 'Improvement_School_Rank' in merged_df.columns: agg_dict['Improvement_School_Rank'] = 'mean'
        
        class_summary = class_group.agg(agg_dict).reset_index()
        
        rename_dict = {
            'Total_Score_Monthly': 'Avg_Score_Monthly',
            'Total_Score_Midterm': 'Avg_Score_Midterm',
            'Delta_Total_Score': 'Avg_Score_Change',
            'Improvement_School_Rank': 'Avg_Rank_Improvement'
        }
        class_summary = class_summary.rename(columns=rename_dict)
    else:
        class_summary = pd.DataFrame()

    # Subject Level Analysis (Global)
    print("Performing Subject Analysis...")
    subject_summary_data = []
    for sub in subjects:
        score_col_monthly = f"{sub}_分数_Monthly"
        score_col_midterm = f"{sub}_分数_Midterm"
        if score_col_monthly in merged_df.columns and score_col_midterm in merged_df.columns:
            mean_monthly = merged_df[score_col_monthly].mean()
            mean_midterm = merged_df[score_col_midterm].mean()
            subject_summary_data.append({
                'Subject': sub,
                'Avg_Score_Monthly': mean_monthly,
                'Avg_Score_Midterm': mean_midterm,
                'Delta': mean_midterm - mean_monthly
            })
    subject_summary = pd.DataFrame(subject_summary_data)

    # Write to Excel
    output_file = 'analysis_result.xlsx'
    print(f"Writing results to {output_file}...")
    with pd.ExcelWriter(output_file) as writer:
        merged_df.to_excel(writer, sheet_name='Student_Comparison', index=False)
        class_summary.to_excel(writer, sheet_name='Class_Summary', index=False)
        subject_summary.to_excel(writer, sheet_name='Subject_Summary', index=False)
        
    print("Analysis complete!")

if __name__ == "__main__":
    analyze()
