import pandas as pd
import numpy as np

def clean_dataframe(df, exam_name):
    # Handle multi-level columns
    # The header is complex. Let's inspect the first few rows to understand better.
    # Based on previous inspection:
    # Row 0: Title (skip)
    # Row 1: Subject names (e.g., "语文", "数学")
    # Row 2: Metrics (e.g., "分数", "联考排名", "学校排名", "班级排名")
    
    # Forward fill the top level column names (Subjects)
    new_columns = []
    current_subject = None
    
    # Iterate through the columns. The dataframe should have MultiIndex columns if read with header=[0, 1]
    # Let's adjust how we read it.
    pass

def load_data(filepath):
    print(f"Loading {filepath}...")
    # Read with header=[1, 2] to capture the subject and metric rows
    # Skip the very first row (title)
    df = pd.read_excel(filepath, header=[1, 2])
    
    # Flatten columns
    new_cols = []
    for col in df.columns:
        subject = col[0]
        metric = col[1]
        
        if "Unnamed" in str(subject):
            subject = ""
        if "Unnamed" in str(metric):
            metric = ""
            
        # Handling the specific structure observed
        # If subject is valid and metric is unnamed, it might be a single column subject like "姓名"
        if subject and not metric:
            new_cols.append(subject)
        elif not subject and metric:
            new_cols.append(metric) # Should not happen based on structure
        elif subject and metric:
            new_cols.append(f"{subject}_{metric}")
        else:
            new_cols.append(f"Unknown_{len(new_cols)}")
            
    # The above logic is a bit simplistic for merged cells.
    # Let's look at the raw data again.
    # The "Unnamed" usually appears when a cell is merged horizontally.
    # Pandas read_excel with header=[1,2] produces a MultiIndex.
    # Level 0 has "语文", "nan", "nan", "nan"...
    # Level 1 has "分数", "联考排名", "学校排名", "班级排名"...
    
    # Let's try to fix the MultiIndex columns
    # Forward fill the level 0 (Subjects)
    levels = df.columns.levels
    codes = df.columns.codes
    
    # We can reconstruct the columns
    new_columns = []
    last_subject = None
    
    for i in range(len(df.columns)):
        col = df.columns[i]
        subject = col[0]
        metric = col[1]
        
        if "Unnamed" in str(subject):
            subject = last_subject
        else:
            last_subject = subject
            
        if "Unnamed" in str(metric):
            metric = ""
            
        if subject and metric:
            new_columns.append(f"{subject}_{metric}")
        elif subject:
            new_columns.append(subject)
        else:
            new_columns.append(metric) # Fallback
            
    df.columns = new_columns
    
    # Clean up column names
    # "姓名" might be "姓名_nan" or just "姓名" depending on how it was read
    # Let's standardize essential columns
    # Map typical columns
    col_map = {}
    for col in df.columns:
        if "姓名" in col: col_map[col] = "Name"
        elif "学号" in col: col_map[col] = "StudentID"
        elif "考号" in col: col_map[col] = "ExamID"
        elif "班级" in col and "排名" not in col: col_map[col] = "Class"
        elif "总分" in col and "排名" not in col: col_map[col] = "Total_Score"
        elif "总分" in col and "排名" in col: 
             if "班级" in col: col_map[col] = "Total_Class_Rank"
             elif "学校" in col: col_map[col] = "Total_School_Rank"
             elif "联考" in col: col_map[col] = "Total_Joint_Rank"
    
    df = df.rename(columns=col_map)
    
    # Drop rows that are clearly not data (e.g. repeated headers or empty rows)
    df = df.dropna(subset=['Name', 'StudentID'])
    
    return df

if __name__ == "__main__":
    try:
        df1 = load_data('diyiciyuekao.xlsx')
        print("Loaded First Monthly Exam Data:")
        print(df1.head())
        print(df1.columns.tolist())
        
        df2 = load_data('qizhognchengji.xlsx')
        print("\nLoaded Midterm Exam Data:")
        print(df2.head())
        print(df2.columns.tolist())
        
    except Exception as e:
        print(f"An error occurred: {e}")
