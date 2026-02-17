import pandas as pd
import json
import numpy as np

def load_data():
    filename = 'analysis_result.xlsx'
    print(f"Loading data from {filename}...")
    try:
        # Load all sheets
        df_students = pd.read_excel(filename, sheet_name='Student_Comparison')
        df_class = pd.read_excel(filename, sheet_name='Class_Summary')
        df_subject = pd.read_excel(filename, sheet_name='Subject_Summary')
        return df_students, df_class, df_subject
    except Exception as e:
        print(f"Error loading data: {e}")
        return None, None, None

def convert_to_json():
    df_students, df_class, df_subject = load_data()
    
    if df_students is None:
        return

    # 1. Global Stats
    global_stats = {
        'total_students': int(len(df_students)),
        'avg_score_monthly': float(df_students['Total_Score_Monthly'].mean()),
        'avg_score_midterm': float(df_students['Total_Score_Midterm'].mean()),
        'score_distribution': {
            'monthly': df_students['Total_Score_Monthly'].dropna().tolist(),
            'midterm': df_students['Total_Score_Midterm'].dropna().tolist()
        }
    }

    # 2. Subject Stats
    subject_stats = df_subject.to_dict(orient='records')

    # 3. Class Stats
    # Sort by class name
    if 'Class_Midterm' in df_class.columns:
        df_class = df_class.sort_values('Class_Midterm')
    class_stats = df_class.to_dict(orient='records')
    
    # 4. Top/Bottom Improvers
    # Top 5 Improvement in School Rank (Positive is better in our calc? 
    # Wait, in analyze_data_full.py:
    # merged_df['Improvement_School_Rank'] = merged_df['Total_School_Rank_Monthly'] - merged_df['Total_School_Rank_Midterm']
    # If Monthly Rank 10, Midterm Rank 5. 10-5 = 5. Positive = Improvement.
    
    top_improvers = df_students.nlargest(5, 'Improvement_School_Rank')[['Name_Midterm', 'Class_Midterm', 'Improvement_School_Rank']].to_dict(orient='records')
    
    # Bottom 5 (Smallest improvement, i.e., largest negative number)
    bottom_improvers = df_students.nsmallest(5, 'Improvement_School_Rank')[['Name_Midterm', 'Class_Midterm', 'Improvement_School_Rank']].to_dict(orient='records')

    # 5. Students Data (Optimized for search)
    # We need a list of students with their details.
    # Structure: { name: "Name", class: "Class", scores: {...}, ranks: {...} }
    
    students_list = []
    
    subjects = ['语文', '数学', '英语', '生物', '道德与法治', '历史', '地理']
    
    for _, row in df_students.iterrows():
        student_data = {
            'name': row['Name_Midterm'],
            'student_id': row['StudentID'],
            'class': row['Class_Midterm'],
            'total_score_monthly': row['Total_Score_Monthly'] if pd.notna(row['Total_Score_Monthly']) else 0,
            'total_score_midterm': row['Total_Score_Midterm'] if pd.notna(row['Total_Score_Midterm']) else 0,
            'total_rank_monthly': row['Total_Joint_Rank_Monthly'] if pd.notna(row['Total_Joint_Rank_Monthly']) else 0, # Use Joint Rank for wider comparison? Or School Rank? User asked for "联考排名"
            'total_rank_midterm': row['Total_Joint_Rank_Midterm'] if pd.notna(row['Total_Joint_Rank_Midterm']) else 0,
            'rank_change': row['Total_Joint_Rank_Monthly'] - row['Total_Joint_Rank_Midterm'] if pd.notna(row['Total_Joint_Rank_Monthly']) and pd.notna(row['Total_Joint_Rank_Midterm']) else 0,
            'subjects': []
        }
        
        for sub in subjects:
            # Need to find correct columns. In analysis_result.xlsx, they are like "语文_联考排名_Monthly"
            rank_monthly = row.get(f'{sub}_联考排名_Monthly', 0)
            rank_midterm = row.get(f'{sub}_联考排名_Midterm', 0)
            
            # Handle NaN
            rank_monthly = int(rank_monthly) if pd.notna(rank_monthly) else 0
            rank_midterm = int(rank_midterm) if pd.notna(rank_midterm) else 0
            
            change = 0
            if rank_monthly > 0 and rank_midterm > 0:
                change = rank_monthly - rank_midterm
            
            student_data['subjects'].append({
                'name': sub,
                'rank_monthly': rank_monthly,
                'rank_midterm': rank_midterm,
                'change': change
            })
            
        students_list.append(student_data)

    final_data = {
        'global_stats': global_stats,
        'subject_stats': subject_stats,
        'class_stats': class_stats,
        'top_improvers': top_improvers,
        'bottom_improvers': bottom_improvers,
        'students': students_list
    }
    
    # Custom JSON encoder for numpy types
    class NpEncoder(json.JSONEncoder):
        def default(self, obj):
            if isinstance(obj, np.integer):
                return int(obj)
            if isinstance(obj, np.floating):
                return float(obj)
            if isinstance(obj, np.ndarray):
                return obj.tolist()
            return super(NpEncoder, self).default(obj)

    output_path = 'dashboard/data.json'
    print(f"Exporting to {output_path}...")
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(final_data, f, ensure_ascii=False, cls=NpEncoder)
    print("Done!")

if __name__ == "__main__":
    convert_to_json()
