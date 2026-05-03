import json
import numpy as np
import pandas as pd


def load_data():
    filename = "analysis_result.xlsx"
    print(f"Loading data from {filename}...")
    try:
        exams = pd.read_excel(filename, sheet_name="Exams_Metadata")
        students = pd.read_excel(filename, sheet_name="Student_Comparison")
        class_summary = pd.read_excel(filename, sheet_name="Class_Summary")
        subject_summary = pd.read_excel(filename, sheet_name="Subject_Summary")
        return exams, students, class_summary, subject_summary
    except Exception as e:
        print(f"Error loading data: {e}")
        return None, None, None, None


def to_native(value):
    if pd.isna(value):
        return None
    if isinstance(value, (np.integer,)):
        return int(value)
    if isinstance(value, (np.floating, float)):
        return float(value)
    return value


def convert_to_json():
    exams_df, students_df, class_df, subject_df = load_data()
    if exams_df is None:
        return

    exams_df = exams_df.sort_values("order").reset_index(drop=True)
    exams_df = exams_df.replace({np.nan: None})
    exams = exams_df[["order", "exam_key", "exam_name", "source_file"]].to_dict(orient="records")

    score_distribution = {}
    for exam in exams:
        key = exam["exam_key"]
        score_col = f"总分__分数__{key}"
        score_distribution[key] = {
            "exam_name": exam["exam_name"],
            "scores": students_df[score_col].dropna().tolist() if score_col in students_df.columns else [],
        }

    latest_exam = exams[-1]
    previous_exam = exams[-2] if len(exams) >= 2 else exams[-1]

    latest_rank_col = f"总分__联考排名__{latest_exam['exam_key']}"
    prev_rank_col = f"总分__联考排名__{previous_exam['exam_key']}"
    compare_df = students_df.copy()
    compare_df["rank_change"] = compare_df[prev_rank_col] - compare_df[latest_rank_col]
    compare_df = compare_df.dropna(subset=["rank_change"])

    top_improvers = (
        compare_df.nlargest(5, "rank_change")[["Display_Name", "Display_Class", "rank_change"]]
        .rename(columns={"Display_Name": "name", "Display_Class": "class"})
        .to_dict(orient="records")
    )
    bottom_improvers = (
        compare_df.nsmallest(5, "rank_change")[["Display_Name", "Display_Class", "rank_change"]]
        .rename(columns={"Display_Name": "name", "Display_Class": "class"})
        .to_dict(orient="records")
    )

    subjects = sorted(subject_df["subject"].dropna().unique().tolist())
    students = []

    for _, row in students_df.iterrows():
        exam_stats = []
        for exam in exams:
            key = exam["exam_key"]
            exam_stats.append({
                "exam_key": key,
                "exam_name": exam["exam_name"],
                "total_score": to_native(row.get(f"总分__分数__{key}")),
                "total_joint_rank": to_native(row.get(f"总分__联考排名__{key}")),
                "total_school_rank": to_native(row.get(f"总分__学校排名__{key}")),
                "class_name": row.get(f"Class__{key}") if f"Class__{key}" in row else None,
                "school_name": row.get(f"School__{key}") if f"School__{key}" in row else None,
            })

        subject_stats = []
        for subject in subjects:
            subject_exam_stats = []
            for exam in exams:
                key = exam["exam_key"]
                subject_exam_stats.append({
                    "exam_key": key,
                    "exam_name": exam["exam_name"],
                    "score": to_native(row.get(f"{subject}__分数__{key}")),
                    "joint_rank": to_native(row.get(f"{subject}__联考排名__{key}")),
                    "school_rank": to_native(row.get(f"{subject}__学校排名__{key}")),
                })
            subject_stats.append({
                "name": subject,
                "exam_stats": subject_exam_stats,
            })

        latest_change = None
        if pd.notna(row.get(prev_rank_col)) and pd.notna(row.get(latest_rank_col)):
            latest_change = to_native(row.get(prev_rank_col) - row.get(latest_rank_col))

        students.append({
            "name": row["Display_Name"],
            "student_id": row["StudentID"],
            "class": row["Display_Class"],
            "school": row.get("Display_School"),
            "exam_stats": exam_stats,
            "subjects": subject_stats,
            "latest_rank_change": latest_change,
        })

    subject_df = subject_df.replace({np.nan: None})
    class_df = class_df.replace({np.nan: None})

    final_data = {
        "exams": exams,
        "global_stats": {
            "total_students": int(len(students_df)),
            "score_distribution": score_distribution,
        },
        "subject_stats": subject_df.to_dict(orient="records"),
        "class_stats": class_df.to_dict(orient="records"),
        "comparison_window": {
            "from_exam_key": previous_exam["exam_key"],
            "from_exam_name": previous_exam["exam_name"],
            "to_exam_key": latest_exam["exam_key"],
            "to_exam_name": latest_exam["exam_name"],
        },
        "top_improvers": top_improvers,
        "bottom_improvers": bottom_improvers,
        "students": students,
    }

    output_path = "dashboard/data.json"
    print(f"Exporting to {output_path}...")
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(final_data, f, ensure_ascii=False)
    print("Done!")


if __name__ == "__main__":
    convert_to_json()
