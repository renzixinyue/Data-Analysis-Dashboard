import pandas as pd


EXAMS = [
    {"key": "seven_up_monthly", "name": "七上第一次月考", "file": "diyiciyuekao.xlsx"},
    {"key": "seven_up_midterm", "name": "七上期中", "file": "qizhognchengji.xlsx"},
    {"key": "seven_down_midterm", "name": "七下期中", "file": "2026年七下期中考试排行榜.xlsx"},
]
CLASS_AVG_FILES = {
    "seven_down_midterm": "2026年七下期中考试班级平均分.xlsx",
}

META_MAP = {
    "姓名": "Name",
    "学号": "StudentID",
    "考号": "ExamID",
    "班级": "Class",
    "学校": "School",
    "标签": "Tag",
}

META_FIELDS = {"Name", "StudentID", "ExamID", "Class", "School", "Tag"}
RANK_METRIC = "学校排名"
JOINT_RANK_METRIC = "联考排名"


def flatten_columns(df):
    columns = []
    last_subject = None
    for col in df.columns:
        subject = str(col[0]).replace(" ", "").replace("\u3000", "").strip()
        metric = str(col[1]).replace(" ", "").replace("\u3000", "").strip()
        if "Unnamed" in subject or subject == "nan":
            subject = last_subject
        else:
            last_subject = subject
        if "Unnamed" in metric or metric == "nan":
            metric = ""
        columns.append(f"{subject}_{metric}" if subject and metric else (subject or metric))
    df.columns = columns
    return df


def normalize_student_id(series):
    return (
        series.astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
    )


def load_exam(exam):
    print(f"Loading {exam['file']} ...")
    df = pd.read_excel(exam["file"], header=[1, 2])
    df = flatten_columns(df)

    renamed = {}
    subject_order = []

    for col in df.columns:
        if col in META_MAP:
            renamed[col] = META_MAP[col]
            continue

        if "_" not in col:
            renamed[col] = col
            continue

        subject, metric = col.split("_", 1)
        if subject not in META_MAP and subject not in subject_order and subject != "总分":
            subject_order.append(subject)
        renamed[col] = f"{subject}__{metric}"

    df = df.rename(columns=renamed)
    df = df.dropna(subset=["Name", "StudentID"])
    df["StudentID"] = normalize_student_id(df["StudentID"])

    for col in df.columns:
        if col not in META_FIELDS and col != "StudentID":
            df[col] = pd.to_numeric(df[col], errors="coerce")

    rank_subjects = ["总分"] + subject_order
    for subject in rank_subjects:
        score_col = f"{subject}__分数"
        joint_rank_col = f"{subject}__{JOINT_RANK_METRIC}"
        if score_col in df.columns and joint_rank_col not in df.columns:
            df[joint_rank_col] = (
                df[score_col]
                .rank(ascending=False, method="min")
            )

    renamed_for_exam = {}
    for col in df.columns:
        if col == "StudentID":
            continue
        renamed_for_exam[col] = f"{col}__{exam['key']}"

    df = df.rename(columns=renamed_for_exam)
    return df, subject_order


def first_non_null(row, columns):
    for col in columns:
        if col in row and pd.notna(row[col]) and str(row[col]).strip():
            return row[col]
    return ""


def load_official_class_summary(exam):
    avg_file = CLASS_AVG_FILES.get(exam["key"])
    if not avg_file:
        return None

    df = pd.read_excel(avg_file, sheet_name="平均分", header=[1, 2])
    df = df.iloc[1:].copy()
    df.columns = [
        "class_name",
        "applicant_count",
        "student_count",
        "teacher_name",
        "avg_total_score",
        "avg_score_rate",
        "class_rank",
        "avg_score_diff",
        "score_std",
        "max_score",
        "min_score",
    ]
    df = df[df["class_name"].astype(str).str.fullmatch(r"\d+")].copy()
    df["class_name"] = df["class_name"].astype(str).str.strip()
    for col in ["student_count", "avg_total_score", "class_rank"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    return df[["class_name", "avg_total_score", "student_count", "class_rank"]]


def build_outputs():
    exam_frames = []
    subject_order = []

    for exam in EXAMS:
        exam_df, exam_subjects = load_exam(exam)
        exam_frames.append(exam_df)
        for subject in exam_subjects:
            if subject not in subject_order:
                subject_order.append(subject)

    merged_df = exam_frames[0]
    for exam_df in exam_frames[1:]:
        merged_df = pd.merge(merged_df, exam_df, on="StudentID", how="inner")
    merged_df = merged_df.copy()

    name_cols = [f"Name__{exam['key']}" for exam in reversed(EXAMS)]
    class_cols = [f"Class__{exam['key']}" for exam in reversed(EXAMS)]
    school_cols = [f"School__{exam['key']}" for exam in reversed(EXAMS)]

    merged_df["Display_Name"] = merged_df.apply(lambda row: first_non_null(row, name_cols), axis=1)
    merged_df["Display_Class"] = merged_df.apply(lambda row: first_non_null(row, class_cols), axis=1)
    merged_df["Display_School"] = merged_df.apply(lambda row: first_non_null(row, school_cols), axis=1)

    latest_score_col = f"总分__分数__{EXAMS[-1]['key']}"
    if latest_score_col in merged_df.columns:
        merged_df = merged_df.sort_values(latest_score_col, ascending=False)

    overall_rows = []
    class_rows = []
    subject_rows = []

    for order, exam in enumerate(EXAMS, start=1):
        key = exam["key"]
        total_score_col = f"总分__分数__{key}"
        total_rank_col = f"总分__{JOINT_RANK_METRIC}__{key}"
        class_col = f"Class__{key}"

        overall_rows.append({
            "order": order,
            "exam_key": key,
            "exam_name": exam["name"],
            "source_file": exam["file"],
            "rank_metric": JOINT_RANK_METRIC,
            "student_count": int(merged_df[total_score_col].notna().sum()) if total_score_col in merged_df.columns else 0,
            "avg_total_score": merged_df[total_score_col].mean() if total_score_col in merged_df.columns else None,
            "avg_total_rank": merged_df[total_rank_col].mean() if total_rank_col in merged_df.columns else None,
        })

        if class_col in merged_df.columns and total_score_col in merged_df.columns:
            class_agg = load_official_class_summary(exam)
            if class_agg is None:
                class_agg = (
                    merged_df.groupby(class_col, dropna=True)
                    .agg(
                        avg_total_score=(total_score_col, "mean"),
                        avg_total_rank=(total_rank_col, "mean") if total_rank_col in merged_df.columns else (total_score_col, "size"),
                        student_count=(total_score_col, "count"),
                    )
                    .reset_index()
                    .rename(columns={class_col: "class_name"})
                )
                class_agg["class_rank"] = (
                    class_agg["avg_total_score"]
                    .rank(ascending=False, method="min")
                )
            else:
                class_agg["avg_total_rank"] = None

            for _, row in class_agg.iterrows():
                class_rows.append({
                    "exam_key": key,
                    "exam_name": exam["name"],
                    "class_name": row["class_name"],
                    "avg_total_score": row["avg_total_score"],
                    "avg_total_rank": row["avg_total_rank"],
                    "student_count": int(row["student_count"]),
                    "class_rank": int(row["class_rank"]) if pd.notna(row["class_rank"]) else None,
                })

        for subject in subject_order:
            score_col = f"{subject}__分数__{key}"
            rank_col = f"{subject}__{JOINT_RANK_METRIC}__{key}"
            if score_col not in merged_df.columns:
                continue
            subject_rows.append({
                "exam_key": key,
                "exam_name": exam["name"],
                "subject": subject,
                "avg_score": merged_df[score_col].mean(),
                "avg_rank": merged_df[rank_col].mean() if rank_col in merged_df.columns else None,
            })

    exams_df = pd.DataFrame(overall_rows)
    class_df = pd.DataFrame(class_rows)
    subject_df = pd.DataFrame(subject_rows)

    preferred_cols = ["StudentID", "Display_Name", "Display_Class", "Display_School"]
    for exam in EXAMS:
        key = exam["key"]
        preferred_cols.extend(
            [
                f"Name__{key}",
                f"Class__{key}",
                f"School__{key}",
                f"总分__分数__{key}",
                f"总分__{JOINT_RANK_METRIC}__{key}",
            ]
        )
        for subject in subject_order:
            preferred_cols.extend(
                [
                    f"{subject}__分数__{key}",
                    f"{subject}__{JOINT_RANK_METRIC}__{key}",
                ]
            )

    remaining = [col for col in merged_df.columns if col not in preferred_cols]
    student_df = merged_df[[col for col in preferred_cols if col in merged_df.columns] + remaining]
    return exams_df, student_df, class_df, subject_df


def analyze():
    exams_df, student_df, class_df, subject_df = build_outputs()
    output_file = "analysis_result.xlsx"
    print(f"Writing results to {output_file} ...")
    with pd.ExcelWriter(output_file) as writer:
        exams_df.to_excel(writer, sheet_name="Exams_Metadata", index=False)
        student_df.to_excel(writer, sheet_name="Student_Comparison", index=False)
        class_df.to_excel(writer, sheet_name="Class_Summary", index=False)
        subject_df.to_excel(writer, sheet_name="Subject_Summary", index=False)
    print("Analysis complete!")


if __name__ == "__main__":
    analyze()
