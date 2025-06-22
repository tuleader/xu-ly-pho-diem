import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# Đọc dữ liệu từ file Excel
file_path = 'form.xlsx'
df = pd.read_excel(file_path, sheet_name='BXH')

# Cột điểm là cột E (thường là 'Unnamed: 4')
score_col = df.columns[4]

# Chuyển '9,75' → '9.75' và ép thành float
scores_raw = df[score_col].astype(str).str.replace(',', '.', regex=False)
scores = pd.to_numeric(scores_raw, errors='coerce').dropna()

# Thống kê cơ bản
total_students = len(scores)
average_score = scores.mean()
median_score = scores.median()
mode_score = scores.mode().iloc[0] if not scores.mode().empty else None
students_below_1 = (scores <= 1).sum()
percentage_below_1 = (students_below_1 / total_students) * 100
students_below_5 = (scores < 5).sum()
percentage_below_5 = (students_below_5 / total_students) * 100

# Thống kê số lượng thí sinh theo từng mức điểm
score_counts = scores.value_counts().sort_index()

# Tạo DataFrame để vẽ biểu đồ
score_data = pd.DataFrame({
    'Score': score_counts.index,
    'Number of Students': score_counts.values
})

# Biểu đồ phổ điểm


def plot_vertical_distribution():
    plt.figure(figsize=(14, 8), dpi=274)
    x_indexes = np.arange(len(score_data)) * 1.2
    plt.bar(x_indexes, score_data['Number of Students'],
            width=0.8, edgecolor='black', color='blue')
    plt.xlabel('Điểm')
    plt.ylabel('Số lượng thí sinh')
    plt.title('Đề Kiểm Tra Toàn Diện Chương 1 - Thầy VNA - Đề số 2')

    # Ghi nhãn từng cột
    for i, v in enumerate(score_data['Number of Students']):
        plt.text(x_indexes[i], v + 1, str(v),
                 ha='center', va='bottom', fontsize=9)

    # Hiển thị các mức điểm đầy đủ
    plt.xticks(
        x_indexes, [f"{score:.2f}" for score in score_data['Score']], rotation=90)
    plt.grid(axis='y', linestyle='--', alpha=0.5)
    plt.tight_layout()
    plt.savefig('vertical_distribution.png', dpi=274)
    plt.show()

# Bảng thống kê tổng quát


def plot_statistics_table():
    stats = {
        "Statistic": [
            "Tổng số thí sinh",
            "Điểm trung bình",
            "Trung vị",
            "Số thí sinh đạt điểm ≤ 1",
            "Số thí sinh dưới trung bình (< 5)",
            "Mốc điểm phổ biến nhất (mode)"
        ],
        "Value": [
            total_students,
            f"{average_score:.2f}",
            f"{median_score:.2f}",
            f"{students_below_1} ({percentage_below_1:.2f}%)",
            f"{students_below_5} ({percentage_below_5:.2f}%)",
            f"{mode_score:.2f}" if mode_score is not None else "N/A"
        ]
    }
    stats_df = pd.DataFrame(stats)
    fig, ax = plt.subplots(figsize=(8, 4), dpi=480)
    ax.axis('off')
    ax.table(cellText=stats_df.values, colLabels=stats_df.columns,
             cellLoc='center', loc='center')
    plt.tight_layout()
    plt.savefig('so_lieu_thong_ke.png', dpi=480)
    plt.show()

# Bảng tổng hợp điểm cao


def display_score_summary_image():
    summary = {
        'Điểm': ['Tổng số thí sinh', 'Điểm 10', 'Điểm 9.75', 'Điểm 9.50', 'Điểm 9.25', 'Điểm 9.00'],
        'Số lượng': [
            len(scores),
            score_counts.get(10, 0),
            score_counts.get(9.75, 0),
            score_counts.get(9.5, 0),
            score_counts.get(9.25, 0),
            score_counts.get(9.0, 0),
        ]
    }
    summary_df = pd.DataFrame(summary)
    fig, ax = plt.subplots(figsize=(8, 4), dpi=480)
    ax.axis('off')
    table = ax.table(cellText=summary_df.values,
                     colLabels=summary_df.columns, cellLoc='center', loc='center')
    table.auto_set_font_size(False)
    table.set_fontsize(12)
    table.scale(1.2, 1.2)
    plt.savefig('score_summary_image.png', bbox_inches='tight', dpi=480)
    plt.show()


# Gọi hàm xuất kết quả
plot_vertical_distribution()
plot_statistics_table()
display_score_summary_image()
