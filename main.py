import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# Hàm làm tròn điểm đến 0.25 gần nhất
def round_to_nearest_quarter(score):
    return round(score * 4) / 4

# Đọc dữ liệu từ file Excel
data_path = r'/home/tuleader/Documents/phodiem/data.xlsx'  # Thay bằng đường dẫn đến file của bạn
data = pd.read_excel(data_path)

# Chuẩn bị dữ liệu
scores = data['Unnamed: 5'].dropna()
scores = pd.to_numeric(scores, errors='coerce').dropna()
rounded_scores = scores.apply(round_to_nearest_quarter)
rounded_score_counts = rounded_scores.value_counts().sort_index()

# Thống kê
total_students = scores.size
average_score = scores.mean()
median_score = scores.median()
mode_score = scores.mode().iloc[0] if not scores.mode().empty else None
students_below_1 = (scores <= 1).sum()
percentage_below_1 = (students_below_1 / total_students) * 100
students_below_5 = (scores < 5).sum()
percentage_below_5 = (students_below_5 / total_students) * 100

# Tạo dữ liệu cho biểu đồ
rounded_score_data = pd.DataFrame({
    'Score': rounded_score_counts.index,
    'Number of Students': rounded_score_counts.values
})

# Hàm vẽ biểu đồ phổ điểm theo chiều dọc
def plot_vertical_distribution():
    plt.figure(figsize=(12, 8))
    plt.bar(rounded_score_data['Score'], rounded_score_data['Number of Students'], width=0.2, edgecolor='black', color='blue', align='center')
    plt.xlabel('Điểm')
    plt.ylabel('Số lượng thí sinh')
    plt.title('Phổ điểm Đề thi khảo sát - Số 01 - Lớp 12 - Thầy VNA')

    for i, v in enumerate(rounded_score_data['Number of Students']):
        plt.text(rounded_score_data['Score'][i], v + 3, str(v), ha='center', color='black')

    plt.grid(True, linestyle='--')
    plt.xticks(np.arange(0, 10.25, 0.25), rotation=90)
    plt.tight_layout()
    plt.savefig('vertical_distribution.png')
    plt.show()

# Hàm vẽ biểu đồ phổ điểm theo chiều ngang
def plot_horizontal_distribution():
    plt.figure(figsize=(14, 10))
    plt.barh(rounded_score_data['Score'], rounded_score_data['Number of Students'], height=0.2, edgecolor='black', color='blue', align='center')
    plt.ylabel('Điểm')
    plt.xlabel('Số lượng thí sinh')
    plt.title('Phổ điểm Đề thi khảo sát - Số 01 - Lớp 12 - Thầy VNA')

    for i, v in enumerate(rounded_score_data['Number of Students']):
        plt.text(v + 3, rounded_score_data['Score'][i], str(v), va='center', color='black')

    plt.grid(True, linestyle='--')
    plt.yticks(np.arange(0, 10.25, 0.25))
    plt.tight_layout()
    plt.savefig('horizontal_distribution.png')
    plt.show()

# Hàm tạo bảng số liệu thống kê
def plot_statistics_table():
    stats_data = {
        "Statistic": [
            "Tổng số thí sinh", 
            "Điểm trung bình", 
            "Trung vị", 
            "Số thí sinh đạt điểm <=1", 
            "Số thí sinh đạt điểm dưới trung bình (<5)", 
            "Mốc điểm trung bình có nhiều thí sinh đạt được nhất"
        ],
        "Value": [
            total_students, 
            f"{average_score:.2f}", 
            f"{median_score:.2f}", 
            students_below_1, 
            f"{students_below_5} ({percentage_below_5:.2f}%)", 
            f"{mode_score:.2f}"
        ]
    }

    stats_df = pd.DataFrame(stats_data)

    fig, ax = plt.subplots(figsize=(8, 4))
    ax.axis('tight')
    ax.axis('off')
    ax.table(cellText=stats_df.values, colLabels=stats_df.columns, cellLoc='center', loc='center')

    plt.tight_layout()
    plt.savefig('statistics_table.png')
    plt.show()

# Gọi các hàm để vẽ biểu đồ và in bảng số liệu thống kê
plot_vertical_distribution()
plot_horizontal_distribution()
plot_statistics_table()
