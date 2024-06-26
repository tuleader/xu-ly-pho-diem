import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# Hàm làm tròn điểm đến 0.25 gần nhất
def round_to_nearest_quarter(score):
    return round(score * 4) / 4

# Đọc dữ liệu từ file Excel
data_path = r'form.xlsx'  # Thay bằng đường dẫn đến file của bạn
data = pd.read_excel(data_path)

# Chuẩn bị dữ liệu
scores = data['Unnamed: 4'].dropna()
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
    plt.figure(figsize=(14, 8), dpi=274)  # Kích thước ảnh là 14x8 inch, độ phân giải 274 dpi để đạt 3840x2160 pixel
    plt.bar(rounded_score_data['Score'], rounded_score_data['Number of Students'], width=0.2, edgecolor='black', color='blue', align='center')
    plt.xlabel('Điểm')
    plt.ylabel('Số lượng thí sinh')
    plt.title('Phổ Điểm Đề Chuẩn Cấu Trúc Ma Trận 2024 - Đề số 5 - Thầy VNA')

    for i, v in enumerate(rounded_score_data['Number of Students']):
        plt.text(rounded_score_data['Score'][i], v + 3, str(v), ha='center', color='black')

    plt.grid(True, linestyle='--')
    plt.xticks(np.arange(0, 10.25, 0.25), rotation=90)
    plt.tight_layout()
    plt.savefig('pho_diem_doc.png', dpi=274)  # Đảm bảo độ phân giải ảnh là 3840x2160
    plt.show()

# Hàm vẽ biểu đồ phổ điểm theo chiều ngang
def plot_horizontal_distribution():
    plt.figure(figsize=(14, 10), dpi=274)  # Kích thước ảnh là 14x10 inch, độ phân giải 274 dpi để đạt 3840x2160 pixel
    plt.barh(rounded_score_data['Score'], rounded_score_data['Number of Students'], height=0.2, edgecolor='black', color='blue', align='center')
    plt.ylabel('Điểm')
    plt.xlabel('Số lượng thí sinh')
    plt.title('Phổ Điểm Đề Chuẩn Cấu Trúc Ma Trận 2024 - Đề số 5 - Thầy VNA')

    for i, v in enumerate(rounded_score_data['Number of Students']):
        plt.text(v + 3, rounded_score_data['Score'][i], str(v), va='center', color='black')

    plt.grid(True, linestyle='--')
    plt.yticks(np.arange(0, 10.25, 0.25))
    plt.tight_layout()
    plt.savefig('pho_diem_ngang.png', dpi=274)  # Đảm bảo độ phân giải ảnh là 3840x2160
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

    fig, ax = plt.subplots(figsize=(8, 4), dpi=480)  # Kích thước ảnh là 8x4 inch, độ phân giải 480 dpi để đạt 3840x2160 pixel
    ax.axis('tight')
    ax.axis('off')
    ax.table(cellText=stats_df.values, colLabels=stats_df.columns, cellLoc='center', loc='center')

    plt.tight_layout()
    plt.savefig('so_lieu_thong_ke.png', dpi=480)  # Đảm bảo độ phân giải ảnh là 3840x2160
    plt.show()

# Hàm thống kê điểm theo mẫu và xuất ra ảnh
def display_score_summary_image():
    score_summary = {
        'Điểm': ['Tổng số thí sinh', 'Số học sinh đạt điểm 10', 'Số học sinh đạt điểm 9.75', 'Số học sinh đạt điểm 9.50', 'Số học sinh đạt điểm 9.25', 'Số học sinh đạt điểm 9.00'],
        'Số lượng': [len(scores), rounded_score_counts.get(10, 0), rounded_score_counts.get(9.75, 0), rounded_score_counts.get(9.5, 0), rounded_score_counts.get(9.25, 0), rounded_score_counts.get(9, 0)]
    }
    summary_df = pd.DataFrame(score_summary)
    
    # Tạo ảnh của bảng
    fig, ax = plt.subplots(figsize=(8, 4), dpi=240)  # Kích thước ảnh là 8x4 inch, độ phân giải 240 dpi để đạt 1920x1080 pixel
    ax.axis('tight')
    ax.axis('off')
    table = ax.table(cellText=summary_df.values, colLabels=summary_df.columns, cellLoc='center', loc='center')
    table.auto_set_font_size(False)
    table.set_fontsize(12)
    table.scale(1.2, 1.2)
    
    # Lưu ảnh
    plt.savefig('score_summary_image.png', bbox_inches='tight', dpi=240)  # Đảm bảo độ phân giải ảnh là 1920x1080
    plt.show()
    
# Gọi các hàm để vẽ biểu đồ và in bảng số liệu thống kê
plot_vertical_distribution()
plot_horizontal_distribution()
plot_statistics_table()
display_score_summary_image()