import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# Đọc dữ liệu từ file Excel
data_path = r'Đề Kiểm Tra Toàn Diện Số 1 - Lớp 12 - Thầy VNA (Câu trả lời).xlsx' 
data = pd.read_excel(data_path)

# Chuẩn bị dữ liệu
scores = data['Unnamed: 4'].dropna()
scores = pd.to_numeric(scores, errors='coerce').dropna()

# Đọc các mức điểm độc đáo từ dữ liệu và sắp xếp chúng
unique_scores = np.sort(scores.unique())

# Tính số lượng thí sinh cho mỗi mức điểm độc đáo
score_counts = scores.value_counts().reindex(unique_scores, fill_value=0)

# Tạo dữ liệu cho biểu đồ
score_data = pd.DataFrame({
    'Score': unique_scores,
    'Number of Students': score_counts.values
})

# Hàm vẽ biểu đồ phổ điểm theo chiều dọc
def plot_vertical_distribution():
    plt.figure(figsize=(14, 8), dpi=137)  # Kích thước ảnh là 14x8 inch, độ phân giải 137 dpi để đạt 1920x1080 pixel
    # Sử dụng chỉ số của mảng cho trục x để đảm bảo khoảng cách đều
    x_indexes = np.arange(len(unique_scores)) * 1.2  # Tăng khoảng cách giữa các cột
    plt.bar(x_indexes, score_data['Number of Students'], width=0.8, edgecolor='black', color='blue', align='center')
    plt.xlabel('Điểm')
    plt.ylabel('Số lượng thí sinh')
    plt.title('Đề Kiểm Tra Toàn Diện Số 1 - Lớp 12 - Thầy VNA')
    
    # Thêm nhãn cho từng cột
    for i, v in enumerate(score_data['Number of Students']):
        plt.text(x_indexes[i], v + 1, str(v), ha='center', color='black')

    plt.grid(True, linestyle='--')
    # Đặt nhãn trục x bằng các giá trị điểm duy nhất và xoay 90 độ
    plt.xticks(x_indexes, labels=[f"{score:.2f}" for score in unique_scores], rotation=90)
    plt.tight_layout()
    plt.savefig('vertical_distribution.png', dpi=137)  # Đảm bảo độ phân giải ảnh là 1920x1080
    plt.show()

# Hàm thống kê điểm theo mẫu và xuất ra ảnh
def display_score_summary_image():
    score_summary = {
        'Điểm': ['Tổng số thí sinh', 'Số học sinh đạt điểm 10', 'Số học sinh đạt điểm 9.75', 'Số học sinh đạt điểm 9.50', 'Số học sinh đạt điểm 9.25', 'Số học sinh đạt điểm 9.00'],
        'Số lượng': [len(scores), score_counts.get(10, 0), score_counts.get(9.75, 0), score_counts.get(9.5, 0), score_counts.get(9.25, 0), score_counts.get(9, 0)]
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

# Gọi hàm để vẽ biểu đồ
plot_vertical_distribution()

# Gọi hàm để hiển thị bảng thống kê điểm dưới dạng ảnh
display_score_summary_image()
