import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# Đọc dữ liệu từ file Excel
data_path = r'D:\Docs_Github\xu-ly-pho-diem\Đề thi khảo sát - Số 01 - Lớp 12 - Thầy VNA (Câu trả lời) (1).xlsx'
data = pd.read_excel(data_path)

# Chuẩn bị dữ liệu
scores = data['Unnamed: 5'].dropna()
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
    plt.figure(figsize=(14, 8))  # Tăng kích thước biểu đồ để dễ dàng xem các nhãn
    # Sử dụng chỉ số của mảng cho trục x để đảm bảo khoảng cách đều
    x_indexes = np.arange(len(unique_scores)) * 1.5  # Tăng khoảng cách giữa các cột
    plt.bar(x_indexes, score_data['Number of Students'], width=0.8, edgecolor='black', color='blue', align='center')
    plt.xlabel('Điểm')
    plt.ylabel('Số lượng thí sinh')
    plt.title('Phổ điểm Đề thi khảo sát - Số 01 - Lớp 12 - Thầy VNA')
    
    # Thêm nhãn cho từng cột
    for i, v in enumerate(score_data['Number of Students']):
        plt.text(x_indexes[i], v + 1, str(v), ha='center', color='black')

    plt.grid(True, linestyle='--')
    # Đặt nhãn trục x bằng các giá trị điểm duy nhất và xoay 90 độ
    plt.xticks(x_indexes, labels=[f"{score:.2f}" for score in unique_scores], rotation=90)
    plt.tight_layout()
    plt.savefig('vertical_distribution.png')
    plt.show()

# Gọi hàm để vẽ biểu đồ
plot_vertical_distribution()
