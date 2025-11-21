import math
import pandas as pd  # Thư viện xử lý Excel

# === 1. KHAI BÁO THAM SỐ KHO HÀNG ===
d_slot = 1.2  # Chiều dài mỗi ô kệ
d_row = 0.3  # Bề dày kệ
d_cross = 2.0  # Khoảng cách lối đi (vertical) giữa các cụm (A-B, B-C)
d_aisle = 1.0  # Chiều rộng lối đi (giữa dãy 1-2, 3-4...)(horizontal)
BLOCK_WIDTH = 6 * d_slot # Chiều rộng vật lý của 1 cụm (6 kệ liền nhau)
Block_map = {'A': 1, 'B': 2, 'C': 3} # Bản đồ giá trị các cụm

# === 2. CÁC HÀM XỬ LÝ LOGIC ===
def parse_ten_ke(ten_ke): #Dịch tên kệ (VD: '15B-4') thành 3 phần: n=15, M='B', z=4
    vi_tri_gach = ten_ke.rfind('-')
    nM = ten_ke[:vi_tri_gach]
    z = ten_ke[vi_tri_gach + 1:]
    M = nM[-1]  # Chữ cái cuối cùng là Cụm (A, B, C)
    n = nM[:-1]  # Phần số còn lại là Dãy (1-20)
    return int(n), M, int(z)

def tinh_d_doc(n1, n2): #Tính khoảng cách di chuyển DỌC giữa dãy n1 và n2.Bao gồm bề dày dãy (row) và lối đi (aisle).
    delta = abs(n2 - n1)
    dist = (delta * d_row) + (math.floor(delta / 2) * d_aisle) #Công thức: (số dãy * 0.3) + (số lối đi * 1.0)
    if delta % 2 != 0: # Nếu số dãy đi qua là lẻ (dư 1), cộng thêm 0.3m bề dày dãy
        dist += 0.3
    return dist

def tinh_khoang_cach(ke_1, ke_2):
    # Hàm cốt lõi: Tính khoảng cách ngắn nhất giữa 2 kệ bất kỳ. CẤU TRÚC: if - elif - else để tránh lỗi logic.
    n1, M1, z1 = parse_ten_ke(ke_1)
    n2, M2, z2 = parse_ten_ke(ke_2)

    # --- TRƯỜNG HỢP 1: CÙNG CỤM, CÙNG DÃY ---
    if M1 == M2 and n1 == n2:
        return abs(z1 - z2) * d_slot

    # --- TRƯỜNG HỢP 2: CÙNG CỤM (Nhưng Khác Dãy) ---
    # Dùng 'elif' (Else If): Chỉ chạy vào đây nếu TH1 sai VÀ M1 == M2
    elif M1 == M2:
        # 1. Tính cách đi vòng (An toàn nhất)
        d_doc = tinh_d_doc(n1, n2)
        path_1 = (abs(z1 - 1) * d_slot) + d_doc + (abs(z2 - 1) * d_slot)
        path_6 = (abs(z1 - 6) * d_slot) + d_doc + (abs(z2 - 6) * d_slot)
        ket_qua = min(path_1, path_6)

        # 2. Kiểm tra BĂNG QUA LỐI ĐI (Cross-Aisle) - Chỉ cho cặp đối diện
        if abs(n1 - n2) == 1:
            n_min = min(n1, n2)
            if n_min % 2 == 0:  # Dãy chẵn nhìn sang lẻ -> Có lối đi
                path_cross = (abs(z1 - z2) * d_slot) + d_aisle
                ket_qua = min(ket_qua, path_cross)
        return ket_qua

    # --- TRƯỜNG HỢP 3: KHÁC CỤM (M1 != M2) ---
    # Dùng 'else': Bắt buộc chạy vào đây khi cả 2 trường hợp trên đều sai.
    else:
        val_1 = Block_map[M1]
        val_2 = Block_map[M2]
        so_khe_ho = abs(val_1 - val_2)
        # Vì M1 khác M2 nên so_khe_ho tối thiểu là 1 -> so_cum_giua >= 0
        so_cum_giua = so_khe_ho - 1
        dist_cross_total = (so_khe_ho * d_cross) + (so_cum_giua * BLOCK_WIDTH)

        if val_1 < val_2:  # Hướng TIẾN (VD: A -> C)
            d_out = abs(z1 - 6) * d_slot
            d_doc = tinh_d_doc(n1, n2)
            d_in = abs(z2 - 1) * d_slot
            return d_out + dist_cross_total + d_doc + d_in

        else:  # Hướng LÙI (VD: C -> A)
            d_out = abs(z1 - 1) * d_slot
            d_doc = tinh_d_doc(n1, n2)
            d_in = abs(z2 - 6) * d_slot
            return d_out + dist_cross_total + d_doc + d_in

# === 3. TẠO DANH SÁCH KỆ (Theo thứ tự quét ngang) ===
def tao_danh_sach_ke():
    danh_sach = []
    # Duyệt theo DÃY trước (1->20), rồi mới duyệt CỤM (A->B->C)
    for day in range(1, 21):
        for cum in ['A', 'B', 'C']:
            for ke in range(1, 7):
                ten_ke = f"{day}{cum}-{ke}"
                danh_sach.append(ten_ke)
    return danh_sach

# === 4. CHẠY VÒNG LẶP VÀ XUẤT EXCEL ===
def main():
    print("Đang tạo danh sách kệ...")
    danh_sach_ke = tao_danh_sach_ke()
    n = len(danh_sach_ke)
    print(f"Tổng số kệ: {n} (Kiểm tra: 360?)")
    print("Đang tính toán ma trận khoảng cách (sẽ mất khoảng vài giây)...")
    ma_tran_du_lieu = []

    # Vòng lặp tính toán 360x360
    for i in range(n):
        ke_nguon = danh_sach_ke[i]
        hang_hien_tai = []

        for j in range(n):
            ke_dich = danh_sach_ke[j]

            if i == j:
                kc = 0.0
            else:
                kc = tinh_khoang_cach(ke_nguon, ke_dich)

            hang_hien_tai.append(round(kc, 2))

        ma_tran_du_lieu.append(hang_hien_tai)

    print("Tính toán xong. Đang xuất ra file Excel...")

    # Tạo DataFrame và lưu file
    df = pd.DataFrame(ma_tran_du_lieu, index=danh_sach_ke, columns=danh_sach_ke)
    ten_file = 'Ma_Tran_Khoang_Cach_Final.xlsx'
    df.to_excel(ten_file)
    print(f"✅ XONG! File '{ten_file}' đã được tạo thành công.")

if __name__ == "__main__":
    main()