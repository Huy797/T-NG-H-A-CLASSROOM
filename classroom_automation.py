import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd


class ClassroomAutomation:

    def __init__(self, excel_path):
        """Khởi tạo cấu hình hệ thống và nạp dữ liệu học viên."""
        self.excel_path = excel_path
        self.df = None
        self.sender_email = "email_cua_ban@gmail.com"
        self.sender_password = (
            "xxxx xxxx xxxx xxxx"
        )

    def load_data(self):
        """Đọc dữ liệu từ file Excel.

        Nếu file chưa tồn tại hoặc bị lỗi định dạng (file rỗng), tự động tạo
        mới.
        """
        default_df = pd.DataFrame(
            columns=["Ho_Ten", "Email", "So_Buoi_Nghi", "Diem_Danh_Gia"]
        )

        if not os.path.exists(self.excel_path):
            print(f"[!] File {self.excel_path} chưa tồn tại. Đang tạo mới...")
            self.df = default_df
            self.df.to_excel(self.excel_path, index=False)
            return True
        try:
            self.df = pd.read_excel(self.excel_path)
            if self.df.empty and len(self.df.columns) == 0:
                print("[!] File Excel đang bị rỗng. Đang tự động cấu trúc lại...")
                self.df = default_df
                self.df.to_excel(self.excel_path, index=False)
            else:
                print(
                    f"[+] Nạp dữ liệu thành công! Hiện có {len(self.df)} học viên."
                )
            return True
        except Exception as e:
            print(
                f"[!] Phát hiện file Excel lỗi format ({e}). Đang tự động tạo lại file mới chuẩn..."
            )
            self.df = default_df
            self.df.to_excel(self.excel_path, index=False)
            return True

    def input_students(self):
        """Giao diện nhập liệu học viên từ bàn phím trực tiếp trên VS Code."""
        print("\n=== CHƯƠNG TRÌNH NHẬP DỮ LIỆU HỌC VIÊN ===")
        print("(Nhập 'q' tại mục Tên học viên để dừng nhập và lưu lại)\n")

        while True:
            ho_ten = input("1. Nhập họ và tên học viên: ").strip()
            if ho_ten.lower() == "q":
                break
            if not ho_ten:
                print("[-] Tên không được để trống. Vui lòng nhập lại.")
                continue

            email = input("2. Nhập email học viên: ").strip()

            while True:
                try:
                    so_buoi_nghi = int(
                        input("3. Nhập số buổi nghỉ: ").strip() or 0
                    )
                    break
                except ValueError:
                    print("[-] Số buổi nghỉ phải là một chữ số. Nhập lại:")

            while True:
                try:
                    diem_dg = float(
                        input("4. Nhập điểm đánh giá (0-10): ").strip() or 0.0
                    )
                    break
                except ValueError:
                    print("[-] Điểm đánh giá phải là số (Ví dụ: 7.5). Nhập lại:")

            new_row = {
                "Ho_Ten": ho_ten,
                "Email": email,
                "So_Buoi_Nghi": so_buoi_nghi,
                "Diem_Danh_Gia": diem_dg,
            }

            self.df = pd.concat(
                [self.df, pd.DataFrame([new_row])], ignore_index=True
            )
            print(f"[+] Đã thêm tạm thời học viên: {ho_ten}\n")

        try:
            self.df.to_excel(self.excel_path, index=False)
            print(
                f"\n[🎉] ĐÃ LƯU THÀNH CÔNG! File Excel hiện có tổng cộng {len(self.df)} học viên."
            )
        except Exception as e:
            print(f"[-] Lỗi không thể ghi file Excel: {e}")

    def send_email(self, receiver_email, subject, body):
        """Hàm core xử lý việc kết nối SMTP và gửi Email."""
        msg = MIMEMultipart()
        msg["From"] = self.sender_email
        msg["To"] = receiver_email
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain", "utf-8"))

        try:
            with smtplib.SMTP("smtp.gmail.com", 587) as server:
                server.starttls()
                server.login(self.sender_email, self.sender_password)
                server.send_message(msg)
            print(f"[+] Đã gửi email thành công tới: {receiver_email}")
            return True
        except Exception as e:
            print(f"[-] Thất bại khi gửi tới {receiver_email}. Lỗi: {e}")
            return False

    def auto_remind_attendance(self, max_allowed_absent=3):
        """Tự động lọc và gửi mail cảnh báo học viên nghỉ quá số buổi."""
        print(
            f"\n[!] Bắt đầu quét học viên nghỉ quá {max_allowed_absent} buổi..."
        )
        if self.df is None or self.df.empty:
            print("[-] Không có dữ liệu học viên để xử lý email.")
            return

        absent_students = self.df[self.df["So_Buoi_Nghi"] >= max_allowed_absent]
        if absent_students.empty:
            print("[+] Tuyệt vời! Không có học viên nào nghỉ quá giới hạn.")
            return

        for index, row in absent_students.iterrows():
            name = row["Ho_Ten"]
            email = row["Email"]
            absent_count = row["So_Buoi_Nghi"]

            subject = f"[CẢNH BÁO CHUYÊN CẦN] Lớp học Coding - Học viên {name}"
            body = (
                f"Chào {name},\n\n"
                f"Hệ thống ghi nhận bạn đã nghỉ {absent_count} buổi học (Giới hạn: {max_allowed_absent} buổi).\n"
                f"Vui lòng liên hệ lại với Trợ giảng lớp để bù bài tập.\n\n"
                f"Trân trọng,\nĐội ngũ Trợ giảng."
            )
            self.send_email(email, subject, body)

    def export_excel_report(self, output_path="bao_cao_can_ho_tro.xlsx"):
        """Tự động phân loại và xuất file báo cáo học viên cần hỗ trợ đặc

        biệt.
        """
        print("\n[!] Đang khởi tạo báo cáo danh sách cần hỗ trợ...")
        if self.df is None or self.df.empty:
            return

        condition = (self.df["So_Buoi_Nghi"] >= 3) | (
            self.df["Diem_Danh_Gia"] < 6.0
        )
        report_df = self.df[condition]

        try:
            report_df.to_excel(output_path, index=False)
            print(f"[+] Đã xuất file báo cáo thành công ra: {output_path}")
        except Exception as e:
            print(f"[-] Lỗi khi xuất báo cáo: {e}")

if __name__ == "__main__":
    bot = ClassroomAutomation("danh_sach_hoc_vien.xlsx")

    if bot.load_data():
        bot.input_students()
        bot.auto_remind_attendance(max_allowed_absent=3)
        bot.export_excel_report("hoc_vien_yeu_kem_can_ho_tro.xlsx")