import sys
import os
import json

from GenExcel import GenExcel
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel,
    QPushButton, QFileDialog, QMessageBox
)

class CSVExcelApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CSV → Excel Generator")
        self.setGeometry(300, 300, 400, 200)

        self.layout = QVBoxLayout()

        self.label_csv = QLabel()
        self.layout.addWidget(self.label_csv)

        self.label_excel = QLabel()
        self.layout.addWidget(self.label_excel)

        self.btn_browse_csv = QPushButton("เลือกไฟล์ CSV")
        self.btn_browse_csv.clicked.connect(self.browse_csv)
        self.layout.addWidget(self.btn_browse_csv)

        self.btn_browse_excel = QPushButton("เลือกไฟล์ Excel")
        self.btn_browse_excel.clicked.connect(self.browse_excel)
        self.layout.addWidget(self.btn_browse_excel)

        self.btn_generate = QPushButton("สร้าง Excel")
        self.btn_generate.clicked.connect(self.generate_excel)
        self.layout.addWidget(self.btn_generate)

        self.setLayout(self.layout)

        self.config_path = "config.json"
        self.config = {}
        self.csv_path = None
        self.excel_path = None

        self.load_and_validate_config()

    def load_and_validate_config(self):
        if not os.path.exists(self.config_path):
            QMessageBox.critical(self, "ไม่พบ config", "ไม่พบ config.json")
            return

        with open(self.config_path, "r", encoding="utf-8") as f:
            self.config = json.load(f)

        # ตรวจสอบ csv_path
        self.csv_path = self.config.get("csv_path", "")
        if not self.csv_path or not os.path.exists(self.csv_path):
            self.csv_path = ""
            self.config["csv_path"] = ""
            self.label_csv.setText("ยังไม่ได้เลือกไฟล์ CSV")
        else:
            self.label_csv.setText(f"เลือกแล้ว: {self.csv_path}")

        # ตรวจสอบ excel_path
        self.excel_path = self.config.get("excel_path", "")
        if not self.excel_path or not os.path.exists(self.excel_path):
            self.excel_path = ""
            self.config["excel_path"] = ""
            self.label_excel.setText("ยังไม่ได้เลือกไฟล์ Excel")
        else:
            self.label_excel.setText(f"เลือกแล้ว: {self.excel_path}")

        # อัปเดต config กลับไป
        with open(self.config_path, "w", encoding="utf-8") as f:
            json.dump(self.config, f, ensure_ascii=False, indent=2)

    def browse_csv(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "เลือกไฟล์ CSV", "", "CSV Files (*.csv)")
        if file_name:
            self.csv_path = file_name
            self.label_csv.setText(f"เลือกแล้ว: {file_name}")
            self.config["csv_path"] = file_name
            self.save_config()

    def browse_excel(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "เลือกไฟล์ XLSX", "", "XLSX Files (*.xlsx)")
        if file_name:
            self.excel_path = file_name
            self.label_excel.setText(f"เลือกแล้ว: {file_name}")
            self.config["excel_path"] = file_name
            self.save_config()

    def save_config(self):
        with open(self.config_path, "w", encoding="utf-8") as f:
            json.dump(self.config, f, ensure_ascii=False, indent=2)

    def generate_excel(self):
        if not self.csv_path or not os.path.exists(self.csv_path):
            QMessageBox.warning(self, "ผิดพลาด", "กรุณาเลือกไฟล์ CSV ที่ถูกต้องก่อน")
            return
        if not self.excel_path or not os.path.exists(self.excel_path):
            QMessageBox.warning(self, "ผิดพลาด", "กรุณาเลือกไฟล์ Excel ที่ถูกต้องก่อน")
            return

        try:
            # ✅ สร้าง instance ของ GenExcel
            generator = GenExcel(self.csv_path, self.excel_path, self.config_path)
            output_file = generator.generateExcel()  # ✅ เรียกใช้งาน

            QMessageBox.information(
                self, "สำเร็จ", f"สร้าง Excel สำเร็จแล้ว:\n{output_file}"
            )

        except Exception as e:
            QMessageBox.critical(self, "เกิดข้อผิดพลาด", str(e))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = CSVExcelApp()
    window.show()
    sys.exit(app.exec())
