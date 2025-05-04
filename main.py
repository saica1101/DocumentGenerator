import sys
import os
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QLineEdit, QPushButton, QDateEdit, QFormLayout, QDialog, QStackedWidget, QTableWidget, QTextEdit # type: ignore
from PyQt5.QtCore import QDate, QTime # type: ignore
import shutil
import openpyxl # type: ignore
from openpyxl.styles import Alignment # type: ignore
from openpyxl import load_workbook # type: ignore
import tempfile
import logging
import subprocess
import sqlite3

# ログ設定
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

class DatabaseManager:
    def __init__(self, db_name="documents.db"):
        # データベースファイルのパスを絶対パスで指定
        self.db_path = os.path.abspath(db_name)
        self.conn = None
        self.cursor = None

    def connect(self):
        try:
            # データベースファイルが存在しない場合は新規作成
            self.conn = sqlite3.connect(self.db_path)
            self.cursor = self.conn.cursor()

            # テーブルが存在しない場合は作成
            self.cursor.execute('''
                CREATE TABLE IF NOT EXISTS company_info (
                    id INTEGER PRIMARY KEY,
                    company_name TEXT,
                    postal_code TEXT,
                    address TEXT,
                    address_detail TEXT,
                    phone_number TEXT,
                    contact_person TEXT,
                    account_type TEXT,
                    bank_branch TEXT,
                    account_number TEXT,
                    account_name TEXT
                )
            ''')
            self.conn.commit()
            logging.info("Database connected successfully.")
        except sqlite3.Error as e:
            logging.error(f"Error connecting to database: {e}")
            raise

    def close(self):
        try:
            if self.conn:
                self.cursor.close()
                self.conn.close()
                self.conn = None
                logging.info("Database connection closed.")
        except sqlite3.Error as e:
            logging.error(f"Error closing database connection: {e}")
            raise

    def get_company_info(self):
        """自社情報を取得する"""
        try:
            self.connect()
            self.cursor.execute("SELECT * FROM company_info WHERE id = 1")
            row = self.cursor.fetchone()
            if row:
                # カラム名と値をペアにした辞書として返す
                columns = [column[0] for column in self.cursor.description]
                return dict(zip(columns, row))
            else:
                logging.warning("Company info not found.")
                return None  # None を返すように修正 
        except sqlite3.Error as e:
            logging.error(f"Error getting company info: {e}")
            raise
        finally:
            self.close()

    def update_company_info(self, info):
        """自社情報を更新する"""
        try:
            self.connect()
            self.cursor.execute("SELECT * FROM company_info WHERE id = 1")  # id = 1 のレコードを検索 
            existing_data = self.cursor.fetchone()

            if existing_data:
                self.cursor.execute('''
                    UPDATE company_info SET
                        company_name = ?, postal_code = ?, address = ?, address_detail = ?,
                        phone_number = ?, contact_person = ?, account_type = ?, bank_branch = ?,
                        account_number = ?, account_name = ?
                    WHERE id = 1
                ''', (info.get("company_name"), info.get("postal_code"), info.get("address"),
                      info.get("address_detail"), info.get("phone_number"), info.get("contact_person"),
                      info.get("account_type"), info.get("bank_branch"), info.get("account_number"),
                      info.get("account_name")))
                self.conn.commit()
                logging.info("Company info updated.")
            else:
                self.cursor.execute('''
                    INSERT INTO company_info (
                        id, company_name, postal_code, address, address_detail,
                        phone_number, contact_person, account_type, bank_branch,
                        account_number, account_name
                    ) VALUES (1, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (info.get("company_name"), info.get("postal_code"), info.get("address"),
                      info.get("address_detail"), info.get("phone_number"), info.get("contact_person"),
                      info.get("account_type"), info.get("bank_branch"), info.get("account_number"),
                      info.get("account_name")))
                self.conn.commit()
                logging.info("Company info inserted.")
        except sqlite3.Error as e:
            logging.error(f"Error updating company info: {e}")
            raise
        finally:
            self.close()

    def delete_company_info(self):
        """自社情報を削除する (通常は使用しない)"""
        try:
            self.connect()
            self.cursor.execute("DELETE FROM company_info WHERE id = 1")
            self.conn.commit()
            logging.warning("Company info deleted.")
        except sqlite3.Error as e:
            logging.error(f"Error deleting company info: {e}")
            raise
        finally:
            self.close()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("書類作成アプリケーション")
        self.setGeometry(100, 100, 640, 560)
        self.db_manager = DatabaseManager()
        self.stack = QStackedWidget()
        self.setCentralWidget(self.stack)
        self.create_top_menu()
        self.create_estimate_screen()
        self.create_invoice_screen()
        self.create_receipt_screen()
        self.remarks_text = ""

    def create_top_menu(self):
        top_menu = QWidget()
        layout = QVBoxLayout()
        # ボタン作成
        estimate_button = QPushButton("見積書作成")
        invoice_button = QPushButton("請求書作成")
        receipt_button = QPushButton("領収書作成")
        settings_button = QPushButton("自社情報変更")
        exit_button = QPushButton("終了")
        # ボタンクリック時の処理
        estimate_button.clicked.connect(lambda: self.stack.setCurrentWidget(self.estimate_screen))
        invoice_button.clicked.connect(lambda: self.stack.setCurrentWidget(self.invoice_screen))
        receipt_button.clicked.connect(lambda: self.stack.setCurrentWidget(self.receipt_screen))
        settings_button.clicked.connect(self.open_settings_dialog)
        exit_button.clicked.connect(self.exit_process)
        # レイアウトにボタン追加
        layout.addWidget(estimate_button)
        layout.addWidget(invoice_button)
        layout.addWidget(receipt_button)
        layout.addWidget(settings_button)
        layout.addWidget(exit_button)
        top_menu.setLayout(layout)
        self.stack.addWidget(top_menu)

    def create_document_screen(self, document_type):
        screen = QWidget()
        layout = QVBoxLayout()
        input_layout = QFormLayout()
        company_name_field = QLineEdit()
        layout.addLayout(input_layout)
        remarks_button = QPushButton("備考を入力")
        remarks_button.clicked.connect(self.open_remarks_dialog)
        layout.addWidget(remarks_button)
        table = QTableWidget()
        table.setColumnCount(6)  # 列数を修正
        table.setHorizontalHeaderLabels(["摘要", "数量", "単位", "単価", "値引", "税率 (%)"])  # ヘッダーラベルを修正
        layout.addWidget(table)
        add_row_button = QPushButton("行を追加")
        add_row_button.clicked.connect(lambda: self.add_table_row(table))
        layout.addWidget(add_row_button)
        delete_row_button = QPushButton("選択した行を削除")
        delete_row_button.clicked.connect(lambda: self.delete_table_row(table))
        layout.addWidget(delete_row_button)
        generate_button = QPushButton("PDF生成")
        layout.addWidget(generate_button)
        back_button = QPushButton("戻る")
        back_button.clicked.connect(lambda: self.stack.setCurrentIndex(0))
        layout.addWidget(back_button)
        screen.setLayout(layout)
        return screen, input_layout, company_name_field, table, generate_button

    def create_estimate_screen(self):
        self.estimate_screen, estimate_input_layout, self.estimate_company_name, self.estimate_table, generate_button = self.create_document_screen("見積書")
        self.estimate_subject = QLineEdit()
        self.estimate_expiry_date = QDateEdit()
        self.estimate_expiry_date.setCalendarPopup(True)
        self.estimate_expiry_date.setDate(QDate.currentDate())
        self.estimate_delivery_date = QLineEdit()
        self.estimate_delivery_place = QLineEdit()
        self.estimate_transaction_method = QLineEdit()
        estimate_input_layout.addRow("取引先企業名:", self.estimate_company_name)
        estimate_input_layout.addRow("件名:", self.estimate_subject)
        estimate_input_layout.addRow("有効期限:", self.estimate_expiry_date)
        estimate_input_layout.addRow("納入期日:", self.estimate_delivery_date)
        estimate_input_layout.addRow("納入場所:", self.estimate_delivery_place)
        estimate_input_layout.addRow("取引方法:", self.estimate_transaction_method)
        generate_button.clicked.connect(lambda: self.generate_document("見積書"))
        self.stack.addWidget(self.estimate_screen)

    def create_invoice_screen(self):
        self.invoice_screen, invoice_input_layout, self.invoice_company_name, self.invoice_table, generate_button = self.create_document_screen("請求書")
        self.invoice_subject = QLineEdit()
        self.invoice_expiry_date = QDateEdit()
        self.invoice_expiry_date.setCalendarPopup(True)
        self.invoice_expiry_date.setDate(QDate.currentDate())
        self.invoice_delivery_date = QLineEdit()
        self.invoice_delivery_place = QLineEdit()
        self.invoice_transaction_method = QLineEdit()
        invoice_input_layout.addRow("取引先企業名:", self.invoice_company_name)
        invoice_input_layout.addRow("件名:", self.invoice_subject)
        invoice_input_layout.addRow("有効期限:", self.invoice_expiry_date)
        invoice_input_layout.addRow("納入期日:", self.invoice_delivery_date)
        invoice_input_layout.addRow("納入場所:", self.invoice_delivery_place)
        invoice_input_layout.addRow("取引方法:", self.invoice_transaction_method)
        generate_button.clicked.connect(lambda: self.generate_document("請求書"))
        self.stack.addWidget(self.invoice_screen)

    def create_receipt_screen(self):
        self.receipt_screen, receipt_input_layout, self.receipt_company_name, self.receipt_table, generate_button = self.create_document_screen("領収書")
        self.receipt_period_duration = QDateEdit()
        self.receipt_period_duration.setCalendarPopup(True)
        self.receipt_period_duration.setDate(QDate.currentDate())
        self.receipt_delivery_place = QLineEdit()
        self.receipt_transaction_method = QLineEdit()
        receipt_input_layout.addRow("取引先企業名:", self.receipt_company_name)
        receipt_input_layout.addRow("法定保存期限:", self.receipt_period_duration)
        receipt_input_layout.addRow("納入場所:", self.receipt_delivery_place)
        receipt_input_layout.addRow("取引方法:", self.receipt_transaction_method)
        generate_button.clicked.connect(lambda: self.generate_document("領収書"))
        self.stack.addWidget(self.receipt_screen)

    def open_remarks_dialog(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("備考の入力")
        dialog_layout = QVBoxLayout()
        remarks_field = QTextEdit()
        remarks_field.setPlainText(self.remarks_text)
        dialog_layout.addWidget(remarks_field)
        save_button = QPushButton("保存")
        save_button.clicked.connect(lambda: self.save_remarks(dialog, remarks_field))
        dialog_layout.addWidget(save_button)
        dialog.setLayout(dialog_layout)
        dialog.exec_()

    def save_remarks(self, dialog, remarks_field):
        self.remarks_text = remarks_field.toPlainText()
        dialog.accept()

    def add_table_row(self, table):
        if isinstance(table, QTableWidget):
            row_position = table.rowCount()
            table.insertRow(row_position)

    def delete_table_row(self, table):
        selected_rows = table.selectionModel().selectedRows()
        for row in sorted(selected_rows, reverse=True):
            table.removeRow(row.row())

    def generate_document(self, document_type="見積書"):
        # テーブルと入力フィールドの取得
        if document_type == "見積書":
            table = self.estimate_table
            company_name_field = self.estimate_company_name
            subject_field = self.estimate_subject
            expiry_date_field = self.estimate_expiry_date
            delivery_date_field = self.estimate_delivery_date
            delivery_place_field = self.estimate_delivery_place
            transaction_method_field = self.estimate_transaction_method
        elif document_type == "請求書":
            table = self.invoice_table
            company_name_field = self.invoice_company_name
            subject_field = self.invoice_subject
            expiry_date_field = self.invoice_expiry_date
            delivery_date_field = self.invoice_delivery_date
            delivery_place_field = self.invoice_delivery_place
            transaction_method_field = self.invoice_transaction_method
        elif document_type == "領収書":
            table = self.receipt_table
            company_name_field = self.receipt_company_name
            period_duration_field = self.receipt_period_duration
            delivery_place_field = self.receipt_delivery_place
            transaction_method_field = self.receipt_transaction_method
        else:
            raise ValueError(f"Unknown document type: {document_type}")

        template_path = os.path.join(os.path.dirname(__file__), "Templates", f"{document_type}_テンプレート.xlsx")
        temp_dir = tempfile.mkdtemp()
        temp_file = os.path.join(temp_dir, f"{document_type}.xlsx")
        shutil.copy(template_path, temp_file)
        workbook_openpyxl = load_workbook(temp_file)
        sheet = workbook_openpyxl.active
        company_info = self.db_manager.get_company_info()

        # 自社情報
        sheet["F5"].value = company_info.get("company_name", "")
        sheet["G6"].value = company_info.get("postal_code", "")
        sheet["G7"].value = company_info.get("address", "")
        sheet["G8"].value = company_info.get("address_detail", "")
        sheet["G9"].value = company_info.get("phone_number", "")
        sheet["G10"].value = company_info.get("contact_person", "")
        if document_type == "請求書":
            sheet["G11"].value = company_info.get("account_type", "")
            sheet["G12"].value = company_info.get("bank_branch", "")
            sheet["G13"].value = company_info.get("account_number", "")
            sheet["G14"].value = company_info.get("account_name", "")

        # 書類情報
        sheet["A2"].value = company_name_field.text()
        sheet["A2"].alignment = Alignment(vertical="bottom", horizontal="center")
        if document_type in ["見積書", "請求書"]:
            sheet["B5"].value = company_name_field.text()
            sheet["B5"].alignment = Alignment(vertical="center")
            sheet["B6"].value = QDate.currentDate().toString("yyyy/MM/dd")
            sheet["B6"].alignment = Alignment(vertical="center")
            sheet["B7"].value = expiry_date_field.date().toString("yyyy/MM/dd")
            sheet["B8"].value = delivery_date_field.text()
            sheet["B8"].alignment = Alignment(vertical="center")
            sheet["B9"].value = delivery_place_field.text()
            sheet["B9"].alignment = Alignment(vertical="center")
            sheet["B10"].value = transaction_method_field.text()
            sheet["B10"].alignment = Alignment(vertical="center")
        elif document_type == "領収書":
            sheet["B5"].value = QDate.currentDate().toString("yyyy/MM/dd")
            sheet["B5"].alignment = Alignment(vertical="center")
            sheet["B6"].value = QDate.currentDate().toString("yyyy/MM/dd")
            sheet["B6"].alignment = Alignment(vertical="center")
            sheet["B7"].value = period_duration_field.date().toString("yyyy/MM/dd")
            sheet["B7"].alignment = Alignment(vertical="center")
            sheet["B8"].value = delivery_place_field.text()
            sheet["B8"].alignment = Alignment(vertical="center")
            sheet["B9"].value = transaction_method_field.text()
            sheet["B9"].alignment = Alignment(vertical="center")

        docRow = 16
        sumTax = 26
        taxRow = 28
        remarksPos = 33
        if document_type == "請求書":
            docRow += 1
            sumTax += 1
            taxRow += 1
            remarksPos += 1

        # テーブルデータ
        total_excluding_tax = 0
        total_tax = 0
        tax_10_total = 0
        tax_8_total = 0
        tax_0_total = 0
        for row in range(table.rowCount()):
            summary = table.item(row, 0).text() if table.item(row, 0) else ""
            quantity = table.item(row, 1).text() if table.item(row, 1) else "0"
            unit = table.item(row, 2).text() if table.item(row, 2) else ""
            unit_price = table.item(row, 3).text() if table.item(row, 3) else "0"
            discount = table.item(row, 4).text() if table.item(row, 4) else "0"
            tax_rate = float(table.item(row, 5).text()) / 100 if table.item(row, 5) else 0

            sheet[f"A{docRow + row}"].value = summary
            sheet[f"D{docRow + row}"].value = quantity
            sheet[f"D{docRow + row}"].alignment = Alignment(vertical="center", horizontal="right")
            sheet[f"E{docRow + row}"].value = unit
            sheet[f"F{docRow + row}"].value = unit_price
            sheet[f"F{docRow + row}"].alignment = Alignment(vertical="center", horizontal="right")
            sheet[f"G{docRow + row}"].value = discount
            sheet[f"G{docRow + row}"].alignment = Alignment(vertical="center", horizontal="right")
            sheet[f"H{docRow + row}"].value = tax_rate
            sheet[f"H{docRow + row}"].number_format = '0%'
            sheet[f"H{docRow + row}"].alignment = Alignment(vertical="center")

            # 小計計算
            subtotal = (float(quantity) * float(unit_price)) - float(discount or 0)
            sheet[f"I{docRow + row}"].value = subtotal
            sheet[f"I{docRow + row}"].alignment = openpyxl.styles.alignment.Alignment(vertical="center")

        # 合計金額の計算と埋め込み
        total_excluding_tax = sum(float(sheet[f"I{docRow + row}"].value or 0) for row in range(table.rowCount()))
        total_tax = sum(float(sheet[f"I{docRow + row}"].value or 0) * float(sheet[f"H{docRow + row}"].value or 0) for row in range(table.rowCount()))
        total_including_tax = total_excluding_tax + total_tax

        sheet[f"I{sumTax}"].value = total_excluding_tax  # 税抜合計
        sheet[f"I{sumTax+1}"].value = total_tax  # 消費税合計
        sheet[f"I{sumTax+2}"].value = total_including_tax  # 総合計

        # 税率ごとの合計金額を計算して埋め込み
        tax_10_total = sum(sheet[f"I{docRow + row}"].value or 0 for row in range(table.rowCount()) if sheet[f"H{docRow + row}"].value == 0.1)
        tax_10_tax = tax_10_total * 0.1
        tax_8_total = sum(sheet[f"I{docRow + row}"].value or 0 for row in range(table.rowCount()) if sheet[f"H{docRow + row}"].value == 0.08)
        tax_8_tax = tax_8_total * 0.08
        tax_0_total = sum(sheet[f"I{docRow + row}"].value or 0 for row in range(table.rowCount()) if sheet[f"H{docRow + row}"].value in [0, ""])

        sheet[f"B{taxRow}"].value = tax_10_total  # 税率10%商品の税抜合計
        sheet[f"C{taxRow}"].value = tax_10_tax  # 税率10%商品の消費税額合計
        sheet[f"B{taxRow + 1}"].value = tax_8_total  # 税率8%商品の税抜合計
        sheet[f"C{taxRow + 1}"].value = tax_8_tax  # 税率8%商品の消費税額合計
        sheet[f"B{taxRow + 2}"].value = tax_0_total  # 税率0%商品の税抜合計
        sheet[f"C{taxRow + 2}"].value = 0  # 税率0%商品の消費税額合計

        # 備考欄の内容を埋め込む
        if self.remarks_text:
            sheet[f"A{remarksPos}"].value = self.remarks_text
            sheet[f"A{remarksPos}"].alignment = openpyxl.styles.alignment.Alignment(vertical="top", horizontal="left")

        # 保存
        workbook_openpyxl.save(temp_file)
        print(f"Excelファイルが正常に保存されました: {temp_file}")

        # LibreOfficeのパスを明示的に指定
        libreoffice_path = r"C:\Program Files\LibreOffice\program\soffice.exe"

        # LibreOfficeを使用して指定シートのみをPDFに変換
        def convert_sheet_to_pdf_with_libreoffice(input_file, output_dir, sheet_name):
            try:
                # LibreOfficeで指定シートをPDFに変換
                subprocess.run([
                    libreoffice_path, "--headless", "--convert-to", "pdf:calc_pdf_Export", input_file,
                    "--outdir", output_dir, f"--infilter=calc:sheet={sheet_name}"
                ], check=True)
                logging.info(f"Successfully converted sheet '{sheet_name}' in {input_file} to PDF using LibreOffice.")
            except subprocess.CalledProcessError as e:
                logging.error(f"Error converting sheet '{sheet_name}' in {input_file} to PDF: {e}")
                raise
            except FileNotFoundError:
                logging.error("LibreOffice executable not found. Please check the path.")
                raise

        # PDF保存先
        pdf_output_dir = os.path.dirname(temp_file)
        convert_sheet_to_pdf_with_libreoffice(temp_file, pdf_output_dir, document_type)

        # PDFを指定のフォルダに移動
        base_dir = os.path.dirname(os.path.abspath(__file__))
        sub_dir = "見積書" if document_type == "見積書" else ("請求書" if document_type == "請求書" else "領収書")
        company_dir = os.path.join(base_dir, sub_dir, company_name_field.text(), QDate.currentDate().toString('yyyy'), QDate.currentDate().toString('MM'))
        os.makedirs(company_dir, exist_ok=True)
        final_pdf_path = os.path.join(company_dir, f"{QDate.currentDate().toString('yyyyMMdd')}{QTime.currentTime().toString('hhmm')}.pdf")
        shutil.move(temp_file.replace(".xlsx", ".pdf"), final_pdf_path)

        # 一時ファイルとフォルダを削除
        shutil.rmtree(temp_dir)
        
        # PDF生成後に変数を初期化
        company_name_field.clear()
        if not document_type == "領収書":
            subject_field.clear()
            expiry_date_field.setDate(QDate.currentDate())
            delivery_date_field.clear()
        delivery_place_field.clear()
        transaction_method_field.clear()
        self.remarks_text = ""

        # テーブルウィジェットの内容をクリア
        self.estimate_table.setRowCount(0)
        self.invoice_table.setRowCount(0)
        self.receipt_table.setRowCount(0)
        table.setRowCount(0)
        
        print(f"PDFが正常に保存されました: {final_pdf_path}")

    def add_settings_button(self):
        # 自社情報変更ボタン
        settings_button = QPushButton("自社情報を変更")
        settings_button.clicked.connect(self.open_settings_dialog)
        self.layout().addWidget(settings_button)

    def update_company_info(self):
        """データベースから自社情報を再読み込みしてインスタンス変数を更新する"""
        self.company_info = self.db_manager.get_company_info()
        if self.company_info:
            logging.debug("自社情報が更新されました:")
            for key, value in self.company_info.items():
                logging.debug(f"{key}: {value}")
        else:
            logging.debug("自社情報の更新に失敗しました: データベースから情報を取得できませんでした。")

    def open_settings_dialog(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("自社情報設定")
        layout = QFormLayout(dialog)

        company_name_field = QLineEdit()  # 変数として保持
        layout.addRow("会社名:", company_name_field)

        address_line1 = QLineEdit()      # 変数として保持
        layout.addRow("郵便番号:", address_line1)

        address_line2 = QLineEdit()      # 変数として保持
        layout.addRow("住所1:", address_line2)

        address_line3 = QLineEdit()      # 変数として保持
        layout.addRow("住所2:", address_line3)

        phone_number_field = QLineEdit()       # 変数として保持
        layout.addRow("電話番号:", phone_number_field)

        contact_person_field = QLineEdit()   # 変数として保持
        layout.addRow("担当者:", contact_person_field)

        bank_account_type_field = QLineEdit() # 変数として保持
        layout.addRow("口座種別:", bank_account_type_field)

        bank_name_branch_field = QLineEdit()  # 変数として保持
        layout.addRow("銀行名・支店名:", bank_name_branch_field)

        account_number_field = QLineEdit()   # 変数として保持
        layout.addRow("口座番号:", account_number_field)

        account_holder_name_field = QLineEdit() # 変数として保持
        layout.addRow("口座名義:", account_holder_name_field)

        # 既存の自社情報をフィールドに設定
        company_info = self.db_manager.get_company_info()
        if company_info:
            company_name_field.setText(company_info.get("company_name", ""))
            address_line1.setText(company_info.get("postal_code", ""))
            address_line2.setText(company_info.get("address", ""))
            address_line3.setText(company_info.get("address_detail", ""))
            phone_number_field.setText(company_info.get("phone_number", ""))
            contact_person_field.setText(company_info.get("contact_person", ""))
            bank_account_type_field.setText(company_info.get("account_type", ""))
            bank_name_branch_field.setText(company_info.get("bank_branch", ""))
            account_number_field.setText(company_info.get("account_number", ""))
            account_holder_name_field.setText(company_info.get("account_name", ""))

        save_button = QPushButton("保存")
        save_button.clicked.connect(lambda: self.save_company_info(
            dialog,
            company_name_field,
            address_line1,
            address_line2,
            address_line3,
            phone_number_field,
            contact_person_field,
            bank_account_type_field,
            bank_name_branch_field,
            account_number_field,
            account_holder_name_field
        ))
        layout.addWidget(save_button)

        dialog.exec_()

    def save_company_info(
        self,
        dialog,
        company_name_field,
        address_line1,
        address_line2,
        address_line3,
        phone_number_field,
        contact_person_field,
        bank_account_type_field,
        bank_name_branch_field,
        account_number_field,
        account_holder_name_field
    ):
        # 自社情報を保存する処理
        new_info = {
            "company_name": company_name_field.text(),
            "postal_code": address_line1.text(),
            "address": address_line2.text(),
            "address_detail": address_line3.text(),
            "phone_number": phone_number_field.text(),
            "contact_person": contact_person_field.text(),
            "account_type": bank_account_type_field.text(),
            "bank_branch": bank_name_branch_field.text(),
            "account_number": account_number_field.text(),
            "account_name": account_holder_name_field.text()
        }
        self.db_manager.update_company_info(new_info)
        self.update_company_info()
        dialog.accept()

    def closeEvent(self, event):
        # アプリ終了時にデータベース接続を閉じる
        self.db_manager.close()
        event.accept()

    def exit_process(self):
        # アプリケーションを終了する処理
        self.db_manager.close()
        self.close()

class DocumentApp:
    def __init__(self, db_manager):
        self.db_manager = db_manager

    def open_settings_dialog(self):
        pass

if __name__ == "__main__":
    db_manager = DatabaseManager()
    company_info = db_manager.get_company_info()
    if company_info:
        print("自社情報:", company_info)
    else:
        print("自社情報が見つかりませんでした。")

    app = QApplication(sys.argv)
    window = MainWindow()
    if not db_manager.get_company_info():
        window.open_settings_dialog()  # 自社情報がなければ設定ウィンドウを開く
    window.show()
    sys.exit(app.exec_())

def save_company_info():
    db = DatabaseManager()

    # 自社情報の入力例
    company_info = {
        "company_name": "株式会社サンプル",
        "postal_code": "123-4567",
        "address": "東京都新宿区",
        "address_detail": "1-2-3 サンプルビル",
        "phone_number": "03-1234-5678",
        "contact_person": "山田 太郎",
        "account_type": "普通",
        "bank_branch": "新宿支店",
        "account_number": "1234567",
        "account_name": "カ）サンプル"
    }

    # データベースに保存
    db.add_company_info(company_info)
    print("自社情報を保存しました。")

    # 保存した情報を取得して表示
    saved_info = db.get_company_info()
    print("保存された自社情報:", saved_info)

class DatabaseManager:
    def __init__(self, db_name="documents.db"):
        self.db_name = db_name
        self.conn = None

    def connect(self):
        """データベースに接続し、カーソルを作成する"""
        try:
            self.conn = sqlite3.connect(self.db_name)
            self.cursor = self.conn.cursor()
            logging.info("Database connection established.")
        except sqlite3.Error as e:
            logging.error(f"Database connection error: {e}")
            raise

    def close(self):
        """データベース接続を閉じる"""
        if self.conn:
            self.conn.close()
            logging.info("Database connection closed.")

    def create_table(self):
        """テーブルを作成する"""
        try:
            self.connect()
            self.cursor.execute('''
                CREATE TABLE IF NOT EXISTS company_info (
                    id INTEGER PRIMARY KEY,
                    company_name TEXT,
                    postal_code TEXT,
                    address TEXT,
                    address_detail TEXT,
                    phone_number TEXT,
                    contact_person TEXT,
                    account_type TEXT,
                    bank_branch TEXT,
                    account_number TEXT,
                    account_name TEXT
                )
            ''')
            self.conn.commit()  # テーブル作成をコミット 
            logging.info("Table 'company_info' created (if not exists).")
        except sqlite3.Error as e:
            logging.error(f"Table creation error: {e}")
            raise
        finally:
            self.close()

    def add_company_info(self, info):
        """自社情報を追加する"""
        try:
            self.connect()
            self.cursor.execute('''
                INSERT INTO company_info (company_name, postal_code, address, address_detail,
                                        phone_number, contact_person, account_type, bank_branch,
                                        account_number, account_name)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (info.get("company_name"), info.get("postal_code"), info.get("address"),
                  info.get("address_detail"), info.get("phone_number"), info.get("contact_person"),
                  info.get("account_type"), info.get("bank_branch"), info.get("account_number"),
                  info.get("account_name")))
            self.conn.commit()  # データ追加をコミット 
            logging.info("Company info added.")
        except sqlite3.Error as e:
            logging.error(f"Error adding company info: {e}")
            raise
        finally:
            self.close()

    def get_company_info(self):
        """自社情報を取得する"""
        try:
            self.connect()
            self.cursor.execute("SELECT * FROM company_info LIMIT 1")
            row = self.cursor.fetchone()
            if row:
                columns = [column[0] for column in self.cursor.description]
                return dict(zip(columns, row))
            return None
        except sqlite3.Error as e:
            logging.error(f"Error getting company info: {e}")
            raise
        finally:
            self.close()

    def update_company_info(self, info):
        """自社情報を更新する"""
        try:
            self.connect()
            self.cursor.execute('''
                UPDATE company_info SET
                    company_name = ?, postal_code = ?, address = ?, address_detail = ?,
                    phone_number = ?, contact_person = ?, account_type = ?, bank_branch = ?,
                    account_number = ?, account_name = ?
                WHERE id = 1
            ''', (info.get("company_name"), info.get("postal_code"), info.get("address"),
                  info.get("address_detail"), info.get("phone_number"), info.get("contact_person"),
                  info.get("account_type"), info.get("bank_branch"), info.get("account_number"),
                  info.get("account_name")))
            self.conn.commit()  # データ更新をコミット 
            logging.info("Company info updated.")
        except sqlite3.Error as e:
            logging.error(f"Error updating company info: {e}")
            raise
        finally:
            self.close()

    def delete_company_info(self):
        """自社情報を削除する (通常は使用しない)"""
        try:
            self.connect()
            self.cursor.execute("DELETE FROM company_info WHERE id = 1")
            self.conn.commit()  # データ削除をコミット 
            logging.warning("Company info deleted.")
        except sqlite3.Error as e:
            logging.error(f"Error deleting company info: {e}")
            raise
        finally:
            self.close()
