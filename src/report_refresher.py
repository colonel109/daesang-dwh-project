import win32com.client as win32
import pywintypes 
from pathlib import Path
import time, questionary
from datetime import datetime
from shutil import copy
from questionary import Choice


class ReportRefresher:
    """
    Mở và làm mới dữ liệu trong các file mẫu (template_report) sau đó đổi tên và di chuyển sang thư mục kết quả (refreshed_report)
    """
    
    def __init__(self, report_folder: Path):
        self.template_path = Path(report_folder / "template_report") # Thư mục chứa file mẫu
        self.result_path = Path(report_folder / "refreshed_report") # Thư mục chứa file kết quả sau khi refresh file mẫu
    
    def run_process(self):
        file = self.file_selector()
        if file:
            self.process_file(
                selected_paths = file
            ) 
       
    def file_selector(self):
        """
        Chọn các file template, trả về một list các đường dẫn 
        """
        # Hiển thị tên file đối với title và trả về các path tương ứng
        choices = [
            Choice(title=str(file.stem), value=str(file)) 
            for file in self.template_path.glob("*.xlsx")
        ] 
        
        selected_report = questionary.checkbox(
            "Chọn báo cáo cần refresh dữ liệu",
            choices = choices 
        ).ask()

        if not selected_report:
            return
        
        # Lấy danh sách các file đã tạo trước đó trong thư mục result
        selected_file_name = [Path(file).stem for file in selected_report]
        refreshed_file_name = [file for file in self.result_path.glob("*")]

        # Xoá các file kết quả nếu báo cáo gốc của nó được chọn để làm mới dữ liệu
        if refreshed_file_name:
            for f in selected_file_name:
                for rf in refreshed_file_name:
                    if f in rf.stem:
                        rf.unlink()
        
        return selected_report
    
    def refresh_option_prompt(self):
        """
        Tuỳ chọn giúp tránh phải refresh các view nặng
        """
        promt_pl_data = questionary.select(
            "Bạn có muốn làm mới cả dữ liệu lỗ lãi (nếu có)?",
            choices = ["Không", "Có"],
            default = "Không"
        ).ask()

        if promt_pl_data == "Không":
            refresh_all = False
        elif promt_pl_data == "Có":
            refresh_all = True
        
        return refresh_all

    def process_file(self, selected_paths: list):
        """
        Làm mới dữ liệu, đổi tên file theo ngày, giờ và di chuyển đến thư mục kết quả
        """
        refresh_all = self.refresh_option_prompt()
        
        # Mở excel
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False 
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False
    
        has_error = False # Flag đánh dấu process có lỗi
        file_has_error = [] # Danh sách chứa tên file gặp lỗi
        process_order = 1 
        for path in selected_paths:
            file_path = Path(path)
            file_name = file_path.stem

            if not file_path.is_file():
                print(f"Không tìm thấy đường dẫn {path}")
                continue

            print(f"{process_order}/ Đang làm mới dữ liệu file: {file_name}") 

            # Làm mới dữ liệu
            wb = excel.Workbooks.Open(str(file_path))            
            time.sleep(3)

            try:
                if not refresh_all: 
                    excluded_query = "pl_data"
                    for query in wb.Connections:
                        if not excluded_query in query.Name:
                            print(f"\t{query.Name}")
                            query.Refresh()
                else:
                    print("\tĐang làm mới toàn bộ bảng")
                    wb.RefreshAll()
                    excel.CalculateUntilAsyncQueriesDone()
                    wb.Save()
                    wb.Close()

                    # Tạo file kèm nhãn thời gian sau khi thành công
                    timestamp = datetime.now().strftime("%d.%m.%Y_%H.%M.%S")
                    dst = self.result_path / f"({timestamp}) {file_name}{file_path.suffix}"
                    copy(file_path, dst)

                print("\t\u2705 đã tạo báo cáo thành công")

            except pywintypes.com_error as e:
                print(f"\t\u274C Tạo báo cáo thất bại, đang đóng file:\n{e}")
                wb.Close()
                has_error = True
                file_has_error.append(file_name)
                continue # Chuyển qua xử lí file tiếp theo

            process_order += 1
                           
        if has_error:
            print("Quá trình xử lí file gặp lỗi, các báo cáo gặp lỗi:")

            n=1
            for f in file_has_error:
                print(f"{n}: {f}\n")
                n+=1
        
        excel.Quit()