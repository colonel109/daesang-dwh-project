import json


class ConfigLoader:
    def __init__(self, config_path):
        with open(config_path, encoding="utf-8") as config_file:
            self.config_data: dict = json.load(config_file)

        self.col_req_dict = None
        self.col_exclusive_dict = None
        self.col_use = None
        self.rename_dict = None
        self.dtype_dict = None
        self.col_optional = None

        self.config_loader()

    def config_loader(self):
        """
        Đọc dữ liệu từ biến config_data và truyền vào các biến lưu trữ
        """
       
        # Trả về dict tên cột chuẩn hoá và tên cột gốc 
        # dạng {"tên cột chuẩn hoá": "tên cột gốc"}
        self.col_req_dict = {
            col_name: data["raw_name"]
            for col_name, data in self.config_data.items()
        }

        # Trả về dict dạng {"tên cột chuẩn hoá": "true/false"}
        self.col_exclusive_dict = {
            col_name: data["exclusive"]
            for col_name, data in self.config_data.items() if data.get("exclusive")
        }
            
        # Trả về set chứa các raw name của các cột, vì trong 1 file sẽ không có 2 cột trùng nhau nên
        # có thể gom tất cả lại
        self.col_use = {
            col for dict_data in self.config_data.values() 
            for col in dict_data["raw_name"]
        }

        # Trả về dict có dạng {"tên cột gốc": "tên cột chuẩn hoá"}
        self.rename_dict = {
            raw_name: new_col_name 
            for new_col_name, data in self.config_data.items()
            for raw_name in data["raw_name"] 
        }

        self.col_optional = {
            col_name: data["optional"]
            for col_name, data in self.config_data.items() if data.get("optional")
        }

        self.dtype_dict = {col: dict_data["dtype"] for col, dict_data in self.config_data.items()}