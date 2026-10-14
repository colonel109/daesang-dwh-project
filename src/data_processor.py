import pandas as pd
from pathlib import Path
from sqlalchemy import text
import re

from src.helper.config_loader import ConfigLoader


class SalesDataProcessor:
    """
    Hàm này chịu trách nhiệm đọc và lấy dữ liệu từ các một list các path
    """

    def __init__(self, engine):
        self.engine = engine
        self.config_data = ConfigLoader(r"configs/sales_data_config.json")
        
    def read_excel(self, file_path_list: list):
        """
        Đọc và xử lí dữ liệu của các file trong list các path được truyền vào 
        """

        completed_dfs = []
        has_error = False # Đánh dấu batch này có file bị lỗi không
        not_mapped_columns = [] # Lưu trữ danh sách các cột chưa được map

        if file_path_list:
            for file in file_path_list:
                # Đọc file
                file_name = file.stem
                suffix = Path(file).suffix

                file_type = self.check_file_type(file_name)
                if not file_type:
                    print(f"Không thể xác định phân loại cho tệp: {file}, huỷ chương trình...")
                    return
                
                # Đọc file excel
                print(f"Đang xử lí file {file_name}")
                if suffix == ".xlsx" and file_type:
                    try:
                        has_missing_col = False # Đánh dấu file có thiếu cột không
                        raw_col_list = pd.read_excel(file, nrows=0).columns.tolist() # Lấy tên cột có trong file gốc

                        for col_name, raw_col_name_list in self.config_data.col_req_dict.items():                           
                            # Trường hợp thiếu cột: cột bắt buộc không nằm trong cả hai loại file order / ar_invoice
                            col_exclusive_compare = self.config_data.col_exclusive_dict.get(col_name) # Trả về file_type
                            if col_exclusive_compare and col_exclusive_compare != file_type: # Nếu cột khác với file_type thì có thể bỏ qua 
                                continue
                            
                            # Trường hợp không tìm thấy cột nào được map với tên cột trong file gốc
                            can_skip = self.config_data.col_optional.get(col_name)
                            if not any(col in raw_col_name_list for col in raw_col_list) and not can_skip:
                                not_mapped_info = {
                                    "Tên file ": file.stem,
                                    "Tên cột chưa map": col_name
                                }
                                not_mapped_columns.append(not_mapped_info)
                                has_missing_col = True
                                
                        if has_missing_col:
                            has_error = True # Trả về lỗi chung, batch này sẽ không được import vào database
                            continue

                        print("Đã map đủ tất cả các cột bắt buộc!")
                        df: pd.DataFrame = pd.read_excel(
                            file,
                            usecols=[col for col in self.config_data.col_use if col in raw_col_list], 
                            dtype="string", 
                            engine="openpyxl"
                        )

                        # Đổi tên cột
                        df = df.rename(columns=self.config_data.rename_dict)

                        # Drop dòng trống
                        df = df.dropna(subset=['business_partner_code'])

                        # Chuyển định dạng
                        for col, dtype in self.config_data.dtype_dict.items():
                            if col in df.columns:
                                if dtype == "numeric":
                                    df[col] = pd.to_numeric(df[col], errors="coerce")

                                if dtype == "date":
                                    df[col] = pd.to_datetime(df[col], errors="coerce", format="%d.%m.%Y")

                        if not df.empty:
                            df["source_file"] = file_name
                            df["scenario_name"] = file_type
                            completed_dfs.append(df)
                    except Exception as e:
                        print(e)
        
        if has_missing_col or has_error: 
            result = {"Cột còn thiếu": not_mapped_columns}
            return result, has_error
            
        result = pd.concat(completed_dfs, ignore_index=True)
        return result, has_error
    
    def check_file_type(self, file_name):
        file_name_lower = file_name.lower()
        # Value của dict này cần quy định chính xác với value trong configs/sales_data_config.json
        dict_data= {
            "ar_invoice": ["ar", "invoice"],
            "order": ["order", "or"] 
        }

        for file_type, kw_list in dict_data.items():
            if any(kw in file_name_lower for kw in kw_list):
                return file_type
        return None
    
    def copy_to_main_table(self, target_table):
        """
        Lấy dữ liệu từ bảng staging trong database và import vào bảng đích
        """
        
        query = text(
            f"""
            INSERT INTO main.{target_table} (
                business_partner_key, product_key, source_product_key,
                document_type_key, sell_type_key,
                cost_center_key, profit_center_key,
                warehouse_key, return_type_key,
                uom_key, tax_key, scenario_key, posting_date, due_date,
                demand_date, final_delivery_date, document_no, tax_number,
                weight, uom_per_quantity, uom_quantity,
                quantity, quantity_packaging_uom, box,
                base_price, discount_percent, discounted_price,
                discounted_price_tax, sales_amount, sales_amount_fc,
                tax_amount, tax_amount_fc, sales_amount_tax,
                sales_amount_fc_tax, return_quantity, remark, order_type_key, location_key, business_place_key
            )
            SELECT
                business_partner_key,
                COALESCE(p.product_key, 0) AS product_key,
                COALESCE(pm.product_key, 0) AS source_product_key,
                COALESCE(dts.document_type_key, 0) AS document_type_key,
                COALESCE(sts.sell_type_key, 0) AS sell_type_key,
                COALESCE(ccs.cost_center_key, 0) AS cost_center_key,
                COALESCE(pcs.profit_center_key, 0) AS profit_center_key,
                COALESCE(ws.warehouse_key, 0) AS warehouse_key,
                COALESCE(rts.return_type_key, 0) AS return_type_key,
                COALESCE(us.uom_key, 0) AS uom_key,
                COALESCE(ts.tax_key, 0) AS tax_key,
                s.scenario_key,
                t.posting_date, t.due_date,
                t.demand_date, t.final_delivery_date,
                t.document_no, t.tax_number,
                t.weight, t.uom_per_quantity,
                t.uom_quantity, t.quantity,
                t.quantity_packaging_uom, t.box,
                t.base_price, t.discount_percent,
                t.discounted_price, t.discounted_price_tax,
                t.sales_amount, t.sales_amount_fc,
                t.tax_amount, t.tax_amount_fc,
                t.sales_amount_tax, t.sales_amount_fc_tax,
                t.return_quantity,
                t.remark,
                COALESCE(ots.order_type_key, 0) AS order_type_key,
                COALESCE(st.location_key, 0) AS location_key,
                COALESCE(bps.business_place_key, 0) AS business_place_key
            FROM staging.transactions_staging t
            JOIN main.business_partner bp 
                ON t.business_partner_code = bp.business_partner_code 
            AND (t.posting_date >= bp.valid_from AND (t.posting_date <= bp.valid_to OR bp.valid_to IS NULL))
            JOIN main.business_partner_group bpg ON bp.bp_group_key = bpg.bp_group_key
            JOIN main.scenario s ON t.scenario_name = s.scenario_name
            LEFT JOIN main.products p ON t.product_code = p.product_code 
            LEFT JOIN main.products pm ON t.source_product_code = pm.product_code
            LEFT JOIN main.store_location st ON TRIM(t.location_code) = st.location_code AND st.bp_group_key = bpg.bp_group_key
            LEFT JOIN staging.document_type_staging dts ON t.document_type_name = dts.raw_value
            LEFT JOIN staging.sell_type_staging sts ON t.sell_type_name = sts.raw_value
            LEFT JOIN staging.cost_center_staging ccs ON t.cost_center_code = ccs.raw_value
            LEFT JOIN staging.profit_center_staging pcs ON t.profit_center_code = pcs.raw_value
            LEFT JOIN staging.warehouse_staging ws ON t.warehouse_code = ws.raw_value
            LEFT JOIN staging.return_type_staging rts ON t.return_type_code = rts.raw_value
            LEFT JOIN staging.uom_staging us ON t.uom_name = us.raw_value
            LEFT JOIN staging.tax_staging ts ON t.tax_code = ts.raw_value
            LEFT JOIN staging.business_place_staging bps ON t.business_place_name = bps.raw_value
            LEFT JOIN staging.order_type_staging ots ON t.order_type_name = ots.raw_value
            """
            )

        with self.engine.begin() as conn:
            conn.execute(query)


class ProfitAndLossProcessor:
    def __init__(self, engine):
        self.engine = engine
        self.config_data = ConfigLoader(r"configs/pl_data_config.json")

    def read_excel(self, file_path_list):

        completed_dfs = []
        has_error = False # Đánh dấu batch này có file bị lỗi không
        not_mapped_columns = [] # Lưu trữ danh sách các cột chưa được map

        if file_path_list:
            for file in file_path_list:
                # Đọc file
                file_name = file.stem
                suffix = Path(file).suffix
                
                # Đọc file excel
                print(f"Đang xử lí file {file_name}")
                if suffix == ".xlsx":
                    try:
                        has_missing_col = False # Đánh dấu file có thiếu cột không
                        raw_col_list = pd.read_excel(file, nrows=0).columns.tolist() # Lấy tên cột có trong file gốc

                        for col_name, raw_col_name_list in self.config_data.col_req_dict.items():
                            
                            # Trường hợp không tìm thấy cột nào được map với tên cột trong file gốc
                            if not any(col in raw_col_name_list for col in raw_col_list): 
                                not_mapped_info = {
                                    "Tên file ": file.stem,
                                    "Tên cột chưa map": col_name
                                }
                                not_mapped_columns.append(not_mapped_info)
                                has_missing_col = True
                                
                        if has_missing_col:
                            has_error = True # Trả về lỗi chung, batch này sẽ không được import vào database
                            continue

                        print("Đã map đủ tất cả các cột bắt buộc!")
                        df: pd.DataFrame = pd.read_excel(
                            file,
                            usecols=[col for col in self.config_data.col_use if col in raw_col_list], 
                            dtype="string", 
                            engine="openpyxl"
                        )

                        # Đổi tên cột
                        df = df.rename(columns=self.config_data.rename_dict)

                        # Drop dòng trống
                        df = df.dropna(subset=['business_partner_code'])

                        # Tạo date từ tên file 
                        extracted_date = re.search(r"\d{2}\.\d{4}", file_name).group().strip()
                        df["date"] = "01." + extracted_date
                        df["date"] = pd.to_datetime(df["date"], errors="coerce", format="%d.%m.%Y")

                        # Chuyển định dạng
                        for col, dtype in self.config_data.dtype_dict.items():
                            if col in df.columns:
                                if dtype == "numeric":
                                    df[col] = pd.to_numeric(df[col], errors="coerce")
                        
                        if not df.empty:
                            df["source_file"] = file_name
                            completed_dfs.append(df)
                    except Exception as e:
                        print(e)
        
        if has_missing_col or has_error: 
            result = {"Cột còn thiếu": not_mapped_columns}
            return result, has_error
            
        result = pd.concat(completed_dfs, ignore_index=True)
        return result, has_error

    def copy_to_main_table(self):
        """
        Lấy dữ liệu từ bảng staging trong database và import vào bảng đích
        """
        
        query = text(
            f"""
            INSERT INTO main.profit_and_loss (
                profit_center_key,
                business_partner_key,
                product_key,
                "date",
                a_non_operating_expenses,
                a_non_operating_income,
                administrative_expense,
                advertise_s,
                bank_transfer_fee_s,
                communication_a,
                communication_s,
                cost_of_good_sold,
                cost_of_good_sold_s,
                delivery_maint_a,
                depreciation_a,
                depreciation_s,
                entertainment_a,
                entertainment_s,
                expenses_from_provisions_a,
                export_transportation_s,
                extraordinary_expenses,
                extraordinary_incomes,
                financial_expense,
                financial_rev_exp,
                financial_revenue,
                fx_loss_realized,
                gross_profit,
                house_renting_a,
                income_tax,
                land_renting_expense_a,
                management_service_fee_a,
                net_income,
                office_salaries_a,
                offices_supplies_a,
                offices_supplies_s,
                oil_fee_s,
                operating_expense,
                operating_income,
                ordinary_income,
                other_expenses_a,
                other_expenses_by_cash_s,
                other_repair_a,
                other_services_s,
                porters_fee_s,
                provision_for_decline,
                quantity_inventory_uom,
                quantity_net,
                repair_s,
                salary_allowance_s,
                sale_promotions_item,
                sale_promotions_s,
                sales_net,
                school_medical_fee_a,
                selling_expense,
                tax_fees_and_charges_a,
                tools_equipment_s,
                training_education_a,
                transportation_s,
                travelling_a,
                travelling_s,
                water_lighting_a,
                water_lighting_s
            )
            SELECT
                COALESCE(pc.profit_center_key, 0) AS profit_center_key,
                bp.business_partner_key,
                p.product_key,
                pal.date,
                COALESCE(pal.a_non_operating_expenses, 0) AS a_non_operating_expenses,
                COALESCE(pal.a_non_operating_income, 0) AS a_non_operating_income,
                COALESCE(pal.administrative_expense, 0) AS administrative_expense,
                COALESCE(pal.advertise_s, 0) AS advertise_s,
                COALESCE(pal.bank_transfer_fee_s, 0) AS bank_transfer_fee_s,
                COALESCE(pal.communication_a, 0) AS communication_a,
                COALESCE(pal.communication_s, 0) AS communication_s,
                COALESCE(pal.cost_of_good_sold, 0) AS cost_of_good_sold,
                COALESCE(pal.cost_of_good_sold_s, 0) AS cost_of_good_sold_s,
                COALESCE(pal.delivery_maint_a, 0) AS delivery_maint_a,
                COALESCE(pal.depreciation_a, 0) AS depreciation_a,
                COALESCE(pal.depreciation_s, 0) AS depreciation_s,
                COALESCE(pal.entertainment_a, 0) AS entertainment_a,
                COALESCE(pal.entertainment_s, 0) AS entertainment_s,
                COALESCE(pal.expenses_from_provisions_a, 0) AS expenses_from_provisions_a,
                COALESCE(pal.export_transportation_s, 0) AS export_transportation_s,
                COALESCE(pal.extraordinary_expenses, 0) AS extraordinary_expenses,
                COALESCE(pal.extraordinary_incomes, 0) AS extraordinary_incomes,
                COALESCE(pal.financial_expense, 0) AS financial_expense,
                COALESCE(pal.financial_rev_exp, 0) AS financial_rev_exp,
                COALESCE(pal.financial_revenue, 0) AS financial_revenue,
                COALESCE(pal.fx_loss_realized, 0) AS fx_loss_realized,
                COALESCE(pal.gross_profit, 0) AS gross_profit,
                COALESCE(pal.house_renting_a, 0) AS house_renting_a,
                COALESCE(pal.income_tax, 0) AS income_tax,
                COALESCE(pal.land_renting_expense_a, 0) AS land_renting_expense_a,
                COALESCE(pal.management_service_fee_a, 0) AS management_service_fee_a,
                COALESCE(pal.net_income, 0) AS net_income,
                COALESCE(pal.office_salaries_a, 0) AS office_salaries_a,
                COALESCE(pal.offices_supplies_a, 0) AS offices_supplies_a,
                COALESCE(pal.offices_supplies_s, 0) AS offices_supplies_s,
                COALESCE(pal.oil_fee_s, 0) AS oil_fee_s,
                COALESCE(pal.operating_expense, 0) AS operating_expense,
                COALESCE(pal.operating_income, 0) AS operating_income,
                COALESCE(pal.ordinary_income, 0) AS ordinary_income,
                COALESCE(pal.other_expenses_a, 0) AS other_expenses_a,
                COALESCE(pal.other_expenses_by_cash_s, 0) AS other_expenses_by_cash_s,
                COALESCE(pal.other_repair_a, 0) AS other_repair_a,
                COALESCE(pal.other_services_s, 0) AS other_services_s,
                COALESCE(pal.porters_fee_s, 0) AS porters_fee_s,
                COALESCE(pal.provision_for_decline, 0) AS provision_for_decline,
                COALESCE(pal.quantity_inventory_uom, 0) AS quantity_inventory_uom,
                COALESCE(pal.quantity_net, 0) AS quantity_net,
                COALESCE(pal.repair_s, 0) AS repair_s,
                COALESCE(pal.salary_allowance_s, 0) AS salary_allowance_s,
                COALESCE(pal.sale_promotions_item, 0) AS sale_promotions_item,
                COALESCE(pal.sale_promotions_s, 0) AS sale_promotions_s,
                COALESCE(pal.sales_net, 0) AS sales_net,
                COALESCE(pal.school_medical_fee_a, 0) AS school_medical_fee_a,
                COALESCE(pal.selling_expense, 0) AS selling_expense,
                COALESCE(pal.tax_fees_and_charges_a, 0) AS tax_fees_and_charges_a,
                COALESCE(pal.tools_equipment_s, 0) AS tools_equipment_s,
                COALESCE(pal.training_education_a, 0) AS training_education_a,
                COALESCE(pal.transportation_s, 0) AS transportation_s,
                COALESCE(pal.travelling_a, 0) AS travelling_a,
                COALESCE(pal.travelling_s, 0) AS travelling_s,
                COALESCE(pal.water_lighting_a, 0) AS water_lighting_a,
                COALESCE(pal.water_lighting_s, 0) AS water_lighting_s
            FROM staging.profit_and_loss_staging pal
            LEFT JOIN main.profit_center pc USING (profit_center_code)
            LEFT JOIN main.business_partner bp
                ON pal.business_partner_code = bp.business_partner_code 
                AND (pal.date >= bp.valid_from AND (pal.date <= bp.valid_to OR bp.valid_to IS NULL))
            LEFT JOIN main.product p USING (product_code);
            """
            )

        with self.engine.begin() as conn:
            conn.execute(query)


class BusinessPartnerUpdater:
    def __init__(self, engine, folder_path: Path):
        self.engine = engine
        self.data_folder = folder_path / "business_partner_data"
        self.config_data= ConfigLoader(r"configs/list_of_bp_config.json")
    
    def check_file(self):
        """
        Kiểm tra việc folder chỉ có đúng 1 duy nhất file thông tin khách hàng, trả về path của file nếu thoả mãn điều kiện 1 file duy nhất
        """
        file_list = [file for file in self.data_folder.glob("*.xlsx")] 
        if not len(file_list) == 1:
            print(f"Phát hiện {len(file_list)} trong thư mục, vui lòng chỉ lấy duy nhất 1 file")
            return
        
        return file_list[0]
    
    def read_data(self):
        """
        Đọc file dữ liệu khách hàng
        """
        file = self.check_file()
        has_error = False 
        raw_col_list = pd.read_excel(file, nrows=0).columns.tolist() # Lấy tên cột có trong file gốc
        not_mapped_columns = [] # Biến lưu trữ các cột chưa được map raw name

        for col_name, raw_col_name_list in self.config_data.col_req_dict.items():
            
            # Trường hợp không tìm thấy cột nào được map với tên cột trong file gốc
            if not any(col in raw_col_name_list for col in raw_col_list): 
                not_mapped_info = {
                    "Tên file ": file.stem,
                    "Tên cột chưa map": col_name
                }
                not_mapped_columns.append(not_mapped_info)
                has_error = True
        
        if has_error:
            return not_mapped_columns, has_error

        print("Đã map đủ tất cả các cột bắt buộc!")
        df: pd.DataFrame = pd.read_excel(
            file,
            usecols=[col for col in self.config_data.col_use if col in raw_col_list], 
            dtype="string", 
            engine="openpyxl"
        )

        # Đổi tên cột
        df = df.rename(columns=self.config_data.rename_dict)
    
        # Drop dòng không đạt
        df = df.dropna(subset=['business_partner_code'])

        # Chuyển đổi định dạng
        for col, dtype in self.config_data.dtype_dict.items():
            if col in df.columns:
                if dtype == "date":
                    df[col] = pd.to_datetime(df[col], format="%d.%m.%Y")
        
        return df, has_error
    
    def update_data(self):
        query = text(
            """
            UPDATE main.business_partner bp
            SET (creation_date, ship_to_street, bill_to_street) = (bps.creation_date, bps.ship_to_street, bps.bill_to_street)
            FROM staging.business_partner_staging bps 
            WHERE bp.business_partner_code = bps.business_partner_code
        """
        )

        with self.engine.begin() as conn:
            conn.execute(query)
        
        print("Cập nhật dữ liệu thành công!")