import pandas as pd
from sqlalchemy import text
from src.database.structure import Base


class DatabaseController:
    def __init__(self, engine):
        self.engine = engine 

    def insert_dataframe(self, df: pd.DataFrame, table_name: str, schema: str, chunksize: int = 10000):
        """
        Hàm này nhận các dataframe và nhập vào database, có thể dùng chung giữa các pipeline
        """
        
        if df.empty:
            return

        df.to_sql(
            name=table_name,
            con=self.engine,
            schema=schema,
            if_exists="append", # KHÔNG ĐƯỢC THAY ĐỔI: Sử dụng append để database tự map tên cột và bỏ trống cột mà dataframe không có
            index=False,
            chunksize=chunksize,
            method=self._psql_insert_copy
        )
    
    @staticmethod
    def _psql_insert_copy(table, conn, keys, data_iter):
        dbapi_conn = conn.connection
        with dbapi_conn.cursor() as cur:
            columns = ", ".join([f'"{k}"' for k in keys])
            table_name = f'"{table.schema}"."{table.name}"' if table.schema else f'"{table.name}"'
            with cur.copy(f"COPY {table_name} ({columns}) FROM STDIN") as copy:
                for row in data_iter:
                    copy.write_row(row)
    
    def truncate_table(self, target_table: str, target_schema: str):
        target = f"{target_schema}.{target_table}" 
        print(f"Đang xoá dữ liệu của bảng {target}")
        sql = text(f"TRUNCATE TABLE {target} RESTART IDENTITY")
        with self.engine.begin() as conn:
            conn.execute(sql)
    
    def refresh_view(self, target_view: str, target_schema: str):
        target = f"{target_schema}.{target_view}"
        print(f"Đang làm mới dữ liệu của view: {target}")
        sql = text(f"REFRESH MATERIALIZED VIEW {target}")
        with self.engine.begin() as conn:
            conn.execute(sql)

        print(f"Làm mới thành công!")