import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from urllib.parse import quote_plus
import os
import traceback
import sys
import warnings
import threading

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkcalendar import DateEntry
from sqlalchemy import create_engine, text


# TextRedirector
class TextRedirector:
    def __init__(self, widget): self.widget = widget
    def write(self, str_):
        self.widget.configure(state="normal")
        self.widget.insert("end", str_)
        self.widget.see("end")
        self.widget.update_idletasks()
        self.widget.configure(state="disabled")
    def flush(self): pass

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("设备三灯时间统计工具 v1.0")
        self.root.geometry("650x650")
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # Help Guide
        guide_frame = ttk.LabelFrame(main_frame, text="📌 使用说明", padding="10")
        guide_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        help_text = (
            "1. 选择统计的开始日期和结束日期\n"
            "2. 选择导出文件夹\n"
            "3. 点击「开始导出」按钮\n"
            "4. 等待处理完成后，在导出文件夹中查看 Excel 文件\n"
            "──────────────────────────────\n"
            "💡 数据说明:\n"
            "• 统计时间窗口为每天 08:00 ~ 次日 08:00\n"
            "• 绿灯时间 = 设备运行中 (run)\n"
            "• 黄灯时间 = 设备停机 (stop)\n"
            "• 红灯时间 = 设备离线 (offline)"
        )
        ttk.Label(guide_frame, text=help_text, justify=tk.LEFT, wraplength=600).pack(fill='x')

        # Settings
        settings_frame = ttk.LabelFrame(main_frame, text="报表设置", padding="10")
        settings_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        settings_frame.columnconfigure(1, weight=1)

        ttk.Label(settings_frame, text="开始日期:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.start_date_entry = DateEntry(settings_frame, date_pattern='y-mm-dd', width=12)
        self.start_date_entry.grid(row=0, column=1, sticky="w", padx=5, pady=5)

        ttk.Label(settings_frame, text="结束日期:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.end_date_entry = DateEntry(settings_frame, date_pattern='y-mm-dd', width=12)
        self.end_date_entry.grid(row=1, column=1, sticky="w", padx=5, pady=5)

        ttk.Label(settings_frame, text="导出文件夹:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.output_path_var = tk.StringVar()
        self.output_path_entry = ttk.Entry(settings_frame, textvariable=self.output_path_var)
        self.output_path_entry.grid(row=2, column=1, sticky="ew", padx=5, pady=5)
        ttk.Button(settings_frame, text="浏览...", command=self.browse_dir).grid(row=2, column=2, padx=5, pady=5)

        # Buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=2, column=0, sticky="e", pady=(0, 10))
        self.run_button = ttk.Button(btn_frame, text="开始导出", command=self.start_process)
        self.run_button.pack(side="left", padx=5)
        ttk.Button(btn_frame, text="退出", command=root.quit).pack(side="left", padx=5)

        # Log Area
        log_frame = ttk.LabelFrame(main_frame, text="运行日志", padding="10")
        log_frame.grid(row=3, column=0, sticky="nsew")
        main_frame.rowconfigure(3, weight=1)
        self.log_text = tk.Text(log_frame, wrap=tk.WORD, state="disabled")
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text['yscrollcommand'] = scrollbar.set
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    def browse_dir(self):
        path = filedialog.askdirectory()
        if path:
            self.output_path_var.set(path)

    def start_process(self):
        start_date = self.start_date_entry.get_date()
        end_date = self.end_date_entry.get_date()
        output_path = self.output_path_var.get()

        if start_date > end_date:
            messagebox.showerror("错误", "开始日期不能晚于结束日期！")
            return
        if not output_path:
            messagebox.showerror("错误", "请选择导出文件夹！")
            return

        self.run_button.config(state="disabled")
        # Clear Log
        self.log_text.configure(state="normal")
        self.log_text.delete('1.0', tk.END)
        self.log_text.configure(state="disabled")

        # Run in thread
        thread = threading.Thread(target=self.process_task, args=(start_date, end_date, output_path))
        thread.daemon = True
        thread.start()

    def process_task(self, start_date, end_date, output_path):
        original_stdout = sys.stdout
        sys.stdout = TextRedirector(self.log_text)
        try:
            print("=" * 20 + " 开始任务 " + "=" * 20)
            engine = self.get_db_engine()
            with engine.connect() as conn:
                print("数据库连接成功。")
                print("正在获取设备列表...")
                query = "SELECT MachineID, MachineNo, MachineName FROM machine"
                machine_df = pd.read_sql(query, conn)
                machine_df.loc[machine_df['MachineNo'] == 'W045-2', 'MachineNo'] = 'W049'

                all_data = []
                current_date = start_date
                while current_date <= end_date:
                    day_data = self.fetch_day_data(conn, machine_df, current_date)
                    if day_data is not None:
                        all_data.append(day_data)
                    current_date += timedelta(days=1)

            if all_data:
                final_df = pd.concat(all_data, ignore_index=True)
                self.export_to_excel(final_df, output_path, start_date, end_date)
                print("\n✅ 任务完成！文件已保存至: " + output_path)
                messagebox.showinfo("完成", "报表生成成功！")
            else:
                print("\n⚠️ 所选日期范围内没有查询到数据。")
                messagebox.showinfo("提示", "未查询到数据。")

        except Exception as e:
            error_msg = f"发生错误:\n{traceback.format_exc()}"
            print(error_msg)
            messagebox.showerror("错误", f"处理失败: {str(e)}\n详情请查看日志。")
        finally:
            sys.stdout = original_stdout
            self.run_button.config(state="normal")

    def get_db_engine(self):
        try:
            db_user = "fkmes"
            db_password_raw = "fk@123"
            db_host = "192.168.0.37"
            db_port = "3306"
            db_name = "fms_test"
            db_password = quote_plus(db_password_raw)
            conn_str = f"mysql+pymysql://{db_user}:{db_password}@{db_host}:{db_port}/{db_name}?charset=utf8mb4"
            return create_engine(conn_str, pool_recycle=3600, connect_args={"connect_timeout": 5})
        except Exception as e:
            raise ConnectionError(f"数据库引擎创建失败: {e}")

    def fetch_day_data(self, conn, machine_df, target_date):
        print(f"正在查询 {target_date.strftime('%Y-%m-%d')} 的数据...")
        start_time = datetime.combine(target_date, datetime.min.time()) + timedelta(hours=8)
        end_time = start_time + timedelta(days=1)
        
        query = text("""
            SELECT StatusMachineId, StatusStartTime, StatusEndTime, StatusDes 
            FROM machinestatus 
            WHERE StatusStartTime < :end_time AND StatusEndTime > :start_time
        """)
        
        df = pd.read_sql(query, conn, params={'start_time': start_time, 'end_time': end_time})
        
        if df.empty:
            print(f"  {target_date.strftime('%Y-%m-%d')}: 无数据。")
            return None

        # Process Times
        df['StatusStartTime'] = pd.to_datetime(df['StatusStartTime'])
        df['StatusEndTime'] = pd.to_datetime(df['StatusEndTime'])
        
        # Clip to window
        df['EffStart'] = df['StatusStartTime'].clip(lower=start_time)
        df['EffEnd'] = df['StatusEndTime'].clip(upper=end_time)
        df['Duration'] = (df['EffEnd'] - df['EffStart']).dt.total_seconds()
        
        # Pivot Table
        df['StatusDes'] = df['StatusDes'].str.lower().str.strip()
        summary = df.groupby(['StatusMachineId', 'StatusDes'])['Duration'].sum().unstack(fill_value=0)
        
        # Ensure columns exist
        for col in ['run', 'stop', 'offline']:
            if col not in summary.columns:
                summary[col] = 0
                
        summary = summary.rename(columns={
            'run': '绿灯时间',
            'stop': '黄灯时间',
            'offline': '红灯时间'
        })
        
        # Merge with Machine Info
        result = pd.merge(machine_df, summary, left_on='MachineID', right_index=True, how='left')
        result.fillna(0, inplace=True)
        
        # Seconds to Hours
        result['绿灯时间'] = result['绿灯时间'] / 3600
        result['黄灯时间'] = result['黄灯时间'] / 3600
        result['红灯时间'] = result['红灯时间'] / 3600
        
        result['日期'] = target_date.strftime('%Y-%m-%d')
        
        print(f"  {target_date.strftime('%Y-%m-%d')}: 处理完成。")
        return result[['日期', 'MachineNo', 'MachineName', '绿灯时间', '黄灯时间', '红灯时间']]

    def export_to_excel(self, df, output_path, start_date, end_date):
        file_name = f"设备三灯统计_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx"
        full_path = os.path.join(output_path, file_name)
        
        # Calculate total
        df['合计'] = df['绿灯时间'] + df['黄灯时间'] + df['红灯时间']
        
        # Sort
        df = df.sort_values(by=['日期', 'MachineNo'])
        
        # Reorder columns
        df = df.rename(columns={'MachineNo': '设备代码', 'MachineName': '设备名称'})
        
        print(f"正在生成 Excel: {full_path}")
        
        writer = pd.ExcelWriter(full_path, engine='xlsxwriter')
        df.to_excel(writer, sheet_name='三灯时间统计', index=False, startrow=1)
        
        workbook = writer.book
        worksheet = writer.sheets['三灯时间统计']
        
        # Formats
        header_fmt = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#DDEBF7',
            'border': 1
        })
        num_fmt = workbook.add_format({'num_format': '0.00', 'border': 1, 'align': 'center'})
        str_fmt = workbook.add_format({'border': 1, 'align': 'center'})
        
        # Write Header
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_fmt)
            
        # Format Columns
        for i, col in enumerate(df.columns):
            max_len = max(df[col].astype(str).map(len).max(), len(col)) + 4
            if col in ['绿灯时间', '黄灯时间', '红灯时间', '合计']:
                worksheet.set_column(i, i, max_len, num_fmt)
            else:
                worksheet.set_column(i, i, max_len, str_fmt)
                
        writer.close()
        print("Excel 生成成功。")

if __name__ == "__main__":
    warnings.filterwarnings("ignore")
    root = tk.Tk()
    app = App(root)
    root.mainloop()
