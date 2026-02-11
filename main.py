# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import openpyxl
import pyodbc
import traceback
import datetime
import copy
import shutil  # 用于文件复制
import os      # 用于路径处理
from collections import defaultdict
from openpyxl.styles import Font
from tkcalendar import DateEntry

# ============== 用户配置区 ==============
def get_best_sql_driver():
    try:
        installed_drivers = [d for d in pyodbc.drivers()]
    except Exception:
        return "SQL Server"

    driver_preference = [
        "ODBC Driver 18 for SQL Server", "ODBC Driver 17 for SQL Server",
        "ODBC Driver 13 for SQL Server", "SQL Server Native Client 11.0",
        "SQL Server"
    ]
    for drv in driver_preference:
        if drv in installed_drivers: return drv
    return "SQL Server"


CURRENT_DRIVER = get_best_sql_driver()
# 数据库连接 (只读权限)
DB_CONN_STRING = (
    f"DRIVER={{{CURRENT_DRIVER}}};SERVER=192.168.0.117;DATABASE=FQD;"
    "UID=zhitan;PWD=Zt@forcome;TrustServerCertificate=yes;"
)

# 基础数据从第4行开始
ROW_IDX_DATA_START = 4

COL_NAME_WORKSHOP = "车间"
COL_NAME_WO_TYPE = "单别"
COL_NAME_WO_NO = "工单单号"

# ============== 应用程序类 ==============
class DailyPlanAvailabilityApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"排程齐套分析 (Openpyxl最终版+单位) - {CURRENT_DRIVER}")
        self.root.geometry("1000x650")

        self.file_path = tk.StringVar()
        self.sheet_name = tk.StringVar()
        self.selected_workshop = tk.StringVar()
        self.date_column_map = {}
        self.col_map_main = {}

        self._create_widgets()

    def _create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 1. 文件选择
        file_frame = ttk.LabelFrame(main_frame, text="1. 数据源 (程序将自动备份原文件)", padding="5")
        file_frame.pack(fill=tk.X, pady=5)
        ttk.Entry(file_frame, textvariable=self.file_path, width=50).pack(side=tk.LEFT, padx=5)
        ttk.Button(file_frame, text="浏览Excel...", command=self._select_file).pack(side=tk.LEFT, padx=5)
        ttk.Label(file_frame, text="   工作表:").pack(side=tk.LEFT)
        self.sheet_combo = ttk.Combobox(file_frame, textvariable=self.sheet_name, state="disabled", width=15)
        self.sheet_combo.pack(side=tk.LEFT, padx=5)
        self.sheet_combo.bind("<<ComboboxSelected>>", self._on_sheet_selected)

        # 2. 筛选设置
        filter_frame = ttk.LabelFrame(main_frame, text="2. 分析范围设置 (汇总多天排产)", padding="10")
        filter_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(filter_frame, text="开始日期:").pack(side=tk.LEFT)
        self.date_start = DateEntry(filter_frame, width=12, background='darkblue', foreground='white', borderwidth=2,
                                    date_pattern='yyyy/mm/dd')
        self.date_start.pack(side=tk.LEFT, padx=5)

        ttk.Label(filter_frame, text="结束日期:").pack(side=tk.LEFT)
        self.date_end = DateEntry(filter_frame, width=12, background='darkblue', foreground='white', borderwidth=2,
                                    date_pattern='yyyy/mm/dd')
        self.date_end.pack(side=tk.LEFT, padx=5)

        ttk.Label(filter_frame, text="选择车间:").pack(side=tk.LEFT, padx=(30, 5))
        self.workshop_combo = ttk.Combobox(filter_frame, textvariable=self.selected_workshop, state="disabled",
                                           width=20)
        self.workshop_combo.pack(side=tk.LEFT, padx=5)

        # 3. 执行按钮
        action_frame = ttk.LabelFrame(main_frame, text="3. 执行", padding="10")
        action_frame.pack(fill=tk.X, pady=10)
        
        btn_text = "备份 -> 计算汇总齐套率 -> 写入A列 (含单位)"
        ttk.Button(action_frame, text=btn_text, command=self._run_analysis_aggregated).pack(fill=tk.X, padx=100)

        self.log_text = tk.Text(main_frame, height=15, state="disabled", font=("Consolas", 9), bg="#F0F0F0")
        self.log_text.pack(fill=tk.BOTH, expand=True, pady=5)

    def _log(self, msg):
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")
        self.root.update_idletasks()

    def _select_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls *.xlsm")])
        if path:
            self.file_path.set(path)
            try:
                wb = openpyxl.load_workbook(path, read_only=True)
                self.sheet_combo['values'] = wb.sheetnames
                if wb.sheetnames:
                    self.sheet_combo.current(0)
                    self._on_sheet_selected(None)
                self.sheet_combo.config(state="readonly")
                wb.close()
            except Exception as e:
                messagebox.showerror("错误", f"无法打开文件: {e}")

    def _on_sheet_selected(self, event):
        file_path = self.file_path.get()
        sheet_name = self.sheet_name.get()
        if not file_path or not sheet_name: return
        try:
            wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
            ws = wb[sheet_name]

            self.col_map_main = {}
            scan_rows = [3, 2]
            for r in scan_rows:
                for idx, cell in enumerate(ws[r], start=1):
                    val = str(cell.value).strip() if cell.value else ""
                    if val and val not in self.col_map_main:
                        self.col_map_main[val] = idx

            self.date_column_map = {}
            for cell in ws[3]:
                val = cell.value
                dt = self._parse_excel_date(val)
                if dt: self.date_column_map[dt] = cell.column

            col_ws_idx = self.col_map_main.get(COL_NAME_WORKSHOP)
            workshops = set()
            if col_ws_idx:
                for row in ws.iter_rows(min_row=ROW_IDX_DATA_START, min_col=col_ws_idx, max_col=col_ws_idx, values_only=True):
                    if row[0]: workshops.add(str(row[0]).strip())

            self.workshop_combo['values'] = ["全部车间"] + sorted(list(workshops))
            self.workshop_combo.current(0)
            self.workshop_combo.config(state="readonly")

            self._log(f"就绪: 找到 {len(self.date_column_map)} 个日期列。")
            wb.close()
        except Exception as e:
            traceback.print_exc()
            self._log(f"扫描失败: {e}")

    def _parse_excel_date(self, val):
        if val is None: return None
        try:
            if isinstance(val, datetime.datetime): return val.date()
            if isinstance(val, datetime.date): return val
            if isinstance(val, (int, float)):
                return (datetime.datetime(1899, 12, 30) + datetime.timedelta(days=int(val))).date()
            if isinstance(val, str):
                try:
                    return datetime.datetime.strptime(val.strip(), "%Y/%m/%d").date()
                except:
                    pass
            return None
        except:
            return None

    def _create_backup(self, file_path):
        """创建文件备份"""
        try:
            dir_name = os.path.dirname(file_path)
            base_name = os.path.basename(file_path)
            name_part, ext_part = os.path.splitext(base_name)
            
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_name = f"{name_part}_备份_{timestamp}{ext_part}"
            backup_path = os.path.join(dir_name, backup_name)
            
            shutil.copy2(file_path, backup_path)
            self._log(f"已创建备份文件: {backup_name}")
            return True
        except Exception as e:
            self._log(f"备份失败: {e}")
            messagebox.showerror("备份失败", f"无法创建备份文件，操作已取消。\n{e}")
            return False

    def _run_analysis_aggregated(self):
        start_date = self.date_start.get_date()
        end_date = self.date_end.get_date()
        
        if end_date < start_date:
            messagebox.showerror("日期错误", "结束日期不能早于开始日期")
            return

        file_path = self.file_path.get()
        sheet_name = self.sheet_name.get()
        target_workshop = self.selected_workshop.get()

        if not file_path: return

        # 找出范围内所有的列
        target_cols = []
        curr = start_date
        while curr <= end_date:
            if curr in self.date_column_map:
                target_cols.append(self.date_column_map[curr])
            curr += datetime.timedelta(days=1)
            
        if not target_cols:
            messagebox.showwarning("日期无效", "所选日期范围内没有在Excel第3行找到对应的日期列。")
            return

        msg = (f"即将分析并修改文件：\n\n文件: {file_path}\n"
               f"日期: {start_date} 至 {end_date}\n\n"
               "程序将自动创建一份备份文件，然后修改原文件的A列。\n"
               "是否继续？")
        
        confirm = messagebox.askyesno("确认修改", msg)
        if not confirm: return

        # --- 执行备份 ---
        self._log("正在创建备份...")
        if not self._create_backup(file_path):
            return

        try:
            self._log(f"正在加载源文件 (Openpyxl)...")
            wb = openpyxl.load_workbook(file_path)
            ws = wb[sheet_name]

            self._log(f"正在汇总排产数据 (包含 {len(target_cols)} 天)...")
            plans = self._extract_data_aggregated(ws, target_cols, target_workshop)
            
            if not plans:
                messagebox.showinfo("无数据", "所选范围内没有排产数量 > 0 的工单。")
                return

            all_wo_keys = list(set([p['wo_key'] for p in plans]))
            
            self._log(f"查询ERP数据 (工单数: {len(all_wo_keys)})...")
            static_wo_data = self._fetch_erp_data(all_wo_keys)
            
            all_parts = set()
            for w in static_wo_data.values():
                for b in w['bom']: all_parts.add(b['part'])
            
            static_inventory = self._fetch_inventory(list(all_parts))

            running_inv = copy.deepcopy(static_inventory)
            running_wo_issued = defaultdict(float)
            for k, v in static_wo_data.items():
                for b in v['bom']:
                    running_wo_issued[(k[0], k[1], b['part'])] = b['iss']

            self._log("开始计算总齐套逻辑...")
            results = self._simulate_logic(plans, static_wo_data, running_inv, running_wo_issued)

            self._log("正在写入 A 列...")
            font_style = Font(name="微软雅黑", size=9)
            
            count = 0
            for r in results:
                row_idx = r['row_idx']
                rate_str = f"{r['rate']:.0%}"
                msg = r['msg'].replace("\n", "; ")
                final_str = f"{rate_str}：{r['achievable']}：{r['net_demand']}：{msg}"
                
                cell = ws.cell(row=row_idx, column=1)
                cell.value = final_str
                cell.font = font_style
                count += 1

            self._log(f"保存文件...")
            wb.save(file_path)
            wb.close()
            
            messagebox.showinfo("完成", f"分析完成！\n已备份原文件。\n结果已写入 {count} 行到 A 列。")
            self._log("全部完成。")

        except Exception as e:
            traceback.print_exc()
            self._log(f"错误: {e}")
            messagebox.showerror("运行错误", f"发生错误，文件未保存。\n{e}")

    def _extract_data_aggregated(self, ws, col_indices, filter_ws):
        c_ws = self.col_map_main.get(COL_NAME_WORKSHOP)
        c_type = self.col_map_main.get(COL_NAME_WO_TYPE)
        c_no = self.col_map_main.get(COL_NAME_WO_NO)

        if not c_type: c_type = 5
        if not c_no: c_no = 6

        data = []
        for row in ws.iter_rows(min_row=ROW_IDX_DATA_START):
            try:
                total_qty = 0
                for col_idx in col_indices:
                    if col_idx <= len(row):
                        val = row[col_idx - 1].value
                        if isinstance(val, (int, float)) and val > 0:
                            total_qty += val
                
                if total_qty > 0:
                    int_qty = int(round(total_qty))
                    
                    val_ws = row[c_ws - 1].value if (c_ws and c_ws <= len(row)) else None
                    curr_ws = str(val_ws).strip() if val_ws else "未分类"
                    
                    if filter_ws != "全部车间" and curr_ws != filter_ws: continue

                    wt = row[c_type - 1].value
                    wn = row[c_no - 1].value
                    
                    if wt and wn:
                        data.append({
                            'wo_key': (str(wt).strip(), str(wn).strip()),
                            'qty': int_qty, 
                            'row_idx': row[0].row
                        })
            except:
                continue
        return data

    def _fetch_erp_data(self, keys):
        if not keys: return {}
        conditions = [f"(TA.TA001='{t}' AND TA.TA002='{n}')" for t, n in keys]
        data = defaultdict(lambda: {'total': 0, 'bom': []})
        batch_size = 200
        for i in range(0, len(conditions), batch_size):
            batch = conditions[i:i + batch_size]
            sql = f"""
                SELECT RTRIM(TA.TA001) t, RTRIM(TA.TA002) n, TA.TA015 total,
                       RTRIM(TB.TB003) p, ISNULL(RTRIM(MB.MB002),'') name, 
                       ISNULL(RTRIM(MB.MB004),'') unit, TB.TB004 req, TB.TB005 iss
                FROM MOCTA TA
                INNER JOIN MOCTB TB ON TA.TA001=TB.TB001 AND TA.TA002=TB.TB002
                LEFT JOIN INVMB MB ON TB.TB003=MB.MB001
                WHERE {" OR ".join(batch)}
            """
            try:
                with pyodbc.connect(DB_CONN_STRING) as conn:
                    df = pd.read_sql(sql, conn)
                    for _, r in df.iterrows():
                        data[(r['t'], r['n'])]['total'] = float(r['total'])
                        data[(r['t'], r['n'])]['bom'].append({
                            'part': r['p'], 'name': r['name'], 'unit': r['unit'],
                            'req': float(r['req']), 'iss': float(r['iss'])
                        })
            except:
                pass
        return data

    def _fetch_inventory(self, parts):
        if not parts: return {}
        inv = {}
        parts = list(set(parts))
        batch_size = 500
        for i in range(0, len(parts), batch_size):
            p_str = ",".join(f"'{p}'" for p in parts[i:i + batch_size])
            sql = f"SELECT RTRIM(MC001) p, SUM(MC007) q FROM INVMC WHERE MC001 IN ({p_str}) GROUP BY MC001"
            try:
                with pyodbc.connect(DB_CONN_STRING) as conn:
                    df = pd.read_sql(sql, conn)
                inv.update(pd.Series(df.q.values, index=df.p).to_dict())
            except:
                pass
        return inv

    def _simulate_logic(self, plans, wo_data, running_inv, running_wo_issued):
        results = []
        for p in plans:
            key = p['wo_key']
            plan_qty = p['qty']
            info = wo_data.get(key)
            row_idx = p['row_idx']

            res = {
                'row_idx': row_idx,
                'rate': 0.0, 'achievable': 0,
                'net_demand': 0, 'msg': ""
            }

            if not info or not info['bom']:
                res['msg'] = "无ERP信息"
                results.append(res)
                continue

            max_possible_by_erp_limit = 999999
            for b in info['bom']:
                unit_use = b['req'] / info['total'] if info['total'] > 0 else 0
                if unit_use <= 0: continue
                current_issued = running_wo_issued.get((key[0], key[1], b['part']), 0)
                remain_issue_qty = max(0, b['req'] - current_issued)
                possible_sets = int(remain_issue_qty // unit_use)
                if possible_sets < max_possible_by_erp_limit:
                    max_possible_by_erp_limit = possible_sets

            final_net_demand_int = min(plan_qty, max_possible_by_erp_limit)
            
            min_material_rate = 1.0
            min_possible_sets = 999999
            short_details = []
            to_deduct_full = {}

            for b in info['bom']:
                unit_use = b['req'] / info['total'] if info['total'] > 0 else 0
                if unit_use <= 0: continue

                part_net_demand = final_net_demand_int * unit_use
                stock = running_inv.get(b['part'], 0)

                if part_net_demand > 0:
                    effective_stock = max(0, stock)
                    part_rate = effective_stock / part_net_demand
                    if part_rate > 1.0: part_rate = 1.0
                    if part_rate < min_material_rate:
                        min_material_rate = part_rate

                can_do = int(max(0, stock) // unit_use)
                min_possible_sets = min(min_possible_sets, can_do)

                if stock < part_net_demand - 0.0001:
                    diff = part_net_demand - stock
                    # ============== 修改点：增加了 b['unit'] ==============
                    short_details.append(f"{b['name']}缺{diff:g}{b['unit']}")

                full_demand = plan_qty * unit_use 
                to_deduct_full[b['part']] = full_demand

            achievable = min(final_net_demand_int, min_possible_sets)

            res['rate'] = min_material_rate
            res['achievable'] = achievable
            res['net_demand'] = final_net_demand_int
            
            if final_net_demand_int == 0 and plan_qty > 0:
                 res['msg'] = "工单已结案"
            elif min_material_rate >= 0.999:
                 res['msg'] = "齐套"
            else:
                 res['msg'] = ",".join(short_details)

            results.append(res)

            for part, qty in to_deduct_full.items():
                if part not in running_inv: running_inv[part] = 0.0
                running_inv[part] -= qty
                running_wo_issued[(key[0], key[1], part)] += qty

        return results

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = DailyPlanAvailabilityApp(root)
        root.mainloop()
    except Exception as e:
        import tkinter.messagebox
        tkinter.messagebox.showerror("启动失败", str(e))
