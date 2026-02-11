# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import openpyxl
import pyodbc
import traceback
import datetime
import copy
import shutil
import os
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
        self.root.title(f"排程齐套分析 (最终逻辑重构版) - {CURRENT_DRIVER}")
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
        filter_frame = ttk.LabelFrame(main_frame, text="2. 分析范围设置 (按开工日期排序扣减库存)", padding="10")
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
        
        btn_text = "备份 -> 模拟推演 -> 写入A列 (含前缀描述)"
        ttk.Button(action_frame, text=btn_text, command=self._run_analysis_logic_v2).pack(fill=tk.X, padx=100)

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

    def _run_analysis_logic_v2(self):
        """
        新的逻辑入口：
        1. 提取工单，并确定工单在范围内的'开工日期'用于排序。
        2. 按照排序后的顺序，依次扣减库存。
        3. 回写A列。
        """
        start_date = self.date_start.get_date()
        end_date = self.date_end.get_date()
        
        if end_date < start_date:
            messagebox.showerror("日期错误", "结束日期不能早于开始日期")
            return

        file_path = self.file_path.get()
        sheet_name = self.sheet_name.get()
        target_workshop = self.selected_workshop.get()

        if not file_path: return

        # 找出范围内所有的列，以及它们对应的日期
        target_cols_map = {} # {col_idx: date}
        curr = start_date
        while curr <= end_date:
            if curr in self.date_column_map:
                target_cols_map[self.date_column_map[curr]] = curr
            curr += datetime.timedelta(days=1)
            
        if not target_cols_map:
            messagebox.showwarning("日期无效", "所选日期范围内没有在Excel第3行找到对应的日期列。")
            return
        
        sorted_target_cols = sorted(target_cols_map.keys())

        msg = (f"即将执行逻辑重构版分析：\n\n文件: {file_path}\n"
               f"日期: {start_date} 至 {end_date}\n\n"
               "逻辑：按工单开工日期顺序扣减库存 -> 回写A列\n"
               "是否继续？")
        
        confirm = messagebox.askyesno("确认修改", msg)
        if not confirm: return

        if not self._create_backup(file_path): return

        try:
            self._log(f"加载文件 (Openpyxl)...")
            wb = openpyxl.load_workbook(file_path)
            ws = wb[sheet_name]

            self._log(f"提取工单并确定开工顺序...")
            # 提取数据：列表包含 {'wo_key':..., 'start_date':..., 'row_idx':...}
            wo_list = self._extract_data_with_start_date(ws, sorted_target_cols, target_cols_map, target_workshop)
            
            if not wo_list:
                messagebox.showinfo("无数据", "所选范围内没有排产数量 > 0 的工单。")
                return

            # 按开工日期排序，如果日期相同，按行号排序（Excel从上到下）
            wo_list.sort(key=lambda x: (x['start_date'], x['row_idx']))
            
            self._log(f"查询ERP数据 (共 {len(wo_list)} 张工单)...")
            all_wo_keys = list(set([p['wo_key'] for p in wo_list]))
            static_wo_data = self._fetch_erp_data(all_wo_keys)
            
            all_parts = set()
            for w in static_wo_data.values():
                for b in w['bom']: all_parts.add(b['part'])
            
            self._log("查询库存...")
            static_inventory = self._fetch_inventory(list(all_parts))

            # 模拟环境
            running_inv = copy.deepcopy(static_inventory)
            
            # 这里的 running_wo_issued 我们其实不需要模拟扣减它，
            # 因为新逻辑是：需求 = (BOM总 - 已领)。这个“已领”是ERP里的静态值。
            # 我们只需要扣减 running_inv 即可。

            self._log("开始推演 (库存模拟扣减)...")
            results = self._simulate_logic_v2(wo_list, static_wo_data, running_inv)

            self._log("正在回写 A 列...")
            font_style = Font(name="微软雅黑", size=9)
            
            count = 0
            for r in results:
                row_idx = r['row_idx']
                # 格式：齐套率：XX%，最小可生产数：XX，缺料信息：品号,品名,缺XX单位
                rate_str = f"{r['rate']:.0%}"
                
                # 缺料信息处理
                msg = r['msg']
                # 如果 msg 是空的（完全齐套），不需要加前面的冒号
                # 但根据需求，要写“缺料信息：xxx”
                # 如果没缺料，就写“缺料信息：”或留白？通常留空比较好看
                
                detail_str = f"缺料信息：{msg}" if msg else "缺料信息："
                
                final_str = f"齐套率：{rate_str}，最小可生产数：{r['achievable']}，{detail_str}"
                
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

    def _extract_data_with_start_date(self, ws, sorted_col_indices, col_date_map, filter_ws):
        """
        遍历行，找到每行在选定范围内的'最早有数日期'
        """
        c_ws = self.col_map_main.get(COL_NAME_WORKSHOP)
        c_type = self.col_map_main.get(COL_NAME_WO_TYPE)
        c_no = self.col_map_main.get(COL_NAME_WO_NO)

        if not c_type: c_type = 5
        if not c_no: c_no = 6

        data = []
        for row in ws.iter_rows(min_row=ROW_IDX_DATA_START):
            try:
                found_start = False
                first_date = None
                
                # 检查此行在范围内是否有排产
                has_qty = False
                for col_idx in sorted_col_indices:
                    if col_idx <= len(row):
                        val = row[col_idx - 1].value
                        if isinstance(val, (int, float)) and val > 0:
                            has_qty = True
                            if not found_start:
                                found_start = True
                                first_date = col_date_map[col_idx]
                            # 找到第一个日期后，其实可以不用遍历完，但为了确认是否有数，还是遍历吧
                            # 优化：一旦找到 first_date，其实后面只需通过 has_qty 确认即可
                
                if has_qty:
                    val_ws = row[c_ws - 1].value if (c_ws and c_ws <= len(row)) else None
                    curr_ws = str(val_ws).strip() if val_ws else "未分类"
                    
                    if filter_ws != "全部车间" and curr_ws != filter_ws: continue

                    wt = row[c_type - 1].value
                    wn = row[c_no - 1].value
                    
                    if wt and wn:
                        data.append({
                            'wo_key': (str(wt).strip(), str(wn).strip()),
                            'start_date': first_date, # 用于排序
                            'row_idx': row[0].row
                        })
            except:
                continue
        return data

    def _fetch_erp_data(self, keys):
        if not keys: return {}
        conditions = [f"(TA.TA001='{t}' AND TA.TA002='{n}')" for t, n in keys]
        data = defaultdict(lambda: {'status': '', 'total': 0, 'bom': []})
        batch_size = 200
        for i in range(0, len(conditions), batch_size):
            batch = conditions[i:i + batch_size]
            # 增加查询 TA.TA011 (工单状态)
            sql = f"""
                SELECT RTRIM(TA.TA001) t, RTRIM(TA.TA002) n, TA.TA015 total, TA.TA011 status,
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
                        key = (r['t'], r['n'])
                        data[key]['total'] = float(r['total'])
                        data[key]['status'] = str(r['status']).strip()
                        data[key]['bom'].append({
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

    def _simulate_logic_v2(self, wo_list, wo_data, running_inv):
        results = []

        for item in wo_list:
            key = item['wo_key']
            row_idx = item['row_idx']
            info = wo_data.get(key)
            
            res = {
                'row_idx': row_idx,
                'rate': 0.0, 'achievable': 0, 'msg': ""
            }

            if not info or not info['bom']:
                res['msg'] = "无ERP信息"
                results.append(res)
                continue

            # 优先级 1: 工单已完工 (ERP状态为 Y 或 y)
            if info['status'].upper() == 'Y':
                res['rate'] = 1.0
                res['achievable'] = int(info['total']) # 既然完工了，可产数量就是总量
                res['msg'] = "工单已完工"
                results.append(res)
                continue

            # 开始计算 BOM 缺口
            # 需求基准：工单剩余需领量 (Total_Req - Issued)
            # 注意：这里不需要乘以工单数量，因为 TB004 (需领用量) 已经是该工单的总需领用量了
            
            wo_remaining_needs = {} # {part: amount}
            total_remaining_demand = 0
            
            for b in info['bom']:
                rem = max(0, b['req'] - b['iss'])
                if rem > 0:
                    wo_remaining_needs[b['part']] = rem
                    total_remaining_demand += rem

            # 优先级 2: 发料齐套 (无需领料)
            if total_remaining_demand == 0:
                res['rate'] = 1.0
                res['achievable'] = int(info['total'])
                res['msg'] = "发料齐套"
                results.append(res)
                continue

            # 优先级 3 & 4: 仓库齐套 或 缺料
            # 遍历 BOM 计算缺料详情，并扣减库存
            
            min_rate = 1.0
            # 最小可生产数计算：
            # 逻辑：对于每一个料，可支持套数 = (已领 + 仓库库存) // 单耗
            # 这里的单耗 = b['req'] / info['total']
            min_possible_sets = 999999
            
            short_details = []
            is_warehouse_short = False # 是否真缺料

            for b in info['bom']:
                # 单耗
                unit_use = b['req'] / info['total'] if info['total'] > 0 else 0
                if unit_use <= 0: continue
                
                # 该料剩余需求
                part_need = wo_remaining_needs.get(b['part'], 0)
                
                # 当前库存
                stock = running_inv.get(b['part'], 0)
                effective_stock = max(0, stock)
                
                # 计算齐套率 (针对剩余需求)
                if part_need > 0:
                    part_rate = effective_stock / part_need
                    if part_rate > 1.0: part_rate = 1.0
                    if part_rate < min_rate: min_rate = part_rate
                    
                    if effective_stock < part_need - 0.0001:
                        is_warehouse_short = True
                        diff = part_need - effective_stock
                        # 格式：品号,品名,缺数量单位
                        short_details.append(f"{b['part']},{b['name']},缺{diff:g}{b['unit']}")

                # 计算可产数量 (基于总需求 perspective)
                # 可用总数 = 已领 (b['iss']) + 仓库 (effective_stock)
                total_avail_material = b['iss'] + effective_stock
                can_do_sets = int(total_avail_material // unit_use)
                min_possible_sets = min(min_possible_sets, can_do_sets)

                # --- 关键：扣减库存 ---
                # 无论是否缺料，都要把该工单的需求扣掉，因为按日期排在前面的工单有优先权
                if part_need > 0:
                    if b['part'] not in running_inv: running_inv[b['part']] = 0.0
                    running_inv[b['part']] -= part_need
                    # 库存允许扣成负数吗？
                    # 逻辑上，库存被扣完就是0，后续工单看到的就是0。
                    # 如果扣成负数，方便追踪缺口，但后续工单判断 effective_stock = max(0, stock) 即可。
                    # 所以直接减是可以的。

            # 汇总结果
            res['achievable'] = min(int(info['total']), min_possible_sets)
            res['rate'] = min_rate

            if not is_warehouse_short:
                # 优先级 3: 仓库齐套
                res['msg'] = "仓库齐套"
                res['rate'] = 1.0 # 既然仓库够，就是100%
                res['achievable'] = int(info['total'])
            else:
                # 优先级 4: 缺料
                res['msg'] = "; ".join(short_details) # 分号分隔不同物料

            results.append(res)

        return results

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = DailyPlanAvailabilityApp(root)
        root.mainloop()
    except Exception as e:
        import tkinter.messagebox
        tkinter.messagebox.showerror("启动失败", str(e))
