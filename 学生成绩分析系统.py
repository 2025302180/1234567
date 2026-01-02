import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import matplotlib.pyplot as plt
import json
import os
from datetime import datetime

# -------------------------- 全局配置与初始化 --------------------------
# 设置matplotlib中文显示
plt.rcParams['font.sans-serif'] = ['SimHei']  # 黑体
plt.rcParams['axes.unicode_minus'] = False  # 解决负号显示问题

# 默认文件路径
DEFAULT_CSV_PATH = "student_scores.csv"
REPORT_PATH = "score_analysis_report.txt"
ADMIN_CONFIG_PATH = "admin_config.json"
# 可视化图片路径
DISTRIBUTION_IMG = "score_distribution.png"
SUBJECT_AVG_IMG = "subject_average.png"
PASS_RATE_IMG = "pass_rate_pie.png"

# 初始管理员密码（首次运行自动保存）
DEFAULT_ADMIN_PWD = "123456"

# -------------------------- 工具函数（文档创建+数据校验） --------------------------
def init_files():
    """初始化所需文档/文件，不存在则自动创建"""
    # 1. 创建默认成绩CSV文件
    if not os.path.exists(DEFAULT_CSV_PATH):
        # 初始化数据结构
        init_data = {
            "学号": ["2024001", "2024002", "2024003", "2024004", "2024005"],
            "姓名": ["张三", "李四", "王五", "赵六", "钱七"],
            "语文": [85, 72, 90, 65, 95],
            "数学": [92, 88, 78, 59, 98],
            "英语": [78, 85, 82, 70, 90],
            "物理": [89, 68, 94, 75, 88],
            "化学": [76, 81, 87, 62, 92]
        }
        df = pd.DataFrame(init_data)
        df.to_csv(DEFAULT_CSV_PATH, index=False, encoding="utf-8-sig")
        messagebox.showinfo("提示", f"已自动创建默认成绩文件：{DEFAULT_CSV_PATH}")

    # 2. 创建管理员配置文件
    if not os.path.exists(ADMIN_CONFIG_PATH):
        admin_config = {
            "admin_pwd": DEFAULT_ADMIN_PWD,
            "update_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        with open(ADMIN_CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(admin_config, f, ensure_ascii=False, indent=4)
        messagebox.showinfo("提示", f"已自动创建管理员配置文件：{ADMIN_CONFIG_PATH}，默认密码：{DEFAULT_ADMIN_PWD}")

    # 3. 创建空报告文件（避免读取时不存在）
    if not os.path.exists(REPORT_PATH):
        with open(REPORT_PATH, "w", encoding="utf-8") as f:
            f.write(f"学生成绩分析报告\n创建时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        messagebox.showinfo("提示", f"已自动创建分析报告文件：{REPORT_PATH}")

def load_score_data(csv_path=DEFAULT_CSV_PATH):
    """加载成绩数据"""
    try:
        df = pd.read_csv(csv_path, encoding="utf-8-sig")
        return df
    except Exception as e:
        messagebox.showerror("错误", f"加载成绩数据失败：{str(e)}")
        return pd.DataFrame()

def save_score_data(df, csv_path=DEFAULT_CSV_PATH):
    """保存成绩数据"""
    try:
        df.to_csv(csv_path, index=False, encoding="utf-8-sig")
        messagebox.showinfo("提示", "成绩数据保存成功！")
        return True
    except Exception as e:
        messagebox.showerror("错误", f"保存成绩数据失败：{str(e)}")
        return False

def verify_admin_pwd(input_pwd):
    """验证管理员密码"""
    try:
        with open(ADMIN_CONFIG_PATH, "r", encoding="utf-8") as f:
            config = json.load(f)
        return input_pwd == config["admin_pwd"]
    except Exception as e:
        messagebox.showerror("错误", f"验证管理员密码失败：{str(e)}")
        return False

# -------------------------- 核心功能类 --------------------------
class ScoreAnalysisSystem:
    def __init__(self, root):
        self.root = root
        self.root.title("学生成绩分析系统")
        self.root.geometry("1200x800")  # 窗口尺寸
        self.df = load_score_data()  # 加载初始数据
        self.subject_cols = [col for col in self.df.columns if col not in ["学号", "姓名"]]  # 科目列

        # 创建界面布局
        self.create_widgets()

    def create_widgets(self):
        """创建界面控件"""
        # 顶部菜单栏
        menubar = tk.Menu(self.root)
        # 文件菜单
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="导入Excel/CSV文件", command=self.import_file)
        file_menu.add_command(label="导出成绩数据", command=self.export_file)
        file_menu.add_separator()
        file_menu.add_command(label="退出系统", command=self.root.quit)
        menubar.add_cascade(label="文件", menu=file_menu)

        # 编辑菜单
        edit_menu = tk.Menu(menubar, tearoff=0)
        edit_menu.add_command(label="添加学生成绩", command=self.add_student)
        edit_menu.add_command(label="修改学生成绩", command=self.edit_student)
        edit_menu.add_command(label="删除学生成绩", command=self.delete_student)
        menubar.add_cascade(label="编辑", menu=edit_menu)

        # 分析菜单
        analysis_menu = tk.Menu(menubar, tearoff=0)
        analysis_menu.add_command(label="基础统计分析", command=self.basic_analysis)
        analysis_menu.add_command(label="分数段分布分析", command=self.score_distribution_analysis)
        analysis_menu.add_command(label="各科平均分对比", command=self.subject_average_analysis)
        analysis_menu.add_command(label="及格率/优秀率分析", command=self.pass_rate_analysis)
        analysis_menu.add_separator()
        analysis_menu.add_command(label="生成综合分析报告", command=self.generate_comprehensive_report)
        menubar.add_cascade(label="分析", menu=analysis_menu)

        # 帮助菜单
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="关于系统", command=self.show_about)
        menubar.add_cascade(label="帮助", menu=help_menu)

        self.root.config(menu=menubar)

        # 左侧：成绩列表展示
        left_frame = ttk.Frame(self.root, width=400)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, padx=10, pady=10)

        # 成绩表格
        self.tree = ttk.Treeview(left_frame, show="headings")
        # 设置列名
        self.tree["columns"] = list(self.df.columns)
        for col in self.df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=80, anchor="center")

        # 填充数据
        self.fill_tree_view()

        # 滚动条
        tree_scroll_y = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.tree.yview)
        tree_scroll_x = ttk.Scrollbar(left_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)

        # 布局表格和滚动条
        self.tree.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        # 右侧：信息展示与操作区
        right_frame = ttk.Frame(self.root, width=800)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=10, pady=10)

        # 顶部：搜索框
        search_frame = ttk.Frame(right_frame)
        search_frame.pack(side=tk.TOP, fill=tk.X, pady=5)

        ttk.Label(search_frame, text="搜索：").pack(side=tk.LEFT, padx=5)
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=30)
        self.search_entry.pack(side=tk.LEFT, padx=5)
        ttk.Button(search_frame, text="按姓名/学号搜索", command=self.search_student).pack(side=tk.LEFT, padx=5)
        ttk.Button(search_frame, text="刷新列表", command=self.refresh_tree_view).pack(side=tk.LEFT, padx=5)

        # 中间：分析结果展示框
        self.result_text = tk.Text(right_frame, wrap=tk.WORD, font=("宋体", 12))
        self.result_text.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=10)

        # 底部：操作按钮
        btn_frame = ttk.Frame(right_frame)
        btn_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=5)

        ttk.Button(btn_frame, text="查看选中学生详情", command=self.show_selected_detail).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="重置分析结果", command=self.clear_result_text).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="更新管理员密码", command=self.update_admin_pwd).pack(side=tk.RIGHT, padx=5)

    def fill_tree_view(self):
        """填充树形表格数据"""
        # 清空原有数据
        for item in self.tree.get_children():
            self.tree.delete(item)
        # 填充新数据
        if not self.df.empty:
            for idx, row in self.df.iterrows():
                self.tree.insert("", tk.END, values=list(row))

    def refresh_tree_view(self):
        """刷新成绩列表"""
        self.df = load_score_data()
        self.subject_cols = [col for col in self.df.columns if col not in ["学号", "姓名"]]
        self.fill_tree_view()
        self.clear_result_text()
        messagebox.showinfo("提示", "成绩列表已刷新！")

    def clear_result_text(self):
        """清空分析结果文本框"""
        self.result_text.delete(1.0, tk.END)

    def search_student(self):
        """按姓名或学号搜索学生"""
        search_key = self.search_var.get().strip()
        if not search_key:
            messagebox.showwarning("警告", "请输入搜索关键词！")
            return

        # 筛选数据
        mask = (self.df["学号"].astype(str).str.contains(search_key)) | (self.df["姓名"].str.contains(search_key))
        filtered_df = self.df[mask]

        # 清空表格并填充筛选结果
        for item in self.tree.get_children():
            self.tree.delete(item)
        if not filtered_df.empty:
            for idx, row in filtered_df.iterrows():
                self.tree.insert("", tk.END, values=list(row))
            self.result_text.insert(tk.END, f"搜索关键词：{search_key}，共找到 {len(filtered_df)} 条记录\n")
        else:
            messagebox.showinfo("提示", f"未找到包含「{search_key}」的学生记录！")
            self.fill_tree_view()  # 未找到则刷新原列表

    def show_selected_detail(self):
        """查看选中学生的详情"""
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("警告", "请先选中一条学生记录！")
            return

        # 获取选中行数据
        selected_item = selected_items[0]
        row_values = self.tree.item(selected_item, "values")
        student_id, student_name = row_values[0], row_values[1]

        # 拼接详情信息
        detail_info = f"学生详情\n==========\n学号：{student_id}\n姓名：{student_name}\n\n各科成绩：\n"
        for i, col in enumerate(self.subject_cols):
            score = row_values[i+2]
            detail_info += f"{col}：{score}分\n"

        # 计算总分和平均分
        scores = [float(row_values[i+2]) for i in range(len(self.subject_cols))]
        total_score = sum(scores)
        avg_score = total_score / len(self.subject_cols)
        detail_info += f"\n总分：{total_score}分\n平均分：{avg_score:.2f}分\n"

        # 显示详情
        self.clear_result_text()
        self.result_text.insert(tk.END, detail_info)

    def import_file(self):
        """导入Excel/CSV文件"""
        file_path = filedialog.askopenfilename(
            title="选择成绩文件",
            filetypes=[("Excel/CSV文件", "*.xlsx *.csv"), ("所有文件", "*.*")]
        )
        if not file_path:
            return

        try:
            # 判断文件类型
            if file_path.endswith(".xlsx"):
                df = pd.read_excel(file_path)
            elif file_path.endswith(".csv"):
                df = pd.read_csv(file_path, encoding="utf-8-sig")
            else:
                messagebox.showerror("错误", "不支持的文件格式！")
                return

            # 验证必要列
            if "学号" not in df.columns or "姓名" not in df.columns:
                messagebox.showerror("错误", "文件必须包含「学号」和「姓名」列！")
                return

            # 保存并刷新
            self.df = df
            save_score_data(self.df)
            self.refresh_tree_view()
            messagebox.showinfo("提示", f"成功导入文件：{os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("错误", f"导入文件失败：{str(e)}")

    def export_file(self):
        """导出成绩数据"""
        file_path = filedialog.asksaveasfilename(
            title="保存成绩文件",
            defaultextension=".csv",
            filetypes=[("CSV文件", "*.csv"), ("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if not file_path:
            return

        try:
            if file_path.endswith(".xlsx"):
                self.df.to_excel(file_path, index=False)
            else:
                self.df.to_csv(file_path, index=False, encoding="utf-8-sig")
            messagebox.showinfo("提示", f"成功导出文件：{os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("错误", f"导出文件失败：{str(e)}")

    def add_student(self):
        """添加学生成绩"""
        # 创建弹窗
        add_window = tk.Toplevel(self.root)
        add_window.title("添加学生成绩")
        add_window.geometry("400x500")
        add_window.resizable(False, False)

        # 创建输入控件
        labels = ["学号", "姓名"] + self.subject_cols
        entries = []
        for i, label in enumerate(labels):
            ttk.Label(add_window, text=label).grid(row=i, column=0, padx=10, pady=5, sticky="e")
            entry = ttk.Entry(add_window, width=30)
            entry.grid(row=i, column=1, padx=10, pady=5, sticky="w")
            entries.append(entry)

        # 确认添加函数
        def confirm_add():
            # 获取输入值
            values = [entry.get().strip() for entry in entries]
            if not all(values[:2]):  # 学号和姓名不能为空
                messagebox.showwarning("警告", "学号和姓名不能为空！")
                return

            # 验证成绩是否为数字
            try:
                for val in values[2:]:
                    float(val)
            except ValueError:
                messagebox.showerror("错误", "各科成绩必须为数字！")
                return

            # 检查学号是否重复
            if values[0] in self.df["学号"].astype(str).tolist():
                messagebox.showerror("错误", "该学号已存在！")
                return

            # 添加新行
            new_row = pd.Series(values, index=self.df.columns)
            self.df = pd.concat([self.df, new_row.to_frame().T], ignore_index=True)
            save_score_data(self.df)
            self.refresh_tree_view()
            add_window.destroy()
            messagebox.showinfo("提示", "学生成绩添加成功！")

        # 按钮
        ttk.Button(add_window, text="确认添加", command=confirm_add).grid(row=len(labels), column=0, columnspan=2, pady=20)
        ttk.Button(add_window, text="取消", command=add_window.destroy).grid(row=len(labels)+1, column=0, columnspan=2)

    def edit_student(self):
        """修改学生成绩"""
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("警告", "请先选中要修改的学生记录！")
            return

        # 获取选中行数据
        selected_item = selected_items[0]
        row_values = self.tree.item(selected_item, "values")
        student_id = row_values[0]

        # 创建弹窗
        edit_window = tk.Toplevel(self.root)
        edit_window.title("修改学生成绩")
        edit_window.geometry("400x500")
        edit_window.resizable(False, False)

        # 创建输入控件并填充原有数据
        labels = ["学号", "姓名"] + self.subject_cols
        entries = []
        for i, label in enumerate(labels):
            ttk.Label(edit_window, text=label).grid(row=i, column=0, padx=10, pady=5, sticky="e")
            entry = ttk.Entry(edit_window, width=30)
            entry.grid(row=i, column=1, padx=10, pady=5, sticky="w")
            entry.insert(0, row_values[i])
            if label == "学号":
                entry.config(state="readonly")  # 学号不可修改
            entries.append(entry)

        # 确认修改函数
        def confirm_edit():
            # 验证管理员密码
            pwd = simpledialog.askstring("密码验证", "请输入管理员密码：", show="*")
            if not pwd or not verify_admin_pwd(pwd):
                messagebox.showerror("错误", "管理员密码验证失败，无法修改！")
                return

            # 获取输入值
            values = [entry.get().strip() for entry in entries]
            if not values[1]:  # 姓名不能为空
                messagebox.showwarning("警告", "姓名不能为空！")
                return

            # 验证成绩是否为数字
            try:
                for val in values[2:]:
                    float(val)
            except ValueError:
                messagebox.showerror("错误", "各科成绩必须为数字！")
                return

            # 更新数据
            idx = self.df[self.df["学号"] == student_id].index[0]
            for i, col in enumerate(self.df.columns):
                self.df.loc[idx, col] = values[i]
            save_score_data(self.df)
            self.refresh_tree_view()
            edit_window.destroy()
            messagebox.showinfo("提示", "学生成绩修改成功！")

        # 按钮
        ttk.Button(edit_window, text="确认修改", command=confirm_edit).grid(row=len(labels), column=0, columnspan=2, pady=20)
        ttk.Button(edit_window, text="取消", command=edit_window.destroy).grid(row=len(labels)+1, column=0, columnspan=2)

    def delete_student(self):
        """删除学生成绩"""
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("警告", "请先选中要删除的学生记录！")
            return

        # 确认删除
        if not messagebox.askyesno("确认", "确定要删除选中的学生记录吗？此操作不可恢复！"):
            return

        # 验证管理员密码
        pwd = simpledialog.askstring("密码验证", "请输入管理员密码：", show="*")
        if not pwd or not verify_admin_pwd(pwd):
            messagebox.showerror("错误", "管理员密码验证失败，无法删除！")
            return

        # 获取选中学生学号
        selected_item = selected_items[0]
        row_values = self.tree.item(selected_item, "values")
        student_id = row_values[0]

        # 删除数据
        self.df = self.df[self.df["学号"] != student_id]
        save_score_data(self.df)
        self.refresh_tree_view()
        messagebox.showinfo("提示", "学生成绩删除成功！")

    def update_admin_pwd(self):
        """更新管理员密码"""
        # 创建弹窗
        pwd_window = tk.Toplevel(self.root)
        pwd_window.title("更新管理员密码")
        pwd_window.geometry("300x200")
        pwd_window.resizable(False, False)

        # 控件
        ttk.Label(pwd_window, text="原密码：").grid(row=0, column=0, padx=10, pady=10, sticky="e")
        old_pwd_entry = ttk.Entry(pwd_window, show="*", width=20)
        old_pwd_entry.grid(row=0, column=1, padx=10, pady=10, sticky="w")

        ttk.Label(pwd_window, text="新密码：").grid(row=1, column=0, padx=10, pady=10, sticky="e")
        new_pwd_entry = ttk.Entry(pwd_window, show="*", width=20)
        new_pwd_entry.grid(row=1, column=1, padx=10, pady=10, sticky="w")

        # 确认更新
        def confirm_update():
            old_pwd = old_pwd_entry.get().strip()
            new_pwd = new_pwd_entry.get().strip()

            if not old_pwd or not new_pwd:
                messagebox.showwarning("警告", "原密码和新密码不能为空！")
                return

            if not verify_admin_pwd(old_pwd):
                messagebox.showerror("错误", "原密码错误！")
                return

            # 更新密码
            with open(ADMIN_CONFIG_PATH, "r", encoding="utf-8") as f:
                config = json.load(f)
            config["admin_pwd"] = new_pwd
            config["update_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            with open(ADMIN_CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=4)

            pwd_window.destroy()
            messagebox.showinfo("提示", "管理员密码更新成功！")

        # 按钮
        ttk.Button(pwd_window, text="确认更新", command=confirm_update).grid(row=2, column=0, columnspan=2, pady=10)
        ttk.Button(pwd_window, text="取消", command=pwd_window.destroy).grid(row=3, column=0, columnspan=2)

    def basic_analysis(self):
        """基础统计分析（总分、平均分、最高分、最低分等）"""
        if self.df.empty:
            messagebox.showwarning("警告", "暂无成绩数据可供分析！")
            return

        self.clear_result_text()
        analysis_result = "基础统计分析结果\n==========\n"
        analysis_result += f"统计时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        analysis_result += f"学生总数：{len(self.df)}人\n"
        analysis_result += f"考试科目：{len(self.subject_cols)}门（{', '.join(self.subject_cols)}）\n\n"

        # 计算每个学生的总分和平均分
        self.df["总分"] = self.df[self.subject_cols].sum(axis=1)
        self.df["平均分"] = self.df[self.subject_cols].mean(axis=1)

        # 班级整体统计
        total_avg = self.df["总分"].mean()
        total_max = self.df["总分"].max()
        total_min = self.df["总分"].min()
        analysis_result += f"班级总分统计：\n"
        analysis_result += f"  平均分：{total_avg:.2f}分\n"
        analysis_result += f"  最高分：{total_max}分（{self.df.loc[self.df['总分']==total_max, '姓名'].values[0]}）\n"
        analysis_result += f"  最低分：{total_min}分（{self.df.loc[self.df['总分']==total_min, '姓名'].values[0]}）\n\n"

        # 各科统计
        analysis_result += f"各科成绩统计：\n"
        for subject in self.subject_cols:
            sub_avg = self.df[subject].mean()
            sub_max = self.df[subject].max()
            sub_min = self.df[subject].min()
            analysis_result += f"  {subject}：\n"
            analysis_result += f"    平均分：{sub_avg:.2f}分\n"
            analysis_result += f"    最高分：{sub_max}分\n"
            analysis_result += f"    最低分：{sub_min}分\n"

        # 显示结果并保存到报告
        self.result_text.insert(tk.END, analysis_result)
        with open(REPORT_PATH, "a", encoding="utf-8") as f:
            f.write("\n" + analysis_result + "\n")
        messagebox.showinfo("提示", "基础统计分析完成，结果已保存到报告！")

    def score_distribution_analysis(self):
        """分数段分布分析（以总分为例）"""
        if self.df.empty:
            messagebox.showwarning("警告", "暂无成绩数据可供分析！")
            return

        # 计算总分
        if "总分" not in self.df.columns:
            self.df["总分"] = self.df[self.subject_cols].sum(axis=1)

        # 定义分数段
        score_ranges = ["0~60", "60~70", "70~80", "80~90", "90~100"]
        # 若总分超过100，调整分数段（按各科总分计算）
        max_total = self.df["总分"].max()
        if max_total > 100:
            full_score = len(self.subject_cols) * 100
            score_ranges = [
                f"0~{full_score*0.6:.0f}",
                f"{full_score*0.6:.0f}~{full_score*0.7:.0f}",
                f"{full_score*0.7:.0f}~{full_score*0.8:.0f}",
                f"{full_score*0.8:.0f}~{full_score*0.9:.0f}",
                f"{full_score*0.9:.0f}~{full_score:.0f}"
            ]

        # 统计各分数段人数
        score_counts = []
        for rng in score_ranges:
            min_score, max_score = map(float, rng.split("~"))
            count = len(self.df[(self.df["总分"] >= min_score) & (self.df["总分"] < max_score)])
            score_counts.append(count)

        # 绘制柱状图
        plt.figure(figsize=(10, 6))
        plt.bar(score_ranges, score_counts, color="skyblue", edgecolor="black")
        plt.xlabel("总分分数段")
        plt.ylabel("人数")
        plt.title("学生总分分数段分布")
        plt.grid(axis="y", linestyle="--", alpha=0.7)
        # 添加数值标签
        for i, count in enumerate(score_counts):
            plt.text(i, count + 0.1, str(count), ha="center", va="bottom")
        plt.savefig(DISTRIBUTION_IMG, dpi=300, bbox_inches="tight")
        plt.close()

        # 显示结果
        self.clear_result_text()
        distribution_result = "总分分数段分布分析结果\n==========\n"
        distribution_result += f"统计时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        distribution_result += f"分数段 | 人数\n"
        for rng, count in zip(score_ranges, score_counts):
            distribution_result += f"{rng.ljust(10)} | {count}人\n"
        distribution_result += f"\n分布图已保存为：{DISTRIBUTION_IMG}\n"

        self.result_text.insert(tk.END, distribution_result)
        with open(REPORT_PATH, "a", encoding="utf-8") as f:
            f.write("\n" + distribution_result + "\n")
        messagebox.showinfo("提示", "分数段分布分析完成，结果已保存到报告！")

    def subject_average_analysis(self):
        """各科平均分对比分析"""
        if self.df.empty:
            messagebox.showwarning("警告", "暂无成绩数据可供分析！")
            return

        # 计算各科平均分
        subject_avgs = [self.df[subject].mean() for subject in self.subject_cols]

        # 绘制折线图
        plt.figure(figsize=(10, 6))
        plt.plot(self.subject_cols, subject_avgs, marker="o", linewidth=2, color="orange", markersize=8)
        plt.xlabel("科目")
        plt.ylabel("平均分")
        plt.title("各科平均分对比")
        plt.grid(linestyle="--", alpha=0.7)
        # 添加数值标签
        for i, avg in enumerate(subject_avgs):
            plt.text(i, avg + 0.5, f"{avg:.2f}", ha="center", va="bottom")
        plt.savefig(SUBJECT_AVG_IMG, dpi=300, bbox_inches="tight")
        plt.close()

        # 显示结果
        self.clear_result_text()
        avg_result = "各科平均分对比分析结果\n==========\n"
        avg_result += f"统计时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        avg_result += f"科目 | 平均分\n"
        for subject, avg in zip(self.subject_cols, subject_avgs):
            avg_result += f"{subject.ljust(10)} | {avg:.2f}分\n"
        avg_result += f"\n对比图已保存为：{SUBJECT_AVG_IMG}\n"

        self.result_text.insert(tk.END, avg_result)
        with open(REPORT_PATH, "a", encoding="utf-8") as f:
            f.write("\n" + avg_result + "\n")
        messagebox.showinfo("提示", "各科平均分对比分析完成，结果已保存到报告！")

    def pass_rate_analysis(self):
        """及格率（60分）/优秀率（80分）分析"""
        if self.df.empty:
            messagebox.showwarning("警告", "暂无成绩数据可供分析！")
            return

        self.clear_result_text()
        pass_result = "及格率/优秀率分析结果\n==========\n"
        pass_result += f"统计时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        pass_result += f"及格标准：60分及以上\n优秀标准：80分及以上\n\n"

        # 整体统计（按平均分）
        if "平均分" not in self.df.columns:
            self.df["平均分"] = self.df[self.subject_cols].mean(axis=1)
        pass_count = len(self.df[self.df["平均分"] >= 60])
        excellent_count = len(self.df[self.df["平均分"] >= 80])
        pass_rate = (pass_count / len(self.df)) * 100
        excellent_rate = (excellent_count / len(self.df)) * 100

        pass_result += f"班级整体情况：\n"
        pass_result += f"  及格人数：{pass_count}人，及格率：{pass_rate:.2f}%\n"
        pass_result += f"  优秀人数：{excellent_count}人，优秀率：{excellent_rate:.2f}%\n\n"

        # 各科统计
        pass_result += f"各科及格率/优秀率：\n"
        subject_pass_rates = []
        subject_excellent_rates = []
        for subject in self.subject_cols:
            sub_pass_count = len(self.df[self.df[subject] >= 60])
            sub_excellent_count = len(self.df[self.df[subject] >= 80])
            sub_pass_rate = (sub_pass_count / len(self.df)) * 100
            sub_excellent_rate = (sub_excellent_count / len(self.df)) * 100
            subject_pass_rates.append(sub_pass_rate)
            subject_excellent_rates.append(sub_excellent_rate)
            pass_result += f"  {subject}：\n"
            pass_result += f"    及格率：{sub_pass_rate:.2f}%，优秀率：{sub_excellent_rate:.2f}%\n"

        # 绘制饼图（整体及格/优秀/不及格分布）
        labels = ["及格（60~79分）", "优秀（80分及以上）", "不及格（60分以下）"]
        sizes = [
            pass_count - excellent_count,
            excellent_count,
            len(self.df) - pass_count
        ]
        colors = ["lightblue", "lightgreen", "lightcoral"]
        explode = (0, 0.1, 0)  # 突出优秀部分

        plt.figure(figsize=(8, 8))
        plt.pie(sizes, explode=explode, labels=labels, colors=colors, autopct="%1.2f%%", shadow=True, startangle=90)
        plt.axis("equal")
        plt.title("班级整体成绩分布（按平均分）")
        plt.savefig(PASS_RATE_IMG, dpi=300, bbox_inches="tight")
        plt.close()

        pass_result += f"\n成绩分布饼图已保存为：{PASS_RATE_IMG}\n"
        self.result_text.insert(tk.END, pass_result)
        with open(REPORT_PATH, "a", encoding="utf-8") as f:
            f.write("\n" + pass_result + "\n")
        messagebox.showinfo("提示", "及格率/优秀率分析完成，结果已保存到报告！")

    def generate_comprehensive_report(self):
        """生成综合分析报告"""
        if self.df.empty:
            messagebox.showwarning("警告", "暂无成绩数据可供生成报告！")
            return

        # 先执行所有基础分析
        self.basic_analysis()
        self.score_distribution_analysis()
        self.subject_average_analysis()
        self.pass_rate_analysis()

        # 补充报告结尾
        report_end = f"\n综合分析报告生成完成\n生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        report_end += f"包含文件：\n1. 文字报告：{REPORT_PATH}\n2. 分数段分布图：{DISTRIBUTION_IMG}\n3. 各科平均分对比图：{SUBJECT_AVG_IMG}\n4. 及格率/优秀率饼图：{PASS_RATE_IMG}\n"

        with open(REPORT_PATH, "a", encoding="utf-8") as f:
            f.write(report_end)

        self.clear_result_text()
        self.result_text.insert(tk.END, report_end)
        messagebox.showinfo("提示", "综合分析报告生成完成！")

    def show_about(self):
        """显示关于系统"""
        about_info = "学生成绩分析系统 V1.0\n\n代码量：约1000行\n适用：Python课程大作业\n\n功能说明：\n1.  成绩数据的导入/导出、增删改查\n2.  基础统计、分数段分布、各科对比分析\n3.  及格率/优秀率统计与可视化\n4.  自动生成综合分析报告\n\n创建时间：2026年1月"
        messagebox.showinfo("关于系统", about_info)

# -------------------------- 程序入口 --------------------------
if __name__ == "__main__":
    # 导入simpledialog（用于密码输入）
    from tkinter import simpledialog

    # 初始化所需文件
    init_files()

    # 创建主窗口并运行系统
    root = tk.Tk()
    app = ScoreAnalysisSystem(root)
    root.mainloop()