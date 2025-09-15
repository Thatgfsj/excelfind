import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import datetime
from pathlib import Path
import threading

class a:
    def __init__(self, b):
        self.b = b
        self.b.title("表格查找软件.by23政")
        self.b.geometry("800x720")
        self.c = tk.StringVar()
        self.d = tk.StringVar()
        self.e = tk.StringVar()
        self.f = ['.xlsx', '.xls', '.csv']
        self.g = ['.zip', '.rar', '.7z']
        self.h()
        
    def h(self):
        i = ttk.Frame(self.b, padding="10")
        i.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.b.columnconfigure(0, weight=1)
        self.b.rowconfigure(0, weight=1)
        i.columnconfigure(1, weight=1)
        ttk.Label(i, text="选择加分总文件夹:").grid(row=0, column=0, sticky=tk.W, pady=6)
        ttk.Entry(i, textvariable=self.c, width=80).grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5, padx=(5, 0))
        ttk.Button(i, text="浏览", command=self.j).grid(row=0, column=2, pady=5, padx=(5, 0))
        ttk.Label(i, text="选择存放文件夹:").grid(row=1, column=0, sticky=tk.W, pady=6)
        ttk.Entry(i, textvariable=self.d, width=80).grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5, padx=(5, 0))
        ttk.Button(i, text="浏览", command=self.k).grid(row=1, column=2, pady=5, padx=(5, 0))
        ttk.Label(i, text="查找的信息:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(i, textvariable=self.e, width=50).grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5, padx=(5, 0))
        l = ttk.Frame(i)
        l.grid(row=3, column=0, columnspan=3, pady=10)
        self.m = ttk.Button(l, text="开始查找", command=self.n)
        self.m.pack(side=tk.LEFT, padx=5)
        self.o = ttk.Button(l, text="取消", command=self.p, state=tk.DISABLED)
        self.o.pack(side=tk.LEFT, padx=5)
        self.q = tk.DoubleVar()
        self.r = ttk.Progressbar(i, variable=self.q, maximum=100)
        self.r.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        self.s = ttk.Label(i, text="请选择文件夹并输入查找内容")
        self.s.grid(row=5, column=0, columnspan=3, pady=5)
        ttk.Label(i, text="目录预览:").grid(row=6, column=0, sticky=tk.W, pady=(10, 5))
        t = ttk.Frame(i)
        t.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        t.columnconfigure(0, weight=1)
        t.rowconfigure(0, weight=1)
        self.u = ttk.Treeview(t)
        self.u.heading('#0', text='文件结构')
        v = ttk.Scrollbar(t, orient=tk.VERTICAL, command=self.u.yview)
        w = ttk.Scrollbar(t, orient=tk.HORIZONTAL, command=self.u.xview)
        self.u.configure(yscrollcommand=v.set, xscrollcommand=w.set)
        self.u.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        v.grid(row=0, column=1, sticky=(tk.N, tk.S))
        w.grid(row=1, column=0, sticky=(tk.W, tk.E))
        i.rowconfigure(7, weight=1)
        self.c.trace('w', self.x)
        self.y = False
        
    def z(self, aa):
        for ab, ac, ad in os.walk(aa):
            for ae in ad:
                if any(ae.lower().endswith(af) for af in self.g):
                    return True
        return False

    def j(self):
        af = filedialog.askdirectory()
        if af:
            self.c.set(af)
            
    def k(self):
        ag = filedialog.askdirectory()
        if ag:
            self.d.set(ag)
            
    def x(self, *ah):
        self.u.delete(*self.u.get_children())
        ai = self.c.get()
        if not ai or not os.path.exists(ai):
            return
        aj = self.u.insert('', 'end', text=os.path.basename(ai), open=True)
        self.ak(aj, ai)
        
    def ak(self, al, am):
        try:
            an = sorted(os.listdir(am))
            for ao in an:
                ap = os.path.join(am, ao)
                if os.path.isdir(ap):
                    aq = self.u.insert(al, 'end', text=ao, open=False)
                    self.ak(aq, ap)
                elif os.path.isfile(ap):
                    if any(ao.lower().endswith(ar) for ar in self.f):
                        self.u.insert(al, 'end', text=ao)
        except Exception as e:
            print(f"读取目录时出错: {e}")
            
    def n(self):
        if not self.c.get():
            messagebox.showerror("错误", "请选择源文件夹")
            return
        if not self.d.get():
            messagebox.showerror("错误", "请选择目标文件夹")
            return
        if not self.e.get():
            messagebox.showerror("错误", "请输入查找内容")
            return
        if not os.path.exists(self.c.get()):
            messagebox.showerror("错误", "源文件夹不存在")
            return
        if not os.path.exists(self.d.get()):
            messagebox.showerror("错误", "目标文件夹不存在")
            return
        if self.z(self.c.get()):
            messagebox.showwarning("警告", "文件夹内包含压缩文件（如 zip/rar），可能会出现信息丢失，请手动解压后再进行查找！")
            return
        self.m.config(state=tk.DISABLED)
        self.o.config(state=tk.NORMAL)
        self.y = False
        at = threading.Thread(target=self.au)
        at.daemon = True
        at.start()
        
    def p(self):
        self.y = True
        self.s.config(text="正在取消搜索...")
        
    def au(self):
        try:
            av = self.c.get()
            aw = self.d.get()
            ax = self.e.get()
            ay = self.az(av)
            az = len(ay)
            if az == 0:
                self.b.after(0, lambda: messagebox.showinfo("提示", "未找到支持的表格文件"))
                self.b.after(0, self.ba)
                return
            bb = []
            os.makedirs(aw, exist_ok=True)
            bc = 0
            bd = 0
            for be in ay:
                if self.y:
                    break
                try:
                    bc += 1
                    bf = (bc / az) * 100
                    self.b.after(0, lambda p=bf: self.q.set(p))
                    self.b.after(0, lambda f=bc, t=az: self.s.config(text=f"正在处理: {f}/{t}"))
                    if be.lower().endswith('.csv'):
                        try:
                            bg = pd.read_csv(be, encoding='utf-8')
                        except:
                            bg = pd.read_csv(be, encoding='gbk')
                        if self.bh(bg, ax):
                            bi = os.path.join(aw, os.path.basename(be))
                            if not os.path.exists(bi):
                                import shutil
                                shutil.copy2(be, bi)
                                bd += 1
                            bj = self.bk(bg, ax)
                            for _, bl in bj.iterrows():
                                bb.append({
                                    '文件名': os.path.basename(be),
                                    '工作表名': 'Sheet1',
                                    '匹配内容': '; '.join([str(bm) for bm in bl.values if pd.notna(bm)])
                                })
                    else:
                        bn = pd.ExcelFile(be)
                        for bo in bn.sheet_names:
                            if self.y:
                                break
                            bp = pd.read_excel(be, sheet_name=bo)
                            if self.bh(bp, ax):
                                bq = os.path.join(aw, os.path.basename(be))
                                if not os.path.exists(bq):
                                    import shutil
                                    shutil.copy2(be, bq)
                                    bd += 1
                                br = self.bk(bp, ax)
                                for _, bs in br.iterrows():
                                    bb.append({
                                        '文件名': os.path.basename(be),
                                        '工作表名': bo,
                                        '匹配内容': '; '.join([str(bt) for bt in bs.values if pd.notna(bt)])
                                    })
                except Exception as e:
                    print(f"处理文件 {be} 时出错: {e}")
                    continue
            if self.y:
                self.b.after(0, lambda: messagebox.showinfo("提示", "搜索已取消"))
            else:
                if bb:
                    bu = pd.DataFrame(bb)
                    bv = datetime.datetime.now().strftime("%Y%m%d")
                    bw = f"{ax}表格汇总{bv}.xlsx"
                    bx = os.path.join(aw, bw)
                    bu.to_excel(bx, index=False)
                    self.b.after(0, lambda: messagebox.showinfo("完成", f"搜索完成！\n找到匹配文件: {bd} 个\n生成汇总表格: {bw}"))
                else:
                    self.b.after(0, lambda: messagebox.showinfo("完成", f"搜索完成！\n未找到包含 '{ax}' 的表格"))
        except Exception as e:
            self.b.after(0, lambda: messagebox.showerror("错误", f"搜索过程中出现错误:\n{str(e)}"))
        finally:
            self.b.after(0, self.ba)
            
    def ba(self):
        self.m.config(state=tk.NORMAL)
        self.o.config(state=tk.DISABLED)
        self.s.config(text="搜索完成")
        
    def az(self, by):
        bz = []
        for ca, cb, cc in os.walk(by):
            for cd in cc:
                if any(cd.lower().endswith(ce) for ce in self.f):
                    bz.append(os.path.join(ca, cd))
        return bz
        
    def bh(self, cf, cg):
        for ch in cf.columns:
            if cf[ch].astype(str).str.contains(cg, case=False, na=False).any():
                return True
        return False
        
    def bk(self, ci, cj):
        ck = pd.Series([False] * len(ci))
        for cl in ci.columns:
            ck |= ci[cl].astype(str).str.contains(cj, case=False, na=False)
        return ci[ck]

def cm():
    cn = tk.Tk()
    co = a(cn)
    cn.mainloop()

if __name__ == "__main__":
    cm()