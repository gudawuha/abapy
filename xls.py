import tkinter as tk
import tkinter.ttk as ttk
from tkinter import *
from tkinter import filedialog
from matplotlib.backends.backend_tkagg import (
	FigureCanvasTkAgg, NavigationToolbar2Tk)
from matplotlib.figure import Figure
import numpy as np
import pandas as pd
import tkinter.messagebox
import matplotlib as mpl
import matplotlib.pyplot as plt
mpl.use('TkAgg')
plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False


class ExcelEditor(tk.Frame):
	def __init__(self, parent=None):
		tk.Frame.__init__(self, parent)
		self.parent = parent
		self.grid()
		self.create_widgets()

	def create_widgets(self):
		self.open_btn = ttk.Button(self, text="Open Excel", command=self.open_excel)
		self.open_btn.grid(row=0, column=0)

		self.open_btn = ttk.Button(self, text="Backbone Curve", command=self.backbone_curve)
		self.open_btn.grid(row=0, column=1)

		self.shear_btn = ttk.Button(self, text="Beam shear&moment", command=self.beam_shear_moment)
		self.shear_btn.grid(row=0, column=2)

		self.table_frame = ttk.Frame(self)
		self.table_frame.grid(row=2, column=0, columnspan=3)

		self.fig_bone = Figure(figsize=(5, 5))
		self.canvas_bone = FigureCanvasTkAgg(self.fig_bone, self.table_frame)
		self.canvas_bone.get_tk_widget().grid(row=1, column=1)
		self.toolbar_bone = NavigationToolbar2Tk(self.canvas_bone, self.table_frame, pack_toolbar=False)
		self.toolbar_bone.update()
		self.toolbar_bone.grid(row=3, column=0, columnspan=2)

		self.fig = Figure(figsize=(7, 5))
		self.canvas = FigureCanvasTkAgg(self.fig, self.table_frame)
		self.canvas.get_tk_widget().grid(row=1, column=3)
		self.toolbar = NavigationToolbar2Tk(self.canvas, self.table_frame, pack_toolbar=False)
		self.toolbar.update()
		self.toolbar.grid(row=3, column=3, columnspan=2)

	def open_excel(self):
		file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls")])
		if file_path:
			self.load_excel(file_path)

	def load_excel(self, file_path):
		self.df = pd.read_excel(file_path, header=1)
		v = IntVar()
		self.s = tk.Scale(self, from_=1, to=len(self.df), resolution=1, orient=tk.HORIZONTAL, length=400, variable=v)
		self.s.grid(row=1, column=2)

	def backbone_curve(self):
		disp_value = self.df.iloc[:, 1]
		rf_value = self.df.iloc[:, 14]
		d_v = np.array(disp_value)
		r_v = np.array(rf_value)
		r_v = [x / 1000 for x in r_v]
		count = [0 for i in range(0, (len(d_v) - 1))]
		for i in range(len(d_v)-2):
			count[i] = np.sign((d_v[i]-d_v[i+1])*(d_v[i+1]-d_v[i+2]))
		count_array = np.array(count)
		peak = np.where(count_array < 0)
		peak = [x + 1 for x in peak]
		peak = np.array(peak)
		peak_list = peak.flatten()  #变为一维
		d_v_peak = disp_value[peak_list]   #提取位移峰值
		rf_v_peak = rf_value[peak_list]   #提取对应的反力
		d_v_peak_c = np.array(d_v_peak)[:, np.newaxis]   #新轴
		rf_v_peak_c = np.array(rf_v_peak)[:, np.newaxis]   #新轴
		bone_curve = np.hstack((d_v_peak_c, rf_v_peak_c))   #合并成二维矩阵
		bone_curve1 = bone_curve[np.lexsort(bone_curve[:, ::-1].T)]   #排序,按第一列
		j = 0
		a = bone_curve1[0][0]
		c = bone_curve1[0][1]
		d = bone_curve1[-1][0]
		temp_curve = bone_curve1
		for i in range(1, bone_curve1.shape[0]):
			if abs(bone_curve1[i][0] - bone_curve1[i - 1][0]) <= 0.002 and bone_curve1[i][0] != d:
				if abs(bone_curve1[i][1]) >= abs(max(bone_curve1[i - 1][1], c, key=abs)):
					a = bone_curve1[i][0]
					c = bone_curve1[i][1]
				else:
					c = max(bone_curve1[i - 1][1], c, key=abs)
					a = bone_curve1[i][0]
			elif bone_curve1[i][0] == d:
				temp_curve[j][0] = a
				temp_curve[j][1] = max(bone_curve1[i][1], c, key=abs)
				j = j + 1
			else:
				temp_curve[j][0] = a
				temp_curve[j][1] = c
				j = j + 1
				c = bone_curve1[i][1]
				i = i + 1
		bone_C = temp_curve[0:j]
		bone_C = np.append(bone_C, [[0, 0]], axis=0)
		bone_C = bone_C[np.lexsort(bone_C[:, ::-1].T)]   #排序,按第一列
		bone_C[:, 1] = [x / 1000 for x in bone_C[:, 1]]
		axb_s = self.fig_bone.add_subplot(111)
		axb_s.set_title(f"bone")
		axb_s.set_ylabel('kN', labelpad=-30, y=1.02, rotation=0)  # 调整y轴标签与y轴的距离，#调整y轴标签的上下位置
		axb_s.set_xlabel('disp')
		axb_s.grid(linestyle='-.')
		axb_s.plot(bone_C[:, 0], bone_C[:, 1], label="骨架曲线", color='r', marker='^')
		axb_s.plot(d_v, r_v, label="滞回曲线", color='yellowgreen')
		axb_s.legend()
		self.fig_bone.set_tight_layout(True)
		self.canvas_bone.draw()
		self.update()

	def beam_shear_moment(self):
		self.fig.clear()
		step = self.s.get()
		step1 = int(step)
		if step1 <= 0:
			tk.messagebox.showerror(title='Error', message='填入正数')
		elif step1 <= len(self.df):
			frame = self.df.iloc[step1 - 1, 0]
			time = np.array(frame)
			t1 = np.round(time, 3)
			cell_value_s = self.df.iloc[step1 - 1, 2:8]     #剪力值
			cell_value_m = self.df.iloc[step1 - 1, 7:13]     #弯矩值
			beam_position = np.array(cell_value_s.index)
			beam_shear = np.array(cell_value_s)
			beam_moment = np.array(cell_value_m)
			beam_shear = [x / 1000 for x in beam_shear]     #转化成kN
			beam_moment = [x / 1000 for x in beam_moment]     #转化成kN*m
			ax_s = self.fig.add_subplot(121)
			ax_s.set_title(f"Time= {t1}")
			ax_s.set_ylabel('kN', labelpad=-30, y=1.02, rotation=0)    #调整y轴标签与y轴的距离，#调整y轴标签的上下位置
			ax_s.set_xlabel('Shear position')
			ax_s.grid(linestyle='-.')
			ax_s.plot(beam_position, beam_shear, marker='s')
			ax_m = self.fig.add_subplot(122)
			ax_m.set_title(f"Time= {t1}")
			ax_m.set_ylabel('kN*m', labelpad=-30, y=1.02, rotation=0)
			ax_m.set_xlabel('Moment position')
			ax_m.grid(linestyle='-.')
			ax_m.get_tightbbox()
			self.fig.set_tight_layout(True)
			ax_m.plot(beam_position, beam_moment, marker='s')
			self.canvas.draw()
			self.update()
		else:
			tk.messagebox.showerror(title='Error', message='超出范围')

if __name__ == "__main__":
	root = tk.Tk()
	root.title("Excel Editor")
	root.geometry('1250x650')
	ExcelEditor(root)
	root.mainloop()
