# coding=utf-8
import tkinter as tk
import re
import xlsxwriter
import tkinter.ttk as ttk
from tkinter import filedialog
import numpy as np
import pandas as pd
from tkinter import messagebox
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
		self.open_btn = ttk.Button(self, text="打开wmass.out", command=self.open_wmass)
		self.open_btn.grid(row=0, column=0)

		self.open_btn = ttk.Button(self, text="生成xls", command=self.Stiffness_Coordinate)
		self.open_btn.grid(row=0, column=1)

		self.table_frame = ttk.Frame(self)
		self.table_frame.grid(row=2, column=0, columnspan=3)

	def open_wmass(self):
		file_path = filedialog.askopenfilename(filetypes=[("yjkout files", "*.out")])
		if file_path:
			self.load_wmass(file_path)

	def load_wmass(self, file_path):
		self.df = pd.read_csv(file_path, sep="\t", encoding="GB18030", header=None)
		self.excelt = file_path[:-9]
		self.wdisp = self.excelt + 'wdisp.out'
		self.feadat = file_path[:-14]
		self.feapath = self.feadat + 'fea.dat'
		self.feadf = pd.read_csv(self.feapath, sep="[", encoding="GB18030", header=None)

	def Stiffness_Coordinate(self):
		wmass = {}
		story = 0
		out = {}
		tower = {}
		for a in range(len(self.df[0])):
			if self.df[0][a].startswith("  Floor No."):
				story = story + 1
				continue
			if self.df[0][a].startswith("  Xstif="):
				wmass[story] = self.df[0][a]
		for b in range(story):
			out[b] = re.split(r'[=(]', wmass[b + 1])  # 按=号，号，\t分段
		start_tower = '                                 楼层属性'
		end_tower = '                           塔属性'
		start_tower_row = self.df[self.df[0] == start_tower].index[0]
		end_tower_row = self.df[self.df[0] == end_tower].index[0]
		for aa in range(start_tower_row + 1, end_tower_row):  # y+偶然偏心
			tower[aa] = self.df[0][aa].split()
		tower = list(tower.values())
		tower =tower[2:]
		#print(tower)
		self.nstr = self.excelt + 'torsional-angle.xls'
		self.wb = xlsxwriter.Workbook(self.nstr)  # 创建excel
		self.wxp = self.wb.add_worksheet('X+偶然偏心')  # 创建sheet
		self.wxm = self.wb.add_worksheet('X-偶然偏心')  # 创建sheet
		self.wyp = self.wb.add_worksheet('Y+偶然偏心')  # 创建sheet
		self.wym = self.wb.add_worksheet('Y-偶然偏心')  # 创建sheet
		self.wxwind = self.wb.add_worksheet('X风')  # 创建sheet
		self.wywind = self.wb.add_worksheet('Y风')  # 创建sheet
		self.wss = self.wb.add_worksheet('刚心坐标')  # 创建sheet
		self.wss.write(0, 0, 'Xstif')
		self.wss.write(0, 1, 'Ystif')
		self.wxp.write(0, 0, 'JmaxD')
		self.wxp.write(0, 1, 'Max-Dx')
		self.wxp.write(0, 2, 'Ave-Dx')
		self.wxp.write(0, 3, 'Ratio-Dx')
		self.wxp.write(0, 4, 'torsional-angle')
		self.wxm.write(0, 0, 'JmaxD')
		self.wxm.write(0, 1, 'Max-Dx')
		self.wxm.write(0, 2, 'Ave-Dx')
		self.wxm.write(0, 3, 'Ratio-Dx')
		self.wxm.write(0, 4, 'torsional-angle')
		self.wyp.write(0, 0, 'JmaxD')
		self.wyp.write(0, 1, 'Max-Dy')
		self.wyp.write(0, 2, 'Ave-Dy')
		self.wyp.write(0, 3, 'Ratio-Dy')
		self.wyp.write(0, 4, 'torsional-angle')
		self.wym.write(0, 0, 'JmaxD')
		self.wym.write(0, 1, 'Max-Dy')
		self.wym.write(0, 2, 'Ave-Dy')
		self.wym.write(0, 3, 'Ratio-Dy')
		self.wym.write(0, 0, 'torsional-angle')
		self.wxwind.write(0, 0, 'JmaxD')
		self.wxwind.write(0, 1, 'Max-Dy')
		self.wxwind.write(0, 2, 'Ave-Dy')
		self.wxwind.write(0, 3, 'Ratio-Dy')
		self.wxwind.write(0, 0, 'torsional-angle')
		self.wywind.write(0, 0, 'JmaxD')
		self.wywind.write(0, 1, 'Max-Dy')
		self.wywind.write(0, 2, 'Ave-Dy')
		self.wywind.write(0, 3, 'Ratio-Dy')
		self.wywind.write(0, 0, 'torsional-angle')
		for e in range(story):  # 输出到excel
			self.wss.write(e + 1, 0, np.array(out[e][1]))	# 输出节点X坐标
			self.wss.write(e + 1, 1, np.array(out[e][3]))	# 输出节点Y坐标
		outXstif = [float(row[1]) for row in list(out.values())]  # 提取节点X坐标
		outYstif = [float(row[3]) for row in list(out.values())]  # 提取节点Y坐标
		outXstif_1 = list(reversed(outXstif))	# 逆序，和wdisp保持一致
		outYstif_1 = list(reversed(outYstif))	# 逆序，和wdisp保持一致
		dispdt = pd.read_csv(self.wdisp, sep="\t", encoding="GB18030", header=None)		# 位移比读取
		wdispxp = {}
		wdispxm = {}
		wdispyp = {}
		wdispym = {}
		wdispxwind = {}
		wdispywind = {}
		Xplus1 = 0
		Xplus2 = 0
		Xplus3 = 0
		Yplus1 = 0
		Yplus2 = 0
		Yplus3 = len(dispdt[0]) - 1
		Xwind1 = 0
		Ywind1 = 0
		start_marker = '*NODE'
		end_marker = '*CONSTRAINT'
		dataxp = {}
		dataxm = {}
		datayp = {}
		dataym = {}
		dataxwind = {}
		dataywind = {}
		nodea = {}
		nodefinal = {}
		for i in range(len(dispdt[0])):
			if "X+ 偶然偏心规定" in dispdt[0][i]:
				Xplus1 = i
			elif "X- 偶然偏心规定" in dispdt[0][i]:
				Xplus2 = i
			elif "Y 方向规定" in dispdt[0][i]:
				Xplus3 = i
			elif "Y+ 偶然偏心规定" in dispdt[0][i]:
				Yplus1 = i
			elif "Y- 偶然偏心规定" in dispdt[0][i]:
				Yplus2 = i
			elif "+X 方向风荷载作用下的楼" in dispdt[0][i]:
				Xwind1 = i
			elif "+Y 方向风荷载作用下的楼" in dispdt[0][i]:
				Ywind1 = i
		Xwind2 = Xwind1 + 2*len(tower)+3
		Ywind2 = Ywind1 + 2*len(tower)+3
		for j in range(Xplus1 + 3, Xplus2 - 2):  # x+偶然偏心
			wdispxp[j] = dispdt[0][j]
		outxp = list(wdispxp.values())
		for m in range(len(outxp)):
			dataxp[m] = outxp[m].split()  # 按 号分段,不包括首尾的空格
		out1 = [item for index, item in dataxp.items() if index % 2 != 0]
		for k in range(Xplus2 + 3, Xplus3 - 2):  # x-偶然偏心
			wdispxm[k] = dispdt[0][k]
		outxm = list(wdispxm.values())
		for l in range(len(outxm)):
			dataxm[l] = outxm[l].split()  # 按 号分段,不包括首尾的空格
		out2 = [item for index, item in dataxm.items() if index % 2 != 0]
		for n in range(Yplus1 + 3, Yplus2 - 2):  # y+偶然偏心
			wdispyp[n] = dispdt[0][n]
		outyp = list(wdispyp.values())
		for o in range(len(outyp)):
			datayp[o] = outyp[o].split()  # 按 号分段,不包括首尾的空格
		out3 = [item for index, item in datayp.items() if index % 2 != 0]
		for p in range(Yplus2 + 3, Yplus3 - 1):  # Y-偶然偏心
			wdispym[p] = dispdt[0][p]
		outym = list(wdispym.values())
		for r in range(len(outym)):
			dataym[r] = outym[r].split()  # 按 号分段,不包括首尾的空格
		out4 = [item for index, item in dataym.items() if index % 2 != 0]
		for pp in range(Xwind1 + 3, Xwind2 - 1):  # X风
			wdispxwind[pp] = dispdt[0][pp]
		outxwind = list(wdispxwind.values())
		for rr in range(len(outxwind)):
			dataxwind[rr] = outxwind[rr].split()  # 按 号分段,不包括首尾的空格
		out5 = [item for index, item in dataxwind.items() if index % 2 != 0]
		for bb in range(Ywind1 + 3, Ywind2 - 1):  # Y风
			wdispywind[bb] = dispdt[0][bb]
		outywind = list(wdispywind.values())
		for cc in range(len(outywind)):
			dataywind[cc] = outywind[cc].split()  # 按 号分段,不包括首尾的空格
		out6 = [item for index, item in dataywind.items() if index % 2 != 0]
		nodexp = [row[0] for row in out1]
		nodexm = [row[0] for row in out2]
		nodeyp = [row[0] for row in out3]
		nodeym = [row[0] for row in out4]
		nodexwind = [row[0] for row in out5]
		nodeywind = [row[0] for row in out6]
		nodexpcor = nodexp
		nodexmcor = nodexm
		nodeypcor = nodeyp
		nodeymcor = nodeym
		nodexwindcor = nodexwind
		nodeywindcor = nodeywind
		for x in range(len(out1)):  # 输出到excel
			self.wxp.write(x + 1, 0, out1[x][0])
			self.wxp.write(x + 1, 1, np.array(out1[x][1]))
			self.wxp.write(x + 1, 2, np.array(out1[x][2]))
			self.wxp.write(x + 1, 3, np.array(out1[x][3]))
		for y in range(len(out2)):  # 输出到excel
			self.wxm.write(y + 1, 0, out2[y][0])
			self.wxm.write(y + 1, 1, np.array(out2[y][1]))
			self.wxm.write(y + 1, 2, np.array(out2[y][2]))
			self.wxm.write(y + 1, 3, np.array(out2[y][3]))
		for z in range(len(out3)):  # 输出到excel
			self.wyp.write(z + 1, 0, out3[z][0])
			self.wyp.write(z + 1, 1, np.array(out3[z][1]))
			self.wyp.write(z + 1, 2, np.array(out3[z][2]))
			self.wyp.write(z + 1, 3, np.array(out3[z][3]))
		for w in range(len(out4)):  # 输出到excel
			self.wym.write(w + 1, 0, out4[w][0])
			self.wym.write(w + 1, 1, np.array(out4[w][1]))
			self.wym.write(w + 1, 2, np.array(out4[w][2]))
			self.wym.write(w + 1, 3, np.array(out4[w][3]))
		for zz in range(len(out5)):  # 输出到excel
			self.wxwind.write(zz + 1, 0, out5[zz][0])
			self.wxwind.write(zz + 1, 1, np.array(out5[zz][1]))
			self.wxwind.write(zz + 1, 2, np.array(out5[zz][2]))
			self.wxwind.write(zz + 1, 3, np.array(out5[zz][3]))
		for ww in range(len(out6)):  # 输出到excel
			self.wywind.write(ww + 1, 0, out6[ww][0])
			self.wywind.write(ww + 1, 1, np.array(out6[ww][1]))
			self.wywind.write(ww + 1, 2, np.array(out6[ww][2]))
			self.wywind.write(ww + 1, 3, np.array(out6[ww][3]))
		start_row = self.feadf[self.feadf[0] == start_marker].index[0]
		end_row = self.feadf[self.feadf[0] == end_marker].index[0]
		for nn in range(start_row + 1, end_row):  # 读取全部节点
			nodea[nn] = self.feadf[0][nn]
		nodeb = list(nodea.values())
		for oo in range(len(nodeb)):
			nodefinal[oo] = nodeb[oo].split()  # 按 号分段,不包括首尾的空格
		nodefinals = list(nodefinal.values())
		for ii in range(len(nodexp)):
			for ln in nodefinals:
				if nodexp[ii] in ln:	 # 读取以节点号开头的行
					nodexpcor[ii] = re.split(r'[=, \s]', ln[1])	# 分段并提取坐标，
		for jj in range(len(nodexm)):
			for ln in nodefinals:
				if nodexm[jj] in ln:	 # 读取以节点号开头的行
					nodexmcor[jj] = re.split(r'[=, \s]', ln[1])
		for kk in range(len(nodeyp)):
			for ln in nodefinals:
				if nodeyp[kk] in ln:	 # 读取以节点号开头的行
					nodeypcor[kk] = re.split(r'[=, \s]', ln[1])	# 分段并提取坐标，
		for mm in range(len(nodeym)):
			for ln in nodefinals:
				if nodeym[mm] in ln:	 # 读取以节点号开头的行
					nodeymcor[mm] = re.split(r'[=, \s]', ln[1])
		for ff in range(len(nodexwind)):
			for ln in nodefinals:
				if nodexwind[ff] in ln:	 # 读取以节点号开头的行
					nodexwindcor[ff] = re.split(r'[=, \s]', ln[1])
		for hh in range(len(nodeywind)):
			for ln in nodefinals:
				if nodeywind[hh] in ln:	 # 读取以节点号开头的行
					nodeywindcor[hh] = re.split(r'[=, \s]', ln[1])
		nodexpycor = [float(row[2]) for row in nodexpcor]	#x取y向坐标，y取x向坐标
		nodexmycor = [float(row[2]) for row in nodexmcor]  # x取y向坐标，y取x向坐标
		nodeypxcor = [float(row[1]) for row in nodeypcor]  # x取y向坐标，y取x向坐标
		nodeymxcor = [float(row[1]) for row in nodeymcor]  # x取y向坐标，y取x向坐标
		nodexwindycor = [float(row[2]) for row in nodexwindcor]
		nodeywindxcor = [float(row[1]) for row in nodeywindcor]  # x取y向坐标，y取x向坐标
		avexpxcor = [float(row[2]) for row in out1]  # 层间平均位移
		avexmxcor = [float(row[2]) for row in out2]
		aveypxcor = [float(row[2]) for row in out3]
		aveymxcor = [float(row[2]) for row in out4]
		avexwindxcor = [float(row[2]) for row in out5]
		aveywindxcor = [float(row[2]) for row in out6]
		deltaxp = [abs(x1 - x2) for x1, x2 in zip(nodexpycor, outYstif_1)]	# 转动中心到边端抗侧力构件的距离
		deltaxm = [abs(x1 - x2) for x1, x2 in zip(nodexmycor, outYstif_1)]
		deltayp = [abs(x1 - x2) for x1, x2 in zip(nodeypxcor, outXstif_1)]
		deltaym = [abs(x1 - x2) for x1, x2 in zip(nodeymxcor, outXstif_1)]
		deltaxwind = [abs(x1 - x2) for x1, x2 in zip(nodexwindycor, outYstif_1)]
		deltaywind = [abs(x1 - x2) for x1, x2 in zip(nodeywindxcor, outXstif_1)]
		ratioxpycor = [float(row[3]) - 1 for row in out1]	# 层间位移比-1
		ratioxmycor = [float(row[3]) - 1 for row in out2]
		ratioypycor = [float(row[3]) - 1 for row in out3]
		ratioymycor = [float(row[3]) - 1 for row in out4]
		ratioxwindycor = [float(row[3]) - 1 for row in out5]
		ratioywindycor = [float(row[3]) - 1 for row in out6]
		torangxp = [y1 * y3 / y2 /1000 for y1, y2, y3 in zip(ratioxpycor, deltaxp, avexpxcor)]	# 计算扭转角
		torangxm = [y1 * y3 / y2 /1000 for y1, y2, y3 in zip(ratioxmycor, deltaxm, avexmxcor)]
		torangyp = [y1 * y3 / y2 /1000 for y1, y2, y3 in zip(ratioypycor, deltayp, aveypxcor)]
		torangym = [y1 * y3 / y2 /1000 for y1, y2, y3 in zip(ratioymycor, deltaym, aveymxcor)]
		torangxwind = [y1 * y3 / y2 / 1000 for y1, y2, y3 in zip(ratioxwindycor, deltaxwind, avexwindxcor)]
		torangywind = [y1 * y3 / y2 / 1000 for y1, y2, y3 in zip(ratioywindycor, deltaywind, aveywindxcor)]
		for xx in range(len(out1)):  # 输出到excel
			self.wxp.write(xx + 1, 4, torangxp[xx])
			self.wxp.write(xx + 1, 5, np.array(tower[xx][0]))
			self.wxp.write(xx + 1, 6, np.array(tower[xx][1]))
		for yy in range(len(out2)):  # 输出到excel
			self.wxm.write(yy + 1, 4, torangxm[yy])
			self.wxm.write(yy + 1, 5, np.array(tower[yy][0]))
			self.wxm.write(yy + 1, 6, np.array(tower[yy][1]))
		for zz in range(len(out3)):  # 输出到excel
			self.wyp.write(zz + 1, 4, torangyp[zz])
			self.wyp.write(zz + 1, 5, np.array(tower[zz][0]))
			self.wyp.write(zz + 1, 6, np.array(tower[zz][1]))
		for ww in range(len(out4)):  # 输出到excel
			self.wym.write(ww + 1, 4, torangym[ww])
			self.wym.write(ww + 1, 5, np.array(tower[ww][0]))
			self.wym.write(ww + 1, 6, np.array(tower[ww][1]))
		for ab in range(len(out5)):  # 输出到excel
			self.wxwind.write(ab + 1, 4, torangxwind[ab])
			self.wxwind.write(ab + 1, 5, np.array(tower[ab][0]))
			self.wxwind.write(ab + 1, 6, np.array(tower[ab][1]))
		for ac in range(len(out5)):  # 输出到excel
			self.wywind.write(ac + 1, 4, torangywind[ac])
			self.wywind.write(ac + 1, 5, np.array(tower[ac][0]))
			self.wywind.write(ac + 1, 6, np.array(tower[ac][1]))
		self.wb.close()
		result = messagebox.showinfo("保存xls", "文件已保存")
		#print(result)

if __name__ == "__main__":
	root = tk.Tk()
	root.title("扭转角计算")
	root.geometry('250x50')
	ExcelEditor(root)
	root.mainloop()
