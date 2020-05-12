import os
import re
import datetime
import pandas as pd
from templatePy.template import Template

class Fattyacid(Template):
	def __init__(self, path, types):
		self.types = types
		os.chdir(path)
		information_file = [i for i in os.listdir('.') if re.search('infor?mation', i, re.I)]
		if information_file:
			self.projectinfo = pd.read_excel(information_file[0], header=None)
		else:
			print('Error:该项目下无项目 information 表')
			exit()
		self.groupvs = self.projectinfo.iloc[3, 1]
		self.weight = self.projectinfo.iloc[4, 1]

	def header(self, paragraphs, start_row=3):
		today = str(datetime.date.today())
		for i in range(3):
			if i in range(0, 2):
				pa = paragraphs[i + start_row].add_run(self.projectinfo.iloc[i, 1])
				self.paragraph_format(pa, size=14, family=u'楷体', bold=True)
			if i in range(2, 3):
				pa = paragraphs[i + start_row].add_run(self.projectinfo.iloc[i, 1])
				self.paragraph_format(pa, size=14, family="Arial", bold=True) #### family = 'Calibri'
		pa = paragraphs[9].add_run(today)
		self.paragraph_format(pa, size=14, family="Arial", bold=True)

	def table_data(self, tables):
		sample_info = self.projectinfo.iloc[:,4:6]
		sample_info.drop(0, inplace=True)
		self.insert_table(sample_info, tables[0], size=10, family_en='Arial')
		if self.types == 'lc':
			data = pd.read_excel(os.path.join('报告及附件', '附件1_中长链脂肪酸结果列表.xlsx'), sheet_name=None)
		elif self.types == 'sc':
			data = pd.read_excel(os.path.join('报告及附件', '附件1_短链脂肪酸结果列表.xlsx'), sheet_name=None)
		calibrationCurve = data['标曲']
		self.insert_table(calibrationCurve, tables[1], size=10, family_en='Arial')
		content = data['含量整合'].iloc[:,:7]
		self.insert_table(content, tables[2], size=10, family_en='Arial')

	def text_png_data(self, paragraphs):
		acid_file, total_scfa = [], []
		if os.path.exists(os.path.join('报告及附件', self.groupvs, 'Boxplot')):
			acid_file = [i for i in os.listdir(os.path.join('报告及附件', self.groupvs, 'Boxplot')) if not re.search('total|pdf', i, re.I)]
			total_scfa = [i for i in os.listdir(os.path.join('报告及附件', self.groupvs, 'Boxplot')) if re.search('total.*.png', i, re.I)]
		for i, p in enumerate(paragraphs):
			if 'sample_weight' in p.text:
				self.text_replace(p, ['sample_weight'], [self.weight], family_ch=u'楷体')

			if '[qc_rsd]' in p.text:
				self.insert_png(p, '[qc_rsd]', os.path.join('报告及附件', 'QC_RSD.png'))
			if '[acid_boxplot]' in p.text:
				if acid_file:
					self.insert_png(p, '[acid_boxplot]', os.path.join('报告及附件', self.groupvs, 'Boxplot', acid_file[0]))
			if '[total_scfa]' in p.text:
				if total_scfa:
					self.insert_png(p, '[total_scfa]', os.path.join('报告及附件', self.groupvs, 'Boxplot', total_scfa[0]))
			if '[barplot_gap]' in p.text:
				self.insert_png(p, '[barplot_gap]', os.path.join('报告及附件', self.groupvs, 'Barplot', 'all_barplot_GAP.png'))
			if '[SMP_barplot]' in p.text:
				self.insert_png(p, '[SMP_barplot]', os.path.join('报告及附件', self.groupvs, 'Barplot', 'S_M_P_barplot.png'))
			if '[N36_plot]' in p.text:
				self.insert_png(p, '[N36_plot]', os.path.join('报告及附件', self.groupvs, 'Barplot', 'N3_N6_barplot.png'))
			if '[heatmap]' in p.text:
				self.insert_png(p, '[heatmap]', os.path.join('报告及附件', self.groupvs, 'Heatmap', 'Heatmap.png'))

	def save(self, document):
		if self.types == 'lc':
			document.save('中长链脂肪酸报告.docx')
			print(f'报告生成完成，报告路径为：{os.path.join(os.getcwd(), "中长链脂肪酸报告.docx")}')
		elif self.types == 'sc':
			document.save('短链脂肪酸报告.docx')
			print(f'报告生成完成，报告路径为：{os.path.join(os.getcwd(), "短链脂肪酸报告.docx")}')





