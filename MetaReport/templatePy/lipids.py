

import os
import re
import datetime
import pandas as pd
from templatePy.template import Template

class Lipids(Template):
	def __init__(self, path):
		os.chdir(path)
		information_file = [i for i in os.listdir('.') if re.search('weight', i, re.I)]
		if information_file:
			self.projectinfo = pd.read_excel(information_file[0], header=None, sheet_name=1)
		else:
			print('Error:该项目下无weight表')
			exit()
		self.groupvs = self.projectinfo.iloc[3, 1]

	def header(self, paragraphs, start_row=6):
		today = str(datetime.date.today())
		for i in range(3):
			if i in range(0, 2):
				pa = paragraphs[i + start_row].add_run(self.projectinfo.iloc[i, 1])
				self.paragraph_format(pa, size=14, family=u'黑体')
			if i in range(2, 3):
				pa = paragraphs[i + start_row].add_run(self.projectinfo.iloc[i, 1])
				self.paragraph_format(pa, size=14, family="Arial") #### family = 'Calibri'
		pa = paragraphs[12].add_run(today)
		self.paragraph_format(pa, size=14, family="Arial")

	def table_data(self, tables):
		sample_info = self.projectinfo.iloc[:,2:]
		sample_info.dropna(axis=0,how='any', inplace=True)
		sample_info.drop(0, inplace=True)
		self.insert_table(sample_info, tables[0], size=12, family_en='Times New Roman')
		class_species = pd.read_excel('Pic_class_species.xlsx')
		self.insert_table(class_species, tables[1], size=12, family_en='Times New Roman')
		pca = pd.read_excel(os.path.join('报告及附件', '附件2 Result', '04. Lipid Concentration Analysis', 'Species', 'Multivariate Statistical Analysis', 'pic_PCA.xlsx'))
		self.insert_table(pca, tables[3], size=12, family_en='Times New Roman')
		if os.path.exists(os.path.join('报告及附件', '附件2 Result', '04. Lipid Concentration Analysis', 'Species', 'Multivariate Statistical Analysis', 'pic_PLSDA.xlsx')):
			plsda = pd.read_excel(os.path.join('报告及附件', '附件2 Result', '04. Lipid Concentration Analysis', 'Species', 'Multivariate Statistical Analysis', 'pic_PLSDA.xlsx'))
			plsda.drop('N', axis=1, inplace=True)
			self.insert_table(plsda, tables[4], size=12, family_en='Times New Roman')
		if os.path.exists(os.path.join('报告及附件', '附件2 Result', '04. Lipid Concentration Analysis', 'Species', 'Multivariate Statistical Analysis', 'pic_OPLSDA.xlsx')):
			oplsda = pd.read_excel(os.path.join('报告及附件', '附件2 Result', '04. Lipid Concentration Analysis', 'Species', 'Multivariate Statistical Analysis', 'pic_OPLSDA.xlsx'))
			oplsda.drop('N', axis=1, inplace=True)
			self.insert_table(oplsda, tables[5], size=12, family_en='Times New Roman')
		pic_diff = pd.read_excel('Pic_diff.xlsx')
		f = lambda x: str(round(float(x), 4))
		tmp = pic_diff.iloc[:, 3:].applymap(f)
		pic_diff = pd.concat([pic_diff.iloc[:, :3], tmp], axis=1)
		self.insert_table(pic_diff, tables[6], size=10.5, family_en='Times New Roman')

	def text_png_data(self, paragraphs):
		pie12 = [i for i in os.listdir(os.path.join('报告及附件', '附件2 Result', '03. Lipid Composition Analysis', self.groupvs)) if re.search('pie.*.png', i, re.I)]
		for i, p in enumerate(paragraphs):
			if 'groupvs' in p.text:
				self.text_replace(p, ['groupvs'], [self.groupvs], size=12, family_ch=u'黑体', family_en='Times New Roman')
			if '[LipidNumber]' in p.text:
				self.insert_png(p, '[LipidNumber]', os.path.join('报告及附件', '附件2 Result', '02. Lipids_Stat', 'LipidNumber.png'))
			if '[pie12]' in p.text:
				if pie12:
					self.insert_png(p, '[pie12]', os.path.join('报告及附件', '附件2 Result', '03. Lipid Composition Analysis', self.groupvs, pie12[0]), os.path.join('报告及附件', '附件2 Result', '03. Lipid Composition Analysis', self.groupvs, pie12[1]))
			if '[dynamicplot]' in p.text:
				self.insert_png(p, '[dynamicplot]', os.path.join('报告及附件', '附件2 Result', '03. Lipid Composition Analysis', self.groupvs, 'dynamicplot.png'))
			if '[total_lipid]' in p.text:
				self.insert_png(p, '[total_lipid]', os.path.join('报告及附件', '附件2 Result', '04. Lipid Concentration Analysis', 'Total Lipids', self.groupvs, 'total_lipid.png'))
			if '[allclass_lipid.GAP1]' in p.text:
				self.insert_png(p, '[allclass_lipid.GAP1]', os.path.join('报告及附件', '附件2 Result', '04. Lipid Concentration Analysis', 'Class', self.groupvs, 'allclass_lipid.GAP1.png'))
			if '[allclass_lipid.GAP2]' in p.text:
				self.insert_png(p, '[allclass_lipid.GAP2]', os.path.join('报告及附件', '附件2 Result', '04. Lipid Concentration Analysis', 'Class', self.groupvs, 'allclass_lipid.GAP2.png'))
			if '[PC_ErrorBar]' in p.text:
				self.insert_png(p, '[PC_ErrorBar]', os.path.join('报告及附件', '附件2 Result', '04. Lipid Concentration Analysis', 'Class', self.groupvs, 'PC_ErrorBar.png'))
			if '[pca]' in p.text:
				self.insert_png(p, '[pca]', os.path.join('报告及附件', '附件2 Result', '04. Lipid Concentration Analysis', 'Species', 'Multivariate Statistical Analysis', self.groupvs, f'{self.groupvs}-PCA.png'))
			if '[plsda]' in p.text:
				self.insert_png(p, '[plsda]', os.path.join('报告及附件', '附件2 Result', '04. Lipid Concentration Analysis', 'Species', 'Multivariate Statistical Analysis', self.groupvs, f'{self.groupvs}-PLS-DA.png'))
			if '[plsda-perm]' in p.text:
				self.insert_png(p, '[plsda-perm]', os.path.join('报告及附件', '附件2 Result', '04. Lipid Concentration Analysis', 'Species', 'Multivariate Statistical Analysis', self.groupvs, f'{self.groupvs}-PLS-DA-Permutation.png'))
			if '[oplsda]' in p.text:
				self.insert_png(p, '[oplsda]', os.path.join('报告及附件', '附件2 Result', '04. Lipid Concentration Analysis', 'Species', 'Multivariate Statistical Analysis', self.groupvs, f'{self.groupvs}-OPLS-DA.png'))
			if '[oplsda-perm]' in p.text:
				self.insert_png(p, '[oplsda-perm]', os.path.join('报告及附件', '附件2 Result', '04. Lipid Concentration Analysis', 'Species', 'Multivariate Statistical Analysis', self.groupvs, f'{self.groupvs}-OPLS-DA-Permutation.png'))
			if '[volcano]' in p.text:
				self.insert_png(p, '[volcano]', os.path.join('报告及附件', '附件2 Result', '04. Lipid Concentration Analysis', 'Species', 'Univariate Statistical Analysis', self.groupvs, 'Volcano_Plot.png'))
			if '[pc_molecular]' in p.text:
				self.insert_png(p, '[pc_molecular]', os.path.join('报告及附件', '附件2 Result', '04. Lipid Concentration Analysis', 'Species', 'Univariate Statistical Analysis', self.groupvs, 'PC_molecular.png'))
			if '[bubble]' in p.text:
				self.insert_png(p, '[bubble]', os.path.join('报告及附件', '附件2 Result', '04. Lipid Concentration Analysis', 'Species', 'Bubble Plot', self.groupvs, 'diff_Bubble.png'))
			if '[heatmap]' in p.text:
				self.insert_png(p, '[heatmap]', os.path.join('报告及附件', '附件2 Result', '04. Lipid Concentration Analysis', 'Species', 'Hierarchical Clustering Analysis', self.groupvs, 'Heatmap.png'))
			if '[cor_heatmap]' in p.text:
				self.insert_png(p, '[cor_heatmap]', os.path.join('报告及附件', '附件2 Result', '04. Lipid Concentration Analysis', 'Species', 'Correlation Analysis', self.groupvs, 'Cor_heatmap.png'))
			if '[circlize]' in p.text:
				self.insert_png(p, '[circlize]', os.path.join('报告及附件', '附件2 Result', '04. Lipid Concentration Analysis', 'Species', 'Correlation Analysis', self.groupvs, 'circlize.png'))
			if '[network]' in p.text:
				self.insert_png(p, '[network]', os.path.join('报告及附件', '附件2 Result', '04. Lipid Concentration Analysis', 'Species', 'Correlation Analysis', self.groupvs, 'network.png'))
			if '[pc_carbon]' in p.text:
				self.insert_png(p, '[pc_carbon]', os.path.join('报告及附件', '附件2 Result', '05. Lipid Chain Length Analysis', self.groupvs, 'PC_carbon_chain.png'))
			if '[pc_class]' in p.text:
				self.insert_png(p, '[pc_class]', os.path.join('报告及附件', '附件2 Result', '06. Lipid Saturation Analysis', self.groupvs, 'PC_class.png'))
			if '[bpc-pos]' in p.text:
				self.insert_png(p, '[bpc-pos]', os.path.join('报告及附件', '附件2 Result', '01. QC', 'BPC-POS.jpg'))
			if '[bpc-neg]' in p.text:
				self.insert_png(p, '[bpc-neg]', os.path.join('报告及附件', '附件2 Result', '01. QC', 'BPC-NEG.jpg'))
			if '[multiScatter]' in p.text:
				self.insert_png(p, '[multiScatter]', os.path.join('报告及附件', '附件2 Result', '01. QC', 'MultiScatter.png'))
			if '[qc]' in p.text:
				self.insert_png(p, '[qc]', os.path.join('报告及附件', '附件2 Result', '01. QC', 'QC.png'))
			if '[t2_plot]' in p.text:
				self.insert_png(p, '[t2_plot]', os.path.join('报告及附件', '附件2 Result', '01. QC', "Hotelling-s T2Range Line Plot.png"))
			if '[mcc]' in p.text:
				self.insert_png(p, '[mcc]', os.path.join('报告及附件', '附件2 Result', '01. QC', 'MCC.png'))
			if '[qc_rsd]' in p.text:
				self.insert_png(p, '[qc_rsd]', os.path.join('报告及附件', '附件2 Result', '01. QC', 'QCRSD_curve.png'))

	def save(self, document):
		document.save('高分辨广谱脂质组绝对定量报告.docx')
		print(f'报告生成完成，报告路径为：{os.path.join(os.getcwd(), "高分辨广谱脂质组绝对定量报告.docx")}')





