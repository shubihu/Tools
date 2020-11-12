# -*- coding: future_fstrings -*-     # should work even without -*-
############
import os
import re
import datetime
import pandas as pd
from templatePy.template import Template
from templatePy.template import extract_top

class Notarget(Template):
	def __init__(self, path):
		os.chdir(path)
		information_file = [i for i in os.listdir('.') if re.search('report.*info', i, re.I)]
		if information_file:
			self.projectinfo = pd.read_csv(information_file[0], sep='\t')
			# if 'xls' in information_file[0]:
			# 	self.projectinfo = pd.read_excel(information_file[0], index_col=0, header=None)
			# elif information_file[0].endswith('csv'):
			# 	self.projectinfo = pd.read_csv(information_file[0], index_col=0, header=None)
			# elif information_file[0].endswith('txt'):
			# 	self.projectinfo = pd.read_csv(information_file[0], index_col=0, sep='\t', header=None)
		else:
			print('Error:该项目下无report_info表')
			exit()

		self.project_name = self.projectinfo['项目名称'][0]
		self.school = self.projectinfo['客户名称'][0]
		self.project_num = self.projectinfo['样品编号'][0]
		with open('groupvs.txt') as g:
			self.groupvs = g.readline().strip().replace('|', '_')

		self.sample_info = pd.read_excel('information.xlsx')

		fj1 = pd.read_excel(os.path.join('报告及附件', '附件1_代谢物定性定量结果表.xlsx'), sheet_name=None)
		pos = fj1['pos'][fj1['pos']['Name'].notnull()]
		neg = fj1['neg'][fj1['neg']['Name'].notnull()]
		self.pos_meta = set(i for i in pos['Name'])
		self.neg_meta = set(i for i in neg['Name'])
		self.total_meta = self.pos_meta | self.neg_meta

		kegg = pd.read_excel(os.path.join('报告及附件', '附件2 Result', '04. Bioinformatics Analysis', 'KEGG Analysis', self.groupvs, 'Kegg.xlsx'), sheet_name=None)
		self.keggEnrich_top3, _ = extract_top(kegg, 'keggEnrich', num=3)
		self.keggId = extract_top(kegg, 'keggID', num=3)[0][0]
		self.kegg_heatmapID = kegg['map2query']['Map_ID'][0]
		# print(self.kegg_heatmapID)

	def header(self, paragraphs):
		today = str(datetime.date.today())
		pa = paragraphs[9].add_run(self.project_name)
		self.paragraph_format(pa, size=14, family=u'微软雅黑')
		pa = paragraphs[10].add_run(self.school)
		self.paragraph_format(pa, size=14, family=u'微软雅黑')
		pa = paragraphs[11].add_run(self.project_num)
		self.paragraph_format(pa, size=14, family="Arial") #### family = 'Calibri'
		pa = paragraphs[12].add_run(today)
		self.paragraph_format(pa, size=14, family="Arial")

	def table_data(self, tables):
		# newname = pd.read_excel('newname.xlsx')
		# sample_total = len([i for i in newname['sample'] if not re.search('qc', i, re.I)])
		sample_total = self.sample_info['数量'].sum()
		groupvs_num = pd.read_csv('groupvs.txt', header=None).shape[0]

		pos_diff = pd.read_excel(os.path.join('报告及附件', '附件2 Result', '03. Difference Analysis', 'Differential Metabolites', '附件1_样本POS_定性.xlsx'))
		pos_diff = pos_diff[pos_diff['Name'].notnull()]
		if 'VIP' in pos_diff.columns:
			pos_diff1 = pos_diff[(pos_diff['VIP'] > 1) & (pos_diff['p-value'] < 0.05)]
			pos_diff2 = pos_diff[(pos_diff['VIP'] > 1) & (pos_diff['p-value'] >= 0.05) & (pos_diff['p-value'] < 0.1)]
		else:
			pos_diff1 = pos_diff[(pos_diff['ANOVA p value'] < 0.05)]
			pos_diff2 = pos_diff[(pos_diff['ANOVA p value'] >= 0.05) & (pos_diff['ANOVA p value'] < 0.1)]
		pos_diff_meta = set(i for i in pos_diff1['Name'])
		neg_diff = pd.read_excel(os.path.join('报告及附件', '附件2 Result', '03. Difference Analysis', 'Differential Metabolites', '附件1_样本NEG_定性.xlsx'))
		neg_diff = neg_diff[neg_diff['Name'].notnull()]
		if 'VIP' in neg_diff.columns:
			neg_diff1 = neg_diff[(neg_diff['VIP'] > 1) & (neg_diff['p-value'] < 0.05)]
			neg_diff2 = neg_diff[(neg_diff['VIP'] > 1) & (neg_diff['p-value'] >= 0.05) & (neg_diff['p-value'] < 0.1)]
		else:
			neg_diff1 = neg_diff[(neg_diff['ANOVA p value'] < 0.05)]
			neg_diff2 = neg_diff[(neg_diff['ANOVA p value'] >= 0.05) & (neg_diff['ANOVA p value'] < 0.1)]
		neg_diff_meta = set(i for i in neg_diff1['Name'])
		diff_total = pos_diff_meta | neg_diff_meta

		table1 = tables[0]
		self.paragraph_format(table1.cell(1,1).paragraphs[0].add_run(str(sample_total)), size = 10.5, family = 'Arial')
		self.paragraph_format(table1.cell(2,1).paragraphs[0].add_run(str(groupvs_num)), size = 10.5, family = 'Arial')
		self.paragraph_format(table1.cell(4,1).paragraphs[0].add_run(str(len(self.total_meta))), size = 10.5, family = 'Arial')
		self.paragraph_format(table1.cell(7,0).paragraphs[0].add_run(str(self.groupvs)), size = 10.5, family = 'Arial')
		self.paragraph_format(table1.cell(7,1).paragraphs[0].add_run(str(len(diff_total))), size = 10.5, family = 'Arial')
		self.paragraph_format(table1.cell(7,2).paragraphs[0].add_run(';'.join(self.keggEnrich_top3)), size = 9, family = 'Arial')

		self.insert_table(self.sample_info, tables[1], size=12, family_en='Times New Roman')

		table3 = tables[2]
		self.paragraph_format(table3.cell(1,1).paragraphs[0].add_run(str(len(self.pos_meta))), size = 10.5, family = 'Arial')
		self.paragraph_format(table3.cell(2,1).paragraphs[0].add_run(str(len(self.neg_meta))), size = 10.5, family = 'Arial')

		pca = pd.read_excel(os.path.join('统计分析', 'pic_PCA.xlsx'))
		pca['Group'] = pca['Group'].apply(lambda x: str(x).replace('|', '_'))
		pos_pca = pca[(pca['Group'] == self.groupvs) & (pca['Polarity'] == 'pos')]
		pos_pca.drop('Polarity', axis=1, inplace=True)
		neg_pca = pca[(pca['Group'] == self.groupvs) & (pca['Polarity'] == 'neg')]
		neg_pca.drop('Polarity', axis=1, inplace=True)
		self.insert_table(pos_pca, tables[3], size=12, family_en='Times New Roman')
		self.insert_table(neg_pca, tables[4], size=12, family_en='Times New Roman')
		
		if os.path.exists(os.path.join('统计分析', 'pic_PLSDA.xlsx')):
			plsda = pd.read_excel(os.path.join('统计分析', 'pic_PLSDA.xlsx'))
			pos_plsda = plsda[(plsda['Group'] == self.groupvs) & (plsda['Polarity'] == 'pos')]
			pos_plsda.drop(['Polarity', 'N'], axis=1, inplace=True)
			neg_plsda = plsda[(plsda['Group'] == self.groupvs) & (plsda['Polarity'] == 'neg')]
			neg_plsda.drop(['Polarity', 'N'], axis=1, inplace=True)
			self.insert_table(pos_plsda, tables[5], size=12, family_en='Times New Roman')
			self.insert_table(neg_plsda, tables[6], size=12, family_en='Times New Roman')
		
		if os.path.exists(os.path.join('统计分析', 'pic_OPLSDA.xlsx')):
			oplsda = pd.read_excel(os.path.join('统计分析', 'pic_OPLSDA.xlsx'))
			pos_oplsda = oplsda[(oplsda['Group'] == self.groupvs) & (oplsda['Polarity'] == 'pos')]
			pos_oplsda.drop(['Polarity', 'N'], axis=1, inplace=True)
			neg_oplsda = oplsda[(oplsda['Group'] == self.groupvs) & (oplsda['Polarity'] == 'neg')]
			neg_oplsda.drop(['Polarity', 'N'], axis=1, inplace=True)
			self.insert_table(pos_oplsda, tables[7], size=12, family_en='Times New Roman')
			self.insert_table(neg_oplsda, tables[8], size=12, family_en='Times New Roman')

		def diff_func(df1, df2, table):
			if 'VIP' in df1.columns:
				colname = ['adduct', 'Name', 'VIP', 'Fold change', 'p-value', 'm/z', 'rt(s)']
				df1 = df1[colname]
				df1.iloc[:, [2, 3, 5]] = df1.iloc[:, [2, 3, 5]].apply(lambda x: round(x, 2))
				df1.iloc[:, [4, 6]] = df1.iloc[:, [4, 6]].apply(lambda x: round(x, 3))
				df1.loc[''] = ''
				df2 = df2[colname]
				df2.iloc[:, [2, 3, 5]] = df2.iloc[:, [2, 3, 5]].apply(lambda x: round(x, 2))
				df2.iloc[:, [4, 6]] = df2.iloc[:, [4, 6]].apply(lambda x: round(x, 3))
			else:
				colname = ['adduct', 'Name', 'ANOVA p value', 'm/z', 'rt(s)']
				df1 = df1[colname]
				df1.insert(2, 'VIP', '' * df1.shape[1])
				df1.insert(3, 'Fold change', '' * df1.shape[1])
				df1.iloc[:, 5] = df1.iloc[:, 5].apply(lambda x: round(x, 2))
				df1.iloc[:, [4, 6]] = df1.iloc[:, [4, 6]].apply(lambda x: round(x, 3))
				df1.loc[''] = ''
				df2 = df2[colname]
				df2.insert(2, 'VIP', '' * df2.shape[1])
				df2.insert(3, 'Fold change', '' * df2.shape[1])
				df2.iloc[:, 5] = df2.iloc[:, 5].apply(lambda x: round(x, 2))
				df2.iloc[:, [4, 6]] = df2.iloc[:, [4, 6]].apply(lambda x: round(x, 3))
			self.insert_table(df1, table, size=9, family_en='Times New Roman', rgbColor='#FFFF00')
			self.insert_table(df2, table, size=9, family_en='Times New Roman', rgbColor='#0099CC')
		diff_func(pos_diff1, pos_diff2, tables[9])
		diff_func(neg_diff1, neg_diff2, tables[10])

	def text_png_data(self, paragraphs):
		pos_tic = [i for i in os.listdir('.') if re.search('(tic)?.*pos.*(tic)?.png', i, re.I)]
		neg_tic = [i for i in os.listdir('.') if re.search('(tic)?.*neg.*(tic)?.png', i, re.I)]

		for i, p in enumerate(paragraphs):
			if 'total_meta' in p.text:
				self.text_replace(p, ['total_meta'], [str(len(self.total_meta))], size=12, family_ch=u'微软雅黑', family_en='Times New Roman')
			if '[pos-tic]' in p.text:
				if pos_tic:
					self.insert_png(p, '[pos-tic]', pos_tic[0])
			if '[neg-tic]' in p.text:
				if neg_tic:
					self.insert_png(p, '[neg-tic]', neg_tic[0])
			if '[qc_pca_pos]' in p.text:
				self.insert_png(p, '[qc_pca_pos]', os.path.join('报告及附件', '附件2 Result', '01. QC', 'QC-POS.png'))
			if '[qc_pca_neg]' in p.text:
				self.insert_png(p, '[qc_pca_neg]', os.path.join('报告及附件', '附件2 Result', '01. QC', 'QC-NEG.png'))
			if '[pos_scatter]' in p.text:
				self.insert_png(p, '[pos_scatter]', os.path.join('报告及附件', '附件2 Result', '01. QC', 'MultiScatter-POS.png'))
			if '[neg_scatter]' in p.text:
				self.insert_png(p, '[neg_scatter]', os.path.join('报告及附件', '附件2 Result', '01. QC', 'MultiScatter-NEG.png'))
			if '[pos_ht2]' in p.text:
				self.insert_png(p, '[pos_ht2]', os.path.join('报告及附件', '附件2 Result', '01. QC', 'Hotelling-s T2Range Line Plot-POS.png'))
			if '[neg_ht2]' in p.text:
				self.insert_png(p, '[neg_ht2]', os.path.join('报告及附件', '附件2 Result', '01. QC', 'Hotelling-s T2Range Line Plot-NEG.png'))
			if '[pos_mcc]' in p.text:
				self.insert_png(p, '[pos_mcc]', os.path.join('报告及附件', '附件2 Result', '01. QC', 'MCC-POS.png'))
			if '[neg_mcc]' in p.text:
				self.insert_png(p, '[neg_mcc]', os.path.join('报告及附件', '附件2 Result', '01. QC', 'MCC-NEG.png'))
			if '[pos_rsd]' in p.text:
				self.insert_png(p, '[pos_rsd]', os.path.join('报告及附件', '附件2 Result', '01. QC', 'QCRSD_curve-POS.png'))
			if '[neg_rsd]' in p.text:
				self.insert_png(p, '[neg_rsd]', os.path.join('报告及附件', '附件2 Result', '01. QC', 'QCRSD_curve-NEG.png'))
			if '[superclass]' in p.text:
				self.insert_png(p, '[superclass]', os.path.join('报告及附件', '附件2 Result', '02. Identified Metabolites_Stat', 'superclass_pie.png'))
			if '[pos_volcano]' in p.text:
				self.insert_png(p, '[pos_volcano]', os.path.join('报告及附件', '附件2 Result', '03. Difference Analysis', 'Univariate Statistical Analysis', self.groupvs, 'pos_Volcano_Plot.png'))
			if '[neg_volcano]' in p.text:
				self.insert_png(p, '[neg_volcano]', os.path.join('报告及附件', '附件2 Result', '03. Difference Analysis', 'Univariate Statistical Analysis', self.groupvs, 'neg_Volcano_Plot.png'))
			if '[pos_volcano_superclass]' in p.text:
				self.insert_png(p, '[pos_volcano_superclass]', os.path.join('报告及附件', '附件2 Result', '03. Difference Analysis', 'Univariate Statistical Analysis', self.groupvs, 'pos_Volcano_Plot_Superclass.png'))
			if '[neg_volcano_superclass]' in p.text:
				self.insert_png(p, '[neg_volcano_superclass]', os.path.join('报告及附件', '附件2 Result', '03. Difference Analysis', 'Univariate Statistical Analysis', self.groupvs, 'neg_Volcano_Plot_Superclass.png'))
			if '[pca_pos]' in p.text:
				self.insert_png(p, '[pca_pos]', os.path.join('报告及附件', '附件2 Result', '03. Difference Analysis', 'Multivariate Statistical Analysis', 'SIMCAP-POS', self.groupvs, f'{self.groupvs}-PCA.png'))
			if '[pca_neg]' in p.text:
				self.insert_png(p, '[pca_neg]', os.path.join('报告及附件', '附件2 Result', '03. Difference Analysis', 'Multivariate Statistical Analysis', 'SIMCAP-NEG', self.groupvs, f'{self.groupvs}-PCA.png'))
			if '[plsda_pos]' in p.text:
				self.insert_png(p, '[plsda_pos]', os.path.join('报告及附件', '附件2 Result', '03. Difference Analysis', 'Multivariate Statistical Analysis', 'SIMCAP-POS', self.groupvs, f'{self.groupvs}-PLS-DA.png'))
			if '[plsda_neg]' in p.text:
				self.insert_png(p, '[plsda_neg]', os.path.join('报告及附件', '附件2 Result', '03. Difference Analysis', 'Multivariate Statistical Analysis', 'SIMCAP-NEG', self.groupvs, f'{self.groupvs}-PLS-DA.png'))
			if '[plsda_permutation_pos]' in p.text:
				self.insert_png(p, '[plsda_permutation_pos]', os.path.join('报告及附件', '附件2 Result', '03. Difference Analysis', 'Multivariate Statistical Analysis', 'SIMCAP-POS', self.groupvs, f'{self.groupvs}-PLS-DA-Permutation.png'))
			if '[plsda_permutation_neg]' in p.text:
				self.insert_png(p, '[plsda_permutation_neg]', os.path.join('报告及附件', '附件2 Result', '03. Difference Analysis', 'Multivariate Statistical Analysis', 'SIMCAP-NEG', self.groupvs, f'{self.groupvs}-PLS-DA-Permutation.png'))
			if '[oplsda_pos]' in p.text:
				self.insert_png(p, '[oplsda_pos]', os.path.join('报告及附件', '附件2 Result', '03. Difference Analysis', 'Multivariate Statistical Analysis', 'SIMCAP-POS', self.groupvs, f'{self.groupvs}-OPLS-DA.png'))
			if '[oplsda_neg]' in p.text:
				self.insert_png(p, '[oplsda_neg]', os.path.join('报告及附件', '附件2 Result', '03. Difference Analysis', 'Multivariate Statistical Analysis', 'SIMCAP-NEG', self.groupvs, f'{self.groupvs}-OPLS-DA.png'))
			if '[oplsda_permutation_pos]' in p.text:
				self.insert_png(p, '[oplsda_permutation_pos]', os.path.join('报告及附件', '附件2 Result', '03. Difference Analysis', 'Multivariate Statistical Analysis', 'SIMCAP-POS', self.groupvs, f'{self.groupvs}-OPLS-DA-Permutation.png'))
			if '[oplsda_permutation_neg]' in p.text:
				self.insert_png(p, '[oplsda_permutation_neg]', os.path.join('报告及附件', '附件2 Result', '03. Difference Analysis', 'Multivariate Statistical Analysis', 'SIMCAP-NEG', self.groupvs, f'{self.groupvs}-OPLS-DA-Permutation.png'))
			if '[pos_fc_plot]' in p.text:
				self.insert_png(p, '[pos_fc_plot]', os.path.join('报告及附件', '附件2 Result', '03. Difference Analysis', 'Differential Metabolites', self.groupvs, 'pos_FC_plot.png'))
			if '[neg_fc_plot]' in p.text:
				self.insert_png(p, '[neg_fc_plot]', os.path.join('报告及附件', '附件2 Result', '03. Difference Analysis', 'Differential Metabolites', self.groupvs, 'neg_FC_plot.png'))
			if '[pos_cluster]' in p.text:
				self.insert_png(p, '[pos_cluster]', os.path.join('报告及附件', '附件2 Result', '04. Bioinformatics Analysis', 'Hierarchical Clustering Analysis', self.groupvs, 'pos', 'cluster_sample.png'))
			if '[pos_hac]' in p.text:
				self.insert_png(p, '[pos_hac]', os.path.join('报告及附件', '附件2 Result', '04. Bioinformatics Analysis', 'Hierarchical Clustering Analysis', self.groupvs, 'pos', 'cluster_HAC.png'))
			if '[neg_cluster]' in p.text:
				self.insert_png(p, '[neg_cluster]', os.path.join('报告及附件', '附件2 Result', '04. Bioinformatics Analysis', 'Hierarchical Clustering Analysis', self.groupvs, 'neg', 'cluster_sample.png'))
			if '[neg_hac]' in p.text:
				self.insert_png(p, '[neg_hac]', os.path.join('报告及附件', '附件2 Result', '04. Bioinformatics Analysis', 'Hierarchical Clustering Analysis', self.groupvs, 'neg', 'cluster_HAC.png'))
			if '[pos_corr]' in p.text:
				self.insert_png(p, '[pos_corr]', os.path.join('报告及附件', '附件2 Result', '04. Bioinformatics Analysis', 'Correlation Analysis', self.groupvs, 'pos', 'CorrPlot.png'))
			if '[neg_corr]' in p.text:
				self.insert_png(p, '[neg_corr]', os.path.join('报告及附件', '附件2 Result', '04. Bioinformatics Analysis', 'Correlation Analysis', self.groupvs, 'neg', 'COrrPlot.png'))
			if '[pos_circlize]' in p.text:
				self.insert_png(p, '[pos_circlize]', os.path.join('报告及附件', '附件2 Result', '04. Bioinformatics Analysis', 'Correlation Analysis', self.groupvs, 'pos', 'circlize.png'))
			if '[neg_circlize]' in p.text:
				self.insert_png(p, '[neg_circlize]', os.path.join('报告及附件', '附件2 Result', '04. Bioinformatics Analysis', 'Correlation Analysis', self.groupvs, 'neg', 'circlize.png'))
			if '[pos_network]' in p.text:
				self.insert_png(p, '[pos_network]', os.path.join('报告及附件', '附件2 Result', '04. Bioinformatics Analysis', 'Correlation Analysis', self.groupvs, 'pos', 'network.png'))
			if '[neg_network]' in p.text:
				self.insert_png(p, '[neg_network]', os.path.join('报告及附件', '附件2 Result', '04. Bioinformatics Analysis', 'Correlation Analysis', self.groupvs, 'neg', 'network.png'))
			if '[pathway]' in p.text:
				self.insert_png(p, '[pathway]', os.path.join('报告及附件', '附件2 Result', '04. Bioinformatics Analysis', 'KEGG Analysis', self.groupvs, 'KEGG Map', f'{self.keggId}.png'))
			if '[kegg_heatmap]' in p.text:
				self.insert_png(p, '[kegg_heatmap]', os.path.join('报告及附件', '附件2 Result', '04. Bioinformatics Analysis', 'KEGG Analysis', self.groupvs, 'KEGG Pathway_Metabolites Heatmap', f'{self.kegg_heatmapID}_heatmap.png'))
			if '[kegg_enrich1]' in p.text:
				self.insert_png(p, '[kegg_enrich1]', os.path.join('报告及附件', '附件2 Result', '04. Bioinformatics Analysis', 'KEGG Analysis', self.groupvs, 'KEGG Enrichment Analysis', 'EnrichmentBubble.png'))
			if '[kegg_enrich2]' in p.text:
				self.insert_png(p, '[kegg_enrich2]', os.path.join('报告及附件', '附件2 Result', '04. Bioinformatics Analysis', 'KEGG Analysis', self.groupvs, 'KEGG Enrichment Analysis', 'EnrichmentBar.png'))
			if '[kegg_dascore1]' in p.text:
				self.insert_png(p, '[kegg_dascore1]', os.path.join('报告及附件', '附件2 Result', '04. Bioinformatics Analysis', 'KEGG Analysis', self.groupvs, 'Differential Abundance Score', 'DAScore_plot.png'))
			if '[kegg_dascore2]' in p.text:
				self.insert_png(p, '[kegg_dascore2]', os.path.join('报告及附件', '附件2 Result', '04. Bioinformatics Analysis', 'KEGG Analysis', self.groupvs, 'Differential Abundance Score', 'DAScore_plot_H2.png'))


			

	def save(self, document):
		document.save('高分辨非靶代谢组学报告.docx')
		print(f'报告生成完成，报告路径为：{os.path.join(os.getcwd(), "高分辨非靶代谢组学报告.docx")}')





