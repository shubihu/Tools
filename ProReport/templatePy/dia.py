# -*- coding: future_fstrings -*-     # should work even without -*-

import os
import re
import sys
import math
import datetime
import pandas as pd
pd.options.mode.chained_assignment = None
from templatePy.template import extract_top
from templatePy.template import Template

class DIA(Template):
	def __init__(self, path):
		os.chdir(path)

		information_file = [i for i in os.listdir('.') if re.search('infor?mation', i, re.I)]
		if information_file:
			if 'xls' in information_file[0]:
				projectinfo = pd.read_excel(information_file[0], index_col=0, header=None)
			elif information_file[0].endswith('csv'):
				projectinfo = pd.read_csv(information_file[0], index_col=0, header=None)
			elif information_file[0].endswith('txt'):
				projectinfo = pd.read_csv(information_file[0], index_col=0, sep='\t', header=None)
		else:
			print('Error:该项目下无project information表')
			sys.exit()

		index_list = projectinfo.index.tolist()
		groupvs = [i for i in index_list if re.search('比较组', i)]

		self.school = projectinfo.loc['委托单位'][1]
		self.project_name = projectinfo.loc['项目名称'][1]
		self.project_num = projectinfo.loc['项目编号'][1]
		self.sample_num = projectinfo.loc['样品数量'][1]
		self.groupvs = projectinfo.loc[groupvs[0]][1]

		origi_record = pd.read_excel('原始记录.xls', sheet_name=0).fillna('')
		self.diff_groupvs = origi_record[origi_record.loc[:, 'group'].str.contains('vs|oneway|twoway')]
		tmp = self.diff_groupvs.iloc[:, 1:].applymap(lambda x: x if '-' in str(x) else int(float(x)))  # applymap对每个元素进行处理
		self.diff_groupvs = pd.concat([self.diff_groupvs.iloc[:, 0], tmp], axis=1)
		self.fc = float(origi_record[origi_record.iloc[:, 0] == '差异倍数'].iloc[0, 1])
		self.total_pro_num = origi_record[origi_record.iloc[:, 0] == '附件1'].iloc[0, 1]
		self.half_pro_num = origi_record[origi_record.iloc[:, 0] == '一半以上存在定量值'].iloc[0, 1]

		sample_pro_group = pd.read_excel('sample_protein_group.xls', header=None)
		rowNum = math.ceil(sample_pro_group.shape[0] / 2)
		self.pro_data = sample_pro_group.iloc[:rowNum,]
		data = sample_pro_group.iloc[rowNum:,]
		list1 = data.iloc[:,0].tolist()
		list1.extend((rowNum - len(list1)) * [''])
		list2 = data.iloc[:,1].tolist()
		list2.extend((rowNum - len(list2)) * [''])
		self.pro_data['sample'] = list1
		self.pro_data['pro'] = list2

		go_file = os.path.join(self.groupvs, 'go', 'GO.xlsx')
		go = pd.read_excel(go_file, sheet_name=None)
		self.bp_top = extract_top(go, 'BP')[0]
		self.mf_top = extract_top(go, 'MF')[0]
		self.cc_top = extract_top(go, 'CC')[0]

		kegg_file = os.path.join(self.groupvs, 'kegg', 'kegg.xlsx')
		kegg = pd.read_excel(kegg_file, sheet_name=None)
		self.keggEnrich_top5, self.pathway = extract_top(kegg, 'keggEnrich')
		self.keggID = extract_top(kegg, 'keggID')[0][0]


	def header(self, paragraphs):
		today = str(datetime.date.today())
		pa = paragraphs[10].add_run(self.school)
		self.paragraph_format(pa, size=12, family=u'微软雅黑')
		pa = paragraphs[11].add_run(self.project_name)
		self.paragraph_format(pa, size=12, family=u'微软雅黑')
		pa = paragraphs[12].add_run(self.project_num)
		self.paragraph_format(pa, size=12, family=u'微软雅黑')
		pa = paragraphs[13].add_run(today)
		self.paragraph_format(pa, size=12, family="微软雅黑")

	def table_data(self, tables):
		# self.paragraph_format(tables[2].cell(1,1).paragraphs[0].add_run(str(self.DDA_protein)), size=10, family=u'微软雅黑')
		# self.paragraph_format(tables[2].cell(1,2).paragraphs[0].add_run(str(self.DDA_peptide)), size=10, family=u'微软雅黑')
		self.insert_table(self.pro_data, tables[2], size=10, family_ch=u'微软雅黑', family_en='微软雅黑')
		self.insert_table(self.diff_groupvs, tables[5], size=10, family_ch=u'微软雅黑', family_en='微软雅黑')

	def text_png_data(self, paragraphs):
		if os.path.exists('WGCNA'):
			color = 'blue'
			if os.path.exists(os.path.join('WGCNA', 'Module_Enrichment', color)):
				wgcna_enrich = os.path.join('WGCNA', 'Module_Enrichment', color)
			else:
				wgcna_enrich = os.listdir(os.path.join('WGCNA', 'Module_Enrichment'))[0]
				color = os.path.basename(wgcna_enrich)

		cv_png = [i for i in os.listdir(os.path.join('Appendix_A', '附件3')) if re.search('cv.*.png', i, re.I)]
		if cv_png:
			CV_median = cv_png[0].split(' ')[1].replace('%.png','')

		for i, p in enumerate(paragraphs):
			if 'groupvs' in p.text:
				self.text_replace(p, ['groupvs'], [self.groupvs])
			if 'sample_num' in p.text:
				self.text_replace(p, ['sample_num'], [str(self.sample_num)], family_ch=u'微软雅黑', family_en='微软雅黑')
			if 'CV_median' in p.text:
				self.text_replace(p, ['CV_median'], [CV_median], family_ch=u'微软雅黑', family_en='微软雅黑')
			if 'total_pro_num' in p.text:
				self.text_replace(p, ['total_pro_num', 'half_pro_num'], [str(int(self.total_pro_num)), str(int(self.half_pro_num))], family_ch=u'微软雅黑', family_en='微软雅黑')
			# if 'upRatio' in p.text:
			# 	self.text_replace(p, ['upRatio', 'upRatio', 'downRatio'], [str(self.fc), str(self.fc), str(round(1 / self.fc, 2))], family_ch=u'微软雅黑', family_en='微软雅黑')
			if 'upRatio' in p.text:
				find_num = len(re.findall('upRatio', p.text))
				self.text_replace(p, ['upRatio'] * find_num, [str(self.fc)] * find_num)
			if 'downRatio' in p.text:
				find_num = len(re.findall('downRatio', p.text))
				self.text_replace(p, ['downRatio'] * find_num, [str(round(1 / self.fc, 2))] * find_num)
			if 'groupvs' in p.text:
				self.text_replace(p, ['groupvs'], [self.groupvs])
			if 'BP-TOP5' in p.text:
				if all([len(self.bp_top) == 0, len(self.mf_top) == 0, len(self.cc_top) == 0]):
					self.text_replace(p, ['，发生了显著性变化'], [''])
				if self.bp_top:
					self.text_replace(p, ['BP-TOP5'], [', '.join(self.bp_top)])
				else:
					self.text_replace(p, ['BP-TOP5等重要生物学过程'], ['无P值小于0.05的显著性生物学过程'])
			if 'MF-TOP5' in p.text:
				if self.mf_top:
					self.text_replace(p, ['MF-TOP5'], [', '.join(self.mf_top)])
				else:
					self.text_replace(p, ['MF-TOP5等分子功能'], ['无P值小于0.05的显著性分子功能'])
			if 'CC-TOP5' in p.text:
				if self.cc_top:
					self.text_replace(p, ['CC-TOP5'], [', '.join(self.cc_top)])
				else:
					self.text_replace(p, ['CC-TOP5等定位蛋白质'], ['无P值小于0.05的显著性定位蛋白'])
			if 'KeggEnrich-top5' in p.text:
				if self.keggEnrich_top5:
					self.text_replace(p, ['KeggEnrich-top5', 'KeggEnrich-top1'], ['，'.join(self.keggEnrich_top5), self.pathway])
				else:
					self.text_replace(p, ['KeggEnrich-top5等重要通路发生了显著变化', 'KeggEnrich-top1'], ['该比较组没有P值小于0.05的显著性富集通路', self.pathway])
			if 'expr_sample_type' in p.text:
				self.text_replace(p, ['expr_sample_type'], [self.sample_type], size=12, family_ch=u'微软雅黑', family_en='微软雅黑', bold=True)

			if "[Protein_Group_Profile.png]" in p.text:
				pro_groupfile = [i for i in os.listdir(os.path.join('Appendix_A', '附件3')) if re.search('profiles.*.png', i, re.I)]
				if pro_groupfile:
					self.insert_png(p, "[Protein_Group_Profile.png]", os.path.join('Appendix_A', '附件3', pro_groupfile[0]))
			if "[Heatmap.png]" in p.text:
				self.insert_png(p, "[Heatmap.png]", os.path.join('Appendix_A', '附件3', 'Heatmap.png'))
			if "[AVERAGE_DATA_POINT_PER_PEAK.png]" in p.text:
				point = [i for i in os.listdir(os.path.join('Appendix_A', '附件3')) if re.search('point.*.png', i, re.I)]
				if point:
					self.insert_png(p, "[AVERAGE_DATA_POINT_PER_PEAK.png]", os.path.join('Appendix_A', '附件3', point[0]))
			if "[Peak_Capacity.png]" in p.text:
				self.insert_png(p, "[Peak_Capacity.png]", os.path.join('Appendix_A', '附件3', 'Peak Capacity.png'))
			if "[iRT.png]" in p.text:
				self.insert_png(p, "[iRT.png]", os.path.join('Appendix_A', '附件3', 'iRT.png'))
			if "[Protein_FDR.png]" in p.text:
				self.insert_png(p, "[Protein_FDR.png]", os.path.join('Appendix_A', '附件3', 'Protein FDR.png'))
			if "[CV.png]" in p.text:
				# cv_png = [i for i in os.listdir(os.path.join('Appendix_A', '附件3')) if re.search('cv.*.png', i, re.I)]
				if cv_png:
					self.insert_png(p, "[CV.png]", os.path.join('Appendix_A', '附件3', cv_png[0]))
				else:
					self.text_replace(p, ["[CV.png]"], ['无CV图'], family_ch=u'微软雅黑', family_en='微软雅黑')
			if "[DELE-QC.png]" in p.text:
				replace = self.insert_png(p, "[DELE-QC.png]", os.path.join("PCA_QC", "PCA.png"))
				if replace == False:
					self.text_replace(p, ["[DELE-QC.png]"],['无QC_PCA图'], family_ch=u'微软雅黑', family_en='微软雅黑')
			if "[Scatterplot_QC.png]" in p.text:
				replace = self.insert_png(p, "[Scatterplot_QC.png]", os.path.join('Appendix_A', '附件3', 'Scatterplot QC.png'))
				if replace == False:
					self.text_replace(p, ["[Scatterplot_QC.png]"],['无Scatterplot_QC图'], family_ch=u'微软雅黑', family_en='微软雅黑')
			if "[Volcano_Plot.png]" in p.text:
				self.insert_png(p, "[Volcano_Plot.png]", os.path.join('Volcano', f'Volcano_Plot_{self.groupvs}.png'))
			if "[venn.png]" in p.text:
				replace = self.insert_png(p, "[venn.png]", os.path.join('Venn', 'Venn.png'))
				if replace == False:
					self.text_replace(p, ["[venn.png]"],['无venn图'], family_ch=u'微软雅黑', family_en='微软雅黑')
			if '[Ratio_Distribution]' in p.text:
				self.insert_png(p, '[Ratio_Distribution]', os.path.join('Ratio_distribution', f'Ratio_Distribution_{self.groupvs}.png'))
			if "[cluster.png]" in p.text:
				if os.path.exists(os.path.join(self.groupvs, 'cluster')):
					self.insert_png(p, "[cluster.png]", os.path.join(self.groupvs, 'cluster', 'cluster1.png'))
				elif os.path.exists(os.path.join(self.groupvs, 'CLUSTER')):
					self.insert_png(p, "[cluster.png]", os.path.join(self.groupvs, 'CLUSTER', 'cluster1.png'))
			if "[go_enrichment.png]" in p.text:
				self.insert_png(p, "[go_enrichment.png]", os.path.join(self.groupvs, 'go', 'GO_Enrichment.png'))
			if "[KEGG_Enrichment.png]" in p.text:
				self.insert_png(p, "[KEGG_Enrichment.png]", os.path.join(self.groupvs, 'kegg', 'KEGG_Enrichment.png'))
			if "[Pathway.png]" in p.text:
				self.insert_png(p, "[Pathway.png]", os.path.join(self.groupvs, 'kegg', 'map', f'{self.keggID}.png'))
			if "[ppi.png]" in p.text:
				self.insert_png(p, "[ppi.png]", os.path.join(self.groupvs, 'ppi', 'ppi.png'))
			if "[pca.png]" in p.text:
				self.insert_png(p, "[pca.png]", os.path.join('PCA', 'PCA_samples.png'))
			if os.path.exists('WGCNA'):
				if "[PlotModule.png]" in p.text:
					self.insert_png(p, "[PlotModule.png]", os.path.join('WGCNA', 'Module', 'PlotModule.png'))
				if "[Module-trait.relationships.png]" in p.text:
					self.insert_png(p, "[Module-trait.relationships.png]", os.path.join('WGCNA', 'Module', 'Module-trait.relationships.png'))
				if "[Eigengene.dendrogram.heatmap.png]" in p.text:
					self.insert_png(p, "[Eigengene.dendrogram.heatmap.png]", os.path.join('WGCNA', 'Eigengene.dendrogram.heatmap.png'))
				if "[ProteinNetwork_heatmap_plot.png]" in p.text:
					self.insert_png(p, "[ProteinNetwork_heatmap_plot.png]", os.path.join('WGCNA', 'Module', 'ProteinNetwork_heatmap_plot.png'))
				if "[module.heatmap_barplot.png]" in p.text:
					png_path = os.path.join('WGCNA', 'Protein_ModulePlot')
					png = [i for i in os.listdir(png_path) if re.search(color, i, re.I)]
					if png:
						self.insert_png(p, "[module.heatmap_barplot.png]", os.path.join(png_path, png[0]))
				if "[WGCNA_GO_enrichment]" in p.text:
					self.insert_png(p, "[WGCNA_GO_enrichment]", os.path.join(wgcna_enrich, 'go', 'GO_Enrichment.png'))
				if "[WGCNA_KEGG_enrichment]" in p.text:
					self.insert_png(p, "[WGCNA_KEGG_enrichment]", os.path.join(wgcna_enrich, 'kegg', 'KEGG_Enrichment.png'))
				if "[moduel_GS_vs_MM]" in p.text:
					protein_corr = [i for i in os.listdir(os.path.join('WGCNA', 'Protein_relationship')) if os.path.isdir(os.path.join('WGCNA', 'Protein_relationship', i))]
					self.insert_png(p, "[moduel_GS_vs_MM]", os.path.join('WGCNA', 'Protein_relationship', protein_corr[0], f'{color}.moduel.GS_vs_MM.png'))

	def save(self, document):
		document.save(f'{self.project_num}{self.school}DIA报告.docx')
		print('报告生成完成')





