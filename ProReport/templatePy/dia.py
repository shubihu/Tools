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
		# try:
		information_file = [i for i in os.listdir('.') if re.search('infor?mation', i, re.I)]
		if information_file:
			projectinfo = pd.read_excel(information_file[0], header=None)
		else:
			print('Error:该项目下无project information表')
			exit()
		self.school = projectinfo.iloc[0, 1]
		self.project_num = projectinfo.iloc[1, 1]
		self.sample_type = projectinfo.iloc[2, 1]
		self.sample_num = projectinfo.iloc[3, 1]
		self.HPRP_num = projectinfo.iloc[4, 1]
		self.DDA_protein = projectinfo.iloc[5, 1]
		self.DDA_peptide = projectinfo.iloc[6, 1]
		self.aver_peak = projectinfo.iloc[7, 1]
		self.peak_capacity = projectinfo.iloc[8, 1]
		self.CV_median = projectinfo.iloc[9, 1]
		self.groupvs = projectinfo.iloc[10, 1]
		self.experiment_method = projectinfo.iloc[11, 1]
		self.DDA_method = projectinfo.iloc[12, 1]
		self.DIA_method = projectinfo.iloc[13, 1]

		origi_record = pd.read_excel('原始记录.xls', sheet_name=0).fillna('')
		self.diff_groupvs = origi_record[origi_record.loc[:, 'group'].str.contains('vs|oneway|twoway')]
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

		# except Exception as e:
		# 	print(e)
		# 	print('请检查project_information中信息是否填写正确')
		# 	# raise
		# 	exit()

	def header(self, paragraphs):
		today = str(datetime.date.today())
		pa = paragraphs[14].add_run(self.school)
		self.paragraph_format(pa, size=12, family=u'微软雅黑')
		pa = paragraphs[19].add_run(self.project_num)
		self.paragraph_format(pa, size=12, family=u'微软雅黑')
		pa = paragraphs[20].add_run(today)
		self.paragraph_format(pa, size=12, family="微软雅黑")

	def table_data(self, tables):
		self.paragraph_format(tables[2].cell(1,1).paragraphs[0].add_run(str(self.DDA_protein)), size=10, family=u'微软雅黑')
		self.paragraph_format(tables[2].cell(1,2).paragraphs[0].add_run(str(self.DDA_peptide)), size=10, family=u'微软雅黑')

		self.insert_table(self.pro_data, tables[3], size=10, family_ch=u'微软雅黑', family_en='微软雅黑')
		self.insert_table(self.diff_groupvs, tables[6], size=10, family_ch=u'微软雅黑', family_en='微软雅黑')

	def text_png_data(self, paragraphs):
		color = 'blue'
		if os.path.exists(os.path.join('WGCNA', 'Module_Enrichment', color)):
			wgcna_enrich = os.path.join('WGCNA', 'Module_Enrichment', color)
		else:
			wgcna_enrich = os.listdir(os.path.join('WGCNA', 'Module_Enrichment'))[0]
			color = os.path.basename(wgcna_enrich)	
		
		for i, p in enumerate(paragraphs):
			if 'groupvs' in p.text:
				self.text_replace(p, ['groupvs'], [self.groupvs])
			if 'sample_type' in p.text:
				self.text_replace(p, ['sample_type', 'sample_num', 'HPRP_num', 'DDA_protein'], [self.sample_type, str(self.sample_num), str(self.HPRP_num), str(self.DDA_protein)], family_ch=u'微软雅黑', family_en='微软雅黑')
			if 'aver_peak' in p.text:
				if float(self.aver_peak) * 10 % 10 == 0:
					self.text_replace(p, ['aver_peak'], [str(int(self.aver_peak))], family_ch=u'微软雅黑', family_en='微软雅黑')
				else:
					self.text_replace(p, ['aver_peak'], [str(self.aver_peak)], family_ch=u'微软雅黑', family_en='微软雅黑')
			if 'peak_capacity' in p.text:
				if float(self.peak_capacity) * 10 % 10 == 0:
					self.text_replace(p, ['peak_capacity'], [str(int(self.peak_capacity))], family_ch=u'微软雅黑', family_en='微软雅黑')
				else:
					self.text_replace(p, ['peak_capacity'], [str(self.peak_capacity)], family_ch=u'微软雅黑', family_en='微软雅黑')
			if 'CV_median' in p.text:
				self.text_replace(p, ['CV_median'], [str(self.CV_median)], family_ch=u'微软雅黑', family_en='微软雅黑')
			if 'total_pro_num' in p.text:
				self.text_replace(p, ['total_pro_num', 'half_pro_num'], [str(self.total_pro_num), str(self.half_pro_num)], family_ch=u'微软雅黑', family_en='微软雅黑')
			if 'upRatio' in p.text:
				self.text_replace(p, ['upRatio', 'upRatio', 'downRatio'], [str(self.fc), str(self.fc), str(round(1 / self.fc, 2))], family_ch=u'微软雅黑', family_en='微软雅黑')
			if 'groupvs' in p.text:
				self.text_replace(p, ['groupvs'], [self.groupvs])
			if 'BP-TOP5' in p.text:
				if all([len(self.bp_top) == 0, len(self.mf_top) == 0, len(self.cc_top) == 0]):
					self.text_replace(p, ['，发生了显著性变化'], [''])
				if self.bp_top:
					self.text_replace(p, ['BP-TOP5'], [self.bp_top])
				else:
					self.text_replace(p, ['BP-TOP5等重要生物学过程'], ['无P值小于0.05的显著性生物学过程'])
			if 'MF-TOP5' in p.text:
				if self.mf_top:
					self.text_replace(p, ['MF-TOP5'], [self.mf_top])
				else:
					self.text_replace(p, ['MF-TOP5等分子功能'], ['无P值小于0.05的显著性分子功能'])
			if 'CC-TOP5' in p.text:
				if self.cc_top:
					self.text_replace(p, ['CC-TOP5'], [self.cc_top])
				else:
					self.text_replace(p, ['CC-TOP5等定位蛋白质'], ['无P值小于0.05的显著性定位蛋白'])
			if 'KeggEnrich-top5' in p.text:
				if self.keggEnrich_top5:
					self.text_replace(p, ['KeggEnrich-top5', 'KeggEnrich-top1'], ['，'.join(self.keggEnrich_top5), self.pathway])
				else:
					self.text_replace(p, ['KeggEnrich-top5等重要通路发生了显著变化', 'KeggEnrich-top1'], ['该比较组没有P值小于0.05的显著性富集通路', self.pathway])
			if 'expr_sample_type' in p.text:
				self.text_replace(p, ['expr_sample_type'], [self.sample_type], size=12, family_ch=u'微软雅黑', family_en='微软雅黑', bold=True)
			if 'Experiment_method' in p.text:
				self.text_replace(p, ['Experiment_method'], [self.experiment_method], family_ch=u'微软雅黑', family_en='微软雅黑')
			if 'DDA_method' in p.text:
				self.text_replace(p, ['DDA_method'], [self.DDA_method], family_ch=u'微软雅黑', family_en='微软雅黑')
			if 'DIA_method' in p.text:
				self.text_replace(p, ['DIA_method'], [self.DIA_method], family_ch=u'微软雅黑', family_en='微软雅黑')

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
				cv_png = [i for i in os.listdir(os.path.join('Appendix_A', '附件3')) if re.search('cv.*.png', i, re.I)]
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
				self.insert_png(p, '[Ratio_Distribution]', os.path.join('Ratio_distribution', f'Ratio_distribution_{self.groupvs}.png'))
			if "[cluster.png]" in p.text:
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
		document.save(f'{self.project_num}{self.school}{self.sample_type}DIA报告.docx')
		print(f'报告生成完成，报告路径为：{os.path.join(os.getcwd(), f"{self.project_num}{self.school}{self.sample_type}DIA报告.docx")}')





