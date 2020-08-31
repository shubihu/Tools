import os
import sys
import re
import pandas as pd
from scipy import stats
import scipy
import math
import numpy as np
from numpy import mean
from numpy import var
from scipy.stats.distributions import t
from statsmodels.stats.multicomp import pairwise_tukeyhsd
from statsmodels.stats.multicomp import MultiComparison

import warnings
warnings.filterwarnings('ignore')

path = sys.argv[1]
os.chdir(path)

stamp_path = [os.path.join('Result', '06_Differential_analysis', 'STAMP'),
              os.path.join('Result', '07_FunctionPrediction', 'COG', 'STAMP'),
              os.path.join('Result', '07_FunctionPrediction', 'KEGG', 'STAMP')]

def ci(seqGroup1, seqGroup2, coverage):
	n1 = len(seqGroup1)
	n2 = len(seqGroup2)

	meanG1 = float(sum(seqGroup1)) / n1
	meanG2 = float(sum(seqGroup2)) / n2
	dp = meanG1 - meanG2

	varG1 = var(seqGroup1, ddof=1)
	varG2 = var(seqGroup2, ddof=1)

	normVarG1 = varG1 / n1
	normVarG2 = varG2 / n2
	unpooledVar = normVarG1 + normVarG2
	sqrtUnpooledVar = math.sqrt(unpooledVar)

	dof = (unpooledVar*unpooledVar) / ( (normVarG1*normVarG1)/(n1-1) + (normVarG2*normVarG2)/(n2-1) )

	tCritical = t.isf(0.5 * (1.0-coverage), dof) # 0.5 factor accounts from symmetric nature of distribution
	lowerCI = dp - tCritical*sqrtUnpooledVar
	upperCI = dp + tCritical*sqrtUnpooledVar
	return lowerCI, upperCI


def plot(path):
	sample_info = pd.read_csv('sample_info.txt', sep='\t')
	if re.search('COG', path, re.I):
		df = pd.read_csv(os.path.join('Result', '07_FunctionPrediction', 'COG', 'cog_predicted_L2.csv'), index_col=0)
	elif re.search('KEGG', path, re.I):
		df = pd.read_csv(os.path.join('Result', '07_FunctionPrediction', 'KEGG', 'kegg_predicted_L2.csv'), index_col=0)
	else:
		df = pd.read_csv(os.path.join(path, 'Genus.csv'), index_col=0)
	for vs in os.listdir(path):
		if vs.count('_vs_') == 1:
			g1, g2 = vs.split('_vs_')
			g1_sample = sample_info[sample_info['Group'] == g1]['SampleID'].tolist()
			g2_sample = sample_info[sample_info['Group'] == g2]['SampleID'].tolist()
			g1_df = df[g1_sample]
			g2_df = df[g2_sample]

			newdf = pd.DataFrame()
			newdf['{}_MeanRel(%)'.format(g1)] = g1_df.mean(axis=1) * 100
			newdf['{}_MeanRel(%)'.format(g2)] = g2_df.mean(axis=1) * 100
			pvalue = [stats.ttest_ind(list(i), list(j), equal_var = False)[1] for i, j in zip(g1_df.values, g2_df.values)]
			newdf['Pvalue'] = pvalue
			dp = [(float(sum(i)) / len(i) - float(sum(j)) / len(j)) * 100 for i, j in zip(g1_df.values, g2_df.values)]
			newdf['Difference between means'] = dp

			interval = [ci(i, j, 0.95) for i, j in zip(g1_df.values, g2_df.values)]
			lower = [i[0] * 100 for i in interval]
			upper = [i[1] * 100 for i in interval]
			newdf['95% lowerCI'] = lower
			newdf['95% upperCI'] = upper
			newdf.sort_values('Pvalue', inplace=True)
			pvalue0 = newdf['Pvalue'][0]

			newdf.index.name = 'Genus'
			out = os.path.join(path, vs, 't-test.csv')
			newdf.to_csv(out)
			if pvalue0 < 0.05:
				os.system('Rscript /home/jbwang/code/stamp/ExtendedErrorBar.R {}'.format(out))
			else:
				with open(os.path.join(path, vs, '备注.txt'), 'w') as f:
					f.write('该比较组无差异')
		elif vs.count('_vs_') > 1:
			g_list = vs.split('_vs_')
			g_sample = sample_info[sample_info['Group'].isin(g_list)]['SampleID'].tolist()
			data = df[g_sample]
			newdf = pd.DataFrame()
			pvalue_list, posthoc_list = [], []
			for i in range(df.shape[0]):
				tmp_df = pd.DataFrame()
				tmp_df['value'] = data.iloc[i]
				tmp_df['group'] = sample_info[sample_info['Group'].isin(g_list)]['Group'].tolist()

				g_values = [np.array(tmp_df[tmp_df['group'] == g]['value']) for g in g_list]
				pvalue = stats.f_oneway(*g_values)[1]
				if str(pvalue) == 'nan':
					pvalue_list.append('NA')
				else:
					pvalue_list.append(pvalue)

				if pvalue < 0.05:
					mc = MultiComparison(tmp_df['value'], tmp_df['group'])
					result = mc.tukeyhsd()
					result = result.summary().data

					post_hoc = ['_vs_'.join(result[i][:2]) for i in range(1, len(result)) if result[i][3] < 0.05]
					if post_hoc:
						post_hoc = ';'.join(post_hoc)
						posthoc_list.append(post_hoc)
					else:
						posthoc_list.append('NA')
				else:
					posthoc_list.append('NA')

			newdf['Genus'] = df.index.tolist()
			newdf['pvalue'] = pvalue_list
			newdf['post_hoc'] = posthoc_list
			out = os.path.join(path, vs, 'anova.csv')
			newdf.to_csv(out, index=None)


if __name__ == '__main__':
	for p in stamp_path:
		plot(p)



 