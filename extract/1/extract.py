# import pandas as pd
import os
# infile = input('请输入文件(直接拖拽即可, 结果文件为[输入文件-Result.csv], 存放路径和输入文件一致)：')
infile = r'E:\develop\Tools\extract\1\20200320-FSV-test.txt'
outfile = f'{os.path.splitext(infile)[0]}-Result.csv'

# columns = ['Sample Name', 'Sample ID', 'Acquisition Date', 'Analyte Peak Name', 'Analyte Units', 'Calculated Concentration (ng/mL)']
columns = ['条码号', '中科新生命条码号', '检测日期', '检验项目', '单位', '结果']

peak_dict = {'VK1': '维生素K1', 'VA': '维生素A', 'VE': '维生素E', '25OHVD3': '25羟基维生素D3', '25OHVD2': '25羟基维生素D2', 'VB1': '维生素B1',
			'VB2': '维生素B2', 'VB3': '维生素B3', 'VB5': '维生素B5', 'VB6': '维生素B6', 'VB7': '维生素B7', 'VB9': '维生素B9', 'VB12': '维生素B12',
			'Clozapine': '氯氮平', 'Desmethylclozapine': '去甲氯氮平', 'Perphenazine': '奋乃静', 'Quetiapine': '喹硫平', 'Sulpiride': '舒必利',
			'Olanzapine': '奥氮平', 'Haloperidol': '氟哌啶醇', 'Paliperidone': '利培酮', '9-hydroxyrisperidone': '9-羟利培酮', 'amisulpride': '氨磺必利',
			'Chlorpromazine': '氯丙嗪', 'Fluphenazine': '氟奋乃静', 'Aripiprazole': '阿立哌唑', 'Dehydro Aripiprazole': '脱氢阿立哌唑',
			'Ziprasidone': '齐拉西酮', 'Duloxetine': '度洛西汀', 'Fluoxetine': '氟西汀', 'Norfluoxetine': '去甲氟西汀', 'mirtazapine': '米氮平',
			'Paroxetine': '帕罗西汀', 'sertraline': '舍曲林', 'Trazodone': '曲唑酮', 'Venlafaxine': '文拉法辛', 'O-desmethylvenlafaxine': 'O-去甲文拉法辛',
			'Citalopram': '西酞普兰', 'Escitalopram': '艾司西酞普兰', 'Bupropion': '安非他酮', 'Hydroxybupropion': '羟安非他酮', 'Fluvoxoxamine': '氟伏沙明',
			'Alprazolam': '阿普唑仑', 'Oxcarbazepine': '奥卡西平', '10-hydroxycarbamazepine': '10-羟卡马西平', 
			'1MHis': '1-甲基组氨酸',
			 '3MHis': '3-甲基组氨酸',
			 '5-Fluorouracil': '5-氟尿嘧啶',
			 '6AHC': '6-氨基己酸',
			 'Aad': 'α-氨基己二酸',
			 'Abu': '2-氨基丁酸',
			 'Ala': '丙氨酸',
			 'Arg': '精氨酸',
			 'Asn': '天门冬酰胺',
			 'Asp': '天门冬氨酸',
			 'Bromazepam': '溴西泮',
			 'CA': '胆酸',
			 'CDCA': '鹅脱氧胆酸',
			 'Carbamazepine': '卡马西平',
			 'Cit': '瓜氨酸',
			 'Clonazepam': '氯硝西泮',
			 'DA': '多巴胺',
			 'DCA': '脱氧胆酸',
			 'Diazepam': '地西泮',
			 'Docetaxel': '多西他赛',
			 'Donepezil': '多奈哌齐',
			 'E': '肾上腺素',
			 'Estazolam': '艾司唑仑',
			 'Fluconazole': '氟康唑',
			 'GABA': 'γ-氨基丁酸',
			 'GCA': '甘氨胆酸',
			 'GCDCA': '甘氨鹅脱氧胆酸',
			 'GDCA': '甘氨脱氧胆酸',
			 'GLCA': '甘氨石胆酸',
			 'GUDCA': '甘氨熊脱氧胆酸',
			 'Glu': '谷氨酸',
			 'Gly': '甘氨酸',
			 'HVA': '高香草酸',
			 'Harg': '同型精氨酸',
			 'His': '组氨酸',
			 'Hpro': '同型脯氨酸',
			 'Hyp': '羟基脯氨酸',
			 'Ile': '异亮氨酸',
			 'Imipenem': '亚胺培南',
			 'Isoniazid': '异烟肼',
			 'KC': '犬尿氨酸',
			 'LCA': '石胆酸',
			 'Leu': '亮氨酸',
			 'Levetiracetam': '左乙拉西坦',
			 'Linezolid': '利奈唑胺',
			 'Lys': '赖氨酸',
			 'MN': '变肾上腺素',
			 'Meropenem': '美洛培南',
			 'Met': '甲硫氨酸',
			 'Methotrexate': '甲氨蝶呤',
			 'Midazolam': '咪达唑仑',
			 'Moxifloxacin': '莫西沙星',
			 'NE': '去甲肾上腺素',
			 'NMN': '去甲变肾上腺素',
			 'Nitrazepam': '硝西泮',
			 'Orn': '鸟氨酸',
			 'Phe': '苯丙氨酸',
			 'Phenobarbital': '苯巴比妥',
			 'Phenytoin sodium': '苯妥英钠',
			 'Pro': '脯氨酸',
			 'Rifampin': '利福平',
			 'Rivaroxaban': '利伐他班',
			 'Sar': '肌氨酸',
			 'Ser': '丝氨酸',
			 'TCA': '牛磺胆酸',
			 'TCDCA': '牛磺鹅脱氧胆酸',
			 'TDCA': '牛磺脱氧胆酸',
			 'TLCA': '牛磺石胆酸',
			 'TUDCA': '牛磺熊脱氧胆酸',
			 'Teicoplanin': '替考拉宁',
			 'Temazepam': '替马西泮',
			 'Thr': '苏氨酸',
			 'Topiramate': '托吡酯',
			 'Trp': '色氨酸',
			 'Tyr': '酪氨酸',
			 'UDCA': '熊脱氧胆酸',
			 'VMA': '香草苦杏仁酸',
			 'Val': '缬氨酸',
			 'Valproic Acid': '丙戊酸',
			 'Vancomycin': '万古霉素',
			 'bAib': '3-氨基异丁酸',
			 'bAla': 'β-丙氨酸',
			 'dabigatran': '达比加群',
			 'efavirenz': '依非韦伦',
			 'entecavir': '恩替卡韦',
			 'lamivudine': '拉米夫定',
			 'lamotrigine': '拉莫三嗪',
			 'lorazepam': '劳拉西泮',
			 'memantine': '美金刚',
			 'oxazepam': '奥沙西泮',
			 'rivastigmine': '卡巴拉汀',
			 'taxol': '紫杉醇',
			 'tenofovir': '替诺福韦',
			 'voriconazole': '伏立康唑',
			 'zolpidem': '唑吡坦'}

def add_line(line, x, y, xy, result=[]):
	unit = ''
	if x in line:
		unit = line[18]
		if '<' in line[75]:
			result.append(0)
		else:
			result.append(float(line[75]))
	if y in line:
		if '<' in line[75]:
			result.append(0)
		else:
			result.append(float(line[75]))

	if result:
		if len(result) % 2 == 0:			
			result = result[-2:]
			add_line = [''] * 3
			add_line.extend([xy, unit])
			# print(add_line)
			calcu = result[0] + result[1]
			add_line.append(str(calcu))
			print(add_line)
			# f2.write(','.join(add_line0[0]) + '\n')






with open(infile) as f1:
	with open(outfile, 'w') as f2:
		f2.write(','.join(columns) + '\n')
		# VD_result, Cl_result, Pa_result, Ar_result, Fl_result, Ve_result, Bu_result, Ox_result = [], [], [], [], [], [], [], []		
		for line in f1:
			if 'Unknown' in line:
				new_line = []
				line = line.split('\t')
				new_line.extend(line[:2])
				new_line.append(line[6])
				new_line.append(peak_dict.get(line[17], line[17]))
				new_line.append(line[18])
				new_line.append(line[75])
				f2.write(','.join(new_line) + '\n')

				add_line(line, '25OHVD3', '25OHVD2', '25羟基维生素D')
				

				# if '25OHVD3' in line:
				# 	unit = line[18]
				# 	if '<' in line[75]:
				# 		VD_result.append(0)
				# 	else:
				# 		VD_result.append(float(line[75]))
				# if '25OHVD2' in line:
				# 	if '<' in line[75]:
				# 		VD_result.append(0)
				# 	else:
				# 		VD_result.append(float(line[75]))

				# if len(VD_result) == 2:
				# 	add_line = [''] * 3
				# 	add_line.extend(['25羟基维生素D', unit])
				# 	# print(add_line)
				# 	calcu = VD_result[0] + VD_result[1]
				# 	add_line.append(str(calcu))
				# 	f2.write(','.join(add_line) + '\n')
				# 	VD_result = []

				# if 'Clozapine' in line:
				# 	unit = line[18]
				# 	if '<' in line[75]:

# os.system('pause')
