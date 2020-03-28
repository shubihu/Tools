# import pandas as pd
import os
infile = input('请输入文件(直接拖拽即可, 结果文件为[输入文件-Result.csv], 存放路径和输入文件一致)：')
# infile = 'C:/Users/wjb/Desktop/20200320-FSV-test.txt'
outfile = f'{os.path.splitext(infile)[0]}-Result.csv'
# data = pd.read_csv(infile, sep='\t')
# data = data[(data['Unnamed: 2']=='Sample Type') | (data['Unnamed: 2']=='Unknown')]
# data.columns = data.iloc[0,].tolist()
# data.drop(data.index[0], inplace=True)

# columns = ['Sample Name', 'Sample ID', 'Acquisition Date', 'Analyte Peak Name', 'Analyte Units', 'Calculated Concentration (ng/mL)']
columns = ['条码号', '中科新生命条码号', '检测日期', '检验项目', '单位', '结果']

peak_dict = {'VK1': '维生素K1', 'VA': '维生素A', 'VE': '维生素E', '25OHVD3': '25羟基维生素D3', '25OHVD2': '25羟基维生素D2', 'VB1': '维生素B1',
			'VB2': '维生素B2', 'VB3': '维生素B3', 'VB5': '维生素B5', 'VB6': '维生素B6', 'VB7': '维生素B7', 'VB9': '维生素B9', 'VB12': '维生素B12'}

# data['Analyte Peak Name'] = data['Analyte Peak Name'].apply(lambda x: peak_dict.get(x))
# data.sort_values(by='Analyte Peak Name', ascending=False, inplace=True)

with open(infile) as f1:
	with open(outfile, 'w') as f2:
		f2.write(','.join(columns) + '\n')
		result = []		
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

				if '25OHVD3' in line:
					unit = line[18]
					if '<' in line[75]:
						result.append(0)
					else:
						result.append(float(line[75]))
				if '25OHVD2' in line:
					if '<' in line[75]:
						result.append(0)
					else:
						result.append(float(line[75]))

				if len(result) == 2:
					add_line = [''] * 3
					add_line.extend(['25羟基维生素D', unit])
					# print(add_line)
					calcu = result[0] + result[1]
					add_line.append(str(calcu))
					f2.write(','.join(add_line) + '\n')
					result = []
