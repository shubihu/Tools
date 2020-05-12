import os
import re
import xlrd
from sys import exit

project = int(input('请输入项目类型(输入对应的序号即可)，\
	\n1: 靶向; \
	\n2: 非靶; \
	\n3: 脂质绝对定量; \
	\n4: 脂质相对定量;\n:'))

path = input('请输入项目路径(直接拖拽即可):')

def fileCheck(project, path):
	if not os.path.exists(os.path.join(path, 'newname.xlsx')):
		print(f'Error:不存在{os.path.join(path, "newname.xlsx")}')
		input('按任意键结束')
		exit()
	if not os.path.exists(os.path.join(path, 'groupvs.txt')):
		print(f'Error:不存在{os.path.join(path, "groupvs.txt")}')
		input('按任意键结束')
		exit()

	if project == 1:
		if not os.path.exists(os.path.join(path, '附件1.xlsx')):
			print('Error:不存在附件1.xlsx')
			input('按任意键结束')
			exit()
	if project == 2:
		if not os.path.exists(os.path.join(path, 'neg-dele-iso.csv')):
			print('Error:不存在neg-dele-iso.csv')
			input('按任意键结束')
			exit()
		if not os.path.exists(os.path.join(path, 'pos-dele-iso.csv')):
			print('Error:不存在pos-dele-iso.csv')
			input('按任意键结束')
			exit()
	if project == 4:
		if not os.path.exists(os.path.join(path, 'data_after_pre_pos.csv')):
			print('Error:不存在data_after_pre_pos.xlsx')
			input('按任意键结束')
			exit()
		if not os.path.exists(os.path.join(path, 'data_after_pre_neg.csv')):
			print('Error:不存在data_after_pre_neg.xlsx')
			input('按任意键结束')
			exit()


def main(project, path):
	fileCheck(project, path)
	groupvs = set()
	with open(os.path.join(path, "groupvs.txt")) as g:
		for line in g:
			if '_vs_' in line:
				line = line.strip().split('_vs_')
				groupvs.add(line[0])
				groupvs.add(line[1])
			elif '|' in line:
				line = line.strip().split('|')
				for i in line:
					groupvs.add(i)

	newname = xlrd.open_workbook(os.path.join(path, "newname.xlsx"))
	table = newname.sheet_by_index(0)

	samples, groups = set(), set()
	for i in range(1, table.nrows):
		samples.add(table.cell_value(i, 0))
		for j in range(1, table.ncols):
			groups.add(table.cell_value(i, j))

	if project == 1:
		data = xlrd.open_workbook(os.path.join(path, '附件1.xlsx'))
		title = set(data.sheet_by_index(0).row_values(0))

	if project == 2:
		with open(os.path.join(path, 'neg-dele-iso.csv')) as f1:
			title1 = set(f1.readline().strip().split(','))
		with open(os.path.join(path, 'pos-dele-iso.csv')) as f1:
			title2 = set(f1.readline().strip().split(','))
		title = title1 & title2

	if project == 3:
		file = [i for i in os.listdir(path) if re.search('pos.*neg', i, re.I)]
		is_file = [i for i in os.listdir(path) if re.search('is', i, re.I)]
		if all([file, is_file]):
			data = xlrd.open_workbook(os.path.join(path, file[0]))
			title = set(data.sheet_by_index(0).row_values(0))
			weight = xlrd.open_workbook(os.path.join(path, 'weight.xlsx'))
			w_table = weight.sheet_by_index(0)
			w_samples = set(w_table.cell_value(i, 0) for i in range(1, w_table.nrows))
			if not w_samples.issubset(title):
				w_sample = w_samples - title
				print(f'weight.xlsx中该样本名{w_sample}有误, 不存在于{file[0]}中')
			is_data = xlrd.open_workbook(os.path.join(path, is_file[0]))
			is_samples = is_data.sheet_by_index(0).row_values(0)
			del is_samples[:is_samples.index('rt') + 1]
			is_samples = set(is_samples)
			if not is_samples.issubset(title):
				is_sample = is_samples - title
				print(f'{is_file[0]}中该样本名{is_sample}有误, 不存在于{file[0]}中')
		else:
			print('Error:不存在pos and neg或 IS 相关文件')
			input('按任意键结束')
			exit()

	if project == 4:
		with open(os.path.join(path, 'data_after_pre_pos.csv')) as p:
			pos_title = set(p.readline().strip().split(','))
		with open(os.path.join(path, 'data_after_pre_neg.csv')) as n:
			neg_title = set(n.readline().strip().split(','))
		title = pos_title & neg_title
		if not groups.issubset(title):
			newgroup = groups - title
			print(f'newname.xlsx中该组名{newgroup}不存在于data_after_pre_neg/pos文件中')

	if not samples.issubset(title):
		sample = samples - title
		print(f'newname.xlsx中该样本名{sample}有误')

	if not groupvs.issubset(groups):
		group = groupvs - groups
		print(f'groupvs.txt中该组名{group}有误, 不存在于newname.xlsx中')

	print('检查完成')
	input('按任意键结束')

if __name__ == '__main__':
	main(project, path)

