from collections import defaultdict
from itertools import chain
import xlrd
from openpyxl import Workbook

with open('MW.txt') as f:
	mw_dict = {i.split('\t')[0]: i.strip().split('\t')[1] for i in f}

def handle1(file):
	with open(file) as f:
		with open(file.replace('.res', '-Result.csv'), 'w') as fw:
			mydict = defaultdict(list)
			for line in f:
				if line:			
					line = line.strip().split('\t')
					if line[0].startswith('S'):
						label1 = line[1].split('.')[2]
						label2 = line[1].split('.')[3].split(' ')[0]
						key = ','.join([label1, label2])
						mydict[key] = list()
					
					if line[0].startswith('P'):
						mydict[key].append(line)

			for k, y in mydict.items():
				if y:
					# tmp = [j for i in y for j in i]
					tmp = list(chain(*y))
					tmp.insert(0, k.split(',')[0])
					tmp.insert(1, k.split(',')[1])
					fw.write(','.join(tmp) + '\n')
				else:
					fw.write(k + '\n')

def isequal(seq1, seq2):
	for k,y in zip(seq1, seq2):
			if k != y:
				mw_minus = abs(float(mw_dict.get(k)) - float(mw_dict.get(y)))
				index = seq1.index(k)
				# index2 = seq2.index(y)
				total1 = sum([float(mw_dict.get(j)) for j in seq1[:index]])
				total2 = sum([float(mw_dict.get(j)) for j in seq2[:index]])
				total_minus = abs(total1 - total2)
				if all([mw_minus < 0.1, total_minus < 0.5]):
					continue
				break
	else:
		return 1
	return 0


def isequal2(seq1, seq2):
	for k,y in zip(seq1, seq2):
			if k != y:
				mw_minus = abs(float(mw_dict.get(k)) - float(mw_dict.get(y)))
				index = seq1.index(k)
				# index2 = seq2.index(y)
				total1 = sum([float(mw_dict.get(j)) for j in seq1[:index]])
				total2 = sum([float(mw_dict.get(j)) for j in seq2[:index]])
				total_minus = abs(total1 - total2)
				if any([mw_minus > 0.1, total_minus > 0.5]):
					break
	else:
		return 1
	return 0


def handle2(file):
	data = xlrd.open_workbook(file)
	wb = Workbook()
	ws = wb.active

	table = data.sheet_by_index(0)
	title = table.row_values(0)
	title.insert(5, 'Equal')
	ws.append(title)
	for i in range(1, table.nrows):
		values = table.row_values(i)
		seq1 = values[3]
		seq2 = values[4]
		if seq1 == seq2:
			values.insert(5, 1)
		else:
			values.insert(5, isequal2(seq1, seq2))
		# values = [str(i) for i in values]
		ws.append(values)
	wb.save(file.replace('.xlsx', '-Result.xlsx'))

					


	