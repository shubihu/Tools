import os
from collections import defaultdict
from itertools import chain

infile = input('请输入文件(直接拖拽即可, 结果文件为[输入文件-Result.csv], 存放路径和输入文件一致)：')
# infile = 'C:/Users/wjb/Desktop/20200320-FSV-test.txt'
outfile = f'{os.path.splitext(infile)[0]}-Result.csv'

def handle(infile):
	with open(infile) as f:
		with open(outfile, 'w') as fw:
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

if __name__ == '__main__':
	handle(infile)


					


	