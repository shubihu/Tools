import urllib
import re
import os
import sys
import time
import requests
import pandas as pd
from http import cookiejar
from multiprocessing.pool import ThreadPool

headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.61 Safari/537.36', 
		'Referer':'http://www.csbio.sjtu.edu.cn/bioinf/virus-multi/'}
hosturl='http://www.csbio.sjtu.edu.cn/bioinf/virus-multi/'
posturl='http://www.csbio.sjtu.edu.cn/cgi-bin/VirusmPLoc.cgi'

def process(infile):
	with open(infile) as f:
		seq_dict = {}
		for line in f:
			if '>' in line:
				pro = line.strip()
			else:
				seq_dict[pro] = line.strip()

		seq_list = [f'{i}\n{j}' for i, j in seq_dict.items()]

	return seq_list

def post_data(seq):
	postData = {'mode': 'string',
			'S1': seq,
			'B1': 'Submit'}

	return postData

def predict(hosturl, posturl, postData, headers):
	#设置一个cookie处理器，它负责从服务器下载cookie到本地，并且在发送请求时带上本地的cookie
	cj = cookiejar.CookieJar()
	cookie_support = urllib.request.HTTPCookieProcessor(cj)
	opener = urllib.request.build_opener(cookie_support, urllib.request.HTTPHandler)
	urllib.request.install_opener(opener)
	#打开登录主页面（他的目的是从页面下载cookie，这样我们在再送post数据时就有cookie了，否则发送不成功）
	urllib.request.urlopen(hosturl)
	#需要给Post数据编码
	postDataEncode = urllib.parse.urlencode(postData).encode(encoding='UTF8')
	#通过urllib2提供的request方法来向指定Url发送我们构造的数据，并完成数据发送过程
	request = urllib.request.Request(posturl, postDataEncode, headers)
	res = urllib.request.urlopen(request)
	df = pd.read_html(res)[-1]
	protein = df.iloc[1, 0]
	locations = ';'.join(df.iloc[1, 1].rstrip('.').split('.'))

	return protein, locations

def predict_thread(seq_list):
	pool = ThreadPool(processes=16)
	postData_list = [post_data(seq) for seq in seq_list]
	thread_list = [pool.apply_async(func=predict, args=(hosturl, posturl, postData, headers)) for postData in postData_list] #参数以元组形式传入

	pool.close()
	pool.join()
	# 获取输出结果
	result_list = [p.get() for p in thread_list]

	return result_list


if __name__ == '__main__':
	# infile = sys.argv[1]
	infile = 'all.fasta'
	seq_list = process(infile)
	start = time.time()
	print('提交请求')
	result_list = predict_thread(seq_list)
	with open('protein2Loc.txt', 'w') as fw:
		fw.write('Protein ID\tSubLocations\n')
		for result in result_list:
			fw.write(f'{result[0]}\t{result[1]}\n')

	end = time.time()
	print(f'所需时间：{str(int(end - start))}秒')


