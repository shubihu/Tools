import urllib
import re
import os
import sys
import time
import requests
from http import cookiejar
from multiprocessing.pool import ThreadPool

headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.61 Safari/537.36', 
		'Referer':'https://www.genome.jp/kegg/tool/map_pathway2.html'}
hosturl='https://www.genome.jp/kegg/tool/map_pathway2.html'
posturl='https://www.genome.jp/kegg-bin/color_pathway_object'

# postData = {'org_type': 'other',
# 			'org': 'hsa',
# 			'other_dbs': '',
# 			'unclassified': 'hsa:10105 yellow\nhsa:5481 red',
# 			's_sample':'',
# 			'color_list': '',
# 			'default': 'pink',
# 			'target': 'alias',
# 			'org_name': 'hsa'}

def post_data(infile, species):
	unclassified = ''
	with open(infile) as f:
		for line in f:
			unclassified += line

	postData = {'org_type': 'other',
			'org': species,
			'other_dbs': '',
			'unclassified': unclassified,
			's_sample':'',
			'color_list': '',
			'default': 'pink',
			'target': 'alias',
			'org_name': species}

	return postData

def get_url_list(hosturl, posturl, postData, headers):
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
	text = str(res.read(), encoding = "utf-8")
	href_list = re.findall('href="(.*\.args)"', text)
	href_list = [f'https://www.genome.jp/{i}' for i in href_list]

	return href_list[0]

def get_png(url):
	try:
		req = requests.get(url, headers=headers)
		if req.status_code == 200:
			png = re.findall('img src="(.*?)"', req.text)
			if png:
				png = f'https://www.genome.jp/{png[0]}'
				return png
	except Exception as e:
		print(e)
		print("Get png url failed:{}".format(url))

def get_png_thread(url_list):
	pool = ThreadPool(processes=16)
	thread_list = [pool.apply_async(func=get_png, args=(url,)) for url in url_list] #参数以元组形式传入

	pool.close()
	pool.join()
	# 获取输出结果
	png_list = [p.get() for p in thread_list]

	return png_list

def download_pic(out_path, png_url):
	try:
		map_name = png_url.split('/')[-1].split('_')[0]
		req_png = urllib.request.Request(png_url, headers=headers)
		res_png = urllib.request.urlopen(req_png)

		filename = os.path.join(out_path, f'{map_name}.png')
		if res_png.getcode() == 200:
			with open(filename, "wb") as f:
				f.write(res_png.read())
	except Exception as e:
		print(e)
		print("Download image failed:{}".format(png_url))

def download_pic_thread(out_path, png_list):
	pool = ThreadPool(processes=16)
	thread_list = [pool.apply_async(func=download_pic, args=(out_path, png_url)) for png_url in png_list]

	pool.close()
	pool.join()

if __name__ == '__main__':
	infile = sys.argv[1]
	species = sys.argv[2]
	out_path = sys.argv[3]
	start = time.time()
	print('提交请求')
	postData = post_data(infile, species)
	url_list = get_url_list(hosturl, posturl, postData, headers)
	# png_url_list = get_png_thread(url_list)
	png_url = get_png(url_list)

	print('请求完成')
	print('开始下载图片')
	# download_pic_thread(out_path, png_url_list)
	download_pic(out_path, png_url)
	print('下载完成')
	end = time.time()
	print(f'所需时间：{str(int(end - start))}秒')


