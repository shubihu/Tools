# -*- coding: future_fstrings -*-     # should work even without -*-


import os
import paramiko
from stat import S_ISDIR as isdir

#服务器信息，主机名（IP地址）、端口号、用户名及密码
hostname = "192.168.130.252"
port = 22
username = "root"
password = "apt123.com"

def get_version():
	client = paramiko.SSHClient()
	client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
	client.connect(hostname, port, username, password, compress=True)
	sftp_client = client.open_sftp()
	remote_file = sftp_client.open("/database/metabolome/MetaReport/newversion.txt")#文件路径
	try:
		for line in remote_file:
			if line.startswith('v'):
				version = line.strip().replace('v', '')
	finally:
		remote_file.close()

	return float(version)

def check_local_dir(local_dir_name):
	if not os.path.exists(local_dir_name):
		os.makedirs(local_dir_name)

def down_from_remote(sftp, remote_dir_name, local_dir_name):
	"""远程下载文件"""
	remote_file = sftp.stat(remote_dir_name)
	if isdir(remote_file.st_mode):
		# 文件夹，不能直接下载，需要继续循环
		check_local_dir(local_dir_name)
		print('开始下载：' + remote_dir_name)
		for remote_file_name in sftp.listdir(remote_dir_name):
			sub_remote = os.path.join(remote_dir_name, remote_file_name)
			sub_remote = sub_remote.replace('\\', '/')
			sub_local = os.path.join(local_dir_name, remote_file_name)
			sub_local = sub_local.replace('\\', '/')
			down_from_remote(sftp, sub_remote, sub_local)
	else:
		# 文件，直接下载
		print('开始下载：' + remote_dir_name)
		sftp.get(remote_dir_name, local_dir_name)

def update(local_dir):
	# 远程文件路径（需要绝对路径）
	remote_dir = '/database/metabolome/MetaReport/'
	# 连接远程服务器
	t = paramiko.Transport((hostname, port))
	t.connect(username=username, password=password)
	sftp = paramiko.SFTPClient.from_transport(t)

	# 远程文件开始下载
	down_from_remote(sftp, remote_dir, local_dir)
	# 关闭连接
	t.close()

