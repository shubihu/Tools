一、 运行环境python3 : 安装依赖 pip install -r requirements.txt (若已安装可忽略)
二、 项目路径下需有project_information文件（非靶除外，非靶需要两个文件，一个**report_info文件及information文件），运行脚本前自行修改该文件里对应的内容，生成报告结果保存在项目路径下。
三、 运行脚本 python MetaReport.py -i 项目路径 -t 项目类型
参数：
	-i 项目路径 （必须参数）
	-t 项目类型 （必须参数）


项目类型参数(ld为lipids，lc为长链脂肪酸，sc为短链脂肪酸， nt为非靶代谢)

注：可执行 python MetaReport.py --help查看帮助

若报关于 future_fstrings 的语法错误，执行 pip3 install future_fstrings即可。