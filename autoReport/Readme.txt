一、 运行环境python3 : 安装依赖 pip install -r requirements.txt (若已安装可忽略)
二、 运行脚本前自行修改脚本所在目录的project_information.xls文件里对应的内容，生成报告结果保存在项目路径下。
三、 运行脚本 python autoReport.py -i 项目路径 -t 项目类型
参数：
	-i 项目路径 （必须参数）
	-t 项目类型 （必须参数）
	-fc FoldChange （可选参数，labfree默认为2，Itraq和TMT为1.2， 若默认可不选该参数）


项目类型参数（ l或L 为labfree , i或I 为Itraq , t或T 为TMT)

注：可执行 python autoReport.py --help查看帮助