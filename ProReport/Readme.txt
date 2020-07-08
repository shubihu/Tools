一、 运行环境python3 : 安装依赖 pip install -r requirements.txt (若已安装可忽略)
二、 项目路径下需有project_information文件（同网页版报告使用同一个信息表），运行脚本前自行修改该文件里对应的内容，生成报告结果保存在项目路径下。
三、 运行脚本 python ProReport.py -i 项目路径 -t 项目类型
参数：
	-i 项目路径 （必须参数）
	-t 项目类型 （必须参数）
	-y或-n 是否出具无生信版报告(-y 是， -n 否，默认否)

项目类型参数(l：labfree; i：Itraq; t：TMT; d：DIA; pl：磷酸化Labfree; gl:泛素化labfree;
			nl:糖基化labfree; sl:琥珀酰化labfree; yl:络氨酸磷酸化labfree; al:乙酰化labfree; ml:丙二酰化labfree;
			pt：磷酸化TMT; at：乙酰化TMT)

注：可执行 python ProReport.py --help查看帮助

代码更新命令，直接执行：python ProReport.py 即可。

若 pip install -r requirements.txt 安装失败，可单独安装各模块，例如：
安装 win32com 命令
python -m pip install pypiwin32
