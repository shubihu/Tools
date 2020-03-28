import os

from templatePy.template import Template

class Labfree(Template):
	
	def save(self, document):
		document.save('LabelFree相对定量蛋白质组学生物信息学分析报告.docx')
		print(f'报告生成完成，报告路径为：{os.path.join(os.getcwd(), "LabelFree相对定量蛋白质组学生物信息学分析报告.docx")}')


class Itraq(Template):
	
	def save(self, document):
		document.save('iTRAQ相对定量蛋白质组学报告.docx')
		print(f'报告生成完成，报告路径为：{os.path.join(os.getcwd(), "iTRAQ相对定量蛋白质组学报告.docx")}')

class TMT(Template):
	
	def save(self, document):
		document.save('TMT相对定量蛋白质组学报告.docx')
		print(f'报告生成完成，报告路径为：{os.path.join(os.getcwd(), "TMT相对定量蛋白质组学报告.docx")}')



