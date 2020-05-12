import os
import sys
if sys.platform == 'win32':
	import win32com.client
from docx import Document
from templatePy.template import Template

class Labfree(Template):
	def __init__(self, path, types):
		super(Labfree, self).__init__(path, types)

	def nobioinfo(self, paragraphs):
		self.delete_paragraph(paragraphs, list(range(101, 185)))

	def save(self, document):
		document.save('LabelFree相对定量蛋白质组学生物信息学分析报告.docx')
		print(f'报告生成完成，报告路径为：{os.path.join(os.getcwd(), "LabelFree相对定量蛋白质组学生物信息学分析报告.docx")}')

	def update(self):
		word = win32com.client.DispatchEx("Word.Application")
		doc = word.Documents.Open(os.path.join(os.getcwd(), "LabelFree相对定量蛋白质组学生物信息学分析报告.docx"))
		doc.TablesOfContents(1).Update()
		doc.Close(SaveChanges=True)
		word.Quit()

class Itraq(Template):
	def __init__(self, path, types):
		super(Itraq, self).__init__(path, types)

	def nobioinfo(self, paragraphs):
		self.delete_paragraph(paragraphs, list(range(90, 175)))

	def save(self, document):
		document.save('iTRAQ相对定量蛋白质组学报告.docx')
		print(f'报告生成完成，报告路径为：{os.path.join(os.getcwd(), "iTRAQ相对定量蛋白质组学报告.docx")}')

	def update(self):
		word = win32com.client.DispatchEx("Word.Application")
		doc = word.Documents.Open(os.path.join(os.getcwd(), "iTRAQ相对定量蛋白质组学报告.docx"))
		doc.TablesOfContents(1).Update()
		doc.Close(SaveChanges=True)
		word.Quit()

class TMT(Template):
	def __init__(self, path, types):
		super(TMT, self).__init__(path, types)

	def nobioinfo(self, paragraphs):
		self.delete_paragraph(paragraphs, list(range(90, 175)))

	def save(self, document):
		document.save('TMT相对定量蛋白质组学报告.docx')
		print(f'报告生成完成，报告路径为：{os.path.join(os.getcwd(), "TMT相对定量蛋白质组学报告.docx")}')

	def update(self):
		word = win32com.client.DispatchEx("Word.Application")
		doc = word.Documents.Open(os.path.join(os.getcwd(), "TMT相对定量蛋白质组学报告.docx"))
		doc.TablesOfContents(1).Update()
		doc.Close(SaveChanges=True)
		word.Quit()

class PhoLabfree(Template):
	def __init__(self, path, types):
		super(PhoLabfree, self).__init__(path, types)
	
	def save(self, document):
		document.save('磷酸化Labfree生物信息学分析报告.docx')
		print(f'报告生成完成，报告路径为：{os.path.join(os.getcwd(), "磷酸化Labfree生物信息学分析报告.docx")}')
