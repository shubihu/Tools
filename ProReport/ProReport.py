import os
import sys
import click
from docx import Document

platform = sys.platform

def labfree(*args):
	input_path, types, document, paragraphs, tables, nobioinfo = args
	from templatePy.project import Labfree
	labfree = Labfree(input_path, types)
	if nobioinfo:
		labfree.nobioinfo(paragraphs)
	labfree.header(paragraphs)
	labfree.table_data(tables)
	labfree.text_png_data(document, paragraphs)
	labfree.save(document)
	if platform == 'win32':
		if nobioinfo:
			labfree.update()

def itrq(*args):
	input_path, types, document, paragraphs, tables, nobioinfo = args
	from templatePy.project import Itraq
	itraq = Itraq(input_path, types)
	if nobioinfo:
		itraq.nobioinfo(paragraphs)
	itraq.header(paragraphs, start_row=13)
	itraq.table_data(tables)
	itraq.text_png_data(document, paragraphs)
	itraq.save(document)
	if platform == 'win32':
		if nobioinfo:
			itraq.update()

def tmt(*args):
	input_path, types, document, paragraphs, tables, nobioinfo = args
	from templatePy.project import TMT
	tmt = TMT(input_path, types)
	if nobioinfo:
		tmt.nobioinfo(paragraphs)
	tmt.header(paragraphs, start_row=13)
	tmt.table_data(tables)
	tmt.text_png_data(document, paragraphs)
	tmt.save(document)
	if platform == 'win32':
		if nobioinfo:
			tmt.update()

def dia(*args):
	input_path, _, document, paragraphs, tables, _ = args
	from templatePy.dia import DIA
	dia = DIA(input_path)
	dia.header(paragraphs)
	dia.table_data(tables)
	dia.text_png_data(paragraphs)
	dia.save(document)

def pholabfree(*args):
	input_path, types, document, paragraphs, tables, _ = args
	from templatePy.project import PhoLabfree
	pholabfree = PhoLabfree(input_path, types)
	pholabfree.header(paragraphs)
	pholabfree.table_data(tables)
	pholabfree.text_png_data(document, paragraphs)
	pholabfree.save(document)


@click.command()
@click.option("--input_path", '-i', required=True, help="项目路径")
@click.option("--types", '-t', required=True, help="项目类型, l或L为labfree; i或I为Itraq; t或T为TMT; d为DIA")
@click.option('--nobioinfo/--no-nobioinfo', '-y/-n', default=False, help='是否出无生信版报告，默认为否')
def main(input_path, types, nobioinfo):
	if nobioinfo:
		print('无生信版报告')
	types = types.lower()
	types_dict = {'l': 'Labelfree报告模板-20200309.docx',
				  'i': 'iTRAQ报告模板-20200325-4标8标.docx',
				  't': 'TMT报告模板-20200325-6标10标16标.docx',
				  'd': 'DIA蛋白质组研究正式实验报告模板.docx',
				  'pl': '磷酸化Labelfree报告模板无激酶分析版.docx'}

	project_dict = {'l': labfree, 'i': itrq, 't': tmt, 'd': dia, 'pl': pholabfree}

	template_file = os.path.join(os.path.abspath(os.path.dirname(sys.argv[0])), 'template', types_dict.get(types))
	document = Document(template_file)
	paragraphs = document.paragraphs
	tables = document.tables

	project_dict.get(types)(input_path, types, document, paragraphs, tables, nobioinfo)

if __name__ == '__main__':
	main()
	

