import os
import sys
import click
from docx import Document

def lipids(*args):
	input_path, _, document, paragraphs, tables = args
	from templatePy.lipids import Lipids
	lipids = Lipids(input_path)
	lipids.header(paragraphs)
	lipids.table_data(tables)
	lipids.text_png_data(paragraphs)
	lipids.save(document)

def fattyacid(*args):
	input_path, types, document, paragraphs, tables = args
	from templatePy.fattyacid import Fattyacid
	fa = Fattyacid(input_path, types)
	fa.header(paragraphs)
	fa.table_data(tables)
	fa.text_png_data(paragraphs)
	fa.save(document)

@click.command()
@click.option("--input_path", '-i', required=True, help="项目路径")
@click.option("--types", '-t', required=True, help="项目类型, ld为lipids, lc为长链脂肪酸，sc为短链脂肪酸")
def main(input_path, types):
	types = types.lower()
	types_dict = {'ld': '高分辨广谱脂质组绝对定量报告模板-FINAL.docx',
				  'lc': '中长链脂肪酸报告模板.docx',
				  'sc': '短链脂肪酸报告模板.docx'}

	project_dict = {'ld': lipids, 'lc': fattyacid, 'sc': fattyacid}

	template_file = os.path.join(os.path.abspath(os.path.dirname(sys.argv[0])), 'template', types_dict.get(types))
	document = Document(template_file)
	paragraphs = document.paragraphs
	tables = document.tables

	project_dict.get(types)(input_path, types, document, paragraphs, tables)

if __name__ == '__main__':
	main()
	

