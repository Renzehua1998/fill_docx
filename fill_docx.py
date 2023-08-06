import pandas as pd
from docx import Document  # pip install python-docx
import comtypes.client
import fitz  # pip install pymupdf
import PySimpleGUI as sg
import os

# word转pdf
def word2pdf(docx_name):
    file_path = os.getcwd() + '/' + docx_name
    word = comtypes.client.CreateObject('Word.Application')
    word.Visible = 0
    file_name = file_path.split('.')[0]
    pdf_file = f'{file_name}.pdf'
    w2p = word.Documents.Open(file_path)
    w2p.SaveAs(pdf_file, FileFormat=17)
    w2p.Close()
    return pdf_file

# pdf转图片
def pdf2image(pdf_path, dpi):
    file_name = pdf_path.split('.')[0]
    img_path = f'{file_name}' +'{}.png'
    pdf = fitz.open(pdf_path)
    for page_num in range(len(pdf)):
        page = pdf[page_num]
        pix = page.get_pixmap(matrix=fitz.Matrix(dpi/72, dpi/72))
        pix.save(img_path.format(page_num), "png")

def transform(template_path, data_path, progress_bar, dpi):
    # 读取Excel文件
    data = pd.read_excel(data_path)
    column_names = data.columns

    size = len(data)
    # 遍历Excel数据
    for index, row in data.iterrows():
        name = row[column_names[0]]

        # 读取Word模板
        template = Document(template_path)

        children = template.element.body.iter()

        for child in children:
            # 通过类型判断目录
            if child.tag.endswith('txbx'):
                for ci in child.iter():
                    if ci.tag.endswith('main}r'):
                        for col in column_names:
                            if f'[{col}]' in ci.text:
                                ci.text = ci.text.replace(f'[{col}]', str(row[col]))

        # 新建同名文件夹
        os.makedirs(name)
        print(f'已创建 {name}/ 文件夹')

        # 保存填充后的Word文档
        template.save(f'{name}/{name}.docx')
        print(f'已生成 {name}.docx')

        # 转换为PDF
        word2pdf(f'{name}/{name}.docx')
        print(f'已生成 {name}.pdf')

        # 转换为图片
        pdf2image(f'{name}/{name}.pdf', dpi)
        print(f'已生成 {name}.png')

        progress_bar.update(int((index+1)/size*100))

def main():
    # 选择主题
    sg.theme('LightBlue1')

    layout = [
        [sg.Text('文档批量生成工具', font=('微软雅黑', 18)),], \
        [sg.Text('模板路径：', font=('微软雅黑', 10), text_color='blue'), sg.Text('', key='temp_name', size=(30, 1), font=('微软雅黑', 10), text_color='blue')], \
        [sg.Text('数据路径：', font=('微软雅黑', 10), text_color='blue'), sg.Text('', key='data_name', size=(30, 1), font=('微软雅黑', 10), text_color='blue')], \
        [sg.Output(size=(70, 10), font=('微软雅黑', 10))], \
        [sg.FilesBrowse('选择模板', key='template', target='temp_name', file_types=(("Word文档", "*.docx"),)), \
         sg.FilesBrowse('填充数据', key='data', target='data_name', file_types=(("Excel表格", "*.xlsx"),)), \
         sg.Text('dpi分辨率：', font=('微软雅黑', 18)), \
         sg.Combo([300, 600, 1200], default_value=300, key='combo'), \
         sg.Button('转换'), \
         sg.Button('退出'),], \
        [sg.ProgressBar(100, orientation='h', size=(70, 20), key='progressbar')],]
    # 创建窗口
    window = sg.Window("Renz的小工具系列", layout, font=("微软雅黑", 15), default_element_size=(50, 1))
    # 进度条
    progress_bar = window['progressbar']
    # 事件循环
    while True:
        # 窗口的读取，有两个返回值（1.事件；2.值）
        event, values = window.read()
        print(event, values)

        if event == '转换':
            if values['template']:
                if values['data']:
                    transform(values['template'], values['data'], progress_bar, values['combo'])
                    print('------------------------------')
                    print('输出完成！')
                else:
                    print('请输入填充数据！')
            else:
                print('请输入模板文件！')

        if event in (None, '退出'):
            break

    window.close()

if __name__ == '__main__':
    main()
    # transform('成绩证书模板fin.docx', '成绩单.xlsx')