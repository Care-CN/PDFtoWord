from pdf2docx import Converter
import PySimpleGUI as sg


def pdf2word(file_path, s, e):
    file_name = file_path.split('.')[0]
    docx_file = f'{file_name}.docx'
    if s == '':
        s = 0
    else:
        s = int(s) - 1
    if e == '':
        e = None
    else:
        e = int(e)
    p2w = Converter(file_path)
    p2w.convert(docx_file, start=s, end=e)
    p2w.close()
    return docx_file


def main():
    # sg.theme('BlueMono')
    sg.theme('DarkGreen')
    layout = [
        [sg.Text('PDF转Word', font=('微软雅黑', 12)),
         sg.Text('', key='filename', size=(70, 1), font=('微软雅黑', 10), text_color='blue')],
        [sg.Output(size=(80, 10), font=('微软雅黑', 10), key='output')],
        [sg.FileBrowse('选择文件', key='file', target='filename'),
         sg.Text('转换起始页:', font=('微软雅黑', 12)),
         sg.Input('', key='start', size=(3, 1)),
         sg.Text('转换停止页:', font=('微软雅黑', 12)),
         sg.Input('', key='end', size=(3, 1)),
         sg.Button('开始转换'), sg.Button('退出'), sg.Button('使用须知')]
    ]
    window = sg.Window('PDF转Word', layout, font=('微软雅黑', 15), default_element_size=(50, 1))
    while True:
        event, values = window.read()
        if event in (None, '退出'):
            break
        if event == '使用须知':
            window.FindElement('output').Update('')  # 清空输出框
            print('#########################')
            print('使用须知：')
            print('请先选择文件后再点击开始转换；')
            print('"转换起始页"与"转换停止页"为选填内容，默认转换全部页面；')
            print('页数输入错误可能会闪退，闪退后重新打开即可；')
            print('可能存在未知bug。')
            print('#########################')
        if event == '开始转换':
            if values['file'] and values['file'].split('.')[1] == 'pdf':
                file_path = pdf2word(values['file'], values['start'], values['end'])
                if file_path == 0:
                    print('!!!!!!!!!!!!!!!!!!!!!!!!!' + '\n' + 'pdf文件页数输入错误  T^T' + '\n' + '!!!!!!!!!!!!!!!!!!!!!!!!!')
                if values['start'] == '':
                    Ss = '首页'
                else:
                    Ss = '第' + values['start'] + '页'
                if values['end'] == '':
                    Se = '末页'
                else:
                    Se = '第' + values['end'] + '页'
                print('*************************' + '\n' + '转换成功   ' + Ss + '--' + Se + '   ^_^')
                print('word文件位置：', file_path + '\n' + '*************************')
            else:
                print(
                    '!!!!!!!!!!!!!!!!!!!!!!!!!' + '\n' + '未选取文件或文件非pdf格式  T^T\n请先选择正确的文件' + '\n' + '!!!!!!!!!!!!!!!!!!!!!!!!!')
    window.close()


if __name__ == '__main__':
    main()
