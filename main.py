import os
from mutagen.mp3 import MP3

def main():
    url = input('请输入音频文件夹所在路径\n(不输入则默认为当前MP3文件夹下的数据)\n')
    addr = url if url else os.getcwd() + '\MP3\\'
    print('将执行获取音频文件信息来源：\n')
    print('>>> ' + addr + ' <<<\n')
    run = input('确认执行此路径扫描？（确认请输入：yes，任意键取消！）\n')
    if run != 'yes':
        print('退出程序\n')
        return
    else:
        print('执行扫描程序\n')
        get_file(addr)

def get_file(addr):
    list = []
    for root, dirs, files in os.walk(addr):
        for mp3 in files:
            if '.mp3' in mp3:
                _file_name = root + '/' + mp3
                _time = int(MP3(_file_name).info.length)
                m, s = divmod(_time, 60)
                h, m = divmod(m, 60)
                _shijian = '%02d:%02d:%02d' % (h, m, s)
                _root = root.replace(addr, '\\')
                _data = {
                    'url': _root,
                    'name': mp3,
                    'time': _shijian
                }
                list.append(_data)
    print('扫描结束...执行导出EXECL......\n')
    writeExecl(list)

def writeExecl(row):
    path = input('请输入Execl文件名（任意键默认名demo）\n')
    _path = path if path else 'demo'
    import xlsxwriter
    workbook = xlsxwriter.Workbook(_path + '.xlsx')
    wroksheet = workbook.add_worksheet('音频时间')
    for i, v in enumerate(row):
        wroksheet.write(i, 0, v['url'])
        wroksheet.write(i, 1, v['name'])
        wroksheet.write(i, 2, v['time'])
    workbook.close()
    print('程序执行结束，请查看本地xlsx文件\n')


if __name__ == '__main__':
    main()

# 生成指令
# pyinstaller -F main.py