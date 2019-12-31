import os
import sys
import comtypes.client
from PyPDF2 import PdfFileReader, PdfFileWriter
from os.path import isfile, join, abspath, splitext
from datetime import datetime

# https://docs.microsoft.com/en-us/office/vba/api/word.wdsaveformat
wdFormatPDF = 17
# https://docs.microsoft.com/en-us/deployoffice/compat/office-file-format-reference
word_support = [
    '.doc', '.docm', '.docx', '.dot', '.dotm', 
    '.dotx', '.htm', '.html', '.mht', '.mhtml', 
    '.odt','.rtf', '.txt', '.wps'
]
powerpoint_support = [
    '.bmp', '.emf', '.gif', '.jpg', '.mp4', '.odp', 
    '.png', '.pot', '.potm', '.potx', '.ppa', '.ppam',
    '.pps', '.ppsm', '.ppsx', '.ppt', '.pptm', 
    '.pptx', '.thmx', '.tif', '.wmf', '.wmv'
]
excel_support = [
    '.csv', '.dbf', '.dif', '.prn', '.slk', 
    '.xla', '.xlam', '.xls', '.xlsb', '.xlsm', 
    '.xlsx', '.xlt', '.xltm', '.xltx', '.xlw', 
    '.xml', '.xps', '.ods'
]


def merge_pdfs(paths: list, output: str):
    '''
    credit:
    http://www.blog.pythonlibrary.org/2018/04/11/splitting-and-merging-pdfs-with-python/
    '''
    pdf_writer = PdfFileWriter()
    for index, path in enumerate(paths):
        print('{i}. {f}'.format(i=index+1, f=path))
        pdf_reader = PdfFileReader(path)
        for page in range(pdf_reader.getNumPages()):
            # Add each page to the writer object
            pdf_writer.addPage(pdf_reader.getPage(page))
    # Write out the merged PDF
    with open(output, 'wb') as out:
        pdf_writer.write(out)
    out.close()
    print('Finished!')
    pass


def contains_non_pdf(paths: list) -> bool:
    return not all([p.endswith('.pdf') for p in paths])


def get_ext(fname: str) -> str:
    return splitext(fname)[1]


def args_to_paths(args: list, non_pdf_flag: bool) -> list:
    if non_pdf_flag:
        # launch microsoft word/powerpoint/excel applications
        word = comtypes.client.CreateObject('Word.Application')
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        excel = comtypes.client.CreateObject('Excel.Application')
    ret_paths = []
    for arg in args:
        arg = abspath(arg)
        if arg.endswith('.pdf') is False:
            print('saving file [{ef}] as pdf..'.format(ef=arg), end='')
            fext = get_ext(arg)
            try:
                if fext in word_support:
                    # try open with Microsoft Word
                    word.Visible = True
                    doc = word.Documents.Open(arg)
                    doc.SaveAs(arg+'.pdf', FileFormat=wdFormatPDF)
                    doc.Close()
                    print('Done')
                elif fext in powerpoint_support:
                    # try open with Microsoft Powerpoint
                    powerpoint.Visible = True
                    doc = powerpoint.Presentations.Open(arg)
                    doc.SaveAs(arg+'.pdf', 32)
                    doc.Close()
                    print('Done')
                elif fext in excel_support:
                    # try open with Microsoft Excel
                    excel.Visible = True
                    doc = excel.Workbooks.Open(arg)
                    doc.ExportAsFixedFormat(0, arg+'.pdf', 1, 0)
                    doc.Close()
                    print('Done')
                else:
                    sys.exit('file [{ef}] is not supported.'.format(ef=arg))
            except Exception as e:
                print(e)
                sys.exit('cannot save file [{ef}] as pdf.'.format(ef=arg))
            arg = arg+'.pdf'
        if isfile(arg):
            ret_paths.append(arg)
        else:
            sys.exit('file [{ef}] not found.'.format(ef=arg))
    if non_pdf_flag:       
        # quit applications
        word.Quit()
        powerpoint.Quit()
        excel.Quit()
    return ret_paths


def main():
    args = []
    if len(sys.argv) == 1:
        while True:
            arg = input('add(or drag) file, then hit <Enter> ([1]type "dd" to delete last file [2] leave empty to finish): ')
            if arg == 'dd':
                args.pop()
            elif arg:
                # default local volume label
                volume_label = 'C:'
                split_args = arg.split(':')
                for index, a in enumerate(split_args):
                    if index == 0:
                        volume_label = a[-1]+':'
                        continue
                    elif index != len(split_args)-1:
                        a = a[:-1]
                    a = a.strip().strip('\"')
                    if a:
                        a = join(volume_label, a)
                        args.append(os.path.normpath(a))
            else:
                break
            print()
            for index, arg in enumerate(args):
                print('{i}. {f}'.format(i=index+1, f=arg))
            print()
        # sys.exit('Usage: python3 merge_pdf.py <file_path> [<file_path> ...]')
    else:
        args = [os.path.normpath(p) for p in sys.argv[1:]]
    if len(args) < 2:
        sys.exit('merge-pdf requires at least 2 files.')
    non_pdf_flag = contains_non_pdf(args)
    pdf_paths = args_to_paths(args=args, non_pdf_flag=non_pdf_flag)
    now_time = datetime.now().strftime('%Y%m%d%H%M%S')
    outfile = 'merged-' + now_time + '.pdf'
    merge_pdfs(pdf_paths, outfile)


if __name__ == '__main__':
    main()
