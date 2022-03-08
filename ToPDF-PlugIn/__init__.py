"""
====================================
 :mod:`your.demo.helloworld`
====================================
.. moduleauthor:: Your Name <user@modify.me>
.. note::

Description
===========
Your Demo plugin module sample
"""
# Authors
# ===========
#
# * Your Name
#
# Change Log
# --------
#
#  * [2019/03/08]
#     - add icon
#  * [2018/10/28]
#     - starting

################################################################################
import sys
from alabs.common.util.vvargs import ModuleContext, func_log, \
    ArgsError, ArgsExit, get_icon_path
import os
import comtypes.client
from win32com import client
from img2pdf import convert


################################################################################
@func_log
def to_pdf(mcxt, argspec):
    mcxt.logger.info('>>>starting...')

    f_route = argspec.folder_route

    wdFormatPDF = 17

    word_to_pdf = []
    excel_to_pdf = []
    img_to_pdf = []

    for file in os.listdir(f_route):
        if file.endswith(".docx") or file.endswith(".doc"):
            word_to_pdf.append(file)
        elif file.endswith(".xlsx") or file.endswith(".xls"):
            excel_to_pdf.append(file)
        elif file.endswith(".jpg") or file.endswith(".JPG"):
            file = f_route + "\\" + file
            img_to_pdf.append(file)

    if argspec.wordfile is True:
        try:
            for i in word_to_pdf:
                if '~$' not in i:
                    out_file_temp = f_route + "\\" + i.split('.')[0]
                    i = f_route + "\\" + i
                    word = comtypes.client.CreateObject('Word.Application')
                    doc = word.Documents.Open(i)
                    doc.SaveAs(out_file_temp, FileFormat=wdFormatPDF)
                    doc.Close()
                    word.Quit()
        except:
            pass

    if argspec.excelfile is True:
        try:
            for i in excel_to_pdf:
                if '~$' not in i:
                    out_file_temp = f_route + "\\" + i.split('.')[0]
                    i = f_route + "\\" + i
                    xlApp = client.Dispatch("Excel.Application")
                    excel = xlApp.Workbooks.Open(i)
                    ws = excel.Worksheets[0]
                    ws.Visible = 1
                    ws.ExportAsFixedFormat(0, out_file_temp)
                    excel.Close()
        except:
            pass

    if argspec.jpgimage is True:
        try:
            if argspec.pdf_name == '' or argspec.pdf_name is None:
                write_folder = f_route + "\\" + 'out.pdf'
            else:
                write_folder = f_route + "\\" + argspec.pdf_name + '.pdf'

            with open(write_folder, "wb") as f:
                pdf = convert(img_to_pdf)
                f.write(pdf)
        except:
            pass


    mcxt.logger.info('>>>end...')
    return 0


################################################################################
def _main(*args):
    """
    Build user argument and options and call plugin job function
    :param args: user arguments
    :return: return value from plugin job function
    """
    with ModuleContext(
            owner='jeon',
            group='demo',
            version='1.3',
            platform=['windows', 'darwin', 'linux'],
            output_type='text',
            display_name='To PDF Plugin',
            icon_path=get_icon_path(__file__),
            description='PDF Converter',
    ) as mcxt:
        # ##################################### for app dependent parameters
        mcxt.add_argument('folder_route',
                          display_name='Folder Route',
                          input_method='folderread',
                          help='credentials.json location')
        mcxt.add_argument('pdf_name',
                          display_name='JPG to PDF name',
                          help='If this field is blank, the file name is output')
        mcxt.add_argument('--wordfile',
                          display_name='Word File',
                          action='store_true',
                          help='credentials.json location')
        mcxt.add_argument('--excelfile',
                          display_name='Excel File',
                          action='store_true',
                          help='credentials.json location')
        mcxt.add_argument('--jpgimage',
                          display_name='Jpg File',
                          action='store_true',
                          help='credentials.json location')
        argspec = mcxt.parse_args(args)
        return to_pdf(mcxt, argspec)


################################################################################
def main(*args):
    try:
        return _main(*args)
    except ArgsError as err:
        sys.stderr.write('Error: %s\nPlease -h to print help\n' % str(err))
    except ArgsExit as _:
        pass
