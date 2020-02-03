'''Author : Ryuodan'''

import win32com.client

def PPT_to_PDF(infile_path, outfile_path):
    '''convert from PPT to PDF file format'''
    powerpoint = win32com.client.Dispatch("Powerpoint.Application")
    pdf = powerpoint.Presentations.Open(infile_path, WithWindow=False)
    pdf.SaveAs(outfile_path, 32)
    pdf.Close()
    powerpoint.Quit()


if __name__ == "__main__":
    path_of_inputfile = 'test.ppt'
    path_of_outputfile = 'test.pdf'
    print(f'Converting {path_of_inputfile} to {path_of_outputfile}')
    PPT_to_PDF(infile_path=path_of_inputfile, outfile_path=path_of_outputfile)