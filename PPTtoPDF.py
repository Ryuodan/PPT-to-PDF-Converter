import win32com.client
import os
in_file=input("Enter the path of the file")
out_file=os.path.splitext(in_file)[0]     # desktop\file.pptx 
powerpoint=win32com.client.Dispatch("Powerpoint.Application")
pdf=powerpoint.Presentations.Open(in_file,WithWindow=False)
pdf.SaveAs(out_file,32)
pdf.Close()
powerpoint.Quit()