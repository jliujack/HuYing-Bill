from win32com.client import Dispatch
import sys
import os
import handleBill.assist as assist

def copyToTargetPath(sourcePath, targetPath):
  '''复制到指定路径,doc文件将转换为docx'''
  try:
    word = Dispatch('Word.Application')
    if (sourcePath[0] == '.'):
      sourcePath = os.path.abspath(sourcePath)
    if (targetPath[0] == '.'):
      targetPath = os.path.abspath(targetPath)
    doc = word.Documents.Open(sourcePath)
    doc.SaveAs(targetPath, FileFormat=12)
    doc.Close()
    word.Quit()
  except :
    raise assist.CustomError("请勿打开word！")
  
