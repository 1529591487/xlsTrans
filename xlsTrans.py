# -*- coding: utf-8 -*-
"""
@author: liuzhiwei

@Date:  2020/12/12
"""

import os
import logging

import win32com.client

logger = logging.getLogger('Sun')
logging.basicConfig(level=20,
                    # format="[%(name)s][%(levelname)s][%(asctime)s] %(message)s",
                    format="[%(levelname)s][%(asctime)s] %(message)s",
                    datefmt='%Y-%m-%d %H:%M:%S'  # 注意月份和天数不要搞乱了，这里的格式化符与time模块相同
                    )


def getFiles(dir, suffix, ifsubDir=True):  # 查找根目录，文件后缀
    res = []
    for root, directory, files in os.walk(dir):  # =>当前根,根下目录,目录下的文件
        for filename in files:
            name, suf = os.path.splitext(filename)  # =>文件名,文件后缀
            if suf.upper() == suffix.upper():
                res.append(os.path.join(root, filename))  # =>吧一串字符串组合成路径
        if False is ifsubDir:
            break
    return res


class formatTrans:
    def __init__(self, infoDict, filePath):
        self.infoDict = infoDict
        self.filePath = filePath
        self.fileOperator = None

        self.init_excelOperator()
        self.convert_files_in_folder(self.filePath)
        self.quit()
        os.system('pause')

    def quit(self):
        if None is not self.fileOperator:
            self.fileOperator.Quit()

    def init_excelOperator(self):
        try:
            self.fileOperator = win32com.client.DispatchEx("Excel.Application")
            self.fileOperator.Visible = False
        except Exception as e:
            logger.error(str(e))

    def format_trans(self, inputFileName):
        # https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype

        infoDict = self.infoDict
        formatType = infoDict['formatType']
        outputFileName = self.getNewFileName(infoDict['name'], inputFileName)

        if '' == outputFileName:
            return
        inputFileName = inputFileName.replace('/', '\\')
        outputFileName = outputFileName.replace('/', '\\')
        if '' == outputFileName:
            return
        if None is self.fileOperator:
            return
        logger.info('开始转换：[{0}]'.format(inputFileName))
        deck = self.fileOperator.Workbooks.Open(inputFileName)

        try:
            deck.SaveAs(outputFileName, formatType)
            logger.info('转换完成：[{0}]'.format(outputFileName))
        except Exception as e:
            logger.error(str(e))
        deck.Close()

    def convert_files_in_folder(self, filePath):
        if True is os.path.isdir(filePath):
            srcfiles = []
            dirPath = filePath
            files = os.listdir(dirPath)
            srcfilesWithoutDir = [f for f in files if f.endswith((".xls", ".xlsx"))]
            for srcfile in srcfilesWithoutDir:
                srcfiles.append('{0}\\{1}'.format(filePath,srcfile))
        elif True is os.path.isfile(filePath):
            srcfiles = [filePath]
        else:
            self.logError('不是文件夹，也不是文件')
            return

        for srcfile in srcfiles:
            fullpath = os.path.abspath(srcfile)
            self.format_trans(fullpath)

    def getNewFileName(self, newType, filePath):
        try:
            dirPath = os.path.dirname(filePath)
            baseName = os.path.basename(filePath)
            fileName = baseName.rsplit('.', 1)[0]
            suffix = baseName.rsplit('.', 1)[1]
            if newType == suffix:
                logger.warning('文档[{filePath}]类型和需要转换的类型[{newType}]相同'.format(filePath=filePath, newType=newType))
                return ''
            newFileName = '{dir}/{fileName}.{suffix}'.format(dir=dirPath, fileName=fileName, suffix=newType)
            if os.path.exists(newFileName):
                newFileName = '{dir}/{fileName}_new.{suffix}'.format(dir=dirPath, fileName=fileName, suffix=newType)
            return newFileName
        except Exception as e:
            logger.error(str(e))
            return ''


if __name__ == "__main__":
    # https://docs.microsoft.com/en-us/office/vba/api/excel.xlfileformat

    transDict = {}
    transDict.update({1: {'name': 'xlsx', 'formatType': 51}})
    transDict.update({2: {'name': 'xls', 'formatType': 56}})

    hintStr = ''
    for key in transDict:
        hintStr = '{src}{key}:->{type}\n'.format(src=hintStr, key=key, type=transDict[key]['name'])

    while True:
        print(hintStr)
        transFerType = int(input("转换类型:"))
        if None is transDict.get(transFerType):
            logger.error('未知类型')
        else:
            infoDict = transDict[transFerType]
            path = input("文件路径:")
            op = formatTrans(infoDict, path)
