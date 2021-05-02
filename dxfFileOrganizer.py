from zipfile import ZipFile
from pathlib import Path
import os
import shutil


def dxfExtractor():

    baseDir = 'C:\\Users\\sho25\\Desktop'
    fileList = os.listdir(baseDir)

    file_num = 1

    for filename in fileList:
        filepath = Path(os.path.join(baseDir, filename))
        ext = os.path.splitext(filename)[-1].lower()
        if ext == '.zip':
            with ZipFile(filepath, 'r') as zip:

                extractFileLocation = Path(os.path.join(
                    baseDir, f'extractedFile\\{file_num}'))

                file_num += 1
                if extractFileLocation.exists():
                    zip.extractall(extractFileLocation)
                else:
                    os.makedirs(extractFileLocation)
                    zip.extractall(extractFileLocation)

        # print(file_num)
        # print(filename)
