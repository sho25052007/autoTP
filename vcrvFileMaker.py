from pathlib import Path
import os
import shutil


def vcrvMaker():

    homeDir = 'C:\\Users\\sho25\\Desktop'
    baseDir = 'C:\\Users\\sho25\\Desktop\\extractedFile'
    baseDirList = os.listdir(baseDir)

    for file_count in baseDirList:
        targetDir = os.path.join(baseDir, file_count)
        for client in targetDir:
            clientName = os.listdir(targetDir)
            clientPath = os.path.join(targetDir, clientName[0])
            clientPathFormat = Path(clientPath)
            copyPath = Path(os.path.join(
                homeDir, f'vcrvFile\\{clientName[0]}'))
            fileCheck = copyPath.exists()
            if fileCheck:
                pass
                # print(f'{copyPath} already exists.')
            else:
                os.makedirs(copyPath)
            clientFileList = os.listdir(clientPath)
            for filename in clientFileList:
                completePath = os.path.join(clientPath, filename)
                # print(filename)
                # print(completePath)
                ext = os.path.splitext(filename)[-1].lower()
                if ext == '.crv':
                    shutil.move(
                        completePath, copyPath/filename)
                    break
                # else-continue


def vcrvMakerFull():

    homeDir = 'C:\\Users\\sho25\\Desktop'
    baseDir = 'C:\\Users\\sho25\\Desktop\\extractedFile'
    baseDirList = os.listdir(baseDir)

    for file_count in baseDirList:
        targetDir = os.path.join(baseDir, file_count)
        for client in targetDir:
            clientName = os.listdir(targetDir)
            clientPath = os.path.join(targetDir, clientName[0])
            clientPathFormat = Path(clientPath)
            copyPath = Path(os.path.join(
                homeDir, f'vcrvFile\\{clientName[0]}'))
            fileCheck = copyPath.exists()
            if fileCheck:
                pass
                # print(f'{copyPath} already exists.')
            else:
                os.makedirs(copyPath)
            clientFileList = os.listdir(clientPath)
            for filename in clientFileList:
                completePath = os.path.join(clientPath, filename)
                # print(filename)
                # print(completePath)
                ext = os.path.splitext(filename)[-1].lower()
                if ext == '.crv':
                    shutil.copy(
                        completePath, copyPath/filename)
                else:
                    continue
