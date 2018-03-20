#version=2016.11.30.001
import os, re

def removeEmptyDir(path):
    def isEmptyDir(dpath):
        empty = True
        for file in os.listdir(dpath):
            if file != "." and file != "..":
                fpath = os.path.join(dpath, file)
                if os.path.isdir(fpath):
                    if isEmptyDir(fpath):
                        os.rmdir(fpath)
                    else:
                        empty = False
                else:
                    if os.path.isfile(fpath):
                        empty = False
        return empty
    
    if isEmptyDir(path):
        os.rmdir(path)
def createDirAsPath(path):
    if os.path.isdir(path):
        return True
    paths = os.path.split(path)
    if len(paths[0]):
        if os.path.isdir(paths[0]):
            if len(paths[1]):
                os.mkdir(path)
                return True
        else:
            if createDirAsPath(paths[0]):
                os.mkdir(path)
                return True
    return False
def createDirAsTree(parent, tree):
    if not os.path.isdir(parent):
        if not createDirAsPath(parent):
            return False
    result = True
    for branch in tree:
        npath = os.path.join(parent, branch[1])
        if not os.path.isdir(npath):
            os.mkdir(npath)
        if len(branch[2]):
            result = (result and createDirAsTree(os.path.join(parent, branch[1]), branch[2]))
    return result
def createTreeAsPath(path, fileRegular='', scanSubFolder=True, treeMode=False, relativePath=False, forFile=True, maxloops=100):
    def createTree(root, path, filePattern, scanSubFolder, treeMode, pathStartIndex, forFile, maxloops, curloop):
        branch, flist = [], []
        if forFile:
            for file in os.listdir(path):
                if file[0] != "." and file[0] != "$" and file != "System Volume Information":
                    filepath = os.path.join(path, file)
                    if os.path.isfile(filepath):
                        if ((filePattern == None) or (filePattern.match(file) != None)):
                            if treeMode:
                                flist.append([1, file, []])
                            else:
                                root.append(filepath[pathStartIndex:])
                    elif scanSubFolder == True and maxloops > curloop:
                        if treeMode:
                            subbranch = createTree(branch, filepath, filePattern, scanSubFolder, treeMode, pathStartIndex, forFile, maxloops, curloop+1)
                            if len(subbranch):
                                branch.append([0, file, subbranch])
                        else:
                            createTree(root, filepath, filePattern, scanSubFolder, treeMode, pathStartIndex, forFile, maxloops, curloop+1)  
        else:
            for file in os.listdir(path):
                if file[0] != "." and file[0] != "$" and file != "System Volume Information":
                    filepath = os.path.join(path, file)
                    if os.path.isdir(filepath):
                        if scanSubFolder == True and maxloops > curloop:
                            if treeMode:
                                subbranch = createTree(branch, filepath, filePattern, scanSubFolder, treeMode, pathStartIndex, forFile, maxloops, curloop+1)
                                if len(subbranch):
                                    branch.append([0, file, subbranch])
                                else:
                                    flist.append([0, file, []])
                            else:
                                rlen = len(root)
                                createTree(root, filepath, filePattern, scanSubFolder, treeMode, pathStartIndex, forFile, maxloops, curloop+1)
                                if len(root) == rlen:
                                    root.append(filepath[pathStartIndex:])
                        else:
                            if treeMode:
                                flist.append([0, file, []])
                            else:
                                root.append(filepath[pathStartIndex:])
        if treeMode:
            if len(flist):
                branch.extend(flist)
            return branch
    
    if os.path.isdir(path):
        root = []
        if relativePath:
            pathStartIndex = len(path) + 1
        else:
            pathStartIndex = 0
        if fileRegular == '':
            filePattern = None
        else:
            filePattern = re.compile(fileRegular, re.IGNORECASE)
        if treeMode:
            return createTree(root, path, filePattern, scanSubFolder, treeMode, pathStartIndex, forFile, maxloops, 1)
        else:
            createTree(root, path, filePattern, scanSubFolder, treeMode, pathStartIndex, forFile, maxloops, 1)
            return root
    else:
        return []
def getListFromTree(tree, path=''):
    def jionPath(path, file):
        if len(path):
            return os.path.join(path, file)
        else:
            return file
    
    fileList = []
    for branch in tree:
        if len(branch[2]):
            fileList.extend(getListFromTree(branch[2], jionPath(path, branch[1])))
        else:
            fileList.append(jionPath(path, branch[1]))
    return fileList
def getExtFromPath(path):
    """get extention from path"""
    chk = max(path.rfind("\\"), path.rfind('/'))
    inx = path.rfind('.')
    if inx > -1:
        if inx > chk:
            return path[inx:]
    return ""