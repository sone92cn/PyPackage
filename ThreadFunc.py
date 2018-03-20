from ThreadTool import ReportMessage
from ExcelTool import ExcelReader, ExcelWriter

def LoadExcelFile(args):
    if len(args) == 4:
        fname, group, data_only, read_only = args
        try:
            if read_only:
                ExcelReader(fname, group, data_only)
            else:
                ExcelWriter(fname, group, data_only)
        except:
            return ReportMessage(0, 0, 'Fail to load: ' + fname)  #mid=0, code=0, msg=""
        else:
            return ReportMessage(0, 1, 'Succeed to load: ' + fname)  #mid=0, code=0, msg=""
    else:
        return ReportMessage(0, 0, 'Wrong input!')  #mid=0, code=0, msg=""