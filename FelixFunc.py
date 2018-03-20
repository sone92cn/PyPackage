#encoding='utf-8'
from datetime import date, time, datetime

def myprint(obj, title=None):
    if title != None:
        print(title)
    if type(obj) == type(dict()):
        for o in obj:
            print(o, ':', obj[o])
    elif type(obj) in (type(tuple()), type(list())):
        for o in obj:
            print(o)
    else:
        print(obj)
def getStrFromInt(num, width=0):
    temp = str(num)
    if len(temp) < width:
        temp = ('0'*(width-len(temp))) + temp
    return temp  
def isNumber(text):
    if text[0] in ('+', '-'):
        text = text[1:]
    if text.isdigit():
        return True
    elif '.' in text:
        if (text.find('.') == text.rfind('.')):
            if text.replace('.', '').isdigit():
                return True
    return False
def wipeNone(lst):
    while None in lst:
        lst.remove(None)
    return lst
def getPosList(full_list, part_list, pos_list=[], all_included=False):
    for lst in part_list:
        if lst in full_list:
            pos_list.append(full_list.index(lst))
        elif all_included:
            return False
        else:
            pos_list.append(-1)
    return True
def getSubListByPos(data_list, pos_list, ignore_error=True):
    part_list = []
    inx = len(data_list)
    for i in pos_list:
        if i > -1 and i < inx:
            part_list.append(data_list[i])
        else:
            if ignore_error:
                part_list.append(None)
            else:
                raise Exception('Index out of length error.')
    return part_list
def combineLists(*lists):
    full_list = []
    for lst in lists:
        full_list.extend(lst)
    return full_list
def cleanedFileName(text):
    while '\\' in text:
        text = text.replace('\\', '_a_')
    while '/' in text:
        text = text.replace('/', '_b_')
    while ':' in  text:
        text = text.replace(':', '_c_')
    while '*' in  text:
        text = text.replace('*', '_d_')
    while '?' in  text:
        text = text.replace('?', '_e_')
    while '"' in  text:
        text = text.replace('"', '_f_')
    while '<' in  text:
        text = text.replace('<', '_g_')
    while '>' in  text:
        text = text.replace('>', '_h_')
    while '|' in  text:
        text = text.replace('|', '_i_')
    return text
def cleanedSheetName(text):
    while '\\' in text:
        text = text.replace('\\', '_a')
    while '/' in text:
        text = text.replace('/', '_b')
    while ':' in  text:
        text = text.replace(':', '_c')
    while '*' in  text:
        text = text.replace('*', '_d')
    while '?' in  text:
        text = text.replace('?', '_e')
    while '[' in  text:
        text = text.replace('<', '_g')
    while ']' in  text:
        text = text.replace('>', '_h')
    return text
def createStringFromVar(invar, ignore_error=True):
    if type(invar) == type(None):
        return ''
    elif type(invar) == type(date(1990,1,1)):
        return invar.strftime('%Y-%m-%d')
    elif type(invar) == type(time()):
        return invar.strftime('%H:%M:%S')
    elif type(invar) == type(datetime(1990, 1, 1)):
        return invar.strftime('%Y-%m-%d %H:%M:%S')
    elif type(invar) in  (type(1), type(1.1), type('s')):
        return str(invar)
    else:
        if ignore_error:
            return ''
        else:
            raise Exception('Unknown type error.')
def createStringFromVarList(var_list, sep=',', add_quotes=True, all_included=False, skip_none=False, ignore_error=True):
    type_list = (type(None), type(1), type(1.1), type('s'), type(date(1990,1,1)), type(time()), type(datetime(1990, 1, 1)))
    full_str = ''
    for var in var_list:
        if skip_none:
            if var == None:
                continue
        if full_str != '':
            full_str = full_str + sep
        if type(var) in type_list:
            if add_quotes:
                full_str = full_str + '"' + createStringFromVar(var) + '"'
            else:
                full_str = full_str + createStringFromVar(var)
        elif all_included:
            if ignore_error:
                if add_quotes:
                    full_str = full_str + '""'
            else:
                raise Exception("Not all translated to string error")
        else:
            if add_quotes:
                full_str = full_str + '""'
    return full_str
def resizeRect(rect, screen, getmax=False, rounddown=False):
    r_width, r_height = rect
    s_width, s_height = screen
    if r_width/r_height > s_width/ s_height:
        if getmax:
            n_width = s_width
        else:
            n_width = min(r_width, s_width)
        n_height = r_height / r_width * n_width
    else:
        if getmax:
            n_height = s_height
        else:
            n_height = min(r_height, s_height)
        n_width = r_width / r_height * n_height
    if rounddown:
        return (int(n_width), int(n_height))
    else:
        return (n_width, n_height)