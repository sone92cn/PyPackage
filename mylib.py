import pickle
import pandas as pd

def loadVar(fname):
    file = open(fname, "rb")
    value = pickle.load(file)
    file.close()
    return value

def saveVar(fname, value):
    file = open(fname, "wb")
    pickle.dump(value, file, 2)
    file.close()
    return True

def getUniqueDataFrame(df, col0, col1):
    duplicated_lt = df[col0].duplicated(keep=False) | df[col1].duplicated(keep=False)
    duplicated_df = df[duplicated_lt]

    lst0, lst1 = [], []
    for inx, row in duplicated_df.iterrows():
        if not(row[0] in lst0 or row[1] in lst1):
            lst0.append(row[0])
            lst1.append(row[1])

    df = df[~duplicated_lt]
    if len(lst0):
        df = pd.concat([df, pd.DataFrame(data={col0: lst0, col1: lst1})], ignore_index=True)
    return df

