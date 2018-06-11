import os, re, time, string

LI = ['A', 'B', 'C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S']
INDEX = ['料號', '品名'] 
HEAD = ['不良總數', '變形', '尺寸不良', '髒汙', '烤漆不良', 
        '毛邊', '刮傷', '印刷不良', '缺件', '螺孔不良', '氣密不良', '功能不良',
        '破裂', '滑牙', '其他', '其他(備註)', '原材不良']

import pandas as pd
from numpy import NaN

'''
- create a list of excel files
    - get_files, inputs nothing, outputs a list of names of files
- process excel data to make it look pretty and store in its respective data frame, rename the columns too
- concat all data frame into one
- group them up using material name and number, while preserving my 'other' section
- rename the columns to HEAD
- export the big data frame into excel
'''

def main():
    os.chdir(r'/Users/joseph/Desktop/QC_log')

    #store all of the files in a list, waiting to process
    outter_files = get_outter_files()
    inner_files = get_inner_files()

    #create a list where there are different names to store out files
    df_l = []
    #put parsed outter data in df list
    for file in outter_files:
        df_l.append(parse_outter(file))
    #put parsed inner data in df list
    for file in inner_files:
        df_l.append(parse_inner(file))
    #concatanate all of the df inside the list
    df = concat(df_l)

    #start grouping them! While perserving 'other'
    #making sure that material numbers and 'others' are all non nan values
    df.iloc[:, [0,17]] = df.iloc[:,[0,17]].fillna('')
    #group all of the data using A and B
    j = df.groupby(['A','B']).sum()

    # order the number to see large to small
    j = j.sort_values(by= 'C', ascending=False)

    #change the head of the excel
    j.index.names = INDEX
    j.columns = HEAD
  
    # #color parts of the value red
    # red_df = j.iloc[0:,0:15]
    # product = red_df.style.applymap(color_red)
    # writer = pd.ExcelWriter('kf_test.xlsx')
    # product.to_excel(writer)
    
    #write into excel at somewhere else
    os.chdir(r'/Users/joseph/Desktop/QC_log/test')
    j.to_excel('kf_test.xlsx', 'Sheet1')

def get_outter_files():
    files = []
    for f in os.listdir():
        if (f.endswith('.xlsx') | f.endswith('.xls') | f.endswith('.xml')) and f.startswith ('20') and not f.startswith('~$'):
            files.append(f)
    return(files)

def get_inner_files():
    files = []
    for f in os.listdir():
        if (f.endswith('.xlsx') | f.endswith('.xls') | f.endswith('.xml')) and not f.startswith('20') and not f.startswith('~$'):
            files.append(f)
    return(files)

def parse_outter(file):
    df = pd.read_excel(file)
    #parse out the data parts and rename the head
    df = df.iloc[2:,5:24]
    df.columns = LI
    return(df)

def parse_inner(file):
    df = pd.read_excel(file)
    df.reset_index(level=[0,1,2], inplace=True)
    df = df.iloc[3:,0:19]
    df.columns = LI
    return(df)

def concat(df_list):
    df = df_list[0]
    for l in df_list[1:]:
        df = pd.concat([df,l])
    return(df)

# def color_red(v):
#     color = 'red' if v > 4 else 'black'
#     return 'color: %s' % color

if __name__ == '__main__':
    main()