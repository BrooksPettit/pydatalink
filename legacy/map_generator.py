#Brooks Pettit
#brooks.pettit@live.com
#(832)-580-7346
#2020-04-07

import os
import pandas as pd
import numpy as np

import click
from tqdm import tqdm
import pyfiglet

#define PIDataLink description and unit of measure formulas

def desc_form_str(tag):
    return '=PITagAtt("%s","description","")'%tag

def uom_form_str(tag):
    return '=PITagAtt("%s","uom","")'%tag

def list_to_str(array):
    s=""
    for i in range(0,len(array)):
        if i != range(0,len(array))[-1]:
            s+=array[i]+', '
        else:
            s+=array[i]
    return s

@click.command()
@click.argument('filesystem',type=click.Path(exists=True))
def main(filesystem):
    dirstring=click.format_filename(filesystem)
    workingdir=os.getcwd()
    os.chdir(dirstring)

    full_paths=[]
    filenames=[]
    totalfiles=0
    for root, dirs, files in os.walk(".", topdown=False):
        for name in files:
            totalfiles+=1
            if name[-9:] == '_tags.csv':
                full_paths.append(os.path.join(os.getcwd(),root[2:],name))
                filenames.append(name[:-9])

    master_tag_list=[]
    tag_dict={}
    
    for name in filenames:
        tag_dict[name]=[]
    error_paths=[]
    empty_paths=[]
    duplicates=0

    
    with tqdm(total=len(full_paths),desc="Read tag files",position=1) as pbar:
        i=0
        for path in full_paths:
            df=pd.DataFrame()
            try:
                df=pd.read_csv(path,header=None,skip_blank_lines=True)
            except pd.errors.ParserError:
                error_paths.append(path)
                i+=1
                continue
            except pd.errors.EmptyDataError:
                error_paths.append(path)
                i+=1
                continue
            for item in df.iloc[:,0]:
                tag_dict[filenames[i]].append(item)
                if item not in master_tag_list:
                    master_tag_list.append(item)
                else:
                    duplicates+=1
            pbar.update(1)
            i+=1
    #print(tag_dict)
    #print('/n/n/n/n')

    os.chdir(workingdir)
    file_map={}
    for tag in master_tag_list:
        file_map[tag]=[]
        for file in tag_dict:
            if tag in tag_dict[file]:
                file_map[tag].append(file)

    #print(file_map)

    final_dict={key:list_to_str(file_map[key]) for key in file_map}
    
    
    t=[]
    s=[]
    for a,b in final_dict.items():
        t.append(a)
        s.append(b)

    df_final=pd.DataFrame({'Tag':t,'Files':s})
    print(df_final)
    df_final.to_excel('output10.xlsx')
                 
if __name__ == '__main__':
    main()
