#Brooks Pettit
#brooks.pettit@live.com
#(832)-580-7346
#2020-04-07

import os
import pandas as pd
import numpy as np
import xlUpdate

import click
from tqdm import tqdm
import pyfiglet

#define PIDataLink description and unit of measure formulas

def desc_form_str(tag):
    return '=PITagAtt("%s","description","")'%tag

def uom_form_str(tag):
    return '=PITagAtt("%s","uom","")'%tag

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
                filenames.append(name)

    master_tag_list=[]
    error_paths=[]
    empty_paths=[]
    duplicates=0

    with tqdm(total=len(full_paths),desc="Read tag files",position=1) as pbar:
        for path in full_paths:
            try:
                df=pd.read_csv(path,header=None,skip_blank_lines=True)
            except pd.errors.ParserError:
                error_paths.append(path)
                pass
            except pd.errors.EmptyDataError:
                error_paths.append(path)
                pass
            for item in df.iloc[:,0]:
                if item not in master_tag_list:
                    master_tag_list.append(item)
                else:
                    duplicates+=1
            pbar.update(1)

    descriptions=[desc_form_str(tag) for tag in master_tag_list]
    uoms=[uom_form_str(tag) for tag in master_tag_list]
    output_frame=pd.DataFrame({'Tag':master_tag_list,'Description':descriptions,'Unit':uoms})

    #n = 100  #chunk row size
    #list_df = [output_frame[i:i+n] for i in range(0,output_frame.shape[0],n)]

    os.chdir(workingdir)

    #with pd.ExcelWriter(os.path.join(workingdir,'output.xlsx'),
                    #mode='a',engine='openpyxl') as writer:
        #for df in tqdm(list_df,position=4,desc='PIDataLink Extraction'):
            #df.to_excel('temp.xlsx',index=False)
            #xlUpdate.xlUpdate(os.path.join(workingdir,'temp.xlsx'))
            #df2=pd.read_excel('temp.xlsx',index=False)
            #df2.to_excel(writer,index=False)

    print("\n\nExtracting PIDataLink Data... This may take forever.")
    output_frame.to_excel(os.path.join(workingdir,'temp.xlsx'),index=False)
    xlUpdate.xlUpdate(os.path.join(workingdir,'temp.xlsx'))
    df2=pd.read_excel(os.path.join(workingdir,'temp.xlsx'),index=False)
    df2.to_excel(os.path.join(workingdir,'output.xlsx'),index=False)
    input("Press the enter key to exit.")
                 
if __name__ == '__main__':
    main()
