#Brooks Pettit
#brooks.pettit@live.com
#(832)-580-7346
#2020-03-19

"""A script to generate a list of py tags from all _tag.csv files in a directory tree.
Pass the path ro the filesystem and the target file."""

import os
import pandas as pd
import numpy as np
import datetime
import xlUpdate

import click

@click.command()
@click.argument('filesystem',type=click.Path(exists=True))
@click.argument('target',type=click.Path())
def tagcompile(filesystem,target):
    print('\n\n')
    """Search FILESYSTEM for _tag.csv files; write tag names to TARGET"""
    dirstring=click.format_filename(filesystem)
    output=click.format_filename(target)
    workingdir=os.getcwd()
    os.chdir(dirstring)
    #print(os.getcwd())

    tag_files_list=[]
    filenames=[]
    totalfiles=0
    for root, dirs, files in os.walk(".", topdown=False):
        for name in files:
            totalfiles+=1
            if name[-9:] == '_tags.csv':
                tag_files_list.append(os.path.join(os.getcwd(),root[2:],name))
                filenames.append(name)


    
    master_tag_list=[]
    error_paths=[]
    empty_paths=[]
    duplicates=0

    with click.progressbar(tag_files_list,length=len(tag_files_list),label='Collecting tags') as bar:
        for path in bar:
            try:
                df=pd.read_csv(path,header=None,skip_blank_lines=True)
            except pd.errors.ParserError:
                error_paths.append(path)
                continue
            except pd.errors.EmptyDataError:
                continue
            for item in df.iloc[:,0]:
                if item not in master_tag_list:
                    master_tag_list.append(item)
                else:
                    duplicates+=1

    print("\n\nOf %d files in FILESYSTEM, found %d _tag.csv files"%(totalfiles,len(filenames)))
    print("%d files searched for %d unique tags"%(len(filenames),len(master_tag_list)))
    print("Command returned with %d files not searched"%(len(error_paths)))
    if duplicates!=0:
        print("%f%% tag duplication"%(duplicates*100/(len(master_tag_list)+duplicates)))
    else:
        print("0% tag duplication")

    os.chdir(workingdir)
    dict_init={'tag':master_tag_list}
    df_tags=pd.DataFrame(dict_init)
    
    descriptions=['=PITagAtt("%s","description","")'%(tag) for tag in df_tags['tag']]
    df_tags['description']=descriptions

    uoms=['=PITagAtt("%s","uom","")'%(tag) for tag in df_tags['tag']]
    df_tags['unit']=uoms

    writer=pd.ExcelWriter(output)
    df_tags.to_excel(writer,sheet_name='Tags',index=False)
    writer.save()
    writer.close()

    print("Evaluating excel formulas. This may take several minutes...")
    xlUpdate.xlUpdate(os.path.join(os.getcwd(),output))


if __name__ == '__main__':
    tagcompile()
