#Brooks Pettit
#brooks.pettit@live.com
#(832)-580-7346
#2020-03-19

import os
import pandas as pd
import numpy as np
import datetime
import xlUpdate

import click
import tqdm

date_format_string='%m/%d/%Y %H:%M'
def twa_formula_string(tag,start,stop):
    #pass timestamps as datetime.datetime objects
    #pass tag as a string
    startstring=start.strftime(date_format_string)
    stopstring=stop.strftime(date_format_string)

    return '=PIAdvCalcVal("%s","%s","%s","average","time-weighted",0,1,0,"")'%(tag,startstring,stopstring)

@click.command()
@click.argument('tagfile',type=click.Path(exists=True))
@click.argument('datafile',type=click.Path())
@click.option('--start',prompt='Start timestamp (MM/dd/yyyy hh/mm)')
@click.option('--stepsize',prompt='Step size (minutes)',help='Step size bwtween timestamps in minutes')
@click.option('--steps',prompt='Steps',help='Number of datapoints to pull')
@click.option('--chunksize',default=10,help='Number of tags to handle in memory')
def datapull(tagfile,datafile,start,stepsize,steps,chunksize):
    """Extract PI data from TAGFILE into DATAFILE using Excel, PIDatalink, and python"""
    infile=click.format_filename(tagfile)
    outfile=click.format_filename(datafile)

    df_tags=pd.read_excel(infile)
    print("Generating timestamps...")
    startstamp=datetime.datetime.strptime(start,date_format_string)
    delta=datetime.timedelta(minutes=int(stepsize))
    timestamps=[startstamp+delta*i for i in range(0,int(steps)+1)]
    tags=list(df_tags.loc[:,'tag'])
    print("Generating PI Datalink formulas...")
    table=[[twa_formula_string(tag,ts,ts+delta) for tag in tags] for ts in timestamps]
    print("Generating DataFrame")
    df_data=pd.DataFrame(table,columns=tags)
    df_data['timestamps']=timestamps
    df_data=df_data.set_index('timestamps')
    print("Writing output file...")
    writer=pd.ExcelWriter(outfile)
    df_data.to_excel(writer,sheet_name='Data')
    writer.save()
    writer.close()
    print("Evaluating formulas. This may take several minutes.")
    xlUpdate.xlUpdate(os.path.join(os.getcwd(),outfile))
    input("Press the enter key to exit.")






if __name__ == '__main__':
    datapull()
