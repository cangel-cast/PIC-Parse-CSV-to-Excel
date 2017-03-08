# coding: utf-8
import pandas as pd, datetime, glob, argparse, sys
from os import path as checkpath
from os import makedirs

def prep_sheet(writer, inframe, insheet, dates=[]):
    
    inframe.to_excel(writer,insheet, index=(insheet != 'Raw Data'))
    #Indicate workbook and worksheet for formatting
    workbook = writer.book
    worksheet = writer.sheets[insheet]
    #Find the number of columns to iterate through
    columns = len(inframe.columns)
    rows = inframe.shape[0]
    column_series = pd.Series(range(columns))
    column_list = column_series.tolist()

    for i in column_list:
        #filter for header
        header = inframe[[i]].astype(str).columns.values
        #Get length of header    
        header_len = len(header.max()) + 2
        #Get length of column values
        column_len = inframe[[i]].astype(str).apply(lambda x: x.str.len()).max().max() + 2
        #Choose the greater of the column length or column value length
        if header_len >= column_len:
            worksheet.set_column(i,i,header_len)
        else:
            worksheet.set_column(i,i,column_len)   
    #Set border colors based on group lengths
    if len(dates) > 0:
        format = workbook.add_format()
        format.set_bottom(2)
        format.set_bottom_color('black')
        i = 1
        numdates = len(dates)
        while i <= rows/numdates:
            rownum = (i*numdates)
            worksheet.set_row(rownum, None, format)
            i+=1

def runall(path):
    pathin = path+'/RAW_CSV_DATA/*.csv'
    pathout = path+'/PROCESSED_CSV_DATA/'
    if not checkpath.exists(pathout):
        makedirs(pathout)
    files = glob.glob(pathin)
    df = pd.DataFrame()
    for file in files:
        df = pd.concat([df, pd.read_csv(file,sep=';')])

    df['date'] = pd.to_datetime(df['Snapshot Time Stamp']).apply(lambda x: x.strftime('%Y-%m-%d'))
    df.columns
    all_columns = df.columns
    newnames = {}

    #Parse column names
    for col in df.columns:
        if "fluo" in col.lower() or "signal" in col.lower():
            if "low" in col.lower():
                newnames[col] = 'FLUO Low'
            if "med" in col.lower():
                newnames[col] = 'FLUO Med'
            if "high" in col.lower():
                newnames[col] = 'FLUO High'
            if "no" in col.lower():
                newnames[col] = 'FLUO No'
        if "nir" in col.lower() or "water" in col.lower():
            if "low" in col.lower():
                newnames[col] = 'NIR Low'
            if "med" in col.lower():
                newnames[col] = 'NIR Med'
            if "high" in col.lower():
                newnames[col] = 'NIR High'
        if "yellow" in col.lower():
            newnames[col] = 'Yellow'
        if "green" in col.lower():
            newnames[col] = 'Green'
    df = df.rename(columns=newnames)
    
    #Move date field to beginning
    mid = df['date']
    df.drop(labels=['date','Row No'], axis=1,inplace = True)
    df.insert(0, 'date', mid)
    #Setup Row, Plant, and Plant ID fields
    df.insert(3, 'Row', None)
    df.insert(4, 'Plant', None)
    df.insert(5, 'Plant ID', None)
    df['Row'] = df.apply(lambda row: row['ROI Label'].split('0')[0], axis=1)
    df['Plant'] = df.apply(lambda row: row['ROI Label'].split('0')[1], axis=1)
    df['Plant ID'] = df.apply(lambda row: row['Snapshot ID Tag']+'_'+row['ROI Label'], axis=1)
    df.drop(labels=['ROI Label'], axis=1,inplace = True)

    df2 = df.drop(['Area','Convex Hull Area','Writer Label','Caliper Length','Compactness'],axis=1)
    df3 = df[df['Writer Label'].str.contains('vis_obj_cc')][['Snapshot ID Tag','Row','Plant','Area','Convex Hull Area','Caliper Length','Compactness','date']]
    df4 = pd.merge(df2,df3,on=['Snapshot ID Tag','Row','Plant','date']).groupby(['Snapshot ID Tag', 'Row','Plant', 'date']).max()
    #Raw data table to be written
    raw_data = df4.reset_index()

    df5 = df4.reset_index().groupby(['Snapshot ID Tag','date']).describe()

    df6 = df5.reset_index()

    df_sem = df4.reset_index().groupby(['Snapshot ID Tag','date']).sem().reset_index()

    df_sem['level_2'] = 'sem'

    df_comb = pd.concat([df6,df_sem]).groupby(['Snapshot ID Tag','date','level_2']).first()
    df_comb2 = df_comb.reset_index()
    statistics = df_comb2.drop(['Plant','Plant ID','Row','Snapshot Time Stamp'], axis=1)
    #Statistics table to be written
    statistics = statistics.rename(columns={'level_2':'Statistic'}).groupby(['Snapshot ID Tag','date','Statistic']).first()

    exnames = df6['Snapshot ID Tag'].unique()
    exnames = sorted(exnames, key=lambda s: s.lower())

    visfields = ['Area','Convex Hull Area','Caliper Length','Compactness']
    if "Excentricity" in all_columns:
        visfields.append("Excentricity")
    if "Circumference" in all_columns:
        visfields.append("Circumference")
    dates = df6['date'].unique()
    dates = sorted(dates)

    total_vis = pd.DataFrame()
    for field in visfields:
        interim = pd.DataFrame()
        for name in exnames:
            join1 = df_comb2[df_comb2['Snapshot ID Tag'].str.contains(name) & df_comb2['level_2'].str.contains('mean')][['date',field]].rename(columns={field: name})
            if interim.empty:
                interim = join1
            else:
                interim = pd.merge(interim, join1, how='outer', on=['date'])
        interim.insert(0,'Stat',field)
        total_vis = pd.concat([total_vis, interim])
    total_vis = total_vis.groupby(['Stat','date']).first()

    infields = ['Green','Yellow']
    header=pd.MultiIndex.from_product([exnames,infields],names=['exp','metric'])
    total_vis_metrics = pd.DataFrame('', index=dates, columns = header)
    for name in exnames:
        for date in dates:
            for field in infields:
                newvalue = df_comb2[(df_comb2['date'] == date) & (df_comb2['Snapshot ID Tag'] == name) & (df_comb2['level_2'] == 'mean')][field]
                if not newvalue.empty:
                    newvalue = newvalue.get_value(newvalue.index[0], field)
                    total_vis_metrics.set_value(date,(name,field),newvalue)

    infields = ['NIR Low','NIR Med','NIR High']
    newinfields = ['Low','Med','High']
    header=pd.MultiIndex.from_product([exnames,newinfields],names=['exp','metric'])
    total_nir_metrics = pd.DataFrame('', index=dates, columns = header)
    for name in exnames:
        for date in dates:
            for field in infields:
                newvalue = df_comb2[(df_comb2['date'] == date) & (df_comb2['Snapshot ID Tag'] == name) & (df_comb2['level_2'] == 'mean')][field]
                if not newvalue.empty:
                    newvalue = newvalue.get_value(newvalue.index[0], field)
                    total_nir_metrics.set_value(date,(name,field.split(' ')[1]),newvalue)

    infields = ['FLUO No','FLUO Low','FLUO Med','FLUO High']
    newinfields = ['No','Low','Med','High']
    header=pd.MultiIndex.from_product([exnames,newinfields],names=['exp','metric'])
    total_FLUO_metrics = pd.DataFrame('', index=dates, columns = header)
    for name in exnames:
        for date in dates:
            for field in infields:
                newvalue = df_comb2[(df_comb2['date'] == date) & (df_comb2['Snapshot ID Tag'] == name) & (df_comb2['level_2'] == 'mean')][field]
                if not newvalue.empty:
                    newvalue = newvalue.get_value(newvalue.index[0], field)
                    total_FLUO_metrics.set_value(date,(name,field.split(' ')[1]),newvalue)

    now = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
    #Create excel document with each sheet
    writer = pd.ExcelWriter(pathout+'output_'+now+'.xlsx')
    prep_sheet(writer, raw_data, 'Raw Data')
    prep_sheet(writer, statistics, 'Statistics',[0,0,0,0,0,0,0,0,0])
    prep_sheet(writer, total_vis, 'VIS', dates)
    prep_sheet(writer, total_vis_metrics, 'VIS - Metrics')
    prep_sheet(writer, total_nir_metrics, 'NIR (Water Content)')
    prep_sheet(writer, total_FLUO_metrics, 'Fluorescence')
    writer.save()
    if checkpath.exists(pathout+'output_'+now+'.xlsx'):
        print('output_'+now+'.xlsx')
        sys.exit(0)
    else:
        sys.exit(1)

#Accept system arguments
ap = argparse.ArgumentParser()
ap.add_argument("-p", "--path", required=True, help="path to set of csv files")
args = vars(ap.parse_args())
if "path" in args:
    runall(args['path'])