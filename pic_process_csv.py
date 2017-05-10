# coding: utf-8
import pandas as pd, datetime, glob, argparse, sys
from os import path as checkpath
from os import makedirs

# VIS field settings
# A better way to do this is create these arrays as the columns are renamed
visfields = ['Area','Convex Hull Area','Caliper Length','Compactness']
color_class_fields = [' Color Class 1', ' Color Class 2', ' Color Class 3']
flu_fields = ['FLUO No','FLUO Low','FLUO Med','FLUO High']
nir_fields = ['NIR Low','NIR Med','NIR High']
color_fields = ['Green','Yellow']

# computer SEM too?
sem = True

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

def renamer(df):
	newnames = {}
	color_class_problem_present = False
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
		if "colorclass_01" in col.lower():
				newnames[col] = 'Color Class 1'
				color_class_problem_present = True
		if "colorclass_02" in col.lower():
				newnames[col] = 'Color Class 2'
				color_class_problem_present = True
		if "colorclass_03" in col.lower():
				newnames[col] = 'Color Class 3'
				color_class_problem_present = True
		if "writer" in col.lower():
				newnames[col] = 'Sensor'
	return (df.rename(columns=newnames), color_class_problem_present)

## Parses the Sensor labels from the "Writer Label" column
def parse_writer_label(label):
    if "flu" in label:
        return "FLU"
    if "vis" in label:
        return "VIS"
    if "nir" in label:
        return "NIR"
    if "ir" in label:
        return "IR"

## Used to apply the row value based on the same column - if color_class_problem_present == True
def fix_class_problem(type, row, col):
    if row['Sensor'] == type:
        return row[col]
    else:
        return None

def construct_class_fields(type):
	fields = []
	for color in color_class_fields:
		fields.append(type+color)
	return fields

## Reformats the dataframe into something that makes sense
def reformat(df, color_class_problem_present):
	mid = df['date']
	df.drop(labels=['date', 'Row No'], axis=1, inplace = True)
	df.insert(0, 'date', mid)
	df.insert(3, 'Row', None)
	df.insert(4, 'Plant', None)
	df.insert(5, 'Plant ID', None)
	# Build the sensor label
	df['Sensor'] = df.apply(lambda row: parse_writer_label(row['Sensor']), axis=1)
	if color_class_problem_present:
		df.insert(6, 'FLU Color Class 1', None)
		df['FLU Color Class 1'] = df.apply(lambda row: fix_class_problem('FLU',row, 'Color Class 1'), axis=1)
		df.insert(7, 'NIR Color Class 1', None)
		df['NIR Color Class 1'] = df.apply(lambda row: fix_class_problem('NIR',row, 'Color Class 1'), axis=1)
		df.insert(8, 'FLU Color Class 2', None)
		df['FLU Color Class 2'] = df.apply(lambda row: fix_class_problem('FLU',row, 'Color Class 2'), axis=1)
		df.insert(9, 'NIR Color Class 2', None)
		df['NIR Color Class 2'] = df.apply(lambda row: fix_class_problem('NIR',row, 'Color Class 2'), axis=1)
		df.insert(10, 'FLU Color Class 3', None)
		df['FLU Color Class 3'] = df.apply(lambda row: fix_class_problem('FLU',row, 'Color Class 3'), axis=1)
		df.insert(11, 'NIR Color Class 3', None)
		df['NIR Color Class 3'] = df.apply(lambda row: fix_class_problem('NIR',row, 'Color Class 3'), axis=1)
		df.drop(labels=['Color Class 1', 'Color Class 2', 'Color Class 3'], axis=1, inplace = True)
	df['Row'] = df.apply(lambda row: row['ROI Label'].split('0')[0], axis=1)
	df['Plant'] = df.apply(lambda row: row['ROI Label'].split('0')[1], axis=1)
	# Give each plant a unique id for stats software
	df['Plant ID'] = df.apply(lambda row: row['Snapshot ID Tag']+'_'+row['ROI Label'], axis=1)
	df.drop(labels=['ROI Label'], axis=1,inplace = True)
	return df

def create_sensor_sheet(dates, names, fields, df, stat='mean'):
	header=pd.MultiIndex.from_product([names,fields],names=['exp','metric'])
	this_total = pd.DataFrame('', index=dates, columns = header)
	for name in names:
		for date in dates:
			for field in fields:
				newvalue = df[(df['date'] == date) & (df['Snapshot ID Tag'] == name) & (df['level_2'] == stat)][field]
				if not newvalue.empty:
					newvalue = newvalue.get_value(newvalue.index[0], field)
					this_total.set_value(date,(name,field),newvalue)
	return this_total

def create_vis_sheet(names, fields, df, stat='mean'):
    this_total = pd.DataFrame()
    for field in fields:
        interim = pd.DataFrame()
        for name in names:
            join1 = df[df['Snapshot ID Tag'].str.contains(name) & df['level_2'].str.contains(stat)][['date',field]].rename(columns={field: name})
            if interim.empty:
                interim = join1
            else:
                interim = pd.merge(interim, join1, how='outer', on=['date'])
        interim.insert(0,'Stat',field)
        this_total = pd.concat([this_total, interim])
    this_total = this_total.groupby(['Stat','date']).first()
    return this_total

def runall(path):
	pathin = path+'/RAW_CSV_DATA/*.csv'
	pathout = path+'/PROCESSED_CSV_DATA/'
	if not checkpath.exists(pathout):
		makedirs(pathout)
	files = glob.glob(pathin)

	## Create the empty DataFrame
	df = pd.DataFrame()
	# Iterate files and concatenate all into df
	for file in files:
	    df = pd.concat([df, pd.read_csv(file, sep=";")])
	# Remove extraneous entries where sensor scanned an empty area/well
	df = df[df['Area'] > 0]
	# Parse human readable date - seems to take the longest
	df['date'] = pd.to_datetime(df['Snapshot Time Stamp']).apply(lambda x: x.strftime('%Y-%m-%d'))
	# Rename columns, and check to see if it's the type with strange column type
	df, color_class_problem_present = renamer(df)
	# Capture all sent columns
	all_columns = df.columns
	# Reformat table
	df = reformat(df, color_class_problem_present)

	## Start data collection and statistical analysis
	# Create the xlsx writer and timestamp object first
	now = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
	writer = pd.ExcelWriter(pathout+'output_'+now+'.xlsx')

	# What sensors are we working with?
	sensors_present = df['Sensor'].unique()
	
	# Drop any of these that apply to the VIS camera first to get all other stats values
	drop_fields = ['Area','Convex Hull Area','Sensor','Caliper Length','Compactness']
	if "Excentricity" in all_columns:
		drop_fields.append("Excentricity")
	if "Circumference" in all_columns:
		drop_fields.append("Circumference")
	df2 = df.drop(drop_fields,axis=1)
	# Get VIS data
	visible_sensor_fields = ['Snapshot Time Stamp','Snapshot ID Tag','Row','Plant','Plant ID','Area','Convex Hull Area','Caliper Length','Compactness','date']
	if "Excentricity" in all_columns:
		visible_sensor_fields.append("Excentricity")
	if "Circumference" in all_columns:
		visible_sensor_fields.append("Circumference")
	df3 = df[df['Sensor'].str.contains('VIS')][visible_sensor_fields]
	# Get other sensors too:
	# Merge these back together
	if color_class_problem_present:
		df4 = pd.merge(df3, df[df['Sensor'].str.contains('NIR')][['Snapshot Time Stamp', 'date', 'Snapshot ID Tag', 'Row', 'Plant', 'Plant ID', 'NIR Color Class 1', 'NIR Color Class 2', 'NIR Color Class 3']], on=['Snapshot ID Tag', 'Row', 'Plant', 'date', 'Plant ID'])
		df4 = pd.merge(df4, df[df['Sensor'].str.contains('FLU')][['Snapshot Time Stamp', 'date','Snapshot ID Tag', 'Row', 'Plant', 'Plant ID', 'FLU Color Class 1', 'FLU Color Class 2', 'FLU Color Class 3']], on=['Snapshot ID Tag', 'Row', 'Plant', 'date', 'Plant ID'])
		if 'IR' in sensors_present:
			df4 = pd.merge(df4, df[df['Sensor'].str.contains('IR')][['date','Snapshot ID Tag', 'Row', 'Plant', 'Plant ID', 'IR Color Class 1', 'IR Color Class 2', 'IR Color Class 3']], on=['Snapshot ID Tag', 'Row', 'Plant', 'date', 'Plant ID'])
		df4 = df4.groupby(['date','Snapshot Time Stamp','Snapshot ID Tag', 'Row', 'Plant', 'Plant ID']).mean()
	else:
		## FIX: capture the mean - this drops any non-numerical column like Snapshot Time Stamp which is only referenced in stats below to be removed anyways
		df4 = pd.merge(df2,df3,on=['Snapshot ID Tag','Row','Plant','Plant ID', 'date']).groupby(['Snapshot ID Tag', 'Row','Plant', 'Plant ID', 'date']).mean()
	# Create RAW data table
	raw_data = df4.reset_index()
	# Attach RAW data to worksheet as sheet named 'Raw Data'
	prep_sheet(writer, raw_data, 'Raw Data')
	# Collect stats on all non-VIS data using describe
	# if color_class_problem_present:
	# 	df4 = df4.groupby(['Snapshot ID Tag', 'Row','Plant','date', 'Plant ID']).mean()
	df5 = df4.reset_index().groupby(['Snapshot ID Tag','date']).describe()
	df6 = df5.reset_index()
	# Collect SEM on all non-VIS data using .sem() called 'level_2' for summary/grouping later
	df_sem = df4.reset_index().groupby(['Snapshot ID Tag','date']).sem().reset_index()
	df_sem['level_2'] = 'sem'

	# get date collection for global use
	dates = df6['date'].unique()
	dates = sorted(dates)

	# Join the two stats tables
	df_comb = pd.concat([df6,df_sem]).groupby(['Snapshot ID Tag','date','level_2']).first()
	df_comb2 = df_comb.reset_index()
	# Remove metadata fields and keep all stats
	statistics = df_comb2.drop(['Plant','Plant ID','Row'], axis=1)
	statistics = statistics.rename(columns={'level_2':'Statistic'}).groupby(['Snapshot ID Tag','date','Statistic']).first()
	# Add stats to worksheet as sheet called 'Statistics'
	prep_sheet(writer, statistics, 'Statistics',[0,0,0,0,0,0,0,0,0])

	# Format each sensor data by DATE and Snapshot ID Tag
	exnames = df6['Snapshot ID Tag'].unique()
	exnames = sorted(exnames, key=lambda s: s.lower())
	
	# VIS Stats
	if "Excentricity" in all_columns:
		visfields.append("Excentricity")
	if "Circumference" in all_columns:
		visfields.append("Circumference")
	prep_sheet(writer, create_vis_sheet(exnames, visfields, df_comb2, 'mean'), 'VIS (mean)')
	if sem:
		prep_sheet(writer, create_vis_sheet(exnames, visfields, df_comb2, 'sem'), 'VIS (sem)')

	# Color VIS Stats
	if not color_class_problem_present:
		prep_sheet(writer, create_sensor_sheet(dates, exnames, color_fields, df_comb2, 'mean'), 'VIS (color - mean)')
		if sem:
			prep_sheet(writer, create_sensor_sheet(dates, exnames, color_fields, df_comb2, 'sem'), 'VIS (color - sem)')

	# NIR Stats
	if color_class_problem_present:
		prep_sheet(writer, create_sensor_sheet(dates, exnames, construct_class_fields('NIR'), df_comb2, 'mean'), 'NIR (mean)')
		if sem:
			prep_sheet(writer, create_sensor_sheet(dates, exnames, construct_class_fields('NIR'), df_comb2, 'sem'), 'NIR (sem)')
	else:
		prep_sheet(writer, create_sensor_sheet(dates, exnames, nir_fields, df_comb2, 'mean'), 'NIR (mean)')
		if sem:
			prep_sheet(writer, create_sensor_sheet(dates, exnames, nir_fields, df_comb2, 'sem'), 'NIR (sem)')
	
	# FLU Stats
	if color_class_problem_present:
		prep_sheet(writer, create_sensor_sheet(dates, exnames, construct_class_fields('FLU'), df_comb2, 'mean'), 'FLU (mean)')
		if sem:
			prep_sheet(writer, create_sensor_sheet(dates, exnames, construct_class_fields('FLU'), df_comb2, 'sem'), 'FLU (sem)')
	else:
		prep_sheet(writer, create_sensor_sheet(dates, exnames, flu_fields, df_comb2, 'mean'), 'FLU (mean)')
		if sem:
			prep_sheet(writer, create_sensor_sheet(dates, exnames, flu_fields, df_comb2, 'sem'), 'FLU (sem)')
	
	if 'IR' in sensors_present:
		# FLU Stats
		if color_class_problem_present:
			prep_sheet(writer, create_sensor_sheet(dates, exnames, construct_class_fields('IR'), df_comb2, 'mean'), 'IR (mean)')
			if sem:
				prep_sheet(writer, create_sensor_sheet(dates, exnames, construct_class_fields('IR'), df_comb2, 'sem'), 'IR (sem)')

	# Clasp hands, wiggle toes and ears, pray, meditate, send good feels into the universe and...
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