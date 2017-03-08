# PIC-Parse-CSV-to-Excel
Parse raw CSV files from LemnaTec Scanalyzer Database into PIC Excel Format

Command line usage:

python pic_process_csv.py -p "path_to_folder"

The "path_to_folder" MUST contain a subfolder called "RAW_CSV_DATA" that contains all of the CSV files that will be merged for a single experiment. The script will then create a new folder titled "PROCESSED_CSV_DATA" that will contain an "output.xlsx" file with a timestamp in the filename. 

To run a new batch, simply replace the files in "RAW..." and run it again. The timestamp generation will help ensure that the output file is not accidentally overwritten.
