import xml.etree.ElementTree as ET
import sys
import pandas as pd
from openpyxl import *
import os

def main():
    # Creates tree and root from xml part of pzfx file
    args = sys.argv

    if len(args) < 3:
        sys.exit("Use: " + args[0] + " input_file(or dir) output_file(or dir)")

    input = args[1]
    output = args[2]
    
    if os.path.splitext(input)[1] == '.pzfx' and os.path.splitext(output)[1] == '.xlsx':
        convert_to_excel(input, output)
    elif os.path.isdir(input) and os.path.isdir(output):
        convert_to_excel_dir(input, output)
    else:
        sys.exit("Input and output have to be both files or both directories\nInput files must be '.pzfx' and output files '.xlsx'")


def convert_to_excel_dir(input_dir, output_dir):
    input_files = []
    for file in os.listdir(input_dir):
        if file[-4:] == 'pzfx':
            input_files.append(file)
    
    if len(input_files) == 0:
        sys.exit("No compatible input files in directory, files must be '.pzfx'")
    
    for file in input_files:
        input_file = os.path.join(input_dir, file)
        output_file = os.path.join(output_dir, os.path.splitext(file)[0] + ".xlsx")
        convert_to_excel(input_file, output_file)


def convert_to_excel(input_file, output_file):
    print("Converting file " + input_file + " to " + output_file)
    tree = ET.parse(input_file)
    root = tree.getroot()
    tables = get_tables(root)
    tables_to_excel(tables, output_file)
    print("Convertion done!")


def get_tables(root):
    tables = {}
    # Loop over file
    for table in root:
        # Get tables
        if table.tag[-5:] == "Table":
            tableTitle = ""
            tableData = {}
            # Loop over table parts
            for part in table:
                # Get table title
                if part.tag[-5:] == "Title":
                    tableTitle = part.text
                # Get table columns
                if part.tag[-7:] == "YColumn":
                    columnTitle = ""
                    datapoints = []
                    # Loop over column parts
                    for subpart in part:
                        # Get column header
                        if subpart.tag[-5:] == "Title":
                            columnTitle = subpart.text
                        # Get column data
                        if subpart.tag[-9:] == "Subcolumn":
                            # Loop over data
                            for datafield in subpart:
                                datapoint = datafield.text
                                if (datapoint == None):
                                    continue
                                datapoints.append(datapoint)
                    # Add column (as floats) to table dict
                    tableData[columnTitle] = stringlist_to_floatlist(datapoints)
            # Add table (as DataFrame) to tables dict
            tables[tableTitle] = tableData
    return tables


def tables_to_excel(tables, output_file):

    with pd.ExcelWriter(output_file) as writer:
    
        # Loop over tables and write them to new excel sheet
        for table_title in tables:
            table = tables[table_title]
            table_df = pd.DataFrame()

            # Loop over columns and add them to table df
            for column_title in table:
                column_df = pd.DataFrame({column_title: table[column_title]})
                table_df = pd.concat([table_df,column_df], axis=1, sort=False)
                
            # Flip table and add average and SEM
            table_df_transposed = table_df.transpose()
            table_df_transposed['Average'] = table_df_transposed.mean(axis=1)
            table_df_transposed['SEM'] = table_df_transposed.sem(axis=1)

            table_df_transposed.to_excel(writer, sheet_name=clean_name(table_title), index=True)


def stringlist_to_floatlist(stringlist):
    s_list = [i.replace(",", ".") for i in stringlist]
    f_list = [float(i) for i in s_list]
    return f_list


def clean_name(original_name):
    new_name = original_name.replace(" ", "_").replace("/", "_")
    if len(new_name) > 30:
        new_name = new_name[:30]
    return new_name


if __name__ == '__main__':
    main()

