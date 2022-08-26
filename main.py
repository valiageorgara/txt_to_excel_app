import xlsxwriter
import os
import re

if __name__ == '__main__':

    # Folder Path of txt files
    path = "txt_files"
    # Change the directory
    os.chdir(path)

    # Iterate through all files
    for txt_file in os.listdir():
        # Check whether file is in text format or not
        if txt_file.endswith(".txt"):
            print("Converting txt: " + txt_file + " to excel format.")
            # prints file name without extension
            # print(os.path.splitext(txt_file)[0])

            # store name of file in variable filename
            filename = os.path.splitext(txt_file)[0]

            # Workbook() takes one, non-optional, argument which is the filename that we want to create.
            workbook = xlsxwriter.Workbook(filename + '.xlsx')

            # The workbook object is then used to add new worksheet via the add_worksheet() method.
            worksheet = workbook.add_worksheet()

            # Start from the first cell. Rows and columns are zero indexed.
            row = 0
            column = 0

            # Using readlines()
            file = open(txt_file, 'r')
            Lines = file.readlines()

            count = 0

            skip = True
            # for each line of the txt
            for line in Lines:
                count += 1

                # skip lines until you find the cell configuration
                if 'ADD CELL:' not in line:
                    if skip:
                        continue
                else:
                    skip = False

                # once you've found cell configuration proceed with the below
                # Ignore new lines and comments
                if re.match('^[A-Za-z1-9 ]*:.*;', line):
                    print("Line{}: {}".format(count, line))

                    split_data = line.replace(';', '').split(":")
                    print(split_data)
                    # split data[0] = ADD CELL
                    # split data[1] = LocalCellId=11, CellName="KYL1841", FreqBand=20, UlEarfcnCfgInd=NOT_CFG, DlEarfcn=6200, UlBandWidth=CELL_BW_N50, DlBandWidth=CELL_BW_N50, CellId=11, PhyCellId=15, FddTddInd=CELL_FDD, RootSequenceIdx=606,CELLRADIUS=38975, CustomizedBandWidthCfgInd=NOT_CFG, UePowerMaxCfgInd=NOT_CFG, MultiRruCellFlag=BOOLEAN_FALSE, TxRxMode=2T2R;

                    op_and_mo = split_data[0].split(" ")

                    operation = op_and_mo[0]
                    manage_object = op_and_mo[1]

                    i = 0
                    print(split_data[1])

                    parameters = split_data[1].replace(' ', '').split(',')

                    localcellid_found = False
                    if 'LOCALCELLID' in split_data[1] or 'LocalCellId' in split_data[1]:
                        localcellid_found = True
                        localcellid = parameters[0]
                        localcellid_split = localcellid.split('=')
                        parameters.pop(0)

                    for param in parameters:
                        other_param = param.split('=')
                        ampersand_param = other_param[1].replace('"', '').split('&')
                        for ambersand in ampersand_param:
                            worksheet.write(row, column, operation)
                            column += 1
                            worksheet.write(row, column, manage_object)
                            column += 1

                            if localcellid_found:
                                worksheet.write(row, column, localcellid_split[0])
                                column += 1
                                worksheet.write(row, column, localcellid_split[1])
                                column += 1

                            worksheet.write(row, column, other_param[0])
                            column += 1

                            worksheet.write(row, column, ambersand)

                            row += 1
                            column = 0

                    row += 1
                    column = 0

            workbook.close()
