import xlsxwriter
import os

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
            print(os.path.splitext(txt_file)[0])
            filename = os.path.splitext(txt_file)[0]

            # Workbook() takes one, non-optional, argument
            # which is the filename that we want to create.
            workbook = xlsxwriter.Workbook(filename + '.xlsx')

            # The workbook object is then used to add new
            # worksheet via the add_worksheet() method.
            worksheet = workbook.add_worksheet()

            # Start from the first cell.
            # Rows and columns are zero indexed.
            row = 0
            column = 0

            # Using readlines()
            file = open(txt_file, 'r')
            Lines = file.readlines()

            count = 0

            # Strips the newline character
            for line in Lines:
                count += 1
                if not line.startswith("//") and not line == "\n" and not line == " \n":
                    print("Line{}: {}".format(count, line))

                    split_data = line.replace(';', '').split(":")
                    print(split_data)

                    op_and_mo = split_data[0].split(" ")
                    print(op_and_mo)

                    print()

                    i = 0
                    print(split_data[1])

                    parameters = split_data[1].split(',')
                    len_param = len(parameters)

                    flag = 0
                    if 'LOCALCELLID' in split_data[1] or 'LocalCellId' in split_data[1]:
                        flag = 1
                        localcellid = parameters[0]
                        parameters.pop(0)

                    if flag == 0:
                        worksheet.write(row, column, op_and_mo[0])
                        column += 1
                        worksheet.write(row, column, op_and_mo[1])
                        column += 1
                        for param in parameters:
                            other_param = param.split('=')

                            worksheet.write(row, column, other_param[0])
                            column += 1

                            ampersand_param = other_param[1].split('&')
                            for ambersand in ampersand_param:
                                worksheet.write(row, column, ambersand)
                                column += 1
                    else:
                        for param in parameters:
                            worksheet.write(row, column, op_and_mo[0])
                            column += 1
                            worksheet.write(row, column, op_and_mo[1])
                            column += 1

                            localcellid_split = localcellid.split('=')
                            worksheet.write(row, column, localcellid_split[0])
                            column += 1
                            worksheet.write(row, column, localcellid_split[1])
                            column += 1

                            other_param = param.split('=')

                            worksheet.write(row, column, other_param[0])
                            column += 1

                            ampersand_param = other_param[1].split('&')
                            for ambersand in ampersand_param:
                                worksheet.write(row, column, ambersand)
                                column += 1

                            row += 1
                            column = 0

                    row += 1
                    column = 0
                    # incrementing the value of row by one
                    # with each iterations.

            workbook.close()
