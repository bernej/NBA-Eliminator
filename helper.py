import datetime, xlrd

def read_sheet(path, index):
    path = 'Analytics_Attachment.xlsx'

    workbook = xlrd.open_workbook(path)
    worksheet = workbook.sheet_by_index(index)

    # Change this depending on how many header rows are present
    # Set to 0 if you want to include the header data.
    offset = 0

    rows = []
    for i, row in enumerate(range(worksheet.nrows)):
        if i <= offset:  # (Optionally) skip headers
            continue
        r = []
        for j, col in enumerate(range(worksheet.ncols)):
            if index and col == 0:
                a1 = worksheet.cell_value(rowx=row, colx=col)
                a1_as_datetime = datetime.datetime(*xlrd.xldate_as_tuple(a1, workbook.datemode))
                date = str(a1_as_datetime)
                # print date.split()[0]
                # print 'datetime: %s' % a1_as_datetime
                r.append(date.split()[0])
            elif type(worksheet.cell_value(i, j)) is not float:
                r.append(worksheet.cell_value(i, j).encode("utf-8"))
            else:
                r.append(worksheet.cell_value(i, j))
        rows.append(r)
    print (len(rows))
    print (offset)
    print ('Got %d rows' % (len(rows) - offset))
    # print (rows[0])  # Print column headings
    # print (rows[offset])  # Print first data row sample
    return rows