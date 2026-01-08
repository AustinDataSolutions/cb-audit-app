import os
from io import BytesIO

import pandas as pd
import xlsxwriter


def _build_sortable_filename(uploaded_file):
    original_name = getattr(uploaded_file, "name", "") or "reformatted_audit.xlsx"
    base, ext = os.path.splitext(original_name)
    if not ext:
        ext = ".xlsx"
    return f"{base}_sortable{ext}"


def handle_audit_reformat(uploaded_file):
    file_bytes = uploaded_file.read()
    excel_file = pd.ExcelFile(BytesIO(file_bytes))

    sentences_sheet = excel_file.parse(excel_file.sheet_names[0], header=None)
    categories_sheet = excel_file.parse(excel_file.sheet_names[1], header=None)

    new_sentences_rows = []

    sentence_rows = sentences_sheet.fillna("").values.tolist()
    row_count = len(sentence_rows)
    row_num = 2
    sentence_row_count = row_count
    previous_row = sentence_rows[0]
    while row_num < row_count:
        this_row = sentence_rows[row_num]
        if this_row[0] == "":
            this_row[0] = previous_row[0]
            this_row[1] = previous_row[1]
        new_sentences_rows.append(this_row)
        previous_row = this_row
        row_num += 1


    new_categories_rows = []
    category_rows = categories_sheet.fillna("").values.tolist()
    row_count = len(category_rows)
    row_num = 1
    while row_num < row_count:
        this_row = category_rows[row_num]
        new_categories_rows.append(this_row)
        row_num += 1

    output = BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})

    worksheet = workbook.add_worksheet("Sentences")
    #print("A")

    header = [
              '#',
              'Sentences',
              'Category',
              'Audit',
              'Category Notes',
              'Vote number',
              'Node identifier',
              'Sentiment',
              'Sentiment Audit',
              'Sentiment Notes',
              'Vote number',
              'Tokens',
              ]
    row_num=0
    col_num=0


    bold = workbook.add_format({'bold': True})
    green = workbook.add_format({'font_color': 'green'})
    red = workbook.add_format({'font_color': 'red'})
    percentage = workbook.add_format({'num_format': '0%'})
    #print("B")
    for cell in header:
        worksheet.write(row_num, col_num, cell, bold)
        col_num += 1
    row_num += 1
    rows_to_sort = []
    for report_row in new_sentences_rows:
        row_to_sort = []
        for cell in report_row:
            row_to_sort.append(cell)
        rows_to_sort.append(row_to_sort)

    #print("C")
    rows_to_sort.sort(key=lambda tup: tup[2], reverse=False)
    #print("D")
    for row in rows_to_sort:
        try:
            col_num = 0
            #print(row)
            for cell in row:

                if col_num==7 and cell in ("Positive", "Strongly Positive"):
                    worksheet.write(row_num, col_num, cell, green)
                elif col_num==7 and cell in ("Negative", "Strongly Negative"):
                    worksheet.write(row_num, col_num, cell, red)
                elif col_num==5:
                    worksheet.write(row_num, col_num, '=IF(D'+str(row_num+1)+'="YES", 1, 0)')
                elif col_num==10:
                    worksheet.write(row_num, col_num, '=IF(I'+str(row_num+1)+'="YES", 1, 0)')
                else:
                    worksheet.write(row_num, col_num, cell)

                col_num += 1
                if col_num==4:
                    col_num += 1
                if col_num==9:
                    col_num += 1
            #print("HERE")
            row_num += 1
        except:
            print("ROW ERROR ON "+str(row_num))
            #print(row)
    #print("EE")

    worksheet.set_column('F:G', None, None, {'hidden': True})
    worksheet.set_column('K:K', None, None, {'hidden': True})
    worksheet.set_column(0, 0, 9)
    worksheet.set_column(1, 1, 40)
    worksheet.set_column(2, 2, 30)
    worksheet.set_column(3, 3, 9)
    worksheet.set_column(4, 4, 30)
    worksheet.set_column(7, 7, 9)
    worksheet.set_column(8, 8, 9)
    worksheet.set_column(9, 9, 30)
    #print("D")
    worksheet.data_validation('D2:D'+str(sentence_row_count+1), {'validate': 'list',
                                      'source': ['YES', 'NO']})
    worksheet.data_validation('I2:I'+str(sentence_row_count+1), {'validate': 'list',
                                      'source': ['YES', 'NO']})



    #print("E")
    worksheet = workbook.add_worksheet("Categories")
    #categories_sheet
    header = [
              '#',
              'Category',
              'Precision rate',
              '',
              'Sentiment precision',
              ]
    row_num=0
    col_num=0
    for cell in header:
        worksheet.write(row_num, col_num, cell, bold)
        col_num += 1
    row_num += 1
    rows_to_sort = []
    for report_row in new_categories_rows:
        row_to_sort = []
        for cell in report_row:
            row_to_sort.append(cell)
        rows_to_sort.append(row_to_sort)

    rows_to_sort.sort(key=lambda tup: tup[1], reverse=False)
    for row in rows_to_sort:
        col_num = 0
        for cell in row:
            if col_num not in (2, 4):
                worksheet.write(row_num, col_num, cell)
            elif col_num==2:
                worksheet.write(row_num, col_num, "=IF(COUNTIF(Sentences!G:G, D"+str(row_num+1)+")=0, 1, SUMIF(Sentences!G:G, D"+str(row_num+1)+", Sentences!F:F)/COUNTIF(Sentences!G:G, D"+str(row_num+1)+" ))", percentage)
            elif col_num==4:
                worksheet.write(row_num, col_num, "=IF(COUNTIF(Sentences!G:G, D"+str(row_num+1)+")=0, 1, SUMIF(Sentences!G:G, D"+str(row_num+1)+", Sentences!K:K)/COUNTIF(Sentences!G:G, D"+str(row_num+1)+" ))", percentage)
            col_num += 1
        row_num += 1

    worksheet.set_column('D:D', None, None, {'hidden': True})
    worksheet.set_column(0, 0, 9)
    worksheet.set_column(1, 1, 30)
    worksheet.set_column(2, 2, 14)
    worksheet.set_column(4, 4, 14)
    workbook.close()
    output.seek(0)
    output_filename = _build_sortable_filename(uploaded_file)
    return output, output_filename
