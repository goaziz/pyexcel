import xlsxwriter

workbook = xlsxwriter.Workbook('test.xlsx')
worksheet = workbook.add_worksheet()
cell_format = workbook.add_format({
    'align': 'center',
    'valign': 'vcenter',
    'border': 1,
    'bold': 1})

merge_format = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter'})
merge_format.set_font_size(8)

cell_format.set_font_size(8)
cell_format.set_rotation(90)

excel_column = ['C:C', 'D:D', 'E:E', 'F:F', 'G:G', 'H:H', 'I:I', 'J:J', 'K:K', 'L:L', 'M:M', 'N:N', 'O:O']

worksheet.set_row(8, 30)
worksheet.set_row(9, 30)

for i in excel_column:
    worksheet.set_column(i, 12)

worksheet.set_column('D:D', 30)
worksheet.set_column('C:C', 5)
numbers = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13]

row = 12
col = 0
for col_num, data in enumerate(numbers):
    worksheet.write(12, col_num + 2, data, merge_format)

worksheet.write('C14', 1, merge_format)
worksheet.merge_range('D14:O14', "O'zR IIV Markaziy apparati, QR IIV, Toshkent shahar IIBB va viloyatlar IIBlari apparatlari", merge_format)
worksheet.merge_range('D7:D12', 'Bosh boshqarma, boshqarma, \nbolimi va lavozimlari nomi', merge_format)
worksheet.merge_range('E7:E12', 'Maxsus unvoni \n(xodimlar \ntoifasi)', merge_format)
worksheet.merge_range('F7:N7', 'SHTAT BIRLIKLARI SONI', merge_format)
worksheet.merge_range('F8:N8', 'Shulardan taminlash manbalari boyicha', merge_format)
worksheet.merge_range('F9:F12', 'Jami:', merge_format)
worksheet.merge_range('C7:C12', 'qator soni', cell_format)
worksheet.merge_range('G9:G12', 'Respublika \nbyudjeti\n hisobidan', cell_format)
worksheet.merge_range('H9:H12', 'JIEM \nsistemasining \n1-moddasi \nhisobidan\n (boshqaruv \napparati\n xodimlari)',
                      cell_format)
worksheet.merge_range('I9:I12', 'JIEM \nsistemasining \n1-moddasi \nhisobidan', cell_format)
worksheet.merge_range('J9:J12', 'Milliy byudjet \nhisobidan', cell_format)
worksheet.merge_range('L9:L12', 'Muassasa \nhisobidan', cell_format)
worksheet.merge_range('K9:K12', 'Jamgarma \nhisobidan', cell_format)
worksheet.merge_range('M9:M12', 'Kapital qoyilmalar \nhisobidan', cell_format)
worksheet.merge_range('N9:N12', 'Oz ozini mablag \nbilan taminlovchi \nbolinmalar', cell_format)
worksheet.merge_range('O7:O12', 'Bosh lavozimlar soni', cell_format)

workbook.close()
