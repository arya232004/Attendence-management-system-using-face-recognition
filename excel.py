import xlsxwriter 
workbook = xlsxwriter.Workbook('attendence.xlsx')
# The workbook object is then used to add new
# worksheet via the add_worksheet() method.
worksheet = workbook.add_worksheet()
 
# Use the worksheet object to write
# data via the write() method.
worksheet.write('A1', 'Roll no')
worksheet.write('B1', 'Name')
worksheet.write('C1', 'Face Encoding')
worksheet.write('D1', 'Photo')
worksheet.write('E1', 'Attendence')


 
# Finally, close the Excel file
# via the close() method.
workbook.close()