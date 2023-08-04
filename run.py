import openpyxl

subject=[]
chapter=[]


Excel=openpyxl.load_workbook("Excel_data.xlsx")


sheets_name=Excel.sheetnames

for i in range(len(sheets_name)):
    subject.append({'id':i+1,'name':sheets_name[i]})
# print(sheets_name)
# print(subject)

count = 1
for j in range(len(sheets_name)):
    current_sheet=Excel[sheets_name[j]]
    first_row_values = [cell.value for cell in current_sheet[1]]
    b = [{'id': count+number, 'chapter_name': i,'subject_id':j+1} for number,i in enumerate(first_row_values,start=0)]
    chapter.extend(b)
    count += len(first_row_values) 

# print(first_row_values)
# print(current_sheet)
print(chapter)

