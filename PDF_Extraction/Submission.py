# Importing the libraires
import tabula
import PyPDF2
import os
import fnmatch
import xlrd 
from xlutils.copy import copy

#pdfs_path = "/home/kunal/Desktop/Tasks"
path = "Discharge Summary 1.pdf"

no_of_pdfs = len(fnmatch.filter(os.listdir(), '*.pdf'))
print(no_of_pdfs)

def pdf_to_excel(pt_ctr):
    dt[:]
    dt.info()
    dt.columns
    new_data = dt.rename(columns = {"Unnamed: 0": "Col1","Unnamed: 1": "Col1d","Unnamed: 2": "Col2",
    "DISCHARGE SUMMARY": "Col3","Unnamed: 4": "Col3d","Unnamed: 5": "Col4",}) 
    new_data.info()
    
    new_data.columns
    
    new_data.drop(['Col1d', 'Col3d'], axis=1)
    
    patient_name = new_data.iloc[0,2]
    mrn = new_data.iloc[1,2]
    admitting_consultant = new_data.iloc[2,2]
    admn_date = new_data.iloc[5,2]
    discharge_date = new_data.iloc[5,5]
    visit_no = new_data.iloc[1,5]
    age_sex = new_data.iloc[0,5]
    department = new_data.iloc[4,5]
    
    ad_dt = admn_date.split(" ")
    admission_date = ad_dt[0]
    print(admission_date)
    
    age_data = age_sex.split("/")
    age = age_data[0]
    sex = age_data[1]
    print("Age:",age,"\nSex:",sex)
    
    pdfFileObj = open(pt_ctr, 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
    print(pdfReader.numPages)
    
    pageObj = pdfReader.getPage(0)
    print(pageObj.extractText())
    
    page1_data = pageObj.extractText()
    pageObj = pdfReader.getPage(1)
    page2_data = pageObj.extractText()
    
    def between(value, a, b):
    # Find and validate before-part.
        pos_a = value.find(a)
        if pos_a == -1: return ""
        # Find and validate after part.
        pos_b = value.rfind(b)
        if pos_b == -1: return ""
        # Return middle part.
        adjusted_pos_a = pos_a + len(a)
        if adjusted_pos_a >= pos_b: return ""
        return value[adjusted_pos_a:pos_b]
    
    # Test the methods with this literal.
    discharge_diagnosis = between(page1_data, "Discharge diagnosis :", "Consultants involved :")
    brief_history = between(page1_data, "Brief history and physical on admission :", "Significant Past Medical and Surgical History:")
    course_in_hospital = between(page1_data, "Course in the hospital :", "Procedures Performed :")
    discharge_advise = between(page2_data, "Discharge Medication and advice :", "Follow up /appointment :")
    discharge_instructions = between(page2_data, "Discharge Instructions / When to Obtain Urgent Care :", "ﬁPlease")
    if discharge_instructions == "":
        discharge_instructions = None
    
    print('''MRN : ''',mrn,
      "\n\nVisit No : ",visit_no,
      "\n\nPatient Name : ",patient_name, 
      "\n\nAge : ",age,
      "\n\nSex : ",sex,
      "\n\nAdmitting Consultant : ",admitting_consultant,
      "\n\nAdmission Date : ",admission_date,
      "\n\nDischarge Date : ",discharge_date,
      "\n\nDepartment : ",department,
      "\n\nDischarge Diagnosis : ",discharge_diagnosis,
      "\n\nBrief History : ",brief_history,
      "\n\nCourse in Hospital : ",course_in_hospital,
      "\n\nDischarge Medical and Advise : ",discharge_advise,
      "\n\nDischarge Instructions: ",discharge_instructions
     )
    
    # Give the location of the file 
    loc = ("DS_Format_Automation.xlsx") 
      
    wb = xlrd.open_workbook(loc) 
    sheet = wb.sheet_by_index(0) 
    sheet.cell_value(0, 0) 
    print("Number of rows :",sheet.nrows)
    print("\nNumber of Columns:",sheet.ncols)
    print("\nNumber of sheets :",wb.nsheets) 
    
    # Workbook is created 
    wb = xlrd.open_workbook(loc)
    wb = copy(wb)
    sheet1 = wb.get_sheet(0)
    
    row_count = sheet.nrows + 1
    sheet1.write(sheet.nrows,0 , mrn) 
    sheet1.write(sheet.nrows,1 , visit_no)
    sheet1.write(sheet.nrows,2 , patient_name)
    sheet1.write(sheet.nrows,3 , age)
    sheet1.write(sheet.nrows,4 , sex)
    sheet1.write(sheet.nrows,5 , admitting_consultant)
    sheet1.write(sheet.nrows,6 , admission_date)
    sheet1.write(sheet.nrows,7 , discharge_date)
    sheet1.write(sheet.nrows,8 , department)
    sheet1.write(sheet.nrows,9 , discharge_diagnosis)
    sheet1.write(sheet.nrows,10 , brief_history)
    sheet1.write(sheet.nrows,11 , course_in_hospital)
    sheet1.write(sheet.nrows,12 , discharge_advise)
    sheet1.write(sheet.nrows,13, discharge_instructions)
    
    wb.save('DS_Format_Automation.xlsx')
    return 0


for s in range(1,no_of_pdfs+1):
    c = str(s)
    path_counter = path.replace('1',c)
    dt = tabula.read_pdf(path_counter)
    pdf_to_excel(path_counter)
    print("Done")

