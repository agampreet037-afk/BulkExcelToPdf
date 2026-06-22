# NOTE: To run this program, your system must have Microsoft Excel installed

import win32com.client 
import os
import time

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False

print("Welcome to Bulk Excel to PDF converter.")
def create_first_ws():
    for file in files:
        if file.endswith(".xlsx"):
            file_path = os.path.join(folder_path,file)   
            workbook = excel.Workbooks.Open(file_path)
            worksheet = workbook.Worksheets(1)
            worksheet.PageSetup.Orientation = 2     
            worksheet.PageSetup.Zoom = False
            worksheet.PageSetup.FitToPagesWide = 1
            worksheet.PageSetup.FitToPagesTall = False
            pdf_path = os.path.splitext(file_path)[0] + ".pdf"
            count = 1
            while os.path.exists(pdf_path):
                pdf_path = f"{os.path.splitext(file_path)[0]}" + f'({count}).pdf'
                count = count + 1
            worksheet.ExportAsFixedFormat(
            Type = 0, 
            Filename = pdf_path)   
            print(f"{file} converted to PDF successfully")
            workbook.Close(False)    

def create_a_ws():
    ws_num = int(input("Enter worksheet number: "))
    for file in files:
        if file.endswith(".xlsx"):
            file_path = os.path.join(folder_path,file)   
            workbook = excel.Workbooks.Open(file_path)
            worksheet = workbook.Worksheets(ws_num)
            worksheet.PageSetup.Orientation = 2      
            worksheet.PageSetup.Zoom = False
            worksheet.PageSetup.FitToPagesWide = 1
            worksheet.PageSetup.FitToPagesTall = False
            pdf_path = os.path.splitext(file_path)[0] + ".pdf"
            count = 1
            while os.path.exists(pdf_path):
                pdf_path = f"{os.path.splitext(file_path)[0]}" + f'({count}).pdf'
                count = count + 1
            worksheet.ExportAsFixedFormat(
            Type = 0,  
            Filename = pdf_path)
            print(f"{file} converted to PDF successfully")
            workbook.Close(False)      


def create_all_ws():
    for file in files:
        if file.endswith(".xlsx"):
            file_path = os.path.join(folder_path, file)
            workbook = excel.Workbooks.Open(file_path)


            for worksheet in workbook.Worksheets:
                worksheet.PageSetup.Orientation = 2      
                worksheet.PageSetup.Zoom = False
                worksheet.PageSetup.FitToPagesWide = 1
                worksheet.PageSetup.FitToPagesTall = False

            pdf_path = os.path.splitext(file_path)[0] + ".pdf"
            count = 1
            while os.path.exists(pdf_path):
                pdf_path = f"{os.path.splitext(file_path)[0]}" + f'({count}).pdf'
                count = count + 1
            workbook.ExportAsFixedFormat(
                Type=0,      
                Filename=pdf_path
            )
            print(f"{file} converted to PDF successfully.")
            workbook.Close(False)

       

while True:
        folder_path = input("Enter Path of folder containing Excel Files: ").strip().strip('"')
        if not os.path.exists(folder_path):
            print("Please enter a valid path")
            time.sleep(2)
            continue
        files = os.listdir(folder_path)
        user_input = input("Enter as per your need\n1. Convert first worksheet\n2. Convert worksheet by number\n3. Convert all worksheets\n")
        if user_input == "1":
            create_first_ws()

        elif user_input == "2":
            create_a_ws()

        elif user_input == "3":
            create_all_ws()
                            

        else:
            print("Enter valid input!")
            time.sleep(2)
            continue


        user_resp = input("Do you want to convert any other files?(y/n): ")
        if user_resp.lower() != "y":
            print("Goodbye!")
            break 


excel.Quit()



    


