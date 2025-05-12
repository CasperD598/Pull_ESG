import openpyxl
from openpyxl import Workbook
from tkinter import filedialog
from tkinter import messagebox
import os
import os.path

#Create lists of the .txt files
with open("jobtext.txt", "r") as jobtext_file:
    jobtexts = jobtext_file.read()
    jobtexts = jobtexts.split("\n")

with open("header.txt", "r") as header_file:
    headers = header_file.read()
    headers = headers.split("\n")

#Create list of files in folder
def list_of_files(folder_path):
    return os.listdir(folder_path)
    
#Create list of rows where headers match
def find_match(folder_path, files, headers, jobtexts):
    #Create file to paste result
    result_wb = Workbook()

    #For each file in folder
    for file in files:
        job_row = []
        match_list = []
        tank_match = []
        found_types = []

        data_workbook = openpyxl.load_workbook(f"{folder_path}/{file}")
        data_worksheet = data_workbook.active

        #Find job rows
        for i in range(5, data_worksheet.max_row):
            if data_worksheet.cell(row=i, column=1).value != None:
                job_row.append(i)

        print(job_row)    
        #Find rows that contain header
        for job in job_row:
            job_text = data_worksheet.cell(row=job, column=3).value.lower()
            for header in headers:
                if header in job_text and "tank" not in job_text and "tanks" not in job_text:
                    match_list.append(job_text)
                    break
                elif "tank" in job_text or "tanks" in job_text:
                    tank_match.append(job_text)
                    break
        
        
        result = []
        #print(match)
        for match in match_list:
            print(match)
            for line in range(match[0], job_row[match[1]]+1):
                print(line)
                #for text in jobtexts:
                    #if text in data_worksheet.cell(row=line, column=3).value.lower() and data_worksheet.cell(row=line, column=5).value != "":
                        #result.append([match[1], match[2], data_worksheet.cell(row=line, column=5).value])
        
        #print(result)
            
        """for x in range(5, data_worksheet.max_row):
            if data_worksheet.cell(row=x, column=1).value != None:
                for header in headers:
                    if header in str(data_worksheet.cell(row=x, column=3).value).lower():
                        if str(data_worksheet.cell(row=x, column=3).value) not in found_types:
                            found_types.append(str(data_worksheet.cell(row=x, column=3).value))                         

        #Search for job text in matches
        text_match = []
        for match in match_list:
            for line in range(match[0], job_row[match[1]]+1):
                for text in jobtexts:
                    if text in str(data_worksheet.cell(row=line, column=3).value).lower():
                        text_match.append([line, str(data_worksheet.cell(row=line, column=3).value).lower(), str(data_worksheet.cell(row=line, column=5).value).lower(), match[2], match[3]])
        
        #Paste result
        result_wb.create_sheet(file)
        result_ws = result_wb[file]
        result_ws.cell(row=1, column=1).value = "ROW"
        result_ws.cell(row=1, column=2).value = "TEXT"
        result_ws.cell(row=1, column=3).value = "AMOUNT"
        result_ws.cell(row=1, column=4).value = "JOB HEADER"
        result_ws.cell(row=1, column=5).value = "JOB NUMBER"

        x = 2
        used_row = []
        for match in text_match:
            if match[0] not in used_row:
                result_ws.cell(row=x, column=1).value = match[0]
                result_ws.cell(row=x, column=2).value = match[1]
                result_ws.cell(row=x, column=3).value = match[2]
                if match[4] not in used_row:
                    result_ws.cell(row=x, column=4).value = match[4]
                    result_ws.cell(row=x, column=5).value = match[3]
                used_row.append(match[0])
                used_row.append(match[4])
                x += 1

    result_wb.remove(result_wb["Sheet"])
    result_wb.save(f"{folder_path}/Resultat.xlsx")
    result_wb.close()"""
    #print(found_types)
    if os.path.isfile(f"{folder_path}/Resultat.xlsx"):
        messagebox.showinfo(message=f"Done!\nResult was place in {folder_path}/Resultat.xlsx")  
    else:
        messagebox.showerror(message="An error has occured!")


if __name__ == "__main__":
    #try:
        folder_path = filedialog.askdirectory(title="Hvor ligger filerne?")
        files = list_of_files(folder_path)
        find_match(folder_path, files, headers, jobtexts)
    #except:
       #messagebox.showerror(message="An error has occured!")
    
