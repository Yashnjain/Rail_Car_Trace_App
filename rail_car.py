import os
import re
import sys
import time
import shutil
import requests
import logging
import tkinter
import traceback
import bu_alerts
import numpy as np
import pandas as pd
import xlwings as xw
import customtkinter
from selenium import webdriver
from collections import Counter
from datetime import date,datetime
import xlwings.constants as win32c
from tkinter import messagebox,Tk
import xlwings.constants as win32c
from bu_config import config as buconfig
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.wait import WebDriverWait
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support import expected_conditions as EC
from sqlalchemy.dialects import registry 
registry.register('snowflake', 'snowflake.sqlalchemy', 'dialect')
# Modes: system (default), light, dark
customtkinter.set_appearance_mode("light")
# Themes: blue (default), dark-blue, green
customtkinter.set_default_color_theme("dark-blue")


def on_closing():
        try:
            if messagebox.askokcancel("Quit", "Do you want to quit?"):
                app.destroy()
                sys.exit()
        except Exception as e:
            print(f"Exception caught in on_closin method: {e}")
            logging.info(f"Exception caught in on_closin method: {e}")
            raise e
        
def report_callback_exception(self,exc, val, tb):
        
        msg = traceback.format_exc()
        messagebox.showerror("Error", message=msg)
        app.update()
        logging.exception(str(msg))
        # BU_LOG entry(Failed) in PROCESS_LOG table
        log_json = '[{"JOB_ID": "'+str(job_id)+'","JOB_NAME": "'+str(job_name)+'","CURRENT_DATETIME": "'+str(datetime.now())+'","STATUS": "FAILED"}]'
        bu_alerts.bulog(process_name=job_name,table_name=table_name,status='FAILED',process_owner=process_owner ,row_count=0,log=log_json,database=database,warehouse=warehouse)
        bu_alerts.send_mail(receiver_email = receiver_email,mail_subject =f'JOB FAILED -{job_name}',mail_body = f'{job_name} failed in __main__, Attached logs',attachment_location = logfile)
        

def resource_path(relative_path):
    try:
        """ Get absolute path to resource, works for dev and for PyInstaller """
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        return os.path.join(base_path, relative_path)
    except Exception as e:
        print(f"Exception caught in resource_path method: {e}")
        logging.info(f"Exception caught in resource_pathe method: {e}")
        raise e

def button_function():
    try:
        button_text.set("PROCESSING")
        button.configure(state='disable')
        app.update()
        main()
        button_text.set("Generate Trace Report")
        button.configure(state='normal')
        messagebox.showinfo("INFO",f"Trace Run Successful")
        app.quit()
    except Exception as e:
        print(f"Exception caught in button_function method: {e}")
        logging.info(f"Exception caught in button_function method: {e}")
        raise e

def row_range_calc(filter_col:str, input_sht,wb):
    try:
        sp_lst_row = input_sht.range(f'{filter_col}'+ str(input_sht.cells.last_cell.row)).end('up').row
        if sp_lst_row!=2:
            sp_address= input_sht.api.Range(f"{filter_col}2:{filter_col}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address
        else:
            sp_address="$2:$2"
        sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]        

        row_range = sorted([int(i) for i in list(set(re.findall("\d+",sp_address.replace("$",""))))])

        while row_range[-1]!=sp_lst_row:

            sp_lst_row = input_sht.range(f'{filter_col}'+ str(input_sht.cells.last_cell.row)).end('up').row

            sp_address = sp_address+','+(input_sht.api.Range(f"{filter_col}{row_range[-1]+1}:{filter_col}{sp_lst_row}").SpecialCells(win32c.CellType.xlCellTypeVisible).EntireRow.Address)

            # sp_initial_rw = re.findall("\d+",sp_address.replace("$","").split(":")[0])[0]

            row_range.extend(sorted([int(i) for i in list(set(re.findall("\d+",sp_address.replace("$",""))))]))
            
        sp_address = sp_address.replace("$","").split(",")
        init_list= [list(range(int(i.split(":")[0]), int(i.split(":")[1])+1)) for i in sp_address]
        sublist = []
        flat_list = [item for sublist in init_list for item in sublist]
        return flat_list, sp_lst_row,sp_address
    except Exception as e:
        print(f"Exception caught in row_range_calc method: {e}")
        logging.info(f"Exception caught in row_range_calc method: {e}")
        raise e


def num_to_col_letters(num):
    try:
        letters = ''
        while num:
            mod = (num - 1) % 26
            letters += chr(mod + 65)
            num = (num - 1) // 26
        return ''.join(reversed(letters))
    except Exception as e:
        print(f"Exception caught in num_to_col_letters method: {e}")
        logging.info(f"Exception caught in num_to_col_letters method: {e}")
        raise e

def custum_sort(workbook,worksheet,range1,range2,range3):
    try:
        worksheet.api.Sort.SortFields.Clear()
        worksheet.api.Sort.SortFields.Add2(Key:=worksheet.api.Range(range1), SortOn:=win32c.SortOn.xlSortOnValues, Order:=win32c.SortOrder.xlAscending,CustomOrder:="CO", DataOption:=win32c.SortDataOption.xlSortNormal)
        worksheet.api.Sort.SortFields.Add2(Key:=worksheet.api.Range(range2), SortOn:=win32c.SortOn.xlSortOnValues, Order:=win32c.SortOrder.xlAscending,CustomOrder:="Placed Actual,Placed Construct", DataOption:=win32c.SortDataOption.xlSortNormal)
        a = workbook.app.api.ActiveSheet.Sort
        a.SetRange(Rng=worksheet.api.Range(range3))
        a.Header = win32c.YesNoGuess.xlYes
        a.MatchCase = False
        a.Orientation = win32c.Constants.xlTopToBottom
        a.SortMethod = win32c.SortMethod.xlPinYin
        a.Apply()
    except Exception as e:
        print(f"Exception caught in custum_sort method: {e}")
        logging.info(f"Exception caught in custum_sort method: {e}")
        raise e
    
def interior_coloring(colour_value,cellrange:str,working_sheet,working_workbook):
    try:
        working_sheet.activate()
        if working_sheet.api.AutoFilterMode:
            working_sheet.api.Range(cellrange).SpecialCells(win32c.CellType.xlCellTypeVisible).Select()
        else:
            working_sheet.api.Range(cellrange).Select()
        a = working_workbook.app.selection.api.Interior
        a.Pattern = win32c.Constants.xlSolid
        a.PatternColorIndex = win32c.Constants.xlAutomatic
        a.Color = colour_value
        a.TintAndShade = 0
        a.PatternTintAndShade = 0
    except Exception as e:
        print(f"Exception caught in interior_coloring method: {e}")
        logging.info(f"Exception caught in interior_coloring method: {e}")
        raise e  
    

def remove_existing_files(files_location):
    """_summary_

    Args:
        files_location (_type_): _description_

    Raises:
        e: _description_
    """
    logger.info("Inside remove_existing_files function")
    try:
        files = os.listdir(files_location)
        if len(files) > 0:
            for file in files:
                os.remove(files_location + "\\" + file)
            logger.info("Existing files removed successfully")
        else:
            print("No existing files available to reomve")
        print("Pause")
    except Exception as e:
        print(f"Exception caught in remove_existing_files method: {e}")
        logging.info(f"Exception caught in remove_existing_files method: {e}")
        raise e 
   
def tracereport_dwonload():
    try:
        tracedict={"Test_format_trace":"Text Format - Event Translation","weight_format_trace":"W format - Scale Weight"}
        for key,value in tracedict.items():
            driver.switch_to.window(driver.window_handles[0])
            logger.info("change CLM format trace")
            repo_format = Select(WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "select[name='reportformat']"))))
            repo_format.select_by_visible_text(value)
            logger.info("putting location as 1st and making it ascending")
            WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[value='A'][name='sort_dir_1']"))).click()
            WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/form/table/tbody/tr[11]/td/table/tbody/tr[1]/td[2]/select"))).click()
            WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/form/table/tbody/tr[11]/td/table/tbody/tr[1]/td[2]/select/option[7]"))).click()
            logger.info("running trace")
            time.sleep(1)
            element = driver.find_element(By.CSS_SELECTOR, "input[value='Run']")
            driver.execute_script("arguments[0].scrollIntoView();", element)
            time.sleep(1)
            WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[value='Run']"))).click()
            sec_page = driver.window_handles[1] 
            driver.switch_to.window(sec_page)
            logger.info("selecting download")
            time.sleep(5)
            returnvalue=requests.get(driver.current_url).status_code
            if returnvalue!=200:
                logging.info(f"server is not responding please try again")
                print(f"server is not responding please try again")
                sys.exit(0)
            via = Select(WebDriverWait(driver, 180, poll_frequency=1).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "select[name='deliveryType']"))))
            via.select_by_visible_text("DOWNLOAD")
            # a = Select(WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[2]/td/form/p/select[2]"))))
            # a.select_by_visible_text("Comma Delimited (Spreadsheet)")
            logger.info("downloading report via send button")
            WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[value='Send']"))).click()
            if key=="Test_format_trace":
                des_text =  WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "tbody tr p:nth-child(2)"))).text
                des_text = des_text.replace("\n",'')
            driver.close()
            filesToUpload = os.listdir(os.getcwd() + "\\Raw_Files")
            for file in filesToUpload:
                name =key+"."+file.split(".")[-1]
                shutil.move(files_location+"\\"+file,trace_directory+"\\"+name)
                # form[name='frm'] h3 (text = Track and Trace - Track and Trace Error - Running a Trace)
        return des_text
    except Exception as e:
        print(f"Exception caught in tracereport_dwonload method: {e}")
        logging.info(f"Exception caught in tracereport_dwonload method: {e}")
        raise e 

def combine_reports(des_text,key):
    try:
        global comp_list
        comp_list=[]
        send_check = True
        print("inside combine_reports")
        retry=0
        while retry < 10:
            try:
                tr_wb=xw.Book(trace_directory+"\\"+"Test_format_trace.csv")
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e
        retry=0
        while retry < 10:
            try:
                we_wb=xw.Book(trace_directory+"\\"+"weight_format_trace.csv")
                break
            except Exception as e:
                time.sleep(5)
                retry+=1
                if retry ==10:
                    raise e
        we_wb.activate() 
        we_ws1 = we_wb.sheets[0]
        we_ws1.activate() 
        we_ws1.api.Range(f"E:E").EntireColumn.Delete()
        tr_wb.activate()
        tr_ws1 = tr_wb.sheets[0]  
        tr_ws1.activate()  
        tr_ws1.api.Range(f"M:M").EntireColumn.Delete()
        tr_ws1.api.Range(f"H:H").EntireColumn.Delete()
        we_ws1.range(f"C:E").copy(tr_ws1.range(f"L1"))
        we_wb.close()
        tr_ws1.api.Range("1:1").EntireRow.Insert()
        tr_ws1.range("A1").value=des_text
        # tr_ws1.api.Range(f"A2").AutoFilter(Field:=1)
        last_column_letter_plus1=num_to_col_letters(tr_ws1.range('A2').end('right').last_cell.column+1)
        tr_ws1.range(f"{last_column_letter_plus1}2").value = 'Car_no'
        tr_ws1.range(f"{last_column_letter_plus1}3").value = f'=A3&B3'
        last_row = tr_ws1.range(f'A'+ str(tr_ws1.cells.last_cell.row)).end('up').row
        tr_ws1.api.Range(f"{last_column_letter_plus1}3:{last_column_letter_plus1}{last_row}").Select()
        if last_row!=3:
            tr_wb.app.api.Selection.FillDown()
        tr_wb.app.api.Selection.Copy()
        tr_wb.app.api.Selection._PasteSpecial(Paste=-4163)
        tr_wb.app.api.CutCopyMode=False
        #djfnvj
        db_car_nak=tr_ws1.range(f"{last_column_letter_plus1}2:{last_column_letter_plus1}{last_row}").options(pd.DataFrame,header=1,index=False).value  
        tr_wb.save(final_directory+"\\"+f"Trace_Report_{key}_initial.xlsx")
        if os.path.exists(final_directory+"\\"+f"Trace_Report_{key}.xlsx"):
            print(f"there may be a new rail car for - {key}")
            logging.info(f"database present for - {key}")
            retry=0
            while retry < 10:
                try:
                    db_wb=xw.Book(final_directory+"\\"+f"Trace_Report_{key}.xlsx")
                    break
                except Exception as e:
                    time.sleep(5)
                    retry+=1
                    if retry ==10:
                        raise e
            # db_wb.activate()
            db_ws1 = db_wb.sheets[0] 
            car_no_add = db_ws1.range(f"O1").end('down').address.replace("$","")
            car_last_rw = db_ws1.range(f'O'+ str(db_ws1.cells.last_cell.row)).end('up').row
            car_db_last=db_ws1.range(f"{car_no_add}:O{car_last_rw}").options(pd.DataFrame,header=1,index=False).value
            common = pd.merge(db_car_nak, car_db_last, on=['Car_no'], how='inner')
            comp_db =pd.concat([db_car_nak,common]).drop_duplicates(keep=False) 
            comp_list=list(comp_db['Car_no'])
            time.sleep(1)
            db_wb.close()
            time.sleep(1)
        # tr_wb.save(final_directory+"\\"+f"{in_var}_Trace_Report_{key}.xlsx")
        last_rov = tr_ws1.range(f'A'+ str(tr_ws1.cells.last_cell.row)).end('up').row
        custum_sort(tr_wb,tr_ws1,f"D3:D{last_rov}",f"H3:H{last_rov}",f"A2:O{last_rov}")
        state_types = tr_ws1.range(f"D3:D{last_rov}").value
        tr_ws1.api.Range(f"2:2").Font.Bold = True
        pa_count = 0
        pc_count = 0
        diff_count = 0
        check = None
        if state_types is not None and 'CO' in state_types:
            tr_ws1.api.Range(f"D1").AutoFilter(Field:=4, Criteria1:=["CO"])
            flat_list, sp_lst_row,sp_address = row_range_calc(f"D",tr_ws1,tr_wb)
            if type(tr_ws1.range("H3").expand("down").api.SpecialCells(win32c.CellType.xlCellTypeVisible).Value) == tuple:
                event_types = list(tr_ws1.range("H3").expand("down").api.SpecialCells(win32c.CellType.xlCellTypeVisible).Value)
            else:
                event_types = tr_ws1.range("H3").expand("down").api.SpecialCells(win32c.CellType.xlCellTypeVisible).Value
                check = True   
            if (type(event_types)==list and len(event_types)>0 and ('Placed Actual',) in event_types) or 'Placed Actual' in event_types:
                tr_ws1.api.Range(f"D1").AutoFilter(Field:=8, Criteria1:=["Placed Actual"])
                l1, sp_lst_row1,sp_address1 = row_range_calc(f"D",tr_ws1,tr_wb)
                l1.remove(2)
                if len(l1)>0:
                    #other color value =#FFFF00
                    interior_coloring(colour_value="65535",cellrange=f"A{l1[0]}:N{l1[-1]}",working_sheet=tr_ws1,working_workbook=tr_wb)
                else:
                    interior_coloring(colour_value="65535",cellrange=f"A{l1[0]}:N{l1[0]}",working_sheet=tr_ws1,working_workbook=tr_wb)
                pa_count = len(l1) 
                tr_ws1.api.Range(f"D1").AutoFilter(Field:=8)
                if not check:
                    event_types = list(filter(lambda x: x != ('Placed Actual',), event_types))
            if (type(event_types)==list and len(event_types)>0 and ('Placed Construct',) in event_types) or 'Placed Construct' in event_types:
                tr_ws1.api.Range(f"D1").AutoFilter(Field:=8, Criteria1:=["Placed Construct"])
                l2, sp_lst_row2,sp_address2 = row_range_calc(f"D",tr_ws1,tr_wb)
                l2.remove(2)
                if len(l2)>0:
                    #other color value =#00B050
                    interior_coloring(colour_value="5287936",cellrange=f"A{l2[0]}:N{l2[-1]}",working_sheet=tr_ws1,working_workbook=tr_wb)
                else:
                    interior_coloring(colour_value="5287936",cellrange=f"A{l2[0]}:N{l2[0]}",working_sheet=tr_ws1,working_workbook=tr_wb)
                pc_count = len(l2) 
                tr_ws1.api.Range(f"D1").AutoFilter(Field:=8) 
                if not check:
                    event_types = list(filter(lambda x: x != ('Placed Construct',), event_types)) 
            if (type(event_types)==list and len(event_types)>0 and ('Placed Actual',) not in event_types  and ('Placed Construct',) not in event_types) or (type(event_types)==str and ('Placed Actual' not in event_types  and 'Placed Construct' not in event_types)):
                tr_ws1.api.Range(f"D1").AutoFilter(Field:=8, Criteria1:=["<>Placed Actual"],Operator:=1, Criteria2:=['<>Placed Construct'])
                l3, sp_lst_row3,sp_address3 = row_range_calc(f"D",tr_ws1,tr_wb)
                l3.remove(2)
                if len(l3)>0:
                    #other color value =#9FA459
                    interior_coloring(colour_value="5874847",cellrange=f"A{l3[0]}:N{l3[-1]}",working_sheet=tr_ws1,working_workbook=tr_wb)
                else:
                    interior_coloring(colour_value="5874847",cellrange=f"A{l3[0]}:N{l3[0]}",working_sheet=tr_ws1,working_workbook=tr_wb)
                diff_count = len(l3)
                tr_ws1.api.Range(f"D1").AutoFilter(Field:=8)
        tr_ws1.api.Range(f"D1").AutoFilter(Field:=4) 
        color_dict = {65535:"On Hand",5287936:"PC",5874847:"CO"} 
        if diff_count>0:
            tr_ws1.api.Range("2:2").EntireRow.Insert()
            tr_ws1.api.Range("A2").Value = f"{diff_count} CO"
            interior_coloring(colour_value="5874847",cellrange=f"A2",working_sheet=tr_ws1,working_workbook=tr_wb)         
        if pc_count>0:
            tr_ws1.api.Range("2:2").EntireRow.Insert()
            tr_ws1.api.Range("A2").Value = f"{pc_count} PC"
            interior_coloring(colour_value="5287936",cellrange=f"A2",working_sheet=tr_ws1,working_workbook=tr_wb)
        if pa_count>0:
            tr_ws1.api.Range("2:2").EntireRow.Insert()
            tr_ws1.api.Range("A2").Value = f"{pa_count} On Hand"
            interior_coloring(colour_value="65535",cellrange=f"A2",working_sheet=tr_ws1,working_workbook=tr_wb)
        combinedfile = final_directory+"\\"+f"Trace_Report_{key}.xlsx"
        tr_ws1.autofit()
        tr_ws1.api.Columns("A:A").ColumnWidth = 11.14
        time.sleep(1)
        if os.path.exists(combinedfile):
            os.remove(combinedfile)
        
        color_list = []
        #logic for removing other destination cities    
        tr_ws1.api.Cells.Find(What:="Car_no", After:=tr_ws1.api.Application.ActiveCell,LookIn:=win32c.FindLookIn.xlFormulas,LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByRows, SearchDirection:=win32c.SearchDirection.xlNext).Activate()
        bcell_value = tr_ws1.api.Application.ActiveCell.Address.replace("$","")
        brow_value = int(re.findall("\d+",bcell_value)[0])
        lr = tr_ws1.range(f'A'+ str(tr_ws1.cells.last_cell.row)).end('up').row
        des_column_no = tr_ws1.api.Cells.Find(What:="Destination City", After:=tr_ws1.api.Application.ActiveCell,LookIn:=win32c.FindLookIn.xlFormulas,LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByRows, SearchDirection:=win32c.SearchDirection.xlNext).Column
        des_column_letter = num_to_col_letters(des_column_no)
        new_column_letter = num_to_col_letters(des_column_no+1)
        tr_ws1.api.Range(f"{new_column_letter}:{new_column_letter}").EntireColumn.Insert()
        tr_ws1.range(f"K{brow_value}").value = 'Destination Check'
        tr_ws1.range(f"K{brow_value+1}").value = f'=OR({des_column_letter}{brow_value+1}="JOHNSTOWN",{des_column_letter}{brow_value+1}="LOVELAND",{des_column_letter}{brow_value+1}="GREELEY",{des_column_letter}{brow_value+1}="")'
        if brow_value+1 == lr:
            print("single rail car trace condiiton")
            if tr_ws1.range(f"K{brow_value+1}").value == False:
                tr_ws1.api.Range(f"{brow_value+1}:{brow_value+1}").EntireRow.Delete()  
                send_check = False
        else:
            print("multiple cars for this commodity")
            tr_ws1.range(f"{new_column_letter}{brow_value+1}:{new_column_letter}{lr}").api.Select()
            tr_wb.app.api.Selection.FillDown()
            tr_ws1.api.Range(f"{new_column_letter}{brow_value}").AutoFilter(Field:=f"{des_column_no+1}", Criteria1:=["False"],Operator:=1)
            l4, sp_lst_row4,sp_address4 = row_range_calc(f"{new_column_letter}",tr_ws1,tr_wb)
            l4 = [x for x in l4 if x > brow_value]
            if len(l4)>0:
                print("other destination cities found")
                l4.sort(reverse=True)
                for row in l4:
                    if int(tr_ws1.api.Range(f"A{row}").Interior.Color)!=16777215: 
                        color_list.append(int(tr_ws1.api.Range(f"A{row}").Interior.Color))   
                        tr_ws1.api.Range(f"{row}:{row}").EntireRow.Delete()    
                    else:
                        tr_ws1.api.Range(f"{row}:{row}").EntireRow.Delete()       
            tr_ws1.api.AutoFilterMode=False

        if len(color_list)>0:
            new_dict = {}
            for j in Counter(color_list):
                new_dict[color_dict[j]]=Counter(color_list)[j]
            for key2,value in new_dict.items():
                tr_ws1.api.Range(f"A:A").Select()
                value_values=tr_ws1.api.Range(f"A:A").Find(What:=f"{key2}", After:=tr_ws1.api.Application.ActiveCell,LookIn:=win32c.FindLookIn.xlFormulas,LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByRows, SearchDirection:=win32c.SearchDirection.xlNext).Value    
                diff_amount = int(value_values.split(" ")[0]) - new_dict[key2]
                if diff_amount!=0:
                    tr_ws1.api.Cells.Find(What:=f"{key2}", After:=tr_ws1.api.Application.ActiveCell,
                                        LookIn:=win32c.FindLookIn.xlFormulas,LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByRows,
                                            SearchDirection:=win32c.SearchDirection.xlNext).Value = f"{diff_amount} {key2}" 
                else:
                    del_row = tr_ws1.api.Cells.Find(What:=f"{key2}", After:=tr_ws1.api.Application.ActiveCell,
                                        LookIn:=win32c.FindLookIn.xlFormulas,LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByRows,
                                            SearchDirection:=win32c.SearchDirection.xlNext).Row
                    tr_ws1.api.Range(f"{del_row}:{del_row}").EntireRow.Delete()   

        tr_ws1.api.Range(f"{new_column_letter}:{new_column_letter}").EntireColumn.Delete()
        ###logic end ####
        if len(comp_list)>0:
            for car_no in comp_list:
                tr_ws1.activate()
                try:
                    tr_ws1.api.Cells.Find(What:=car_no, After:=tr_ws1.api.Application.ActiveCell,LookIn:=win32c.FindLookIn.xlFormulas,LookAt:=win32c.LookAt.xlPart, SearchOrder:=win32c.SearchOrder.xlByRows, SearchDirection:=win32c.SearchDirection.xlNext).Activate()
                except:
                    continue     
                bcell_value = tr_ws1.api.Application.ActiveCell.Address.replace("$","")
                brow_value = re.findall("\d+",bcell_value)[0]
                interior_coloring(colour_value="255",cellrange=f"A{int(brow_value)}:N{int(brow_value)}",working_sheet=tr_ws1,working_workbook=tr_wb)

                # check = True 
        ###logic for correct trace number###
        lst_no_rw = tr_ws1.range(f'O'+ str(tr_ws1.cells.last_cell.row)).end('up').row
        first_no_rw = tr_ws1.range(f'O'+ str(tr_ws1.cells.last_cell.row)).end('up').end('up').row
        no_of_cars_traced =  lst_no_rw - first_no_rw  
        replace_value =  re.findall("\d+",des_text)[-1]   
        tr_wb.save(combinedfile)
        time.sleep(1)
        tr_wb.app.quit()
        time.sleep(1)
        nl = '<br>'
        if send_check:
            if key == 'Inbound YC Reload HRW':
                subb = 'Inbound YC/Reload HRW'
                mailbody = f"{nl}<strong>Hello</strong> {nl}{nl} Please find attached rail trace sheet for <strong>{subb}</strong> Cars Location.{nl}{nl}Thank you !!!"
                bu_alerts.send_mail(receiver_email = receiver_email,mail_subject =f'{subb} Cars Location {today_date.strftime("%m-%d-%Y")}',mail_body=F"{mailbody}",multiple_attachment_list = [combinedfile])
            else:
                mailbody = f"{nl}<strong>Hello</strong> {nl}{nl} Please find attached rail trace sheet for <strong>{key}</strong> Cars Location.{nl}{nl}Thank you !!!"    
                bu_alerts.send_mail(receiver_email = receiver_email,mail_subject =f'{key} Cars Location {today_date.strftime("%m-%d-%Y")}',mail_body=F"{mailbody}",multiple_attachment_list = [combinedfile])
    except Exception as e:
        print(f"Exception caught in combine_reports method: {e}")
        logging.info(f"Exception caught in combine_reports method: {e}")
        raise e    

def download_wait(directory, nfiles = None):
    try:
        seconds = 0
        dl_wait = True
        while dl_wait and seconds < 90:
            time.sleep(1)
            dl_wait = False
            files = os.listdir(directory)
            if nfiles and len(files) != nfiles:
                dl_wait = True
            for fname in files:
                print(fname)
                if fname.endswith('.crdownload'):
                    dl_wait = True
                elif fname.endswith('.tmp'):
                    dl_wait = True
            seconds += 1
        return seconds
    except Exception as e:
        print(f"Exception caught in download_wait method: {e}")
        logging.info(f"Exception caught in download_wait method: {e}")
        raise e


def combining_one_file(input_sheet,extracted_directory,empty_cars_directory):
        try:
            if 'empty' in input_sheet:
                extracted_directory = empty_cars_directory
            global wb
            retry=0
            while retry < 10:
                try:
                    wb=xw.Book(extracted_directory+"\\"+input_sheet)
                    break
                except Exception as e:
                    time.sleep(5)
                    retry+=1
                    if retry ==10:
                        raise e
            in_var =  input_sheet.split(".")[0] 
            wb.activate()            
            ws1 = wb.sheets[0]  
            ws1.activate() 
            ws1.cells.unmerge()
            last_row = ws1.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
            if "Enroute" in input_sheet:
                values = ws1.range(f'C3:C{last_row}').value
                or_values = [name.replace(" ","") for name in values]
                ws1.range(f"C3").options(transpose=True).value = or_values
                df=ws1.range(f"A3").expand('table').options(pd.DataFrame,header=0,index=False).value
                df= df[[2,4]]
                df.columns = ["Car_No","Commodity"]
            else:
                if 'Total' in ws1.range("A1").value:
                    ws1.api.Range("1:1").EntireRow.Delete()
                ws1.api.Range("1:1").EntireRow.Insert()
                ws1.range("A1").value = f"=A2&A3"  
                last_column_letter=num_to_col_letters(ws1.range('A2').end('right').last_cell.column)
                ws1.range(f"A1").api.Select()
                wb.app.api.Selection.AutoFill(Destination:=ws1.api.Range(f"A1:{last_column_letter}1"),Type:=win32c.AutoFillType.xlFillDefault)
                ws1.api.Rows("1:1").Select()
                wb.app.api.Selection.Copy()
                wb.app.api.Selection.PasteSpecial(Paste:=-4163, Operation:=-4142, SkipBlanks:=False, Transpose:=False)
                ws1.api.Range("2:3").EntireRow.Delete()
                ws1.api.Range("F:F").EntireColumn.Insert()
                ws1.range('F2').value = f"=D2&E2"
                last_row = ws1.range(f'A'+ str(ws1.cells.last_cell.row)).end('up').row
                ws1.range('F1').value = f"Com_Car_No"
                ws1.range(f"F2:F{last_row}").api.Select()
                wb.app.api.Selection.FillDown()
                df=ws1.range(f"A2").expand('table').options(pd.DataFrame,header=0,index=False).value
                df= df[[5,8]]
                df.columns = ["Car_No","Commodity"]
            if 'empty' in input_sheet:
                df = df[(df['Commodity'] == 'WHEAT') | (df['Commodity'] == 'CORN')]
            wb.app.quit() 
            return df  
        except Exception as e:
            print(f"Exception caught in combining_one_file method: {e}")
            logging.info(f"Exception caught in combining_one_file method: {e}")
            raise e         


def processing_excel(dfs):
    try:
        TRUE_UP_DF = pd.read_excel(test_sheet)
        TRUE_UP_index_dict = {}
        for i,x in TRUE_UP_DF.iterrows():
                TRUE_UP_index_dict.setdefault(TRUE_UP_DF[TRUE_UP_DF.columns[0]][i], []).append(TRUE_UP_DF[TRUE_UP_DF.columns[1]][i])
        for key,value in TRUE_UP_index_dict.items():
            print(f"commodity - {value}")
            # ws1.activate() 
            fil_dfs = dfs[dfs['Commodity'].isin(value)] 
            if key == 'Inbound YC Reload HRW':
                print("Inbound YC Reload HRW found")
                inbound_sheet = os.getcwd() + "\\inbound yc reload hrw\\Inbound YC Reload HRW.xlsx"
                fil_dfs = pd.read_excel(inbound_sheet)
            if len(fil_dfs)>0:
                fil_dfs['Car_No'].to_clipboard(index=False,header=None)
            else:
                print("no values found to filter")
                continue
            logger.info("copying and pasting car numbers")
            driver.switch_to.window(driver.window_handles[0])
            WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "textarea[name='carlist']"))).clear()
            WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "textarea[name='carlist']"))).click() 
            WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "textarea[name='carlist']"))).send_keys(Keys.CONTROL, "v")
            des_text = tracereport_dwonload()
            combine_reports(des_text,key)

        print("opened") 
        time.sleep(1)    
    except Exception as e:
        print(f"Exception caught in processing_excel method: {e}")
        logging.info(f"Exception caught in processing_excel method: {e}")
        raise e  


def login_and_download():  
    '''This function downloads log in to the website'''
    try:
        logging.info('Accesing website')
        driver.get(f"{source_url}")
        time.sleep(5)  
        logging.info('providing id and passwords')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, "txtUserName"))).send_keys(username)
        time.sleep(1)
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, "txtPassword"))).send_keys(password)
        time.sleep(1)
        logging.info('click on Login Button')
        WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, "btnLogin"))).click()
        time.sleep(5)
        dict1={"Enroute":['main_lblenrouteload','main_lblenrouteempty'],"Inbound":['main_lblinboundload','main_lblinboundempty'],"Onhand":['main_lblonhandload','main_lblonhandempty']}
        for key, value in dict1.items():
            car_no =int(WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, value[0]))).text)
            empty_car =int(WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, value[1]))).text)
            if car_no>0:
                time.sleep(1)
                WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, value[0]))).click() 
                time.sleep(1)
                WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, "main_btnExport"))).click()
                driver.back()
                time.sleep(1)
                filesToUpload = os.listdir(os.getcwd() + "\\Raw_Files")
                for file in filesToUpload:
                    name =key+"."+file.split(".")[-1]
                    shutil.move(files_location+"\\"+file,extracted_directory+"\\"+name)
            else:
                logging.info(f"No Loaded railcars for {key}")
            
            if empty_car>0:
                time.sleep(1)
                WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, value[1]))).click() 
                time.sleep(1)
                select_via = Select(WebDriverWait(driver, 180, poll_frequency=1).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#main_ddlLE"))))
                select_via.select_by_visible_text("E")
                logger.info("selecting empty loads via drop down menu")
                time.sleep(1)
                WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, "main_btnSearch"))).click()
                time.sleep(1)
                WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, "main_btnExport"))).click()
                driver.back()
                time.sleep(1)
                filesToUpload = os.listdir(os.getcwd() + "\\Raw_Files")
                for file in filesToUpload:
                    name ="empty"+key+"."+file.split(".")[-1]
                    shutil.move(files_location+"\\"+file,empty_cars_directory+"\\"+name)  
            else:
                logging.info(f"No empty railcars for {key}")           

    except Exception as e:
        print(f"Exception caught in login_and_download method: {e}")
        logging.info(f"Exception caught in login_and_download method: {e}")
        raise e

def login_to_steelroads():  
    '''This function downloads log in to the website'''
    try:
        logging.info('Accesing website')
        received_response = requests.get(steel_roads) 
        if received_response.status_code==200:
            driver.get(f"{steel_roads}")
            time.sleep(5)  
            logging.info('Track and trace window')
            WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.LINK_TEXT, "Track and Trace - Trace your railcar shipments"))).click()
            time.sleep(5)
            try:
                logging.info('Providing id and passwords')
                WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, "okta-signin-username"))).send_keys(steel_username)
                time.sleep(1)
                WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, "okta-signin-password"))).send_keys(steel_password)
                time.sleep(1)
                logging.info('click on Sign In')
                WebDriverWait(driver, 90, poll_frequency=1).until(EC.element_to_be_clickable((By.ID, "okta-signin-submit"))).click()
                time.sleep(5)
            except:
                print("login successfull")
   
        else:
            logging.info(f"vpn issues or server is not available")
            print("vpn issues or server is not available")
            sys.exit()

    except Exception as e:
        print(f"Exception caught in login_to_steelroads method: {e}")
        logging.info(f"Exception caught in login_to_steelroads method: {e}")
        raise e

def movefiles(final_directory):
    try:
        source_folder = final_directory
        destination_folder = os.getcwd() + "\\database_old"

        # Get all the files in the source folder
        files = os.listdir(source_folder)

        # Iterate through the files and copy them to the destination folder
        for file_name in files:
            source_file = os.path.join(source_folder, file_name)
            destination_file = os.path.join(destination_folder, file_name)
            shutil.copy2(source_file, destination_file)
        return destination_folder   
    except Exception as e:
        logging.info(f"copy pasting database failed")
        print(f"copy pasting database failed")
        raise e
    
def main():
    try:
        no_of_rows=0
        remove_existing_files(files_location)
        remove_existing_files(extracted_directory)
        remove_existing_files(trace_directory)
        remove_existing_files(empty_cars_directory)
        destination_folder  = movefiles(final_directory)
        # no_of_rows=
        login_and_download()
        dfs=pd.DataFrame()
        directory_list = [extracted_directory,empty_cars_directory]
        for directory in directory_list:
            for files in os.listdir(directory):
                try:
                    df = combining_one_file(files,extracted_directory,empty_cars_directory)
                    dfs = pd.concat([df,dfs])
                except Exception as e:
                    logging.exception(str(e))
                    raise e
        dfs = dfs.reset_index(drop=True)    
        login_to_steelroads()    
        processing_excel(dfs)    
        locations_list.append(logfile)
        try:
            driver.quit()
        except:
            pass
        remove_existing_files(destination_folder)
        log_json='[{"JOB_ID": "'+str(job_id)+'","CURRENT_DATETIME": "'+str(datetime.now())+'"}]'
        bu_alerts.bulog(process_name= processname,database=database,status='COMPLETED',table_name=table_name,
            row_count=no_of_rows, log=log_json, warehouse=warehouse,process_owner=process_owner)  
        bu_alerts.send_mail(receiver_email = receiver_email,mail_subject =f'JOB SUCCESS - {job_name}',mail_body = f'{job_name} completed successfully, Attached Logs',attachment_location = logfile)
    except Exception as e:
        log_json='[{"JOB_ID": "'+str(job_id)+'","CURRENT_DATETIME": "'+str(datetime.now())+'"}]'
        bu_alerts.bulog(process_name= processname,database=database,status='Failed',table_name=table_name,
            row_count=no_of_rows, log=log_json, warehouse=warehouse,process_owner=process_owner)
        logging.exception(str(e))
        bu_alerts.send_mail(receiver_email = receiver_email,mail_subject =f'JOB FAILED -{job_name}',mail_body = f'{job_name} failed, Attached logs',attachment_location = logfile)


if __name__ == "__main__": 
    
    logging.info("Execution Started")
    time_start=time.time()
    #Global VARIABLES
    locations_list=[]
    body = ''
    dict3={}
    today_date=date.today()
    # log progress --
    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)
    logfile = os.getcwd() + '\\' + 'logs' + '\\' + 'Rail_Car_Log_{}.txt'.format(str(today_date))
    logging.basicConfig(
        level=logging.INFO, 
        format='%(asctime)s [%(levelname)s] - %(message)s',
        filename=logfile)
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    directories_created=["Raw_Files","Logs","Renamed Files","Trace_report","Empty_Rail_Cars"]
    for directory in directories_created:
        path3 = os.path.join(os.getcwd(),directory)  
        try:
            os.makedirs(path3, exist_ok = True)
            print("Directory '%s' created successfully" % directory)
        except OSError as error:
            print("Directory '%s' can not be created" % directory)       
    files_location=os.getcwd() + "\\Raw_Files"
    filesToUpload = os.listdir(os.getcwd() + "\\Raw_Files")
    extracted_directory=os.getcwd() + "\\Renamed Files"
    trace_directory=os.getcwd() + "\\Trace_report"
    final_directory=os.getcwd() + "\\final_report"
    empty_cars_directory=os.getcwd() + "\\Empty_Rail_Cars"
    logging.info('setting paTH TO download')
    path = os.getcwd() + '\\Raw_Files'
    logging.info('SETTING PROFILE SETTINGS FOR FIREFOX')
    profile_path = os.getcwd()+f"\\customProfile"
    # profile = webdriver.FirefoxProfile(profile_directory=profile_path)
    profile = webdriver.FirefoxProfile(profile_path)
    profile.set_preference('browser.download.folderList', 2)
    profile.set_preference('browser.download.dir', path)
    profile.set_preference('browser.download.useDownloadDir', True)
    profile.set_preference('browser.download.viewableInternally.enabledTypes', "")
    profile.set_preference('browser.helperApps.neverAsk.saveToDisk','Portable Document Format (PDF), application/pdf')
    profile.set_preference('pdfjs.disabled', True)
    logging.info('Adding firefox profile')
    test_sheet = os.getcwd() +"\\Car_type_Mapping"+ f'\\mapping details.xlsx'
    current_yr=today_date.year
    current_month=today_date.strftime("%m")
    job_id=np.random.randint(1000000,9999999)

    # Getting credential using bu_config
    credential_dict = buconfig.get_config('RAIL_CAR_AUTOMATION', 'N', other_vert=True)
    receiver_email = credential_dict['EMAIL_LIST']
    job_name = credential_dict['PROJECT_NAME']
    table_name = credential_dict['TABLE_NAME']
    process_owner = credential_dict['IT_OWNER']
    username =  credential_dict["USERNAME"].split(';')[0]
    password = credential_dict["PASSWORD"].split(';')[0]
    steel_username = credential_dict["USERNAME"].split(';')[1]
    steel_password = credential_dict["PASSWORD"].split(';')[1]
    source_url = credential_dict['SOURCE_URL'].split(';')[0]
    steel_roads = credential_dict['SOURCE_URL'].split(';')[1]
    database = credential_dict['DATABASE'].split(";")[0]
    warehouse = credential_dict['DATABASE'].split(";")[1]
    processname = credential_dict['PROJECT_NAME']
    # schema = credential_dict['TABLE_SCHEMA']
    #####################Uncomment for Test############################
    # processname = "RAIL_CAR_AUTOMATION"
    # process_owner = 'Yash Jain'
    # source_url= 'https://www.railconnect.com'
    # steel_roads= 'https://steelroads.railinc.com/index.jsp'
    # steel_username = 'WPJTOWN1'
    # steel_password = 'Wheat010'
    # username= 'gwrwpnt'
    # password = 'Wheat02'
    receiver_email='yashn.jain@biourja.com,ramm@westplainsllc.com'
    # receiver_email='yashn.jain@biourja.com,ramm@westplainsllc.com,bharat.pathak@biourja.com'
    # # check= None
    # #snowflake variables
    # database = ''
    # # Database = "POWERDB_DEV"
    # schema = '' 
    # table_name = ''
    ##################################################################
    

    # BU_LOG entry(started) in PROCESS_LOG table
    log_json = '[{"JOB_ID": "'+str(job_id)+'","JOB_NAME": "'+str(job_name)+'","CURRENT_DATETIME": "'+str(datetime.now())+'","STATUS": "STARTED"}]'
    bu_alerts.bulog(process_name=job_name,table_name=table_name,status='STARTED',process_owner=process_owner ,row_count=0,log=log_json,database=database,warehouse=warehouse)

    app = customtkinter.CTk()  # create CTk window like you do with the Tk window
    app.title("Biourja Renewables")
    app["bg"]= "#e2e1ef"
    biourjaLogo = resource_path('biourjaLogo.png')
    photo = tkinter.PhotoImage(file = biourjaLogo)
    app.iconphoto(False, photo)
    screen_width = app.winfo_screenwidth()
    screen_height = app.winfo_screenheight()
    width2 = 420
    height2 = 190
    x2 = (screen_width/2) - (width2/2)
    y2 = (screen_height/2) - (height2/2)
    app.geometry('%dx%d+%d+%d' % (width2, height2, x2, y2))
    settings_frame = customtkinter.CTkFrame(app, width=50)
    settings_frame.pack(fill=tkinter.X, side=tkinter.TOP, padx=2, pady=2)
    settings_frame.grid_columnconfigure(0, weight=1)
    settings_frame.grid_rowconfigure(3, weight=1)    

    button_text=tkinter.StringVar()
    #text_font=("SF Display",-13))
    button = customtkinter.CTkButton(master=app, textvariable=button_text, command=button_function,width=160,height=36)
    button_text.set("Generate Trace Report")
    button.place(relx=0.5, rely=0.5, anchor=tkinter.CENTER)
    app.protocol("WM_DELETE_WINDOW", on_closing)
    Tk.report_callback_exception = report_callback_exception 
    options = Options()
    options.headless=False
    options.profile = profile
    driver=webdriver.Firefox(executable_path=GeckoDriverManager().install(),options=options)
    app.mainloop()
    # main()
    time_end=time.time()
    logging.info(f'It takes {time_start-time_end} seconds to run')
    
        
