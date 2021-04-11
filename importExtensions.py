from tkinter import filedialog, Tk, Frame, Label, Button, messagebox
import os
from urllib3 import PoolManager

def ExtensionUpdater():
    extnURL = 'https://raw.githubusercontent.com/bhuvannarula/offlineGST-Extensions/master/importExtensions.py'
    try:
        browser2 = PoolManager()
        respupdate = browser2.urlopen(
            'GET', extnURL).data.decode('utf-8')
        if __name__ == '__main__':
            script_path = os.getcwd()
        else:
            script_path = os.getcwd() + '/Extras'
        scriptfilein = open(script_path + '/importExtensions.py', 'r+')
        scriptfileindata = scriptfilein.read()
        if scriptfileindata != respupdate:
            scriptfilein.seek(0)
            scriptfilein.truncate()
            scriptfilein.write(respupdate)
            scriptfilein.close()
            return True
        else:
            scriptfilein.close()
            return False
    except:
        return False

def ExtensionManager(companyGSTINhashed):
    extension_names = ['ext1']
    for extn in extension_names:
        if eval('check_' + extn)(companyGSTINhashed) == True:
            break
    else:
        return False
    return True

def ExtensionExecuter(companyGSTINhashed, cName, sMonth, sale = True):
    extension_names = ['ext1']
    for extn in extension_names:
        if eval('check_' + extn)(companyGSTINhashed) == True:
            break
    else:
        return False
    selected_extn = str(extn)
    eval('execute_' + selected_extn)(cName, sMonth, sale)


# Extension 1 - For Hashed GSTIN : e138392fe8986ff58008bd7e4a62487d6d09f5a001645ab8fa6655266aeef774
def check_ext1(companyGSTINhashed):
    '''
    Function checks if the extension is applicable for selected company, if it is then returns True
    if True is returned, then only import button will be placed.
    '''
    if companyGSTINhashed == 'e138392fe8986ff58008bd7e4a62487d6d09f5a001645ab8fa6655266aeef774':
        return True
    else:
        return False

def execute_ext1(cName, sMonth, sale):
    import openpyxl, csv, os
    
    root = Tk(className=' {} Import'.format(cName))

    frame0 = Frame(root, height = 200, width = 300)
    frame0.pack()

    label1 = Label(frame0, text='{}\n\nImport Files\n\nSelect files for {} month:\n\n'.format(cName, sMonth))
    label1.place(x = 50, y = 20)

    def askfiles():
        filess = filedialog.askopenfilenames()
        if not filess:
            return None
        for item in filess:
            if item.split('.')[-1].lower() not in 'xlsx':
                messagebox.showerror('Error Occured!', 'File(s) selected is/are not Excel file(s). Select again.')
                if not askfiles():
                    return None
        else:
            def startprocessing():
                processing(filess)
                root.destroy()
                root.quit()
                return True
            root.after(100, startprocessing)
        
    button1 = Button(frame0, text='Select Files', command = askfiles)
    button1.place(x = 50, y = 150)
    
    def processing(filess):
        # step 0 : copy the files
        # or maybe not
        # TODO leaving total invoice value blank rn, after all bills, read CSV from start, sort list of rows using invoice no., then find invoice value
        
        stcode = {'35': '35-Andaman and Nicobar Islands', '37': '37-Andhra Pradesh', '12': '12-Arunachal Pradesh', '18': '18-Assam', '10': '10-Bihar', 
                  '04': '04-Chandigarh', '22': '22-Chattisgarh', '26': '26-Dadra and Nagar Haveli', '25': '25-Daman and Diu', '07': '07-Delhi', 
                  '30': '30-Goa', '24': '24-Gujarat', '06': '06-Haryana', '02': '02-Himachal Pradesh', '01': '01-Jammu and Kashmir', 
                  '20': '20-Jharkhand', '29': '29-Karnataka', '32': '32-Kerala', '31': '31-Lakshadweep Islands', '23': '23-Madhya Pradesh', 
                  '27': '27-Maharashtra', '14': '14-Manipur', '17': '17-Meghalaya', '15': '15-Mizoram', '13': '13-Nagaland', '21': '21-Odisha', 
                  '34': '34-Pondicherry', '03': '03-Punjab', '08': '08-Rajasthan', '11': '11-Sikkim', '33': '33-Tamil Nadu', '36': '36-Telangana', 
                  '16': '16-Tripura', '09': '09-Uttar Pradesh', '05': '05-Uttarakhand', '19': '19-West Bengal'}
        newfiles = []
        # step 1 : convert xls to xlsx
        for item in filess:
            filenam = item[::-1].split('.', maxsplit = 1)
            newfilename = filenam[1][::-1] + '.xlsx'
            os.rename(item, newfilename)
            newfiles.append(newfilename)
        
        # step 2 : create CSV file to write data in
        if __name__ == '__main__':
            parentdir = os.getcwd()[::-1].split('/', maxsplit= 1)[1][::-1] # without '/' at end
        else:
            parentdir = os.getcwd()
        csvFileOut = open(parentdir + '/companies/{}/{}/GSTR{}.csv'.format(cName, sMonth, '1' if sale else '2'), 'w+', newline='')
        tempWriter = csv.writer(csvFileOut)
        headerRow = ['GSTIN', 'Supplier Name', 'Invoice Number', 'Invoice Date', 'Invoice Value',
                     'Place Of Supply', 'Invoice Type', 'Rate', 'Taxable Amount', 'Cess Amount']
        tempWriter.writerow(headerRow)
        csvFileOut.flush()
        # CSV Writer is now available to use
        
        # step 3 : open excel file one-by-one using openpyxl module
        rows = []
        mode = 'Sales' if sale else 'Purchase'
        for item in newfiles:
            current_wb = openpyxl.load_workbook(item)
            cur_sheet = current_wb.active
            C7value = cur_sheet['C7'].value
            taxRate = int(float(C7value[-3:-1]))
            unregPOS = '07' if C7value[0] == 'I' else '06'
            ii, hh = 0, 0 # hh = 0 if invoices haven't started, 1 if invoices going on

            while True:
                ii += 1
                # checking invoice status
                temp11 = cur_sheet['D' + str(ii)].value
                if temp11 != mode:
                    if hh == 1:
                        break
                    continue
                elif temp11 == mode:
                    hh = 1
                # now making row from excel for csv file
                temprow = []
                tgstin = cur_sheet['C' + str(ii)].value
                if tgstin == '':
                    tgstin = unregPOS
                    pname = ''
                else:
                    pname = cur_sheet['B' + str(ii)].value
                tempdate0 = cur_sheet['A' + str(ii)].value
                tempdate = tempdate0.strftime("%d/%m/%Y")
                temprow = [
                    tgstin,
                    pname,
                    cur_sheet[('E' if sale else 'F') + str(ii)].value,
                    tempdate,
                    '',
                    stcode[tgstin[:2]],
                    'Regular',
                    taxRate,
                    round(float(cur_sheet[('F' if sale else 'H') + str(ii)].value),2),
                    0
                ]
                rows.append(temprow)
            current_wb.close()
        rows.sort(key = lambda var : var[2])
        
        ih = 0
        while ih < len(rows):
            invnum = rows[ih][2]
            ihst = int(ih)
            totinv = rows[ih][8] * (1 + rows[ih][7]/100)
            while ih < len(rows) - 1:
                ih += 1
                if rows[ih][2] == invnum:
                    totinv += rows[ih][8] * (1 + rows[ih][7]/100)
                else:
                    ihed = int(ih)
                    break
            else:
                ih += 1
            for ik in range(ihst, ihed + 1):
                rows[ik][4] = round(totinv, 2)
        
        tempWriter.writerows(rows)
        csvFileOut.flush()
                    
    root.mainloop()
        