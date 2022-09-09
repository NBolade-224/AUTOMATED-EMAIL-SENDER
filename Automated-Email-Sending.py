

from tkinter import *
from tkinter import filedialog
from pathlib import Path
import win32com.client as win32
import openpyxl, sys, time
MainWindow=Tk()
MainWindow.title("Sending Emails")
MainWindow.geometry("600x500+600+200")
MainWindow.configure(bg="RoyalBlue4")
def NESTSelected():
    NESTWindow=Toplevel(MainWindow)
    NESTWindow.title("Sending emails")
    NESTWindow.geometry("900x800+500+100")
    NESTWindow.configure(bg="RoyalBlue4")
    NESTWindow.attributes('-topmost',1)
    global NESTSendSafety1
    global NESTSendSafety2
    NESTSendSafety1 = True
    NESTSendSafety2 = True
    def NestExcelFile():
        global NESTExcelFIle
        global ExcelFileNEST
        global ExcelSheetNEST
        global AddressListNEST
        global TotalEmailAdressesNEST
        global NESTSendSafety1
        button_exploreExcelNEST.configure(bg='red3')
        button_exploreDOCxNEST.configure(bg='red3')
        button_sendNEST.configure(bg='red3')
        NESTSendSafety1 = True
        labelERRORsfoRNEST.configure(text='')
        AddressListNEST = []
        NESTExcelFIle = filedialog.askopenfilename(parent=NESTWindow, initialdir = "",title = "Select file",filetypes = (("Excel","*xlsx"),("all files","*.*")))
        NESTExcelFIle0 = Path(NESTExcelFIle).stem 
        labelfileopnedNEST.configure(text="File Opened: "+NESTExcelFIle0)
        ExcelFileNEST = openpyxl.load_workbook(NESTExcelFIle)
        ExcelSheetNEST = ExcelFileNEST.active
        for cell in ExcelSheetNEST['A']:
            if cell.value != None and '@' in cell.value:
                AddressListNEST.append(cell.value)
        TotalEmailAdressesNEST = len(AddressListNEST)
        if TotalEmailAdressesNEST > 0:
            pass
        else:
            labelERRORsfoRNEST.configure(text='ERROR: No email address found in Col "N"')
            NESTWindow.update()
            NESTSendSafety1 = True
            return
        NESTWindo1 = Toplevel(NESTWindow)
        NESTWindo1.title('All PDF files in folder, please select another folder if wrong')
        NESTWindo1.geometry("900x700+500+200") 
        NESTWindo1.configure(bg="RoyalBlue4")
        NESTWindo1.attributes('-topmost',1)
        ListofEmailAddressListBoxNEST = Listbox(NESTWindo1, bg="royalblue1",width=80, height=31, selectmode='single', font=('Times', 14))
        ListofEmailAddressListBoxNEST.pack()
        ListofEmailAddressListBoxNEST.insert(END,'All Email Addressses in selected excel file, please choose another file if wrong')
        ListofEmailAddressListBoxNEST.insert(END,'')
        ListofEmailAddressListBoxNEST.insert(END,'Total emails to send to: '+str(TotalEmailAdressesNEST))
        ListofEmailAddressListBoxNEST.insert(END,'')
        for xNEST in AddressListNEST:
            ListofEmailAddressListBoxNEST.insert(END,xNEST)
        labelsentcounterNEST.configure(text="Total Email Addresses to send to: "+str(TotalEmailAdressesNEST))
        NESTSendSafety1 = False
        button_exploreExcelNEST.configure(bg='green4')
    def SelectNESTDocx():
        global NESTSendSafety2
        global enrollmentletterNEST
        button_exploreDOCxNEST.configure(bg='red3')
        button_sendNEST.configure(bg='red3')
        NESTSendSafety2 = True
        enrollmentletterNEST = ''
        labelERRORsfoRNEST.configure(text='')
        if NESTSendSafety1 == True:
            labelERRORsfoRNEST.configure(text='Please select the excel report first')
            button_exploreDOCxNEST.configure(bg='red3')
            return
        try:
            enrollmentletterNEST = filedialog.askopenfilename(parent=NESTWindow, initialdir = "",title = "Select file",filetypes = (("all files","*.*"),))
            filename0NEST = Path(enrollmentletterNEST).stem 
            labelfileopnedDocxNEST.configure(text="File To Be Sent: "+filename0NEST)
        except:
            labelERRORsfoRNEST.configure(text='Please select a valid DOCx file')
            return
        if len(enrollmentletterNEST) == 0:
            NESTSendSafety2 = True
            button_exploreDOCxNEST.configure(bg='red3')
            labelERRORsfoRNEST.configure(text='Please select a valid DOCx file')
            return
        else:
            NESTSendSafety2 = False
            button_exploreDOCxNEST.configure(bg='green4')
            button_sendNEST.configure(bg='blue')
    def SendNEST():
        global NESTSendSafety1
        global NESTSendSafety2
        if NESTSendSafety1 == True or NESTSendSafety2 == True:
            labelERRORsfoRNEST.configure(text='Please select Excel file and DOCx before sending')
            NESTWindow.update()
            return
        else:         
            global TotalEmailAdressesNEST
            global nestlettersent
            global enrollmentletterNEST
            nestlettersent = 0
            for NESTEmails in AddressListNEST:
                outlook = win32.Dispatch('outlook.application')
                mail = outlook.CreateItem(0)
                mail.To = NESTEmails
                mail.Subject = ''
                mail.Body = ''
                mail.Attachments.Add(enrollmentletterNEST)
                mail.Send()
                nestlettersent += 1
                labelsentcounterNEST.configure(text = "Amount of files sent: %d/%d" % (nestlettersent,TotalEmailAdressesNEST))
                time.sleep(1)
                NESTWindow.update()
            labelsentcounterNEST.configure(text = "All emails sent %d/%d" % (nestlettersent,TotalEmailAdressesNEST))
            NESTSendSafety1 = 'On'
            NESTSendSafety2 = 'On'
            button_exploreExcelNEST.configure(bg='red3')
            button_exploreDOCxNEST.configure(bg='red3')
            button_sendNEST.configure(bg='red3')
            NESTWindow.update()
    label_file_explorerNEST = Label(NESTWindow,text = "Email Sender - By Nick",width = 100, height = 4,fg = "white",bg = 'RoyalBlue4', font=('Times', 13))
    button_exploreExcelNEST = Button(NESTWindow,text = "Select Excel File for Sending",bg = 'red3',width = 30, height = 2,command = NestExcelFile, font=('Times', 15, 'bold'), fg = "yellow2")
    button_exploreDOCxNEST = Button(NESTWindow,text = "Select file to be Sent",bg = 'red3',width = 30, height = 2,command = SelectNESTDocx, font=('Times', 15, 'bold'), fg = "yellow2")
    button_exitNEST = Button(NESTWindow,text = "Exit",bg = 'snow4',width = 30,height = 2,command = sys.exit, font=('Times', 15, 'bold'), fg = "black")
    button_sendNEST = Button(NESTWindow,text = "Send",bg = 'red3',width = 30,height = 2,command = SendNEST, font=('Times', 15, 'bold'), fg = "yellow2")
    labelERRORsfoRNEST = Label(NESTWindow,text = "",width = 75, height = 2,fg = "red",bg = 'RoyalBlue4', font=('Times', 16))
    labelfileopnedNEST = Label(NESTWindow,text = "",width = 75, height = 2,fg = "white",bg = 'RoyalBlue4', font=('Times', 16))
    labelsentcounterNEST = Label(NESTWindow,text = "",width = 75, height = 2,fg = "white",bg = 'RoyalBlue4', font=('Times', 16))
    labelfileopnedDocxNEST = Label(NESTWindow,text = "",width = 75, height = 2,fg = "white",bg = 'RoyalBlue4', font=('Times', 16))
    LabelSpace1NEST = Label(NESTWindow,text = "",width = 75, height = 1,fg = "red",bg = 'RoyalBlue4')
    LabelSpace2NEST = Label(NESTWindow,text = "",width = 75, height = 1,fg = "red",bg = 'RoyalBlue4')
    LabelSpace3NEST = Label(NESTWindow,text = "",width = 75, height = 1,fg = "red",bg = 'RoyalBlue4')
    def allpacks():
        label_file_explorerNEST.pack()
        button_exploreExcelNEST.pack()
        LabelSpace1NEST.pack()
        button_exploreDOCxNEST.pack()
        LabelSpace2NEST.pack()
        button_exitNEST.pack()
        LabelSpace3NEST.pack()
        button_sendNEST.pack()
        labelERRORsfoRNEST.pack()
        labelfileopnedNEST.pack()
        labelfileopnedDocxNEST.pack()
        labelsentcounterNEST.pack()
    allpacks()
    NESTWindow.mainloop()   
Labelfill = Label(MainWindow, width = 30, height = 3,bg = 'RoyalBlue4')
Labelfill1 = Label(MainWindow, width = 30, height = 3,bg = 'RoyalBlue4')
NESTWindowButton = Button(MainWindow,text = "Send files",bg = 'RoyalBlue1',width = 30,height = 2, font=('Times', 15, 'bold'), fg = "yellow2",command = NESTSelected) #command = P45Selected
Labelfill.pack()
Labelfill1.pack()
NESTWindowButton.pack()
MainWindow.mainloop()

