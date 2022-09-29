from tkinter import *
from tkinter import filedialog
from pathlib import Path
import win32com.client as win32
import openpyxl, sys, time


## Main Window Properties

MainWindow=Tk()
MainWindow.title("Sending Emails")
MainWindow.geometry("600x500+600+200")
MainWindow.configure(bg="RoyalBlue4")


## Send File Button

def SendFilesButton():


    ## First Sub Window Properties

    SubWindow1=Toplevel(MainWindow)
    SubWindow1.title("SubWindow1")
    SubWindow1.geometry("900x800+500+100")
    SubWindow1.configure(bg="RoyalBlue4")
    SubWindow1.attributes('-topmost',1)


    ## Safety measures to prevent emails being sent before an excel file & attachment file is selected (will be set to False in later functions)

    global ExcelFileSelectedSafety
    global EmailAttachmentSelectedSafety
    ExcelFileSelectedSafety = True
    EmailAttachmentSelectedSafety = True


    ## Select Excel File Button

    def SelectedExcelFileButton():


        ## Global Variables set so that they can be accessed from other functions

        global SelectedExcelFile
        global AddressesFromExcelFile
        global TotalAddresses
        global ExcelFileSelectedSafety


        ## SButton Properties (changes color of buttons depending on whether a file was selected)

        Button_Explorer_Excel.configure(bg='red3')
        button_exploreDOCx.configure(bg='red3')
        button_send.configure(bg='red3')
        ExcelFileSelectedSafety = True
        labelERRORs.configure(text='')


        ## List for Files

        AddressesFromExcelFile = []


        ## Excel Properties

        SelectedExcelFile = filedialog.askopenfilename(parent=SubWindow1, initialdir = "C:\\Users\\nickb\Desktop\\Test Email Sending",title = "Select file",filetypes = (("Excel","*xlsx"),("all files","*.*")))
        SelectedExcelFile0 = Path(SelectedExcelFile).stem 
        labelfileopned.configure(text="File Opened: "+SelectedExcelFile0)
        ExcelFile = openpyxl.load_workbook(SelectedExcelFile)
        ExcelSheet = ExcelFile.active


        ## Listing Email Addresses from Col A
        
        for cell in ExcelSheet['A']:
            if cell.value != None and '@' in cell.value:
                AddressesFromExcelFile.append(cell.value)
        TotalAddresses = len(AddressesFromExcelFile)
        if TotalAddresses > 0:
            pass
        else:
            labelERRORs.configure(text='ERROR: No email address found in Col A')
            SubWindow1.update()
            ExcelFileSelectedSafety = True
            return
        

        NESTWindo1 = Toplevel(SubWindow1)
        NESTWindo1.title('All PDF files in folder, please select another folder if wrong')
        NESTWindo1.geometry("900x700+500+200") 
        NESTWindo1.configure(bg="RoyalBlue4")
        NESTWindo1.attributes('-topmost',1)
        ListofEmailAddressListBoxNEST = Listbox(NESTWindo1, bg="royalblue1",width=80, height=31, selectmode='single', font=('Times', 14))
        ListofEmailAddressListBoxNEST.pack()
        ListofEmailAddressListBoxNEST.insert(END,'All Email Addressses in selected excel file, please choose another file if wrong')
        ListofEmailAddressListBoxNEST.insert(END,'')
        ListofEmailAddressListBoxNEST.insert(END,'Total emails to send to: '+str(TotalAddresses))
        ListofEmailAddressListBoxNEST.insert(END,'')
        for xNEST in AddressesFromExcelFile:
            ListofEmailAddressListBoxNEST.insert(END,xNEST)
        labelsentcounterNEST.configure(text="Total Email Addresses to send to: "+str(TotalAddresses))
        ExcelFileSelectedSafety = False
        Button_Explorer_Excel.configure(bg='green4')


    def SelectNESTDocx():
        global EmailAttachmentSelectedSafety
        global enrollmentletterNEST
        button_exploreDOCx.configure(bg='red3')
        button_send.configure(bg='red3')
        EmailAttachmentSelectedSafety = True
        enrollmentletterNEST = ''
        labelERRORs.configure(text='')
        if ExcelFileSelectedSafety == True:
            labelERRORs.configure(text='Please select the excel report first')
            button_exploreDOCx.configure(bg='red3')
            return
        try:
            enrollmentletterNEST = filedialog.askopenfilename(parent=SubWindow1, initialdir = "",title = "Select file",filetypes = (("all files","*.*"),))
            filename0NEST = Path(enrollmentletterNEST).stem 
            labelfileopnedDocxNEST.configure(text="File To Be Sent: "+filename0NEST)
        except:
            labelERRORs.configure(text='Please select a valid DOCx file')
            return
        if len(enrollmentletterNEST) == 0:
            EmailAttachmentSelectedSafety = True
            button_exploreDOCx.configure(bg='red3')
            labelERRORs.configure(text='Please select a valid DOCx file')
            return
        else:
            EmailAttachmentSelectedSafety = False
            button_exploreDOCx.configure(bg='green4')
            button_send.configure(bg='blue')


    def SendNEST():
        global ExcelFileSelectedSafety
        global EmailAttachmentSelectedSafety
        if ExcelFileSelectedSafety == True or EmailAttachmentSelectedSafety == True:
            labelERRORs.configure(text='Please select Excel file and DOCx before sending')
            SubWindow1.update()
            return
        else:         
            global TotalAddresses
            global nestlettersent
            global enrollmentletterNEST
            nestlettersent = 0
            for NESTEmails in AddressesFromExcelFile:
                outlook = win32.Dispatch('outlook.application')
                mail = outlook.CreateItem(0)
                mail.To = NESTEmails
                mail.Subject = ''
                mail.Body = ''
                mail.Attachments.Add(enrollmentletterNEST)
                mail.Send()
                nestlettersent += 1
                labelsentcounterNEST.configure(text = "Amount of files sent: %d/%d" % (nestlettersent,TotalAddresses))
                time.sleep(1)
                SubWindow1.update()
            labelsentcounterNEST.configure(text = "All emails sent %d/%d" % (nestlettersent,TotalAddresses))
            ExcelFileSelectedSafety = 'On'
            EmailAttachmentSelectedSafety = 'On'
            Button_Explorer_Excel.configure(bg='red3')
            button_exploreDOCx.configure(bg='red3')
            button_send.configure(bg='red3')
            SubWindow1.update()
    

    ## Labels/Buttons


    label_file_explorerNEST = Label(SubWindow1,text = "Email Sender - By Nick",width = 100, height = 4,fg = "white",bg = 'RoyalBlue4', font=('Times', 13))
    Button_Explorer_Excel = Button(SubWindow1,text = "Select Excel File for Sending",bg = 'red3',width = 30, height = 2,command = SelectedExcelFileButton, font=('Times', 15, 'bold'), fg = "yellow2")
    button_exploreDOCx = Button(SubWindow1,text = "Select file to be Sent",bg = 'red3',width = 30, height = 2,command = SelectNESTDocx, font=('Times', 15, 'bold'), fg = "yellow2")
    button_exitNEST = Button(SubWindow1,text = "Exit",bg = 'snow4',width = 30,height = 2,command = sys.exit, font=('Times', 15, 'bold'), fg = "black")
    button_send = Button(SubWindow1,text = "Send",bg = 'red3',width = 30,height = 2,command = SendNEST, font=('Times', 15, 'bold'), fg = "yellow2")
    labelERRORs = Label(SubWindow1,text = "",width = 75, height = 2,fg = "red",bg = 'RoyalBlue4', font=('Times', 16))
    labelfileopned = Label(SubWindow1,text = "",width = 75, height = 2,fg = "white",bg = 'RoyalBlue4', font=('Times', 16))
    labelsentcounterNEST = Label(SubWindow1,text = "",width = 75, height = 2,fg = "white",bg = 'RoyalBlue4', font=('Times', 16))
    labelfileopnedDocxNEST = Label(SubWindow1,text = "",width = 75, height = 2,fg = "white",bg = 'RoyalBlue4', font=('Times', 16))
    LabelSpace1NEST = Label(SubWindow1,text = "",width = 75, height = 1,fg = "red",bg = 'RoyalBlue4')
    LabelSpace2NEST = Label(SubWindow1,text = "",width = 75, height = 1,fg = "red",bg = 'RoyalBlue4')
    LabelSpace3NEST = Label(SubWindow1,text = "",width = 75, height = 1,fg = "red",bg = 'RoyalBlue4')


    ## Packed Labels

    def allpacks():
        label_file_explorerNEST.pack()
        Button_Explorer_Excel.pack()
        LabelSpace1NEST.pack()
        button_exploreDOCx.pack()
        LabelSpace2NEST.pack()
        button_exitNEST.pack()
        LabelSpace3NEST.pack()
        button_send.pack()
        labelERRORs.pack()
        labelfileopned.pack()
        labelfileopnedDocxNEST.pack()
        labelsentcounterNEST.pack()
    allpacks()


    ## Window Loop

    SubWindow1.mainloop()   


## Labels/Buttons

Labelfill = Label(MainWindow, width = 30, height = 3,bg = 'RoyalBlue4')
Labelfill1 = Label(MainWindow, width = 30, height = 3,bg = 'RoyalBlue4')
SubWindow1Button = Button(MainWindow,text = "Send files",bg = 'RoyalBlue1',width = 30,height = 2, font=('Times', 15, 'bold'), fg = "yellow2",command = SendFilesButton) #command = P45Selected


## Packed Labels

Labelfill.pack()
Labelfill1.pack()
SubWindow1Button.pack()


## Window Loop

MainWindow.mainloop()

