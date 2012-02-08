
    # ``````````````````````CSV TO HTML````````````````````````````#
    # ````````````````````By Ryan Kendall``````````````````````````#
    # `````````````````````November 2011```````````````````````````#

import csv
import argparse
from time import time
import datetime
from datetime import datetime
from tkinter import *
import tkinter.messagebox
import tkinter.filedialog
import webbrowser
import sys
from operator import itemgetter

class ListOfExpenses:
    inFile = None
    outName = None
    sortCriteria = None
    columns = [0,1,2,3,4,5,6,7]
    expenses = []
    deptIndex = None
    entIndex = None
    dateIndex = None
    exTypeIndex = None
    exAreaIndex = None
    supplierIndex = None
    amountIndex = None
    transIndex = None
    order = True

    def __init__(self, iF, oN, sC, c):
        self.inFile = iF
        self.outName = oN.replace(".html","")   # clean up user input
        self.sortCriteria = sC
        if c != None:   # user has chosen columns
            cols = []   # Read column numbers 1 by 1
            c = c.replace(",","")
            for cell in c:
                cols.append(int(cell))
            self.columns = cols

    # Methods to help GUI
    # getters and setters

    def getListOfExpenses(self):
        return self.expenses
    def setListOfExpenses(self,loe):
        self.expenses = loe
    def getInFile(self):
        return self.inFile
    def setInFile(self,iF):
        self.inFile = iF
    def getOutName(self):
        return self.outName
    def setOutName(self,oN):
        self.outName = oN
    def getSortCriteria(self):
        return self.sortCriteria
    def setSortCriteria(self,sC):
        self.sortCriteria = sC
    def getColumns(self):
        return self.columns
    def setColumns(self,c):
        self.columns = c
    def setOrder(self,order):
        if order:
            self.order = True
        else:
            self.order= False

    def readInData(self):
        # Removing and Adding file extension
        # (to save checking if they have entered or not)

        self.inFile = self.inFile.replace(".csv","")
        self.inFile += ".csv"
        print("# Opening {0}".format(self.inFile))

        # Try to open file, read in data and add it to self.expenses[]
        try:
            data = csv.reader(open(self.inFile,"r"))
        except IOError:
            # file doesnt exist, close program
            print("\nError opening file, file may not exist. Exiting...")
            sys.exit()

        print("\n# Parsing {0}".format(self.inFile))

        # fill expenses with data from csv
        self.expenses = [[line[i] for i in self.columns] for line in data]

        # get list index of "Transaction" and "Amount Sterling"
        # for conversion later
        if 6 in self.columns:
            indexOfTrans = self.columns.index(6)
        if 7 in self.columns:
            indexOfAmount = self.columns.index(7)

        # Convert Transaction ID and Amount Sterling to int and float
        for x,line in enumerate(self.expenses):
            try:
                self.expenses[x][indexOfAmount] = float(
                                            line[indexOfAmount].replace(",",""))
            except:
                pass
            try:
                self.expenses[x][indexOfTrans] = int(line[indexOfTrans])
            except:
                pass

    def sortData(self):
        self.expenses = sorted(self.expenses,key=itemgetter(self.sortCriteria),
                                                            reverse=self.order)

    def writeHTML(self):
        # Define stylesheet
        styles = """<style type="text/css">
            body{background-color: #2B3856;}#container{width:75%;margin:auto;}
            table{padding-"left":3px;padding-"right":3px;
            border-top:6px solid #000;
            background-color:#fff;}table th{background-color:#f5f5f5;
            font-family:Georgia;font-size:16px;border-bottom:2px solid #a0a0a0;}
            table td{
            font-family: verdana;font-size:12px;padding:3px;border-bottom:1px
            dotted #cfcfcf;border-"right":1px dotted #dfdfdf;}</style>"""

        # List of column headers for selection
        header = ["<th>Department Family</th>","<th>Entity</th>",
        "<th>Date of Payment</th>","<th>Expense Type</th>",
        "<th>Expense Area</th>","<th>Supplier</th>",
        "<th>Transaction number</th>","<th>Amount in Sterling (&pound;)</th>"]

        openPage = """<!doctype html><html>
            <head><title>Data</title>{0}</head><body>
            <div id=\"container\"><table><tr>""".format(styles)

        # Only show certain headers based on columns chosen
        for i in range(8):
            if i in self.columns:
                openPage += header[i]
        openPage += "</tr>"

        listOfTags = [openPage]
        global pageNo               # Global for UI use
        pageNo, x=0, 0


        # x is a counter, for counting data rows
        for x,line in enumerate(self.expenses):
            listOfTags.append("<tr>")
            for i in range(0,len(self.columns)):
                # for as long as there are columns, add html
                try:
                    listOfTags.append("<td>{0}</td>".format(line[i]))
                except IndexError:
                    pass
            listOfTags.append("</tr>")         # end table row

            if x % 100 == 0 and x != 0:        # on the 100th row make a page
                pageNo = int(x*0.01)           # using x counter as page number
                listOfTags.append("</table>")  # finish off page
                if x > 100:
                    # After first page - PREV / NEXT nav
                    listOfTags.append("""<a href=\"{0}{1}.html\">
                            Previous Page</a> | <a href=\"{0}{2}.html\">
                            Next Page</a>""".format(self.outName,pageNo-1,
                                                     pageNo+1))
                else:
                    # First Page - ONLY NEXT nav
                    listOfTags.append("""<a href=\"{0}{1}.html\">Next Page</a>
                        </div></body></html>""".format(self.outName,pageNo+1))
                # Build page filename
                pageName = "{0}{1}.html".format(self.outName,pageNo)

                # Write page
                page = open(pageName, "w")
                page.write("".join(listOfTags))

                # Reset html code
                listOfTags = [openPage]

        # Finalise last page for remaining rows
        if x > 99:
            listOfTags.append("</table><a href=\"{0}{1}.html\">\
                    Previous Page</a></div></body></html>".format(self.outName,
                                                                pageNo-1))
            pageName = "{0}{1}.html".format(self.outName,pageNo)
        else:
            # No navigation needed (only one page)
            listOfTags.append("</table></div></body></html>")
            pageName = "{0}.html".format(self.outName)

        # write page
        page = open(pageName, "w")
        page.write("".join(listOfTags))
        page.close()

        print("# Success! {0} file(s) created".format(pageNo))

    def updateSortCriteria(self, c):
        self.columns = c

        # can only receive gui variable if gui has been used
        if gui:
            self.sortCriteria = varRbStatus.get()

        # if "0" appears in columns, user has chosen to show Dept etc.
        # collect index in which "0" appears for use with sorting (itemgetter)
        # as expenses[x][7] will not refer to Amount Sterling after
        # selecting columns etc.

        # if user selected to sort for "7" (Amount Sterling)
        # then update sortCriteria to account for columns

        try:
            self.deptIndex = int(self.columns.index(0))
            if self.sortCriteria == 0:
                self.sortCriteria = self.deptIndex
        except ValueError:
            pass
        try:
            self.entIndex = int(self.columns.index(1))
            if self.sortCriteria == 1:
                self.setSortCriteria = self.entIndex
        except ValueError:
            pass
        try:
            self.dateIndex = int(self.columns.index(2))
            if self.sortCriteria == 2:
                self.sortCriteria = self.dateIndex
        except ValueError:
            pass
        try:
            self.exTypeIndex = int(self.columns.index(3))
            if self.sortCriteria == 3:
                self.sortCriteria = self.exTypeIndex
        except ValueError:
            pass
        try:
            self.exAreaIndex = int(self.columns.index(4))
            if self.sortCriteria == 4:
                self.sortCriteria = self.exAreaIndex
        except ValueError:
            pass
        try:
            self.supplierIndex= int(self.columns.index(5))
            if self.sortCriteria == 5:
                self.sortCriteria = self.supplierIndex
        except ValueError:
            pass
        try:
            self.transIndex = int(self.columns.index(6))
            if self.sortCriteria == 6:
                self.sortCriteria = self.transIndex
        except ValueError:
            pass
        try:
            self.amountIndex = int(self.columns.index(7))
            if self.sortCriteria == 7:
                self.sortCriteria = self.amountIndex
        except ValueError:
            pass

    # END OF CLASS LISTOFEXPENSES ---------------------------------------- #

    # START OF UI -------------------------------------------------------- #

def chooseFile(loe):
    # File dialog to choose .csv file

    loe.setInFile(tkinter.filedialog.askopenfilename(
                                                title="Choose your CSV File"))

    # update label
    varlblInputFile.set("CSV file Chosen: {0}".format(loe.getInFile()))
    lblInputFile.config(fg='blue')
    loe.readInData()

def nameHTML(loe):
    # File dialog to save your html page

    loe.setOutName(tkinter.filedialog.asksaveasfilename(
                                                    title="Give HTML a name"))

    # update label
    varlblOutputName.set("HTML path and filename: {0}.html".format(
                                        loe.getOutName().replace(".html","")))
    lblOutputName.config(fg='blue')

def fillTxtResults(loe):
    # Fill text box with results of sorting
    exp = loe.getListOfExpenses()
    txtResults.delete(0.0,END)
    txtResults.insert(END, [line for x,line in enumerate(exp) if x < 100])

def radioButtonsChanged(loe):
    # collect info from radio buttons and update class
    loe.setSortCriteria(varRbStatus.get())

def getData(loe):
    # bring together all data from UI elements

    radioButtonsChanged(loe)
    loe.readInData()

    if loe.getColumns() == []:
        # user has disabled all columns
        tkinter.messagebox.showerror(
                                message="Must have at least 1 column chosen!")
    else:
        # at least one column
        if varRbStatus.get() == -1:
            # chosen not to sort
            loe.setSortCriteria(None)
        else:
            # chosen a sort criteria
            loe.setSortCriteria(varRbStatus.get())
            loe.updateSortCriteria(loe.getColumns())
            loe.sortData()

        # update results box
        fillTxtResults(loe)

def updateOrder(loe):
    # Allows user to choose ASC or DESC order for results

    loe.setOrder(varASCorDESCB.get()) # get value from radiobutton
    loe.sortData()  # re sort data based on choice
    fillTxtResults(loe) # update results box

def addSortRadioButtons(loe):
    # A function to create radiobuttons for sorting based on columns chosen

    cols = loe.getColumns()
    loe.setSortCriteria(-1) # default sort criteria to None

    # List of headers for radiobutton labels
    colLabel = ["Dept","Entity","Date of Payment","Expense Type",
                "Expense Area","Supplier","Transaction ID","Amount in Sterling"]

    # always show no sort option
    rSortByNone = Radiobutton(sortFrame, variable=varRbStatus,
                                    value=-1,text="Don't Sort", command=
                                    lambda: radioButtonsChanged(loe))
    rSortByNone.pack()

    # work out which radio buttons to show based on columns
    for i in range(0,len(cols)):
        try:
            print(i)
            rbSortBy = Radiobutton(sortFrame, variable=varRbStatus,
                            value=i,text=colLabel[cols[i]],command=
                            lambda: radioButtonsChanged(loe)).pack()
        except IndexError:
            pass

    rSortByNone.select()    # have no sort selected by default

def quitProgram(app):
    # Let gui close the program to avoid running through to non gui code
    app.quit()
    app.destroy()
    sys.exit()

def updateColumns(loe):
    # take in column choices from checkboxes

    radioButtonsChanged(loe)    # first see which sort is chosen

    # collect checkbox values
    cboxes = [varDeptCB.get(),varEntityCB.get(),varDateCB.get(),
                varExpTypeCB.get(),varExpAreaCB.get(),varSupplierCB.get(),
                varTransCB.get(),varAmountCB.get()]

    newCols = []

    # If checkboxes are clicked add that position to newCols[]
    for x,cell in enumerate(cboxes):
        if cell:
            newCols.append(x)

    print(cboxes)
    loe.setColumns(newCols)
    loe.updateSortCriteria(loe.getColumns())

def guiWriteHTML(loe):
    # Facility for gui to call writeHTML class method
    loe.writeHTML()

    if pageNo == 0:
        tkinter.messagebox.showerror(message="{0} files created,\
                        \nDid you forget to inport the data?".format(pageNo))
    else:
        tkinter.messagebox.showinfo(message="{0} files created".format(pageNo))

def userInterface(self):
    # Define most of the UI here

    app = Tk()
    app.title("CSV to HTML Perfection")
    app.geometry("850x600+200+200")

    # MENU BAR
    menubar = Menu(app)
    fileMenu = Menu(menubar,tearoff=0)
    fileMenu.add_command(label="Choose CSV File",
                                    command= lambda: chooseFile(loe))
    fileMenu.add_command(label="Name HTML File",
                                    command= lambda: nameHTML(loe))
    fileMenu.add_separator()
    fileMenu.add_command(label="Quit", command= lambda: quitProgram(app))
    menubar.add_cascade(label="File",menu=fileMenu)
    app.config(menu=menubar)

    # Filo I/O Group
    ioFrame = LabelFrame(app, text="File I/O")
    ioFrame.pack(fill="both",expand="yes",side="top")

    # CSV LABEL
    global varlblInputFile
    varlblInputFile = StringVar()
    varlblInputFile.set("Current CSV File: {0}".format(loe.getInFile()))
    global lblInputFile
    lblInputFile = Label(ioFrame, textvariable=varlblInputFile,height=2)
    lblInputFile.pack(side="top")

    # LBL HTML FILE
    global varlblOutputName
    varlblOutputName = StringVar()
    varlblOutputName.set("Current HTML Filename: {0}".format(loe.getOutName()))
    global lblOutputName
    lblOutputName = Label(ioFrame, textvariable=varlblOutputName,height=2)
    lblOutputName.pack(side="bottom")

    # Sort Frame
    global sortFrame
    sortFrame = LabelFrame(app, text="Sorting")
    sortFrame.pack(fill="both",expand="yes",side="right")


    # Sort Radio Buttons
    global varRbStatus
    varRbStatus = IntVar()
    addSortRadioButtons(loe)

    # Control Frame
    controlFrame = LabelFrame(app, text="Controls")
    controlFrame.pack(fill="both",expand="yes",side="top")

    global varASCorDESCB
    varASCorDESCB = IntVar()
    cbASCorDESC = Radiobutton(controlFrame, value=1 , text="Ascending",
                    variable=varASCorDESCB,command= lambda: updateOrder(loe))
    cbASCorDESC.pack(side="right")

    cbASCorDESC = Radiobutton(controlFrame, value=0, text="Descending",
                    variable=varASCorDESCB,command= lambda: updateOrder(loe))
    cbASCorDESC.pack(side="right")


    # COLLECT ALL DATA FROM ENTRY
    #-----------------------------
    inButton = Button(controlFrame, text="Submit",
                                                command= lambda: getData(loe))
    inButton.pack(side="left")

    # write html button
    outButton = Button(controlFrame, text="Write HTML",
                                            command= lambda: guiWriteHTML(loe))
    outButton.pack(side="left")

    # Columns Frame
    columnsFrame = LabelFrame(app, text="Choose Columns")
    columnsFrame.pack(fill="both",expand="yes", side="top")

    # Results Frame
    resultsFrame = LabelFrame(app, text="Top 100 Results")
    resultsFrame.pack(fill="both",expand="yes",side="bottom")

    scrollList = Scrollbar(resultsFrame,orient=VERTICAL)
    global txtResults
    txtResults = Text(resultsFrame,width=650,height=300,
                                            yscrollcommand = scrollList.set)

    txtResults.pack()
    scrollList.pack(side="right", fill="x")
    scrollList.config(command=txtResults.yview)




    # Columns checkboxes
    # DEPT
    global varDeptCB
    varDeptCB = BooleanVar()
    cbDept = Checkbutton(columnsFrame, text="Dept", variable=varDeptCB,
                                            command= lambda: updateColumns(loe))
    cbDept.pack(side="left")
    cbDept.select()

    # ENTITY
    global varEntityCB
    varEntityCB = BooleanVar()
    cbEntity = Checkbutton(columnsFrame, text="Entity", variable=varEntityCB,
                                            command= lambda: updateColumns(loe))
    cbEntity.pack(side="left")
    cbEntity.select()

    # DATE OF PAYMENT
    global varDateCB
    varDateCB = BooleanVar()
    cbDate = Checkbutton(columnsFrame, text="Date of Payment",
                        variable=varDateCB,command= lambda: updateColumns(loe))
    cbDate.pack(side="left")
    cbDate.select()

    # EXPENSE TYPE
    global varExpTypeCB
    varExpTypeCB = BooleanVar()
    cbExpType = Checkbutton(columnsFrame, text="Expense Type",
                    variable=varExpTypeCB,command= lambda: updateColumns(loe))
    cbExpType.pack(side="left")
    cbExpType.select()

    # EXPENSE AREA
    global varExpAreaCB
    varExpAreaCB = BooleanVar()
    cbExpArea = Checkbutton(columnsFrame, text="Expense Area",
                    variable=varExpAreaCB,command= lambda: updateColumns(loe))
    cbExpArea.pack(side="left")
    cbExpArea.select()

    # SUPPLIER
    global varSupplierCB
    varSupplierCB = BooleanVar()
    cbSupplier = Checkbutton(columnsFrame, text="Supplier",
                    variable=varSupplierCB,command= lambda: updateColumns(loe))
    cbSupplier.pack(side="left")
    cbSupplier.select()

    # TRANSACTION ID
    global varTransCB
    varTransCB = BooleanVar()
    cbTrans = Checkbutton(columnsFrame, text="Transaction ID",
                        variable=varTransCB,command= lambda: updateColumns(loe))
    cbTrans.pack(side="left")
    cbTrans.select()

    # AMOUNT STERLING
    global varAmountCB
    varAmountCB = BooleanVar()
    cbAmount = Checkbutton(columnsFrame, text="Amount Sterling",
                    variable=varAmountCB,command= lambda: updateColumns(loe))
    cbAmount.pack(side="left")
    cbAmount.select()

    app.mainloop()  # Set off GUI

    # END OF GUI ---------------------------------------------------- #

    # START OF COMMAND LINE ARGS ------------------------------------ #

def argParse():
    # Set up arguments
    parser = argparse.ArgumentParser(prog='CSV2HTML',
            usage="%(prog)s [options]",
            description="Parse a .csv file and create a .html page",
            epilog="Have Fun! ~ Ryan")

    # Add argument - GUI
    parser.add_argument("-gui",  default=False, type=bool,
            help="1 Launches the program in GUI mode. 0 by default")

    # Add argument - Input file
    parser.add_argument("-i",default="data",
            type=str, help="The CSV file to be parsed and/or sorted (file.csv)")

    # Add argument - Name of html results page
    parser.add_argument("-o",default="results",
            type=str, help="The name of the resulting HTML file (file.html)")

    # Add argument - Columns
    parser.add_argument("-cols",default=None,
            type=str, help="""The Columns you wish to see in the results...
                Columns include: Department: 0, Entity: 1,
                Date of Payment: 2, Expense type: 3, Expense area: 4,
                Supplier: 5, Transaction Number: 6, Amount in sterling: 7... 
                Usage -cols 1,2,3,4""")

    # Add argument - Sort Criteria
    parser.add_argument("-cri", default=None, type=int,
            help="""Choices to sort by: Department: 0, Entity: 1,
            Date of Payment: 2, Expense type: 3, Expense area: 4,
            Supplier: 5, Transaction Number: 6, Amount in sterling: 7""")

    # Receive and pass on command line args
    return vars(parser.parse_args())

    # END OF COMMAND LINE ARGS --------------------------------------- #

    # MAIN ----------------------------------------------------------- #

if __name__ == '__main__':

    # recive args
    args = argParse()

    # initialise variables to be sent to class
    inFile = args["i"]
    outName = args["o"]
    sortCriteria = args["cri"]
    columns = args["cols"]
    global gui
    gui = args["gui"]

    # create ListOfExpenses class, with args
    loe = ListOfExpenses(inFile,outName,sortCriteria,columns)

    if gui:
        # user has chosen GUI
        userInterface(loe)
    else:
        # user has not chosen GUI
        loe.readInData()
        if sortCriteria:
            # user chosen to sort
            if columns:
                # user chosen columns
                loe.updateSortCriteria(loe.getColumns())
            # sort
            loe.sortData()
        # write out html page
        loe.writeHTML()

    # END OF MAIN AND CSV2HTML.PY ---------------------------------- #