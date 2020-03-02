#---<===>------<===>------<===>------<===>------<===>------<===>------<===>------<===>------<===>------<===>------<===>------<===>------<===>------<===>------<===>------<===>---#
#    ______ _____          _____        _____ _____  _____ _______ _____ _   _  _____ _______ _____ ____  _   _  _____       _____  _               _   _ _   _ ______ _____     #
#   |  ____|_   _|   /\   |  __ \      |  __ \_   _|/ ____|__   __|_   _| \ | |/ ____|__   __|_   _/ __ \| \ | |/ ____|     |  __ \| |        /\   | \ | | \ | |  ____|  __ \    #
#   | |__    | |    /  \  | |__) |     | |  | || | | (___    | |    | | |  \| | |       | |    | || |  | |  \| | (___       | |__) | |       /  \  |  \| |  \| | |__  | |__) |   #
#   |  __|   | |   / /\ \ |  ___/      | |  | || |  \___ \   | |    | | | . ` | |       | |    | || |  | | . ` |\___ \      |  ___/| |      / /\ \ | . ` | . ` |  __| |  _  /    #
#   | |     _| |_ / ____ \| |          | |__| || |_ ____) |  | |   _| |_| |\  | |____   | |   _| || |__| | |\  |____) |     | |    | |____ / ____ \| |\  | |\  | |____| | \ \    #
#   |_|    |_____/_/    \_\_|          |_____/_____|_____/   |_|  |_____|_| \_|\_____|  |_|  |_____\____/|_| \_|_____/      |_|    |______/_/    \_\_| \_|_| \_|______|_|  \_\   #
#                                                                                                                                                                                #
#                                                                  Copyright 2019 C.C.Gold All Rights Reserved                                                                   #
#                                             Not Affiliated with Fédération Internationale de l'Art Photographique (FIAP) In Any Way                                            #
#                                                                                                                                                                                #
#  PROGRAM NAME : FIAP Distinction Planner                                                                                                                                       #
#        AUTHOR : C.C Gold                                                                                                                                                       #
#  DATE CREATED : 20/03/2019                                                                                                                                                     #
#       VERSION : 0.8.8.0                                                                                                                                                        #
#                                                                                                                                                                                #
#   DESCRIPTION :                                                                                                                                                                #
#   A planner to help art photographers prepare for achieving a FIAP distinction                                                                                                 #
#   For more information see the provided user manual (F1)                                                                                                                       #
#---<===>------<===>------<===>------<===>------<===>------<===>------<===>------<===>------<===>------<===>------<===>------<===>------<===>------<===>------<===>------<===>---#
'''
For posterity I have listed all FIAP distinction requirements here ordered by distinction type - 

AFIAP
    Date (First Acceptance): => 1 Year
    Unique Salons with Acceptance: => 15
    Unique Works with Acceptance: => 15
    Unique Countries with Acceptance: => 5
    Total Acceptances: => 40
    Print Acceptances: => 10% of Total Acceptances (4)

EFIAP
    Date (Attained AFIAP): => 1 Year
    Unique Salons with Acceptance: => 30
    Unique Works with Acceptance: => 50
    Unique Countries with Acceptance: => 20
    Total Acceptances: => 250
    Print Acceptances: => 10% of Total Acceptances (25)

EFIAP Levels B/S/G/P
    Date (Attained EFIAP): => 1 Year
    Total Acceptances: 
        Bronze => 400
        Silver => 700
        Gold => 1000
        Platinum => 1500
    Unique Works with Acceptance:
        Bronze => 100
        Silver => 200
        Gold => 300
        Platinum => 400

EFIAP Levels D1/D2/D3
    Date (Attained EFIAP/p): => 1 Year
    From 2015 onwards:
        Total Awards:
            D1 => 50
            D2 => 100
            D3 => 200
        Unique Works with Award:
            D1 => 15
            D2 => 30
            D3 => 50
        Unique Countries Awarded In:
            D1 => 5
            D2 => 7
            D3 => 10
'''
#Import Libraries
from datetime import date
import pygsheets

#Access user google account
#Walks user through a consent process on first runtime
gc = pygsheets.authorize(client_secret='./data/client_secret.json',credentials_directory="./data")

class MainApplication(tk.Frame):
    
    #---<===>---# Main Controller Function #---<===>---# 

    def __init__(self, parent, sheet):
        tk.Frame.__init__(self, parent)
        
        #---<===>---# Class Attributes #---<===>---# 
               
        #Store google client
        self.gc = sheet

        #Title variables
        self.portElegDist = tk.StringVar()
        self.portEarnedDist = tk.StringVar()
        self.openTitle = tk.StringVar()

        #Runtime Records
        self.whatsOpen = None
        self.drawnRows = {"NewWork":0,"NewSub":0,"SubRec":0,"WorkRec":0}
        self.workRecSheet = None
        self.subRecSheet = None
        self.portRecSheet = None
        self.newWorkTable = []
        self.newSubTable = []
        
        #Table Information
        self.workColumns = 10
        self.editableWorkColumns = 4
        self.subColumns = 10
        self.editableSubColumns = 10
        self.fiapDist = ["None","AFIAP","EFIAP","EFIAP/b","EFIAP/s","EFIAP/g","EFIAP/p","EFIAP/d1","EFIAP/d2","EFIAP/d3"]

        #---<===>---# END Class Attributes #---<===>---# 



        #---<===>---# Spreadsheet Formula Bank #---<===>---# 
        
        #Totals
        self.pfformTotalWorks = "=A2"
        self.pfformTotalSubs = "=A4"
        self.pfformTotalCost = "=SUM(subRec!J2:J)"
        self.pfformWorkRowCheck = "=COUNTA(workRec!B2:B)"
        self.pfformSubRowCheck = "=COUNTA(subRec!B2:B)"

        #Statistics
        self.pfformDateOfFirstSuccess = "=IF(COUNTIF(subRec!D2:D,true)<1,B5,(MIN(FILTER(subRec!C2:C,subRec!D2:D=TRUE))))"
        self.pfformTotalAcceptances = "=SUM(workRec!G2:G)"
        self.pfformUniqueSubWithAcceptance = "=COUNTA(IFERROR(UNIQUE(FILTER(subRec!B2:B,subRec!D2:D=TRUE))))"
        self.pfformSalonsWithSuccess = "=COUNTA(IFERROR(UNIQUE(FILTER(subRec!G2:G,subRec!D2:D=TRUE))))"
        self.pfformCountriesWithSuccess = "=COUNTA(IFERROR(UNIQUE(FILTER(subRec!F2:F,subRec!D2:D=TRUE))))"
        self.pfformUniquePrintAcceptances = "=COUNTA(IFERROR(UNIQUE(FILTER(subRec!B2:B, subRec!D2:D=TRUE,subRec!H2:H=True))))"
        self.pfformTotalAwards = "=COUNTIFS(Arrayformula(IF(subRec!E2:E=FALSE,FALSE,true)),True,ARRAYFORMULA(IF(subRec!C2:C<DATE(2015,1,1),False,True)),true)"
        self.pfformUniqueWorkAwards = "=COUNTA(IFERROR(UNIQUE(FILTER(subRec!B2:B,subRec!E2:E=TRUE,subRec!C2:C>DATE(2015,1,1)))))"
        self.pfformUniqueCountryAwards = "=COUNTA(IFERROR(UNIQUE(FILTER(subRec!F2:F,subRec!E2:E=TRUE,subRec!C2:C>DATE(2015,1,1)))))"
        
        #Eligability Checks
        self.pfformEligibleCheck = "=IF(J4=true,J3,IF(I4=true,I3,IF(H4=true,H3,IF(G4=true,G3,IF(F4=true,F3,IF(E4=true,E3,IF(D4=true,D3,IF(C4=true,C3,IF(B4=true,B3,C5)))))))))"
        self.pfformAFIAP_Check = "=IF(floor(TODAY()-365.25)<E2,False,IF((NOT(K4=C5)),FALSE,IF(H2<15,false,IF(I2<8,false,IF(F2<40,false,IF(G2<15,false,IF(J2<4,false,true)))))))"
        self.pfformEFIAP_Check = "=IF(FLOOR(TODAY()-365.25)<L4,FALSE,IF((NOT(K4=B3)),FALSE,IF(G2<50,false,IF(H2<30,false,IF(I2<20,false,IF(F2<250,false,IF(J2<25,false, true)))))))"
        self.pfformEFIAP_B_Check = "=IF(FLOOR(TODAY()-365.25)<$L$4,FALSE,IF((NOT($K$4=C3)),FALSE,IF($F$2<400,False,IF($G$2<100,false,true))))"
        self.pfformEFIAP_S_Check = "=IF(FLOOR(TODAY()-365.25)<$L$4,FALSE,IF((NOT($K$4=D3)),FALSE,IF($F$2<700,False,IF($G$2<200,false,true))))"
        self.pfformEFIAP_G_Check = "=IF(FLOOR(TODAY()-365.25)<$L$4,FALSE,IF((NOT($K$4=E3)),FALSE,IF($F$2<1000,False,IF($G$2<300,false,true))))"
        self.pfformEFIAP_P_Check = "=IF(FLOOR(TODAY()-365.25)<$L$4,FALSE,IF((NOT($K$4=F3)),FALSE,IF($F$2<1500,False,IF($G$2<400,false,true))))"
        self.pfformEFIAP_D1_Check = "=IF(FLOOR(TODAY()-365.25)<$L$4,FALSE,IF((NOT($K$4=G3)),FALSE,IF($K$2<50,false,IF($L$2<15,false,IF($M$2<5,false,true)))))"
        self.pfformEFIAP_D2_Check = "=IF(FLOOR(TODAY()-365.25)<$L$4,FALSE,IF((NOT($K$4=H3)),FALSE,IF($K$2<100,false,IF($L$2<30,false,IF($M$2<7,false,true)))))"
        self.pfformEFIAP_D3_Check = "=IF(FLOOR(TODAY()-365.25)<$L$4,FALSE,IF((NOT($K$4=I3)),FALSE,IF($K$2<200,false,IF($L$2<50,false,IF($M$2<10,false,true)))))"

        #Work Record Statistics Formulas
        self.wkformSubAsPrint = "=If($B{0}=FALSE,,iferror(if(INDEX(subRec!$H$2:$H,MATCH(B{0}&TRUE,subRec!$B$2:$B&subRec!$H$2:$H,0),0)=TRUE,portRec!$D$5),portRec!$E$5))"
        self.wkformUniqueCountrySubs = "=If($B{0}=FALSE,,ARRAYFORMULA(COUNTUNIQUE(IF(subRec!$B$2:$B=$B{0},subRec!$F$2:$F))-1))"
        self.wkformUniqueSalonSubs = "=If($B{0}=FALSE,,ARRAYFORMULA(COUNTUNIQUE(IF(subRec!$B$2:$B=$B{0},subRec!$G$2:$G))-1))"
        self.wkformAcceptancesNum = "=If(B{0}=FALSE,,ARRAYFORMULA(COUNTIF(IF(subRec!$B$2:$B=$B{0},subRec!$D$2:$D),TRUE)))"
        self.wkformFirstSubDate = "=If($B{0}=FALSE,,Iferror(Min(FILTER(subRec!$C$2:$C,subRec!$B$2:$B=$B{0})),portRec!$B$5))"
        self.wkformAwardsNum = "=If(B{0}=FALSE,,ARRAYFORMULA(COUNTIF(IF(subRec!$B$2:$B=$B{0},subRec!$E$2:$E),TRUE)))"

        #---<===>---# END Spreadsheet Formula Bank #---<===>---# 



        #---<===>---# Display Tables #---<===>---# 
        #Use these to store the location of each element on the center screen
        #and the variable which sets their value for reassignment later
        
        #Records
        self.subRecVars=[]
        self.subRecElements=[]
        self.workRecVars=[]
        self.workRecElements=[]

        #New items
        self.newSubVars=[]
        self.newSubElements=[]
        self.newWorkVars=[]
        self.newWorkElements=[]

        #Portfolio
        self.portVars={}

        #---<===>---# END Display Tables #---<===>---# 



        #---<===>---# Run GUI Creation #---<===>---# 

        #Main GUI creation process
        self.buildGUIStructure()
        self.update()
        
        #Show a display telling the user the program is loading      
        self.loadMessage("Loading...")

        #Open the user's google sheets spreadsheet where records will be kept
        #If one doesn't exist it will be created for them
        self.openSpreadsheet()

        #Find records from the spreadsheet 
        self.buildRecords()

        #Build the top portfolio progress display       
        self.renderPortfolio()

        #Build the header and title displays for the central window
        self.buildTitles()

        #---<===>---# END GUI Creation


        #Open Submission Record screen by default.
        #In the future could save the last open screen and open that by default
        #when the user runs the program next.
        self.openSubRecord()

        #Close the loading popup
        self.loadingDisplay.destroy()

    #---<===>---# END Main Controller Function #---<===>---#  

    def buildRecords(self):
        '''
        Stores various Spreadsheet values for use later.  
        
        Can also be used to rebuild those values after changes have been made.
        '''
        #Find today's date for default values
        today = date.today()
        dateFormatted = today.strftime("%d/%m/%Y")        
        
        #Destroy Existing Records
        self.newWorkTable = []
        self.newSubTable = []        
        
        #Build Portfolio Record Sheet variables
        portSheet = self.portRecSheet.get_all_values(returnas='matrix')
        self.workCount = int(portSheet[1][0])
        print("workCount=",self.workCount)
        self.subCount = int(portSheet[3][0])
        print("subCount=",self.subCount)
        self.portTitles = portSheet[0][2:11]
        print("portTitles=",self.portTitles) 
        self.portValues = portSheet[1][2:11]
        print("portValues=",self.portValues)
        self.portElegDist.set(portSheet[1][1])
        print("portElegDist=",self.portElegDist.get())
        self.portEarnedDist.set(portSheet[3][10])
        print("portEarnedDist=",self.portEarnedDist.get())

        #Build Work Record Sheet variables
        workSheet = self.workRecSheet.get_all_values(returnas='matrix')        
        self.workTable = workSheet[:1+self.workCount][:]
        print("workTable=",workSheet[:1+self.workCount][:])
        self.newWorkTable.append(workSheet[0])
        self.newWorkTable.append(["NEW","",dateFormatted,"None","","","","","",""]) 
        print("newWorkTable=",self.newWorkTable)
        titleListRaw = [subList[1:2] for subList in workSheet]
        del(titleListRaw[-1])
        print("titleListRaw=",titleListRaw)
        self.titleList = ['{}'.format(*item) for item in titleListRaw]
        print("titleList=",self.titleList)

        #Build Submissions Record Sheet variables
        subSheet = self.subRecSheet.get_all_values(returnas='matrix')        
        self.subTable = subSheet[:1+self.subCount][:]
        print("subTable=",self.subTable)
        self.newSubTable.append(subSheet[0])
        self.newSubTable.append(["NEW","Select Title",dateFormatted,"FALSE","FALSE","","","FALSE","","£0.00"])
        print("newSubTable=",self.newSubTable)
        
        #???
        
        #Find count of rows to render
        #self.workCount = int(self.portRecSheet.get_value('A2'))
        #self.subCount = int(self.portRecSheet.get_value('A4'))

        #Build Table of spreadsheet values
        #self.workTable = self.workRecSheet.get_values(start=(1,1),end=(1+self.workCount,self.workColumns))
        #self.subTable = self.subRecSheet.get_values(start=(1,1),end=(1+self.subCount,self.subColumns))

        #Build smaller subtable for new works/records with only column headers
        #Then append a set of default values
        #self.newWorkTable = self.workRecSheet.get_values(start=(1,1),end=(1,self.workColumns))
        #self.newWorkTable.append(["","","","","","","","","",""])

        #self.newSubTable = self.subRecSheet.get_values(start=(1,1),end=(1,self.subColumns))
        #self.newSubTable.append(["","Select Title","","FALSE","","","","FALSE","",""])

        #build list of Work Titles
        #titleListRaw = self.workRecSheet.get_values(start=(2,2),end=(2+self.workCount,2))
        #self.titleList = ['{}'.format(*item) for item in self.titleListRaw]

        #Find Portfolio Values
        #self.portTitles = self.portRecSheet.get_values(start=("C1"),end=("K1"))
        #self.portValues = self.portRecSheet.get_values(start=("C2"),end=("K2"))
        #self.portElegDist.set(self.portRecSheet.get_value("B2"))

    def openSpreadsheet(self):
        '''
        Finds the user's planner spreadsheet and returns it.

        If no such spreadsheet can be found, will construct it before returning it
        '''        
        if "FIAP_PLANNER" not in self.gc.spreadsheet_titles(query=None):
            #If we cannot find a planner spreadsheet, construct it
            
            #Make the spreadsheet, title it then open it
            self.gc.create("FIAP_PLANNER")
            self.plannerSheet = self.gc.open("FIAP_PLANNER")

            #Create seperate worksheets for each record and assign them to class variables
            #Use strict row/column sizes to preserve the user's google drive space
            self.workRecSheet = self.plannerSheet.add_worksheet("workRec",rows=2,cols=self.workColumns)
            self.subRecSheet = self.plannerSheet.add_worksheet("subRec",rows=2,cols=self.subColumns)
            self.portRecSheet = self.plannerSheet.add_worksheet("portRec",rows=5,cols=15)

            #Delete the default sheet, too large by default and requires more google requests to edit
            #it to a usable state than just making new ones and deleting this instead
            delBase = self.plannerSheet.worksheet_by_title("Sheet1")
            self.plannerSheet.del_worksheet(delBase)

            
            #Using a set of default values build spreadsheet structure

            #Work Record Sheet
            self.workRecSheet.update_values((1,1),[
                [None,"Title","Date\nCreated","Used For\nApplication","First\nSubmission Date","Number Of\nAcceptances","Number of\nAwards","Unique Salons\nSubmitted To","Unique Countries\nSubmitted To","Submitted\nAs Print?"],
                ["1",None,None,None,self.wkformFirstSubDate.format(2),self.wkformAcceptancesNum.format(2),self.wkformAwardsNum.format(2),self.wkformUniqueSalonSubs.format(2),self.wkformUniqueCountrySubs.format(2),self.wkformSubAsPrint.format(2)]
                ])

            #Submission Record Sheet
            self.subRecSheet.update_values((1,1),[
                [None,"Title","Date\nSubmitted","Accepted","Recieved\nAward","Country","Salon / Circuit","Submitted\nAs Print?","FIAP\nNumber","Cost"],
                ["1",None,None,None,None,None,None,None,None,None]
                ])

            #Portfolio Record Sheet
            self.portRecSheet.update_values((1,1),[
                ["Work Rows","Elegible For\nDistinction:","Total Work\nRecords","Total Submission\nRecords","Date of\nFirst Success","Total\nAcceptances","Unique Submissions\nWith Acceptances","Salons Entered\nWith Success","Countries Submitted\nIn With Success","Unique Print\nAcceptances","Total\nCost","Total Awards\n(2015 Onwards)","Unique Works\nWith Awards\n(2015 Onwards)","Unique\nCountries\nAwarded In\n(2015 Onward)"],
                [self.pfformWorkRowCheck,self.pfformEligibleCheck,self.pfformTotalWorks,self.pfformTotalSubs,self.pfformDateOfFirstSuccess,self.pfformTotalAcceptances,self.pfformUniqueSubWithAcceptance,self.pfformSalonsWithSuccess,self.pfformCountriesWithSuccess,self.pfformUniquePrintAcceptances,self.pfformTotalCost,self.pfformTotalAwards,self.pfformUniqueWorkAwards,self.pfformUniqueCountryAwards],
                ["Sub Rows","AFIAP","EFIAP","EFIAP/b","EFIAP/s","EFIAP/g","EFIAP/p","EFIAP/d1","EFIAP/d2","EFIAP/d3","Current\nDistinction","Date of\nLast Distinction"],
                [self.pfformSubRowCheck,self.pfformAFIAP_Check,self.pfformEFIAP_Check,self.pfformEFIAP_B_Check,self.pfformEFIAP_S_Check,self.pfformEFIAP_G_Check,self.pfformEFIAP_P_Check,self.pfformEFIAP_D1_Check,self.pfformEFIAP_D2_Check,self.pfformEFIAP_D3_Check,"None","Never"],
                ["Vocabulary","Never","None","Yes","No","AFIAP","EFIAP","EFIAP/b","EFIAP/s","EFIAP/g","EFIAP/p","EFIAP/d1","EFIAP/d2","EFIAP/d3"],
                ])

            #Set cell formats and create model cell variables to apply formats in the future
            self.dateFormat = self.portRecSheet.cell('E2').set_number_format(pygsheets.FormatType.NUMBER,'dd/mm/yyyy')
            self.workRecSheet.cell('C2').set_number_format(pygsheets.FormatType.NUMBER,'dd/mm/yyyy')
            self.workRecSheet.cell('E2').set_number_format(pygsheets.FormatType.NUMBER,'dd/mm/yyyy')
            self.subRecSheet.cell('C2').set_number_format(pygsheets.FormatType.NUMBER,'dd/mm/yyyy')
            self.currencyFormat = self.portRecSheet.cell('K2').set_number_format(pygsheets.FormatType.NUMBER,'£0.00')            

        else:
            #Otherwise just open the existing sheet
            self.plannerSheet = self.gc.open("FIAP_PLANNER")

            #And assign the three record sheets to class variables for later use
            self.workRecSheet = self.plannerSheet.worksheet_by_title("workRec")
            self.subRecSheet = self.plannerSheet.worksheet_by_title("subRec")
            self.portRecSheet = self.plannerSheet.worksheet_by_title("portRec")

            #Find our model cells for applying formats
            self.dateFormat = self.portRecSheet.cell('C2')
            self.currencyFormat = self.portRecSheet.cell('K2') 

    def extendFormulas(self):
        '''
        Extends the formulas on the Work Record Sheet to the next row down
        '''
        
        #Find where the new row is
        newRow = str(self.workCount+2)

        #Assign it formulas, dynamically adjusting cell references
        self.workRecSheet.update_values('E'+newRow+':J'+newRow,[[self.wkformFirstSubDate.format(newRow), self.wkformAcceptancesNum.format(newRow), self.wkformAwardsNum.format(newRow), self.wkformUniqueSalonSubs.format(newRow), self.wkformUniqueCountrySubs.format(newRow), self.wkformSubAsPrint.format(newRow)]])

    def addNewRecord(self,newVars,recCount,recColumns,sheet,checkColumns={None}):
        '''
        Adds a new record based on input on the New Work and New Submission screens

        Takes 4 Arguments with a 5th optional one:

        1: A Table of Tkinter variables from the correct new record screen
        2: A count of the total records (to find the correct row to add to)
        3: A count of the record screen columns
        4: The record sheet to add the record to
        Optional 5: Any columns to treat as a Checkbutton (Output True/False rather than 1/0)
        '''
        #Find the correct row to add the record to
        newRow = recCount+2
        
        #Build empty local lists
        columnRange = []
        newRange = []

        #Loop through all columns, appending TK variable values to a list
        for column in range(recColumns):
            try:
                v = newVars[1][column].get()
            except AttributeError:
                #If the user hasn't added a value to each element, reject their attempt to save
                return
            
            #Ignore the first column, it just has a spacer label
            if column is 0:
                columnRange.append(None)

            #If assigning value from a Checkbutton, switch 1 with "True" and 0 with "False"
            elif column in checkColumns:
                if v is True or v == 1:
                    columnRange.append("TRUE")
                else:
                    columnRange.append("FALSE")
            
            #Append found value to the empty list
            else:
                columnRange.append(str(v))                

        #Create a table with the created list
        newRange.append(columnRange)

        #Add a row to the target sheet
        sheet.add_rows(1)

        #Then save the created table to that sheet targetting the newly created row
        sheet.update_values((newRow,1),newRange)
        
        #If the record we are updating is the Work Record, 
        #also extend it's spreadsheet formulas
        if self.whatsOpen is "NewWork":
            self.extendFormulas()
            self.setValues(self.workTable,self.workRecVars)
        elif self.whatsOpen is "NewSub":
            self.setValues(self.subTable,self.subRecVars)
        else:
            pass

        self.setValues()

    def saveRecord(self,saveVars,recCount,recColumns,recSheet,checkColumns={None}):
        '''
        Saves all changes to the currently open record screen

        Takes 4 Arguments with a 5th optional one:

        1: A Table of Tkinter variables from the correct record screen
        2: A count of the total records
        3: A count of the record screen columns
        4: The record sheet to save the changes to
        Optional 5: Any columns to treat as a Checkbutton (Output True/False rather than 1/0)
        '''
        
        #Build empty local lists
        columnRange = []
        saveRange = []

        #Loop through all rows and columns, appending TK variable values to a list
        for row in range(recCount+1):
            for column in range(recColumns):
                v = saveVars[row][column].get()
                #If assigning value from a Checkbutton, switch 1 with "True" and 0 with "False"
                if column in checkColumns and row != 0:
                    #Ignore the first row, it just has spacer labels
                    if v == True or v == 1:
                        columnRange.append("TRUE")
                    else:
                        columnRange.append("FALSE")

                #Append found value to the empty list        
                else:                    
                    columnRange.append(str(v))
            
            #Append the created list to the table of values for saving
            saveRange.append(columnRange)

            #Reset the created list before continuing the loop
            columnRange=[]

        #Save all found values to the spreadsheet using our created table
        recSheet.update_values((1,1),saveRange)

        self.setValues()

    #---<===>---# END Spreadsheet Building/Access Functions #---<===>---#