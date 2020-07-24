#Zachary Dieter
#Made for the UTMCK
#7/10/20
#This script reads from the JOT form and the JOT form only. This should not be used for anything else. The JOT Form needs to be converted to a CSV for this to work.
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
#If the JOT form's formatting is changed this script will need to be changed.
#-----The things that will need to be changed are "line[#]" as these look at specific columns in the JOT form.
#-----So external managers first and last name are in columns 5,7 so if they get changed to column 9,10. The "line[5], line[7]" need to be changed to "line[9], line[10]"

import csv
import xlsxwriter
from fuzzywuzzy import fuzz

class SVUmtck:
    def __init__(self, email):
        #I am grabbing the email because why not
        #The list for UTSuperVisor names is for misspellings
        #The lwo dictionaries for ExternalNames/Emails are used to link each other tother
        #ExternalSuperVisor names have a key pair of {name:email}
        #ExternalSuperVisor emails have a key pair of {email:[employee]}
        #This is so I can link the correct external supervisor with the correct employees
        self.UtmckEmail = email 
        self.UtmckSuperVisorNames = [] 
        self.ExternalSuperVisorNames = {} 
        self.ExternalSuperVisorEmails = {} 

    #adding a SuperVisorname to the list of possible names
    def add_UTMCKSuperVisorName(self, UTMCKName):
        self.UtmckSuperVisorNames.append(UTMCKName)


    def add_ExternalSuperVisor(self, externalSuperVisorEmail, externalSuperVisorName, employeeName):

        #Checking SuperVisor Email with employees
        #Probably should not have put the logic in the class but it works
        #This does the checking to see if the employee is already in the dictionary in both external dicts
        #So do not do the logic as I already do it for you
        if externalSuperVisorEmail in self.ExternalSuperVisorEmails:
            self.ExternalSuperVisorEmails[externalSuperVisorEmail].append(employeeName)
        else:
            self.ExternalSuperVisorEmails[externalSuperVisorEmail] = []
            self.ExternalSuperVisorEmails[externalSuperVisorEmail].append(employeeName)

        if externalSuperVisorName in self.ExternalSuperVisorEmails:
            self.ExternalSuperVisorNames[externalSuperVisorName].append(externalSuperVisorEmail)
        else:
            self.ExternalSuperVisorNames[externalSuperVisorName] = []
            self.ExternalSuperVisorNames[externalSuperVisorName].append(externalSuperVisorEmail)

    #Call this if you need to print out the whole class 
    def printMe(self):
        print("Email: " + self.UtmckEmail)
        print("Here are my names")
        for UTName in self.UtmckSuperVisorNames:
            print("---" + UTName)

        print("Here are the super visors under me with their employees")

        for ENames in self.ExternalSuperVisorNames:
            print("---" + ENames)
            for Email in self.ExternalSuperVisorNames[ENames]:
                print("+++" + Email) 
                for employee in self.ExternalSuperVisorEmails[Email]:
                    print(">>>" + employee)
            print("\n")

    #Just a helper function to make things clearer
    #Removes a Comma at the end of a string
    def removeLastComma(self, stringToFix):
        return stringToFix[:len(stringToFix) - 1]

   

    #Call this to write one person and there external email, supervisor, and employees 
    #Sheetsetup1 if you wanna call this function
    def WriteMeToExcel1(self, sheetName, currentXPosition, currentYPosition):
        #This is doing what the print does but instead of printing it writes to excel
        #This is using that linking above I described to write the correct people
        superVisorName = ""
        for UTName in self.UtmckSuperVisorNames:
            superVisorName += UTName + ","

        superVisorName = self.removeLastComma(superVisorName)

        for ENames in self.ExternalSuperVisorNames:
            sheetName.write(currentYPosition, currentXPosition,self.UtmckEmail)
            sheetName.write(currentYPosition, currentXPosition+1,superVisorName)
            sheetName.write(currentYPosition, currentXPosition+2,ENames)
            for Email in self.ExternalSuperVisorNames[ENames]:
                sheetName.write(currentYPosition, currentXPosition+3,Email)
                employeeList = ""

                for employee in self.ExternalSuperVisorEmails[Email]:
                    employeeList += employee + ","
                employeeList = self.removeLastComma(employeeList)

                sheetName.write(currentYPosition, currentXPosition+4,employeeList)
            currentYPosition += 1
        return currentYPosition

    #This writes to excel so that all of the external stuff is in one column so I can use mail merge.
    def WriteMeToExcel2(self, sheetName, currentXPosition, currentYPosition):
        superVisorName = ""
        for UTName in self.UtmckSuperVisorNames:
            superVisorName += UTName + ","

        superVisorName = self.removeLastComma(superVisorName)
        sheetName.write(currentYPosition, currentXPosition, self.UtmckEmail)
        sheetName.write(currentYPosition, currentXPosition+1,superVisorName)

        externalInfo = ""
        count = 0
        for ENames in self.ExternalSuperVisorNames:
            for Email in self.ExternalSuperVisorNames[ENames]:
                externalInfo += "Manager: " + ENames + " (" + Email + ")\n"
                for employee in self.ExternalSuperVisorEmails[Email]:
                    externalInfo += str(count) + ". " + employee + "\n"
                    count += 1
                count = 0
        sheetName.write(currentYPosition, currentXPosition+2, externalInfo)                
        return currentYPosition+1



def sheetSetup0():
    #Creates file and worksheet
    OutWorkBook = xlsxwriter.Workbook("DataBaseForm0.xlsx")
    outSheet = OutWorkBook.add_worksheet()


    #Writing Column headers
    outSheet.write("A1", "External Emails")
    return OutWorkBook, outSheet

def sheetSetup1():
    #Creates file and worksheet
    OutWorkBook = xlsxwriter.Workbook("DataBaseForm1.xlsx")
    outSheet = OutWorkBook.add_worksheet()


    #Writing Column headers
    outSheet.write("A1", "UTMCK SuperVisor Email")
    outSheet.write("B1", "UTMCK SuperVisor Names")
    outSheet.write("C1", "External SuperVisor Name")
    outSheet.write("D1", "External SuperVisor Email")
    outSheet.write("E1", "External Employees")
    return OutWorkBook, outSheet
def sheetSetup2():
    #Creates file and worksheet
    OutWorkBook = xlsxwriter.Workbook("DataBaseForm2.xlsx")
    outSheet = OutWorkBook.add_worksheet()


    #Writing Column headers
    outSheet.write("A1", "UTMCK SuperVisor Email")
    outSheet.write("B1", "UTMCK SuperVisor Names")
    outSheet.write("C1", "External Information")
    return OutWorkBook, outSheet

#Function to write to an excel File
#using the sheetSetup1
#These produces a different excel file
#This is the orignal way I was doing it but it does not work for mail merge.
#You can use this to find misspellings if you want
def writeToExcelFile1(UTMCKInfo):

    OutWorkBook, outSheet = sheetSetup1()

    #write data to file
    xPosition = 0
    yPosition = 1

    #Employees
    for Email in UTMCKInfo:
        yPosition = UTMCKInfo[Email].WriteMeToExcel1(outSheet, xPosition, yPosition)


    OutWorkBook.close()


#Function to write to an excel File
#This Setups the excel sheet to be write for a mail merge 
def MailMergeExcelWrite2(UTMCKInfo):

    OutWorkBook, outSheet = sheetSetup2()

    #write data to file
    xPosition = 0
    yPosition = 1

    #Employees
    for Email in UTMCKInfo:
        yPosition = UTMCKInfo[Email].WriteMeToExcel2(outSheet, xPosition, yPosition)


    OutWorkBook.close()

def NamesToLookUpExcelWrite0(NamesListSP, NamesListEE):

    OutWorkBook, outSheet = sheetSetup0()

    xPosition = 0
    yPosition = 1

    NamesListSP = getUniqueList(NamesListSP.split(","))
    NamesListEE = getUniqueList(NamesListEE.split(","))

    for name in NamesListSP:
        outSheet.write(yPosition, xPosition, name)
        yPosition+=1

    for name in NamesListEE:
        outSheet.write(yPosition, xPosition, name)
        yPosition+=1

    OutWorkBook.close()

def RemoveIgnoreListEntries():
    IgnoreList= [] 
    ReturnList = []
    with open("./Vendor Accounts/AccountsIgnore.txt", "r") as f:
       for lines in f.readlines():
            if not lines.strip():
                continue
            else:
                line = lines.strip()
                IgnoreList.append(lines)
    
    with open("./Vendor Accounts/MinireviewedAccounts- Vendor.csv", "r") as f:
        #This is how you read csv's in python
        csv_reader = csv.reader(f)

        #skips the headers
        next(csv_reader)

        for line in csv_reader:
            if line[1] in IgnoreList:
                print("line " + line + " ignored")
                continue
            else:
                ReturnList.append(MakeNameLastFirst(line[0]))

        return ReturnList

#From that list, some people have a () at the end of there name
#This gets rid of it
def MakeNameLastFirst(name):
    #print("S: " + name)
    if("(" in name):
        index = name.find("(")
        name = name[:index - 1]
        #print("f: " + name)
        return name
    else:
        return name

#I want to know if the user is in this list or not
def CheckIfUserEnabled(user, listToCheckAgainst):
    userToBeDeleted = ""
    maxMatch = 0 
    user = user.lower()
    ratio = 0
    for n in listToCheckAgainst:
        ratio = fuzz.ratio(user, n.lower())
        if(ratio >= 90):
            print("Score: " + str(ratio))
            print("User: (" + user + ") (" + n.lower() + ")")
            if(ratio > maxMatch):
                maxMatch = ratio
                userToBeDeleted = n
    if(userToBeDeleted != ""): 
        return userToBeDeleted

    return None


#.lower to make sure there is not mistaken capitalized letters
#.capitalize to make the first letter capitalized
#Helper function to return a string that only has the first letter capitalized
def LowernCap(Name):
    return (Name.lower()).capitalize()

#Just a helper function to make things clearer
#Removes a Comma at the end of a string
def removeLastCommaOutSide(stringToFix):
    return stringToFix[:len(stringToFix) - 1]

#This returns a list with only unique entries
def getUniqueList(l):
    return list(set(l))

def main():

    #TODO: Add argparser
    #ExcelWriterPicker = input("Please pick an Excel writer (0 - 2): ")
    ExcelWriterPicker = 2 

    #Dictionary (key, value)
    DictOfUtmckSV = {}
    NamesToLookUpSPE = ""
    NamesToLookUpEE = ""
    ActiviteAccountsCount = 0
    ListOfActivieAccounts = getUniqueList(RemoveIgnoreListEntries())

    #opening Data.csv as the name f
    with open('Data.csv', 'r') as f:
        #This is how you read csv's in python
        csv_reader = csv.reader(f)

        #skips the headers
        next(csv_reader)

        for line in csv_reader:
            #External Supervisor employees
            #Line 5 is first name external
            #line 7 is last name external
            #line 9 is email external

            #UTMCK point of contact
            #Line 21 is UTMCK point contact first name
            #Line 22 is UTMCK point contact first name
            #line 24 is email external

            #.lower() makes the string all lowercase

            #Just getting the people who answered no
            #if(line[27].lower() == "yes"):
            #    continue

            #Creating the name with email
            EmployeeEmailExternal = LowernCap(line[9])
            EmployeeNameExternal = LowernCap(line[5]) + " " + LowernCap(line[7]) 
            EmployeeNameandEmailExternal = EmployeeNameExternal + " (" + EmployeeEmailExternal + ")"

            #Creates the name for the SuperVisor of external employees
            SuperVisorExternal = LowernCap(line[17]) + " " + LowernCap(line[18])
            #Gets the email for External
            EmailExternal = LowernCap(line[19])

            #Creates the name for the utmck point of contact
            #Makes the Names Capitalized
            SVUFirst = LowernCap(line[21])
            SVULast = LowernCap(line[22])
            SuperVisorUTMCK = LowernCap(line[21]) + " " + LowernCap(line[22]) 

            bValid = CheckIfUserEnabled((SVULast + ", " + SVUFirst), ListOfActivieAccounts)
            if(bValid != None):
                ListOfActivieAccounts.remove(bValid)
                ActiviteAccountsCount +=1
                continue
                

            #Gets the email for UTMCK
            EmailUTMCK = LowernCap(line[24])

            if EmailUTMCK in DictOfUtmckSV:
                temp = DictOfUtmckSV.get(EmailUTMCK)

                if SuperVisorUTMCK not in temp.UtmckSuperVisorNames:
                    temp.add_UTMCKSuperVisorName(SuperVisorUTMCK)

                temp.add_ExternalSuperVisor(EmailExternal, SuperVisorExternal, EmployeeNameandEmailExternal)
                
            else:
                classHolder = SVUmtck(EmailUTMCK)
                classHolder.add_UTMCKSuperVisorName(SuperVisorUTMCK)
                #externalSuperVisorEmail, externalSuperVisorName, employeeName
                classHolder.add_ExternalSuperVisor(EmailExternal, SuperVisorExternal, EmployeeNameExternal)
                DictOfUtmckSV[EmailUTMCK] = classHolder

            if(EmployeeEmailExternal != ""):
                NamesToLookUpEE += EmployeeEmailExternal + ","            

            if(EmailExternal != ""):
                NamesToLookUpSPE += EmailExternal + ","


        NamesToLookUpSPE = removeLastCommaOutSide(NamesToLookUpSPE)
        NamesToLookUpEE = removeLastCommaOutSide(NamesToLookUpEE)

        print("Actitive Accounts Found: " + str(ActiviteAccountsCount)) 

        if(ExcelWriterPicker == 0): 
            NamesToLookUpExcelWrite0(NamesToLookUpSPE, NamesToLookUpEE)
        elif(ExcelWriterPicker == 1):
            writeToExcelFile1(DictOfUtmckSV)
        elif(ExcelWriterPicker == 2):
            MailMergeExcelWrite2(DictOfUtmckSV)
       



if __name__ == "__main__":
    main()
