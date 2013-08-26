import os
import sys
import datetime
import xlrd, xlwt

##################################
## BIG FUCKING NOTE TO SELF!!!! ##
##    [0] = OLD/ARCHIVE FILE    ##
##        [1] = NEW FILE        ##
##################################
    
##################################
## BIG FUCKING NOTE TO SELF!!!! ##
##   [0] = TERM CODE ADMIT      ##
##        [1] = USERID          ##
##    [2] = DATE PROCESSED      ##
##################################


def deleteFile(fpath):
    os.remove(fpath)
    return


def getRunDate():
    now = datetime.datetime.now()
    todayDate = str(now.strftime("%m-%d"))
    return todayDate


def getMaxTermCode(termList):
    mx = 0
    for num in termList:
        if num > mx:
            mx = num
    return mx


def copyToNewArray(list1, list2, offset):
    list1[0].append(list2[0][offset])
    list1[1].append(list2[1][offset])
    list1[2].append(list2[2][offset])

    return list1


def writeNewExcelFile(oldList, diffList):
    fpath = 'Source Files/archive.xls'

    workbook = xlwt.Workbook(encoding = 'ascii')
    sheet = workbook.add_sheet('Sheet')
    sheet.write(0,0,'Term Code Admit')
    sheet.write(0,1,'Userid')
    sheet.write(0,2,'Date Processed')

    for i in range(0, len(oldList[0])):
        sheet.write(i + 1,0,oldList[0][i])
        sheet.write(i + 1,1,oldList[1][i])
        sheet.write(i + 1,2,oldList[2][i])

    i += 2

    for x in range(0, len(diffList[0])):
        sheet.write(i + x,0,diffList[0][x])
        sheet.write(i + x,1,diffList[1][x])
        sheet.write(i + x,2,diffList[2][x])

    workbook.save(fpath)


def writeOutputFile(diffList):
    currTerm = [[],[],[]]
    pastTerm = [[],[],[]]
    maxTerm = getMaxTermCode(diffList[0])
    

    for x in range(0, len(diffList[0])):
        if diffList[0][x] < maxTerm:
            pastTerm = copyToNewArray(pastTerm, diffList, x)

        else:
            currTerm = copyToNewArray(currTerm, diffList, x)

    # now that we're all split up, let's write some shit!
    writeFile = open("/Users/mwl36/Developer/Hyperion Enrollment/Archive/New Additions" + getRunDate() + ".txt", 'w')

    if len(currTerm[0]):
        writeFile.write("NEW STUDENT ORIENTATION\n")
        writeFile.write("======================================\n")
        writeFile.write("Usernames to Enroll:\n")

        for i in range(0, len(currTerm[1])):
            if i == len(currTerm[1]) - 1:
                writeFile.write(currTerm[1][i])
            else:
                writeFile.write(currTerm[1][i] + ", ")

        writeFile.write("\n\n")
        writeFile.write("Emails:\n")

        for i in range(0, len(currTerm[1])):
            if i == len(currTerm[1]) - 1:
                writeFile.write(currTerm[1][i] + "@drexel.edu")
            else:
                writeFile.write(currTerm[1][i] + "@drexel.edu, ")

    if len(pastTerm[0]):     
        writeFile.write("\n======================================\n\n\n\n")
        writeFile.write("TRANSITION STUDENT ORIENTATION\n")
        writeFile.write("======================================\n")
        writeFile.write("Usernames to Enroll:\n")

        for i in range(0, len(pastTerm[1])):
            if i == len(pastTerm[1]) - 1:
                writeFile.write(pastTerm[1][i])
            else:
                writeFile.write(pastTerm[1][i] + ", ")

        writeFile.write("\n\n")
        writeFile.write("Emails:\n")

        for i in range(0, len(pastTerm[1])):
            if i == len(pastTerm[1]) - 1:
                writeFile.write(pastTerm[1][i] + "@drexel.edu")
            else:
                writeFile.write(pastTerm[1][i] + "@drexel.edu, ")

    writeFile.close()
    
    return


def compareStudentLists(oldList, newList):
    nl_length = len(newList[1])
    diffList = [[],[],[]]

    for i in range(0, nl_length):
        
        if not (newList[1][i] in oldList[1]):
            # meaning, we found something in newList that's
            # not already in oldList
            diffList = copyToNewArray(diffList, newList, i)

    return diffList

def getOldStudentInfo(sheet, crow):
    
    termcode = int(sheet.cell_value(crow, 0))
    userid = str(sheet.cell_value(crow, 1))
    dateproc = str(sheet.cell_value(crow, 2))

    return termcode, userid, dateproc


def getNewStudentInfo(sheet, crow):
    now = datetime.datetime.now()

    termcode = int(sheet.cell_value(crow, 35))
    userid = str(sheet.cell_value(crow, 86))
    dateproc = str(now.strftime("%m-%d-%Y"))

    return termcode, userid, dateproc


def buildStudentList(filepath, isOld):
    
    studentList = [[],[],[]]

    workbook = xlrd.open_workbook(filepath)
    sheet = workbook.sheet_by_name('Sheet')
    num_rows = sheet.nrows
    curr_row = 1

    # sheet is now open
    
    if isOld:
        # we know that this file has been formatted
        # and created by our system, so we're safe
        # to assume things about the columns

        while curr_row < num_rows:
            tca, userid, dateproc = getOldStudentInfo(sheet, curr_row)
            studentList[0].append(tca)
            studentList[1].append(userid)
            studentList[2].append(dateproc)
            curr_row +=1

    else:
        # this must be fresh from Hyperion, with
        # tons of extraneous shit we need to remove

        while curr_row < num_rows:
            tca, userid, dateproc = getNewStudentInfo(sheet, curr_row)
            studentList[0].append(tca)
            studentList[1].append(userid)
            studentList[2].append(dateproc)
            curr_row +=1

    print len(studentList[0])

    return studentList

        
def getModDate(filename):
    t = os.path.getmtime(filename)
    return datetime.datetime.fromtimestamp(t)


def idOldVsNew(filelist):
    file1 = filelist[0]
    file2 = filelist[1]
    orgList = ["",""]

    if getModDate(file1) > getModDate(file2):
        #then file1 is newer
        orgList[0] = file2
        orgList[1] = file1
    else:
        orgList = filelist

    return orgList


def getFileList():
    dirList = os.listdir("/Users/mwl36/Developer/Hyperion Enrollment/Source Files/")
    
    # we use 1 and 2 to avoid the damn .DS_Store
    
    dirList[0] = "/Users/mwl36/Developer/Hyperion Enrollment/Source Files/" + dirList[1]
    dirList[1] = "/Users/mwl36/Developer/Hyperion Enrollment/Source Files/" + dirList[2]
    
    result = idOldVsNew(dirList)

    return result


def main():
    fileList = getFileList()

    fileList = [fileList[0], fileList[1]]

    oldList = buildStudentList(fileList[0], True)
    newList = buildStudentList(fileList[1], False)   
    
    diffList = compareStudentLists(oldList, newList)

    writeOutputFile(diffList)

    for item in fileList:
        deleteFile(item)

    writeNewExcelFile(oldList, diffList)

main()
