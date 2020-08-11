#!/usr/bin/python

# import statements
import pyodbc
import win32com.client as win32

#Print entire set in co
def printSet(setToPrint):
    for item in setToPrint:
        print(item)

def multilineCreate(stringToTest):
    output = ''
    for item in stringToTest:
        output += item + '\n'
    return output

def getExpirations(listWWID):
    # connect to INFINITY GLOBAL
    connection = pyodbc.connect(
        'Driver={SQL Server};Server=infinitydb.intel.com,3180;Database=InfinityGlobal;Trusted_Connection=yes;')  # DFIT

    ExpiredSkills = 0
    ListExpiredSkills = dict()
    for item in listWWID:
        ListExpiredSkills = dict()
        cursor = connection.cursor()
        query = (
                    'SELECT [id],[name],[courseCode],sc.wwid,sc.completed,DATEADD(DAY, expirationDays, completed) as Expires FROM [InfinityGlobal].[dbo].[vwSiteFltr_tSkill] left join [InfinityGlobal].[dbo].[vwSiteFltr_tSkillCompletion] sc on id = sc.skillId Where DATEADD(DAY, expirationDays, completed) < GETDATE() and sc.wwid = ' + str(item))
        for row in cursor.execute(query):
            #print('row = %r' % (row,))
            if row[3] not in ListExpiredSkills:
                ListExpiredSkills.__setitem__(row[2], row[1])
        print('done')
        print('You have ' + str(len(ListExpiredSkills)) + ' expired skills')

        strExpiredSkills = 'You have ' + str(ExpiredSkills) + ' expired skills'

        multilineString = multilineCreate(ListExpiredSkills.values())
        print(multilineString)
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        # mail.To = email
        mail.To = 'gregorio.vargas@intel.com'
        mail.Subject = 'Expired Skills'
        mail.Body = 'Report for ' + item + '\n\n' + multilineString

    connection.close()