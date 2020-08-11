import ExpirationsModule
import pyodbc

# connect to EMPLOYEE DATA
connection = pyodbc.connect(
    'Driver={SQL Server};Server=infinitydb.intel.com,3180;Database=EmployeeData;Trusted_Connection=yes;')  # DFIT

cursor = connection.cursor()
query = ('SELECT [id],[name],[isActive] FROM [InfinityGlobal].[dbo].[vwSiteFltr_tArea] where name not like \'%zTraining%\' and name not like \'%xshared%\' and name not like \'%zshared (TTL access)%\' and trainingPOCEntitlement <> \'\' and organizationId = 82 or name = \'VSG\' order by name asc')

for row in cursor.execute(query):
    #email = row[7]
    #FullName.add(row[70])
    #listWWID.add(row[0])
    print(str(row[0]) + ':' + row[1])
print('done')
#print(ExpirationsModule.multilineCreate(listWWID))
#print(ExpirationsModule.multilineCreate(FullName))