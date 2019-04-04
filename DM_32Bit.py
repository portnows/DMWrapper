
# coding: utf-8
import win32com.client as win32
dmdoc = win32.gencache.EnsureDispatch('PCDClient.PCDDocObject')
objGetDoc = win32.gencache.EnsureDispatch('PCDClient.PCDGetDoc')
objGetStream = win32.gencache.EnsureDispatch('PCDClient.PCDGetStream')

login = win32.gencache.EnsureDispatch('PCDClient.PCDLogin')
login.AddLogin(0, 'FY19_ALL_STAFF', 'portnows', 'PASSWORD')
login.Execute();
dst = login.GetDST()

search = win32.gencache.EnsureDispatch('PCDClient.PCDSearch')

search.SetDST(dst)
search.AddSearchLib('FY19_ALL_STAFF')
search.SetSearchObject("VersionsSearch")
search.AddSearchCriteria("%OBJECT_IDENTIFIER", '459381')

search.AddReturnProperty("VERSION_ID")
search.AddReturnProperty("VERSION_LABEL")
search.Execute()
rowsfound = search.GetRowsFound();
i = 0
versionLabel = '1'

while (i < rowsfound):
    search.SetRow((i + 1))
    if (search.GetPropertyValue("VERSION_LABEL") == versionLabel):
        verID = search.GetPropertyValue("VERSION_ID")
        search.ReleaseResults()      
    i = (i + 1)


objGetDoc.AddSearchCriteria("%TARGET_LIBRARY", 'FY19_ALL_STAFF')
objGetDoc.AddSearchCriteria("%DOCUMENT_NUMBER", '459381')
objGetDoc.AddSearchCriteria("%VERSION_ID", verID)

objGetDoc.Execute()
objGetDoc.ErrNumber
objGetDoc.SetRow(1)

objGetStream = objGetDoc.GetPropertyValue("%CONTENT")
fileSize = objGetStream.GetPropertyValue("%ISTREAM_STATSTG_CBSIZE_LOWPART")

bytesRead = 1;

filename = 'C:/Users/portnows/' + 'ThisIsIt' + '.xlsx'
print(type(bytesRead))
print(type(fileSize))

file_to_write = open(filename, 'wb')


while bytesRead != 0:
    stream = objGetStream.Read(fileSize, bytesRead)
    bytes_to_write = stream[0]
    if (bytesRead != 0):
        file_to_write.write(bytes_to_write)
    bytesRead = stream[1]
file_to_write.close()
