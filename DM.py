
# coding: utf-8

import win32com.client as win32

class DM:

    def __init__(self, username, password):
        login = win32.gencache.EnsureDispatch('PCDClient.PCDLogin')
        login.AddLogin(0, 'FY19_ALL_STAFF', username, password)
        login.Execute();
        dst = login.GetDST()
        self.dst = dst
        
        
    def version_search(self, doc_id, version_label = 'max', library = 'FY19_ALL_STAFF'):
    
        search = win32.gencache.EnsureDispatch('PCDClient.PCDSearch')    
        search.SetDST(self.dst)
        search.AddSearchLib('FY19_ALL_STAFF')
        search.SetSearchObject("VersionsSearch")
        search.AddSearchCriteria("%OBJECT_IDENTIFIER", doc_id)

        search.AddReturnProperty("VERSION_ID")
        search.AddReturnProperty("VERSION_LABEL")
        search.Execute()
    
        rowsfound = search.GetRowsFound();


        if version_label == 'max':
            search.SetRow(rowsfound)
            verID = search.GetPropertyValue("VERSION_ID")
            search.ReleaseResults()   
        else:
            i = 0
            while (i < rowsfound):
                search.SetRow((i + 1))
                if (search.GetPropertyValue("VERSION_LABEL") == versionLabel):
                    verID = search.GetPropertyValue("VERSION_ID")
                    search.ReleaseResults()      
                i = (i + 1)
        return verID

    
    def return_doc(self, doc_id, version_label = 'max', library = 'FY19_ALL_STAFF'):
        
        objGetDoc = win32.gencache.EnsureDispatch('PCDClient.PCDGetDoc')
        
        objGetDoc.SetDST(self.dst)
        objGetDoc.AddSearchCriteria("%TARGET_LIBRARY", library)
        objGetDoc.AddSearchCriteria("%DOCUMENT_NUMBER", doc_id)
        
        verID = self.version_search(doc_id, version_label)
        
        objGetDoc.AddSearchCriteria("%VERSION_ID", verID)
        objGetDoc.Execute()
        objGetDoc.SetRow(1)    
        return objGetDoc
    
    def download_doc(self, doc_id,  download_filename, version_label = 'max', library = 'FY19_ALL_STAFF'):
        
        objGetStream = win32.gencache.EnsureDispatch('PCDClient.PCDGetStream')        
        objGetDoc = self.return_doc(doc_id)
        objGetStream = objGetDoc.GetPropertyValue("%CONTENT")
        fileSize = objGetStream.GetPropertyValue("%ISTREAM_STATSTG_CBSIZE_LOWPART")        
        bytesRead = 1;
        file_to_write = open(download_filename, 'wb')        
        while bytesRead != 0:
            stream = objGetStream.Read(fileSize, bytesRead)
            bytes_to_write = stream[0]
            if (bytesRead != 0):
                file_to_write.write(bytes_to_write)
            bytesRead = stream[1]
        file_to_write.close()

dm = DM('portnows', "PASSWORD")

dm.download_doc('459381', 'C:/Users/portnows/WhereAMI.xlsx')

