import win32com.client as win32
import os

class DM:

    def __init__(self, username, password):
        login = win32.gencache.EnsureDispatch('PCDClient.PCDLogin')
        login.AddLogin(0, 'FY21_ALL_STAFF', username, password)
        login.Execute();
        dst = login.GetDST()
        self.dst = dst
               
    def return_all_properties(self, doc_id, version_label = 'max', library  = 'FY19_ALL_STAFF'):
        
        pDocObject =  win32.gencache.EnsureDispatch('PCDClient.PCDDocObject') 
        pDocObject.SetDST(dm.dst)
        pDocObject.SetObjectType("SEARCH")
        pDocObject.SetProperty("%TARGET_LIBRARY", library)
        pDocObject.SetProperty("%OBJECT_IDENTIFIER", doc_id)
        pDocObject.Fetch()

        pDocObject.GetProperties()
        docPropList = pDocObject.GetReturnProperties();
        docPropList.BeginIter();
        propListSize = docPropList.GetSize();
        
        properties = {}
        
        for i in range(0, propListSize):
            properties[docPropList.GetCurrentPropertyName()] = docPropList.GetCurrentPropertyValue()
            docPropList.NextProperty(); 

        return properties   

    def full_search(self, search_info, library = 'FY19_ALL_STAFF'):
        
        dmSearch = win32.gencache.EnsureDispatch('PCDClient.PCDSearch')    
        dmSearch.SetDST(self.dst);
        dmSearch.SetSearchObject("SEARCH");
        dmSearch.AddReturnProperty("DOCNUM");
        dmSearch.AddReturnProperty("DOCNAME");
        
        for key in search_info:
            dmSearch.AddSearchCriteria(key, search_info[key]);
            
        dmSearch.Execute();
        dmSearch.GetRowsFound()
        
        return_info = {}

        for i in range(0, dmSearch.GetRowsFound()):
            dmSearch.NextRow();
            return_info[dmSearch.GetPropertyValue("DOCNAME")] = dmSearch.GetPropertyValue("DOCNUM")
        return return_info
        
        
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
        
        search = win32.gencache.EnsureDispatch('PCDClient.PCDGetDoc')
        
        search.SetDST(self.dst)
        search.AddSearchCriteria("%TARGET_LIBRARY", library)
        search.AddSearchCriteria("%DOCUMENT_NUMBER", doc_id)
        
        verID = self.version_search(doc_id, version_label)
        
        search.AddSearchCriteria("%VERSION_ID", verID)
        search.Execute()
        search.SetRow(1)    
        return search
    
    def download_doc(self, doc_id,  download_filename, version_label = 'max', library = 'FY19_ALL_STAFF'):
        
        objGetStream = win32.gencache.EnsureDispatch('PCDClient.PCDGetStream')        
        search = self.return_doc(doc_id)
        objGetStream = search.GetPropertyValue("%CONTENT")
        fileSize = objGetStream.GetPropertyValue("%ISTREAM_STATSTG_CBSIZE_LOWPART") 
        bytesRead = 1;
        file_to_write = open(download_filename, 'wb')        
        while bytesRead != 0:
            print ('here!')
            stream = objGetStream.Read(fileSize, bytesRead)
            bytes_to_write = stream[0]
            if (bytesRead != 0):
                file_to_write.write(bytes_to_write)
            bytesRead = stream[1]
        file_to_write.close()
        
    def create_profile(self, profile_info):
        
        doc = win32.gencache.EnsureDispatch('PCDClient.PCDDocObject')  
        doc.SetDST(self.dst)
                
        doc.SetObjectType(profile_info['formType'])

        doc.SetProperty("%TARGET_LIBRARY", profile_info['Library'])
        doc.SetProperty("AUTHOR_ID", profile_info['Author'])
        doc.SetProperty("DOCNAME", profile_info['FileName'])
        doc.SetProperty("DOCDATE", profile_info['Date'])

        doc.SetProperty("TYPE_ID", profile_info['DocumentType'])
        doc.SetProperty("JOB_CODE", profile_info['JobCode'])
        doc.SetProperty("ORG_ID", profile_info['Agency'])
        doc.SetProperty("GOAL_ID", profile_info['Goal'])
        doc.SetProperty("TYPIST_ID", profile_info['Typist'])

        doc.SetProperty("ABSTRACT", profile_info['Abstract'])
        doc.SetProperty("APP_ID", profile_info['App'])

        doc.SetProperty("GAORM_CAT_NAME",  profile_info['Cat'])
        doc.SetProperty("GAORM_FUNC_NAME", profile_info['Func'])
        doc.SetProperty("GAORM_ACT_NAME", profile_info['Act'])
        doc.SetProperty("GAORM_PART_NAME", profile_info['Part'])    
        
        doc.SetProperty("%VERIFY_ONLY", "%NO")
        
        doc.Create()
        
        documentNumber = str(doc.GetReturnProperty("%OBJECT_IDENTIFIER"))
        versionID = str(doc.GetReturnProperty("%VERSION_ID"))    
        

        return documentNumber, versionID
        
        
    def upload_doc(self, upload_filepath, profile_info):
        
        docId, versionID = self.create_profile(profile_info)
        
        
        dmPutDoc =  win32.gencache.EnsureDispatch('PCDClient.PCDPutDoc')
        dmPutDoc.SetDST(self.dst);
        dmPutDoc.AddSearchCriteria("%TARGET_LIBRARY", profile_info['Library']);
        dmPutDoc.AddSearchCriteria("%DOCUMENT_NUMBER", docId);
        dmPutDoc.AddSearchCriteria("%VERSION_ID", versionID);
        dmPutDoc.Execute();
        dmPutDoc.NextRow();
        
        dmPutStream = dmPutDoc.GetPropertyValue("%CONTENT")
        objPutStream = win32.gencache.EnsureDispatch('PCDClient.PCDPutStream')  
        filesize = os.path.getsize(upload_filepath)

        with open(upload_filepath, 'rb') as f:
            read_data = f.read()
            dmPutStream.Write(read_data, filesize)
        dmPutStream.SetComplete();     
        
        self.unlock_doc(docId, profile_info)
        
        return docId
        
        
        
    def unlock_doc(self, docId, profile_info):
    
        dmDoc = win32.gencache.EnsureDispatch('PCDClient.PCDDocObject')  
        dmDoc.SetDST(self.dst);
        dmDoc.SetObjectType(profile_info['formType']);
        dmDoc.SetProperty("%TARGET_LIBRARY", profile_info['Library']);
        dmDoc.SetProperty("%OBJECT_IDENTIFIER", docId);
        dmDoc.SetProperty("%STATUS", "%UNLOCK");
        dmDoc.Update();


