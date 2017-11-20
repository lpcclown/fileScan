Function FilesTree(sPath)    
' Loop all file under one folder
    Set oFso = CreateObject("Scripting.FileSystemObject")    
    Set oFolder = oFso.GetFolder(sPath)    
    Set oSubFolders = oFolder.SubFolders    
    Set oFiles = oFolder.Files  
    For Each oFile In oFiles  
		MsgBox("go in loop")
		objCSVFile.Write chr(34) & oFile.Path & chr(34) & vbTab 
		objCSVFile.Write chr(34) & oFile.ParentFolder & chr(34) & vbTab 
		objCSVFile.Write chr(34) & oFile.Name & chr(34) & vbTab 
		objCSVFile.Write chr(34) & oFile.DateCreated & chr(34) & vbTab 
		objCSVFile.Write chr(34) & oFile.DateLastAccessed & chr(34) & vbTab 
		objCSVFile.Write chr(34) & oFile.DateLastModified & chr(34) & vbTab 
		objCSVFile.Write chr(34) & oFile.Size & chr(34) & vbTab
		objCSVFile.Write chr(34) & oFile.Type & chr(34) & vbTab
		objCSVFile.Write chr(34) & oFso.getextensionname(oFile.Path) & chr(34) & vbTab
		' Get file owner
		Set objShell = CreateObject ("Shell.Application")
		Set objFolder = objShell.Namespace (sPath)
		For Each strFileName in objFolder.Items
			if objFolder.GetDetailsOf (strFileName, 0) = oFile.Name Then
				objCSVFile.Write chr(34) & objFolder.GetDetailsOf (strFileName, 10) & chr(34) 
			End If
		Next
		objCSVFile.Writeline
    Next    
        
    For Each oSubFolder In oSubFolders    
        'WScript.Echo oSubFolder.Path    ' Output the subfolder  
        FilesTree(oSubFolder.Path)'recurssion    
    Next    
        
    Set oFolder = Nothing    
    Set oSubFolders = Nothing    
    Set oFso = Nothing    
End Function    

Dim csvFilePath,csvColumns
Const ForWriting = 2
' Create new CSV file 
csvFilePath =".\FileInfoSummary.csv"
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objCSVFile = objFSO.CreateTextFile(csvFilePath, ForWriting, True)
' Write comma delimited list of columns in new CSV file.
csvColumns = chr(34) & "FilePathAndName" & chr(34) & vbTab & _
				chr(34) & "ParentFolder" & chr(34) & vbTab & _
				chr(34) & "Name" & chr(34) & vbTab & _
				chr(34) & "DateCreated" & chr(34) & vbTab & _
				chr(34) & "DateLastAccessed" & chr(34) & vbTab & _
				chr(34) & "DateLastModified" & chr(34) & vbTab & _
				chr(34) & "Size" & chr(34) & vbTab & _
				chr(34) & "Type" & chr(34) & vbTab & _
				chr(34) & "Suffix" & chr(34) & vbTab & _
				chr(34) & "Owner" & chr(34) & vbTab
				
objCSVFile.Write csvColumns
objCSVFile.Writeline
FilesTree("C:\Users\28153\Documents\17_American_Study\02_After_Arrive_US\P14_File_Scan") ' Call the func 