' This script is for testing purposes. It reads .XML file and displays information in the table.
' Rule 1: Read .XML files
' Rule 2: Ignore other (non .XML) files
' Created by Ignas Karaliunas
' Creation date: 2021/08/04
' Last edit date: 2021/08/05

'Declare variables
'Create IDs for items (names, values)
dim nodeID, nodeID1
'Create object and store RQD information


'Load .Xml file
Set objXMLDoc = CreateObject("Microsoft.XMLDOM") 
objXMLDoc.async = False 
objXMLDoc.load("C:\Test\SASKAITA SKIRTA TESTUOTI.xml")

'Create Root
Set Root = objXMLDoc.documentElement 
' 'Use root to obtain all elements within "name" tag
Set NodeList = Root.getElementsByTagName("name") 

'Create two nodes for storing information about names and values
Set colNodes = objXMLDoc.selectNodes ("/document/parameters/par/name")
Set colNodes1 = objXMLDoc.selectNodes ("/document/parameters/par/value")

'Set IDs to 0
nodeID = 0
nodeID1 = 0

'Cycle through both nodes and assign Values to IDs
For Each objNode in colNodes
nodeID += nodeID

For Each objNode in colNodes1
nodeID1 ++

'Cycle through NodeList (tag name "name") and look for tags "Pirkejas", "DokumentoSerija", "DokumentoNr", "SumaSuPVM", "DokumentoData"
For Each objNode in colNodes1
	if objNode.text = "Pirkejas" and nodeID = nodeID1 or objNode.text = "DokumentoSerija" and nodeID = nodeID1  or objNode.text = "DokumentoNr" and nodeID = nodeID1 or objNode.text = "SumaSuPVM" and nodeID = nodeID1 or objNode.text = "DokumentoData" and nodeID = nodeID1  then	
	'Store results into arrays

	end if

'Use array to create name for .pdf file	call it DocName

'Add "C:\Test\Archyvas\" directory to the DocName, call it pdfName

	'If "C:\Test\Archyvas\" doesn't exist create new folder
	' <%
	' dim filesys, newfolder, newfolderpath
	' newfolderpath = "C:\Test\Archyvas\"
	' set filesys=CreateObject("Scripting.FileSystemObject")
	' If Not filesys.FolderExists(newfolderpath) Then
	' Set newfolder = filesys.CreateFolder(newfolderpath)
	' Response.Write("A new folder has been created at: " newfolderpath)
	' End If
	' %>

'DocName = "C:\Test\Archyvas\" + DocName + ".pdf"

'Save object 
'objItem.SaveAs pdfName, olTXT
next