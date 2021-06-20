'Created by Jadin Heaston. Sometime around late 2019.


'Forcing me to define variables, instead of "loosely" using them.
Option Explicit

'Creation of a new zip file delay. This allows time for the zip file to be created before other stuff is done, such as copying to it.
'The default is set to 750ms.
dim newZipDelay
newZipDelay = 750

'Time between copying files. This allows for the file to be fully copied before moving to the next one.
'The default is set to 130ms.
dim copyDelay
copyDelay = 130

'Defining variables related to file paths.
dim zippingFilePath, log, tempZip, baseName, baseName1, tempName 
'Solves a problem of the .xml file being named ".shp.xml".
dim shpXML
'Variable used to create the finished dialog.
dim dialog
'Log keeper!
dim logKeeper
'This needs to be done to browse folders.
dim objFSO, objFile, objFolder
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set log = objFSO.OpenTextFile("Recent Log.txt",2, True, 0)
dim zipFileFound

'If this is set to 0, the log is deleted after use.
logKeeper = 0

'Announcing log creation
log.WriteLine "LOG CREATED"

'Creating handle to browse folders.
dim objShell
Set objShell = CreateObject("Shell.Application")


'Finding the file path for files to be zipped.
Set zippingFilePath = objShell.BrowseForFolder(0, "What folder is data being compressed? (This folder should ONLY contain the shapefiles you wish to have compressed.)", 1, 0)
If Not (zippingFilePath Is Nothing) Then
	log.WriteLine "zippingFilePath: Valid"
	'Updating zippingFilePath to be a file path.
	zippingFilePath = zippingFilePath.Self.path + "\"
	'Setting the collection of files to be within "zippingFilePath".
	Set objFolder = objFSO.GetFolder(zippingFilePath).Files
	
	'Initializing the zip found variable to 0
	zipFileFound = 0
	
	'Cycling through every file in the folder. This will eventually include the zip files depending on the size of the files copied.
	For Each objFile In objFolder
		log.WriteLine "Looking at: " & objFile
		'If the file being looked at has the extension .xml, then create 2 versions: One with the .shp and one without it. The .xml in both is absent.
		If objFSO.GetExtensionName(objFile) = "xml" Then
			shpXML = objFSO.GetBaseName(objFile)
			baseName = Left(shpXML,Len(shpXML)-4)
			log.WriteLine baseName & " (XML file shortened)"
		'If the extension of the file is not a .xml, then take off the extension and create the .zip file path.
		Else
			baseName = objFSO.GetBaseName(objFile)
			tempZip = zippingFilePath & baseName & ".zip"
		End if
		
		'This is comparing if the basename has changed between them. The "for each" goes in alphabetical order, so this should mark a change since they are supposed to be identically named.
		If baseName <> baseName1 Then
			zipFileFound = 0
			log.WriteLine baseName & " Has no zip."
		End If
		'If a zip file hasn't been found, create one!
		If zipFileFound = 0 Then
			NewZip(tempZip)
			log.WriteLine "zip created"
			zipFileFound = 1
		End If
		
		'If the file being looked at is a zip file, then do nothing.
		If objFSO.GetExtensionName(objFile) = "zip" Then
			log.WriteLine "Zip file located. Doing nothing: " & baseName
		'If the shape file is an XML, then copy it with a special name/destination.
		ElseIf objFSO.GetExtensionName(objFile) = "xml" Then
		log.WriteLine "Copying... " & zippingFilePath & baseName & "." & objFSO.GetExtensionName(objFile)
		CopyToZip zippingFilePath & shpXML & "." & objFSO.GetExtensionName(objFile), tempZip
		'Otherwise? Copy it to the zip file
		Else
		log.WriteLine "Copying... " & zippingFilePath & baseName & "." & objFSO.GetExtensionName(objFile)
		CopyToZip zippingFilePath & baseName & "." & objFSO.GetExtensionName(objFile), tempZip
	
		End If
		log.WriteLine "zipped: " & objFile & " " & tempZip
		baseName1 = baseName
	Next
	
'Creating the completion dialog.
dialog = MsgBox("Shapefile compression finished!", 0, "Shapefile compression complete!")

'If the browser dialog is cancelled, kill the script. Give a dialog saying so.
Else
	dialog = MsgBox("Shapefile compression failed. No input given!", 0, "Shapefile compression failed!")
	log.WriteLine "zippingFilePath: No input chosen."
	WScript.Quit
End If

'Closing Log.
log.WriteLine "Closing Log."
log.Close


'If the log isn't wanted, it is deleted here.
if logKeeper = 0 Then
	objFSO.deleteFile("Recent Log.txt")
End If


'Closing script.
WScript.Quit



'Functions!
Private Sub NewZip(pathToZipFile)
	'Variable defining.
   Dim zipFSO
   Set zipFSO = CreateObject("Scripting.FileSystemObject")
   Dim zipFile
   'Creating the text file with the path given to us.
   Set zipFile = zipFSO.CreateTextFile(pathToZipFile)
 
	'Making it a zip file.
   zipFile.Write Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, 0)
	
	'Closing the file.
   zipFile.Close
   Set zipFSO = Nothing
  Set zipFile = Nothing
 
	'Add a sleep to ensure files don't begin transferring before the zip is fully created.
   WScript.Sleep newZipDelay
 
End Sub

Private Sub CopyToZip(fileToCopy, fileDest)
	'Creating shell object
	Dim objShell
	Set objShell = CreateObject("shell.application")
	'Creating "File System Object" Object.
	Dim zipFSO
	Set zipFSO = CreateObject("Scripting.FileSystemObject")
	
	'Count the objects within the folder given, assuming a folder is given at all.
	Dim counter

	Dim zipFolder
	Set zipFolder = objShell.NameSpace(fileDest)

	'Increment the counter.
	counter = zipFolder.Items.Count + 1
	zipFolder.CopyHere(fileToCopy)
	
	'Sleep for 100ms for each object being copied to give time for copying.
	While zipFolder.Items.Count < counter
		WScript.Sleep copyDelay
	Wend

End Sub
