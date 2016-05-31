Option Explicit

Const WINDOW_HANDLE = 0 ' Must ALWAYS be 0

const olMailItem   = 0
Const MY_DOCUMENTS = &H5&
Const MY_PICTURES  = &H27&
Const MY_COMPUTER  = &H11&
const CANCEL_BTN   = 2

Dim cancel
Dim ToAddress, MessageSubject, MessageBody, myRecipient

dim FSO, objFolder
Set FSO = CreateObject("Scripting.FileSystemObject")

ToAddress= Inputbox ("Email you want to send jpgs to",  "To", "john.michaloski@nist.gov")
MessageSubject = "Graduation Pictures from 35mm camera"
MessageBody = ""


'' Need trailing slash!
'objFolder = "C:\Users\michalos\Pictures\Sample Pictures\"
objFolder =  BrowseFolder( "MY PICTURES" , False) & "\"
wscript.echo objFolder

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function BrowseFolder( myStartLocation, blnSimpleDialog )

    Dim numOptions, objFolder, objFolderItem
    Dim objPath, objShell, strPath, strPrompt
   ' Set the options for the dialog window
    strPrompt = "Select a folder:"
    If blnSimpleDialog = True Then
        numOptions = 0      ' Simple dialog
    Else
        numOptions = &H10&  ' Additional text field to type folder path
    End If
    
    ' Create a Windows Shell object
    Set objShell = CreateObject( "Shell.Application" )

    ' If specified, convert "My Computer" to a valid
    ' path for the Windows Shell's BrowseFolder method
    If UCase( myStartLocation ) = "MY COMPUTER" Then
        Set objFolder = objShell.Namespace( MY_COMPUTER )
        Set objFolderItem = objFolder.Self
        strPath = objFolderItem.Path
    ElseIf UCase( myStartLocation ) = "MY PICTURES" Then
        Set objFolder = objShell.Namespace( MY_PICTURES )
        Set objFolderItem = objFolder.Self
        strPath = objFolderItem.Path
    Else
        strPath = myStartLocation
    End If

    Set objFolder = objShell.BrowseForFolder( WINDOW_HANDLE, strPrompt, _
                                              numOptions, strPath )

    ' Quit if no folder was selected
    If objFolder Is Nothing Then
        BrowseFolder = ""
	MsgBox "Quitting"
        WScript.Quit(1)
    End If

    ' Retrieve the path of the selected folder
    Set objFolderItem = objFolder.Self
    objPath = objFolderItem.Path

    ' Return the path of the selected folder
    BrowseFolder = objPath
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SendEmail3(MessageAttachment)

	Dim ol, ns, newMail

	Set ol = WScript.CreateObject("Outlook.Application")
	
	Set ns = ol.getNamespace("MAPI")

	ns.logon "","",true,false

	Set newMail = ol.CreateItem(olMailItem)
	newMail.Subject = MessageSubject
	newMail.Body = MessageBody & vbCrLf

	' validate the recipient, just in case...
	Set myRecipient = ns.CreateRecipient(ToAddress)
	myRecipient.Resolve
	If Not myRecipient.Resolved Then
   		MsgBox "unknown recipient - Quitting"
		WScript.Quit(1) 
	Else
  	 	newMail.Recipients.Add(myRecipient)
  		newMail.Attachments.Add(objFolder & MessageAttachment)
 
   		cancel = MsgBox( "Send " & objFolder & MessageAttachment, 1)
		if cancel = CANCEL_BTN then 
			MsgBox "Quitting"
			WScript.Quit(1) 
		End if

 		newMail.Send
  		
	End If

	Set ol = Nothing
end sub

''''''''''''''''''''''''''''''''
Sub SendPictures(fFolder)
    dim objFolder, colFiles, objFile
    Set objFolder = FSO.GetFolder(fFolder)
    Set colFiles = objFolder.Files
    For Each objFile in colFiles
        If UCase(FSO.GetExtensionName(objFile.name)) = "JPG" Then
            SendEmail3 objFile.Name
        End If
    Next
End Sub

SendPictures  objFolder

