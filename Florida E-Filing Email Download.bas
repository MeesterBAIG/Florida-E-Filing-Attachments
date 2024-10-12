Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" ( _
        ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#Else
    Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" ( _
        ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#End If

Sub DownloadFilesToSubjectFolder()
    Dim olItem As Outlook.MailItem
    Dim selectedItems As Outlook.Selection
    Dim emailBody As String
    Dim emailHTML As String
    Dim subjectText As String
    Dim folderName As String
    Dim baseFolder As String
    Dim linkStart As String
    Dim extractedLink As String
    Dim filePath As String
    Dim emailDate As String
    Dim fileName As String
    Dim downloadStatus As Long
    Dim matches As Object
    Dim i As Long
    Dim linkCounter As Long
    Dim htmlDoc As Object
    Dim htmlAnchors As Object
    Dim anchor As Object
    Dim anchorText As String
    Dim downloadedFiles As String ' To store names of downloaded files
    
    ' Base folder where subject folders will be created
    baseFolder = "C:\CourtDocuments\" ' Adjust this as needed
    
    ' Check if the base folder exists, if not, create it
    If Dir(baseFolder, vbDirectory) = "" Then
        MkDir baseFolder
    End If
    
    ' URL pattern to match
    linkStart = "https://url.avanan.click/v2/r01/___https://www.myflcourtaccess.com/nefdocuments/document.nefdd?nai="
    
    ' Get the selected emails
    Set selectedItems = Application.ActiveExplorer.Selection
    
    ' Check if any emails are selected
    If selectedItems.Count = 0 Then
        MsgBox "No emails selected. Please select at least one email.", vbExclamation
        Exit Sub
    End If
    
    ' Loop through each selected email
    For Each olItem In selectedItems
        If TypeName(olItem) = "MailItem" Then
            ' Get the email subject and remove "SERVICE OF COURT DOCUMENT CASE NUMBER " from it
            subjectText = olItem.Subject
            folderName = Replace(subjectText, "SERVICE OF COURT DOCUMENT CASE NUMBER ", "")
            folderName = CleanFileName(folderName)
            
            ' Define the full path of the folder
            Dim caseFolder As String
            caseFolder = baseFolder & folderName
            
            ' Check if the folder exists, if not, create it
            If Dir(caseFolder, vbDirectory) = "" Then
                MkDir caseFolder
            End If
            
            ' Get the HTML body of the email for hyperlink extraction
            emailHTML = olItem.HTMLBody
            
            ' Create an HTML document object to parse the HTML content
            Set htmlDoc = CreateObject("HTMLFile")
            htmlDoc.Write emailHTML
            
            ' Get all anchor (<a>) tags from the HTML
            Set htmlAnchors = htmlDoc.getElementsByTagName("a")
            
            ' Get the email date in yyyy-mm-dd format
            emailDate = Format(olItem.ReceivedTime, "yyyy-mm-dd")
            
            ' Initialize link counter to track how many links have been processed
            linkCounter = 0
            
            ' Loop through all the anchor tags to find links that match the pattern
            For Each anchor In htmlAnchors
                extractedLink = anchor.href
                anchorText = anchor.innerText ' Get the text associated with the hyperlink
                
                ' Skip the first link in the email
                If linkCounter > 0 Then
                    ' Check if the link starts with the desired pattern
                    If InStr(extractedLink, linkStart) > 0 Then
                        ' Set the save path with the date and anchor text
                        fileName = emailDate & "_" & anchorText & ".pdf"  ' Adjust the file extension as needed
                        filePath = caseFolder & "\" & CleanFileName(fileName)
                        
                        ' Download the file
                        downloadStatus = URLDownloadToFile(0, extractedLink, filePath, 0, 0)
                        
                        If downloadStatus = 0 Then
                            ' Add the downloaded file name to the summary list
                            downloadedFiles = downloadedFiles & fileName & vbCrLf
                        End If
                    End If
                End If
                
                ' Increment the link counter
                linkCounter = linkCounter + 1
            Next anchor
        End If
    Next olItem
    
    ' Display a summary of all the downloaded files at the end
    If downloadedFiles <> "" Then
        MsgBox "Files downloaded successfully:" & vbCrLf & downloadedFiles, vbInformation
    Else
        MsgBox "No files were downloaded.", vbExclamation
    End If
End Sub

' Function to clean the folder name (removes invalid characters)
Function CleanFileName(fileName As String) As String
    Dim invalidChars As String
    Dim i As Integer
    
    invalidChars = "/\:?*""<>|"
    
    ' Remove any invalid characters from the filename
    For i = 1 To Len(invalidChars)
        fileName = Replace(fileName, Mid(invalidChars, i, 1), "")
    Next i
    
    CleanFileName = fileName
End Function