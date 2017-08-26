'Option Explicit

'Private Declare Function URLDownloadToFile Lib "urlmon" Alias _
 '   "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, _
  '  ByVal szFileName As String, ByVal dwReserved As Long, _
   ' ByVal lpfnCB As Long) As Long
    
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

   
Dim InstalledDLNumberArray()
Dim InstalledNameArray()
Dim InstalledVersionArray()
Dim InstalledFolderNameArray()

Dim DownloadNumber()
Dim CarName()
Dim VersionList()

Dim MasterInstallDLNumber()
Dim MasterCarName()
Dim MasterVersion()
Dim MasterFolderName()

Dim UpdateList()
Dim NoUpdateList()
Dim InstallList()
Dim UninstallList()

Dim DownloadArray()

Private Declare Function ShellExecute _
Lib "shell32.dll" _
Alias "ShellExecuteA" ( _
ByVal hwnd As Long, _
ByVal lpOperation As String, _
ByVal lpFile As String, _
ByVal lpParameters As String, _
ByVal lpDirectory As String, _
ByVal nShowCmd As Long) _
As Long

Public Sub Command4_Click()
    ' Use the AsyncRead method to copy the file.
    
'intialize download counter array
ReDim DownloadArray(0 To 2, 0 To 0)
    
If UBound(UpdateList, 2) > 0 Then
For i = 1 To UBound(UpdateList, 2)
    'parse url
  
    'Get http source
    Dim objHttp As Object, strURL As String, strText As String

    strText = ""

    Set objHttp = CreateObject("MSXML2.ServerXMLHTTP")
    strURL = "http://www.racedepartment.com/downloads/" & UpdateList(0, i)
    objHttp.Open "GET", strURL, False
    objHttp.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    objHttp.Send ("")
    strText = objHttp.responseText
    Set objHttp = Nothing
    
    Set myRegExp = New RegExp
    myRegExp.IgnoreCase = True
    myRegExp.Global = True
    myRegExp.Pattern = "href=""(.+version.+)"" "
    Set urlRex = myRegExp.Execute(strText)
    url = urlRex(0).SubMatches(0)
    
    'get download link
    UserControl.AsyncRead "http://www.racedepartment.com/" & url, vbAsyncTypeFile, i
   
Next i
End If

If UBound(InstallList, 2) > 0 Then
For i = UBound(UpdateList, 2) + 1 To UBound(UpdateList, 2) + UBound(InstallList, 2)
    'parse url
  
    'Get http source
    
    strText = ""

    Set objHttp = CreateObject("MSXML2.ServerXMLHTTP")
    strURL = "http://www.racedepartment.com/downloads/" & InstallList(0, i - UBound(UpdateList, 2))
    objHttp.Open "GET", strURL, False
    objHttp.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    objHttp.Send ("")
    strText = objHttp.responseText
    Set objHttp = Nothing
    
    Set myRegExp = New RegExp
    myRegExp.IgnoreCase = True
    myRegExp.Global = True
    myRegExp.Pattern = "href=""(.+version.+)"" "
    Set urlRex = myRegExp.Execute(strText)
    url = urlRex(0).SubMatches(0)
    
    'get download link
    UserControl.AsyncRead "http://www.racedepartment.com/" & url, vbAsyncTypeFile, i
Next i
End If

'delete mod
'remove from data.txt
    
End Sub



Public Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
    ' Yield execution to ensure that the temporary file is written.
DoEvents
DownloadArray(2, CInt(AsyncProp.PropertyName)) = AsyncProp.Value
'run 7zip and folder checks after all have been downloaded
If ProgressBar1.Value = 1 And UBound(DownloadArray, 2) = (UBound(UpdateList, 2) + UBound(InstallList, 2)) Then
    For i = 1 To UBound(DownloadArray, 2)
        'uzip with 7zip
        
        If i <= UBound(UpdateList, 2) Then
            filepathname = UpdateList(0, i)
        Else
            filepathname = InstallList(0, i - UBound(UpdateList, 2))
        End If
                        
       zipextract = Shell("""C:\Program Files\7-Zip\7z.exe"" e -y " & DownloadArray(2, i) & " -oC:\users\ada\desktop\" & filepathname, vbHide)
       'copy to correct directory
        If i <= UBound(UpdateList, 2) Then
            carpathname = UpdateList(1, i)
        Else
            carpathname = InstallList(1, i - UBound(UpdateList, 2))
        End If
        'only change txt file if extraction was successful
        If zipextract > 0 Then
            'find and overwrite data.txt section
            
        Else
           MsgBox (carpathname & " failed to install/update. Please retry.")
        End If
    Next i

End If
    
    
End Sub

Public Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
    ' Display the progress of the file copy using the
    ' BytesRead and BytesMax properties of the AsyncProp object.
    If CInt(AsyncProp.PropertyName) > UBound(DownloadArray, 2) Then
        ReDim Preserve DownloadArray(0 To 2, 0 To CInt(AsyncProp.PropertyName))
    End If
    DownloadArray(0, CInt(AsyncProp.PropertyName)) = AsyncProp.BytesRead
    DownloadArray(1, CInt(AsyncProp.PropertyName)) = AsyncProp.BytesMax
    amountdownload = 0
    totaldownload = 0
    For q = LBound(DownloadArray, 2) To UBound(DownloadArray, 2)
        amountdownload = amountdownload + DownloadArray(0, q)
        totaldownload = totaldownload + DownloadArray(1, q)
    Next q
        
    Dim progress As Double
    If totaldownload > 0 Then
    progress = amountdownload / totaldownload
    Label7.Caption = CLng(amountdownload / 10000) / 100 & " of " & CLng(totaldownload / 10000) / 100 & " MB (" & CInt(100 * progress) & "%)"
    Else
    progress = 0
    End If
    ProgressBar1.Value = progress
    
    
End Sub

Private Sub UserControl_Initialize()
    Label7.Caption = ""
    ' Assign a file to be copied.
    ' Use the Visual Basic 6.0 run-time files package as a test.
    
'Get http source
Dim objHttp As Object, strURL As String, strText As String, i As Integer
ACCarsLoop = 0
strText = ""
i = 1

Do While ACCarsLoop = 0

Set objHttp = CreateObject("MSXML2.ServerXMLHTTP")

strURL = "http://www.racedepartment.com/downloads/categories/ac-cars.6/?page=" & i
objHttp.Open "GET", strURL, False
objHttp.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
objHttp.Send ("")

strText = strText & " " & objHttp.responseText

Set objHttp = Nothing

If i = 1 Then
Set myRegExp = New RegExp
myRegExp.IgnoreCase = True
myRegExp.Global = True
myRegExp.Pattern = "Page 1 of (\d)"
Set PageCountRex = myRegExp.Execute(strText)
PageCount = CInt(PageCountRex(0).SubMatches(0))
End If

i = i + 1
If i > PageCount Then
ACCarsLoop = 1
End If

Loop

'Clipboard.Clear
'Clipboard.SetText (strText)

'get list of all version numbers
Set myRegExp = New RegExp
myRegExp.IgnoreCase = True
myRegExp.Global = True
myRegExp.Pattern = "version"">(.*)<"
Set VersionListRex = myRegExp.Execute(strText)

ReDim VersionList(0 To VersionListRex.Count)
For j = 0 To VersionListRex.Count - 1
    VersionList(j) = VersionListRex(j).SubMatches(0)
Next j
    
'get list of all download numbers
Set myRegExp = New RegExp
myRegExp.IgnoreCase = True
myRegExp.Global = True
myRegExp.Pattern = "\.(.*)/"" class=""faint"
Set DownloadNumberRex = myRegExp.Execute(strText)

ReDim DownloadNumber(0 To DownloadNumberRex.Count)
For j = 0 To DownloadNumberRex.Count - 1
    DownloadNumber(j) = DownloadNumberRex(j).SubMatches(0)
Next j

'get list of all cars names

ReDim CarName(0 To UBound(DownloadNumber))

For k = 0 To UBound(VersionList) - 1
    Set myRegExp = New RegExp
    myRegExp.IgnoreCase = True
    myRegExp.Global = True
    myRegExp.Pattern = DownloadNumber(k) & "/"">(.*)<"
    Set CarNameRex = myRegExp.Execute(strText)
    CarName(k) = CarNameRex(0).SubMatches(0)
    List1.AddItem (CarName(k))
Next k


'import list of installed cars from text file
Open "C:\Users\ADA\Documents\Mod Manager\data.txt" For Input As #1




ReDim InstalledDLNumberArray(0 To 0)

i = 0
Do Until EOF(1)
Input #1, InstalledDLNumber, InstalledName, InstalledVersion, InstalledFolderName

'used on first tab and updated after every click on "install" or "uninstall" buttons
ReDim Preserve InstalledDLNumberArray(0 To i)
ReDim Preserve InstalledNameArray(0 To i)
ReDim Preserve InstalledVersionArray(0 To i)
ReDim Preserve InstalledFolderNameArray(0 To i)

InstalledDLNumberArray(i) = InstalledDLNumber
InstalledNameArray(i) = InstalledName
InstalledVersionArray(i) = InstalledVersion
InstalledFolderNameArray(i) = InstalledFolderName

'used on second tab to reference installed arrays on what mods have been uninstalled
ReDim Preserve MasterInstallDLNumber(0 To i)
ReDim Preserve MasterCarName(0 To i)
ReDim Preserve MasterVersion(0 To i)
ReDim Preserve MasterFolderName(0 To i)

MasterInstallDLNumber(i) = InstalledDLNumber
MasterCarName(i) = InstalledName
MasterVersion(i) = InstalledVersion
MasterFolderName(i) = InstalledFolderName


i = i + 1
Loop

Close #1

'exit if no items on list
If i = 0 Then
Exit Sub
End If

For i = 0 To UBound(InstalledDLNumberArray)
List2.AddItem (InstalledNameArray(i))
Next i
    
End Sub







Private Sub Command1_Click()
If List1.SelCount = 0 Then
MsgBox ("Please select a Car")
Exit Sub
End If

'dont add new car if already in list
For i = 0 To UBound(InstalledDLNumberArray)
    If CInt(InstalledDLNumberArray(i)) = DownloadNumber(List1.ListIndex) Then
        Exit Sub
    End If
Next i
    

'add new car
NewArrayLength = UBound(InstalledDLNumberArray) + 1

ReDim Preserve InstalledDLNumberArray(0 To NewArrayLength)
ReDim Preserve InstalledNameArray(0 To NewArrayLength)
ReDim Preserve InstalledVersionArray(0 To NewArrayLength)
ReDim Preserve InstalledFolderNameArray(0 To NewArrayLength)

InstalledDLNumberArray(NewArrayLength) = DownloadNumber(List1.ListIndex)
InstalledNameArray(NewArrayLength) = CarName(List1.ListIndex)
InstalledVersionArray(NewArrayLength) = VersionList(List1.ListIndex)
InstalledFolderNameArray(NewArrayLength) = "Placeholder"

'refresh car list
List2.Clear

For i = 0 To UBound(InstalledDLNumberArray)
List2.AddItem (InstalledNameArray(i))
Next i


End Sub

Private Sub Command2_Click()
If List2.SelCount = 0 Then
MsgBox ("Please select a Car")
Exit Sub
End If

'create temporary arrays to store all values from original except one being deleted,
Dim temparray1()
Dim temparray2()
Dim temparray3()
Dim temparray4()

ReDim temparray1(0 To UBound(InstalledDLNumberArray) - 1)
ReDim temparray2(0 To UBound(InstalledDLNumberArray) - 1)
ReDim temparray3(0 To UBound(InstalledDLNumberArray) - 1)
ReDim temparray4(0 To UBound(InstalledDLNumberArray) - 1)
j = 0

For i = 0 To UBound(InstalledDLNumberArray)
    If i <> List2.ListIndex Then
        temparray1(j) = InstalledDLNumberArray(i)
        temparray2(j) = InstalledNameArray(i)
        temparray3(j) = InstalledVersionArray(i)
        temparray4(j) = InstalledFolderNameArray(i)
        j = j + 1
    End If
Next i

'clear old arrays and overwrite old arrays with new ones
ReDim InstalledDLNumberArray(0 To UBound(temparray1))
ReDim InstalledNameArray(0 To UBound(temparray2))
ReDim InstalledVersionArray(0 To UBound(temparray3))
ReDim InstalledFolderNameArray(0 To UBound(temparray4))

For i = 0 To UBound(temparray1)
InstalledDLNumberArray(i) = temparray1(i)
InstalledNameArray(i) = temparray2(i)
InstalledVersionArray(i) = temparray3(i)
InstalledFolderNameArray(i) = temparray4(i)
Next i

'refresh car list
List2.Clear

For i = 0 To UBound(InstalledDLNumberArray)
List2.AddItem (InstalledNameArray(i))
Next i

End Sub

Private Sub Command3_Click()
For i = 1 To UBound(UpdateList, 2) + UBound(InstallList, 2)
    'only cancel if not started or not completed
    If DownloadArray(0, i) <> DownloadArray(1, i) Or DownloadArray(0, i) = 0 Then
    UserControl.CancelAsyncRead (i)
    End If
Next i
Label7.Caption = ""
ProgressBar1.Value = 0
End Sub


Private Sub List1_DblClick()
carurl = "http://www.racedepartment.com/downloads/" & DownloadNumber(List1.ListIndex)
Dim r As Long
   r = ShellExecute(0, "open", carurl, 0, 0, 1)
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 1 Then

List3.Clear
List4.Clear
List5.Clear
List6.Clear

ReDim UpdateList(0 To 3, 0 To 0)
ReDim NoUpdateList(0 To 3, 0 To 0)
ReDim InstallList(0 To 3, 0 To 0)
ReDim UninstallList(0 To 1, 0 To 0)

matchfound = 0

'create arrays of cars to be updated, not updated, and installed
For i = 0 To UBound(InstalledDLNumberArray)
    For j = 0 To UBound(MasterInstallDLNumber)
        If InstalledDLNumberArray(i) = MasterInstallDLNumber(j) Then
            For k = 0 To UBound(DownloadNumber)
                If MasterInstallDLNumber(j) = DownloadNumber(k) Then
                    If MasterVersion(j) = VersionList(k) Then
                        'add to no update list
                        LengthNoUpdate = UBound(NoUpdateList, 2) + 1
                        ReDim Preserve NoUpdateList(0 To 3, 0 To LengthNoUpdate)
                        
                        NoUpdateList(0, LengthNoUpdate) = MasterInstallDLNumber(j)
                        NoUpdateList(1, LengthNoUpdate) = MasterCarName(j)
                        NoUpdateList(2, LengthNoUpdate) = MasterVersion(j)
                        NoUpdateList(3, LengthNoUpdate) = MasterFolderName(j)
                        
                        List6.AddItem (NoUpdateList(1, LengthNoUpdate))
                    Else
                        'add to update list
                        LengthUpdate = UBound(UpdateList, 2) + 1
                        ReDim Preserve UpdateList(0 To 3, 0 To LengthUpdate)
                        
                        UpdateList(0, LengthUpdate) = MasterInstallDLNumber(j)
                        UpdateList(1, LengthUpdate) = MasterCarName(j)
                        UpdateList(2, LengthUpdate) = MasterVersion(j)
                        UpdateList(3, LengthUpdate) = MasterFolderName(j)
                        
                        List5.AddItem (UpdateList(1, LengthUpdate))
                    End If
                End If
            Next k
        matchfound = 1
        End If
    Next j
If matchfound = 0 Then
    'add to install list
    LengthInstall = UBound(InstallList, 2) + 1
    ReDim Preserve InstallList(0 To 3, 0 To LengthInstall)
                        
    InstallList(0, LengthInstall) = InstalledDLNumberArray(i)
    InstallList(1, LengthInstall) = InstalledNameArray(i)
    InstallList(2, LengthInstall) = InstalledVersionArray(i)
    InstallList(3, LengthInstall) = InstalledFolderNameArray(i)
                        
    List3.AddItem (InstallList(1, LengthInstall))
Else
    matchfound = 0
End If
Next i

matchfound = 0
'check for items needing uninstalled
For i = 0 To UBound(MasterInstallDLNumber)
    For j = 0 To UBound(InstalledDLNumberArray)
        If MasterInstallDLNumber(i) = InstalledDLNumberArray(j) Then
            matchfound = 1
        End If
    Next j
If matchfound = 0 Then
    'add to uninstall list
    LengthUninstall = UBound(UninstallList, 2) + 1
    ReDim Preserve UninstallList(0 To 1, 0 To LengthUninstall)
                        
    UninstallList(0, LengthUninstall) = MasterInstallDLNumber(i)
    UninstallList(1, LengthUninstall) = MasterCarName(i)
                            
    List4.AddItem (UninstallList(1, LengthUninstall))
Else
    matchfound = 0
End If
Next i

End If
End Sub


