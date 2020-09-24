Attribute VB_Name = "IniAddModule"
Declare Function WritePrivateProfileSection Lib _
"kernel32" Alias "WritePrivateProfileSectionA" _
(ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

Declare Function WritePrivateProfileString Lib _
"kernel32" Alias "WritePrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
ByVal lpFileName As String) As Long

Declare Function GetPrivateProfileSection Lib _
"kernel32" Alias "GetPrivateProfileSectionA" _
(ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, _
ByVal lpFileName As String) As Long

Declare Function GetPrivateProfileString Lib _
"kernel32" Alias "GetPrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'Write function writes passed values to ini file
Public Function WriteINI(ByVal lpAppName As String, ByVal IniKey As String, ByVal IniVal As String, ByVal lpFileName As String)
Dim lonStatus As Long
Dim sBuff As String * 255
Dim sUser As String
Dim I As Integer
Dim mRet As Integer


'This first GetPrivateProfileString is to get the count of names with number associated
'in the ini file. Notice the IniKey & CStr(I + 1) this returns Name1, Name2, Name3 etc...
'Look at the ini file you'll see what I mean.

While GetPrivateProfileString(lpAppName, IniKey & CStr(I + 1), "", sBuff, 255, lpFileName)
 sUser = Left(sBuff, InStr(sBuff, Chr(0)) - 1)
 I = I + 1
 'checking for duplicate name and the default ..
 If sUser = Form1.CboIni.Text Or Form1.CboIni.Text = ".." Then
 mRet = MsgBox("Sorry the user: " & sUser & " is already in the list", vbOKOnly, "Duplicate User Name")
 Exit Function
 End If
 Wend
  
'This statement will write the new name to the Customer Section
'of the ini file
lonStatus = WritePrivateProfileString(lpAppName, IniKey & CStr(I + 1), IniVal, lpFileName)

 

I = 0 'Set I to 0 to start the read of the ini file at the begining


'Not working correctly dose not keep from deleting .. 8/23/01


End Function
'Delete function deletes passed values from ini file
Public Function Delini(ByVal lpAppName As String, ByVal IniKey As String, ByVal IniVal As String, ByVal lpFileName As String, ByVal oldCmbText As String, ByVal cmbIn As Integer)
Dim sBuff As String * 255
Dim sUser As String
Dim I As Integer
Dim lonStatus As Long
I = 0
If oldCmbText = ".." And cmbIn = 0 Then
    Exit Function
End If
'Loops until it find the value you deleted which is oldCmbText
While GetPrivateProfileString(lpAppName, IniKey & CStr(I + 1), "", sBuff, 255, lpFileName)
       sUser = Left(sBuff, InStr(sBuff, Chr(0)) - 1)
         I = I + 1
    
        If sUser = oldCmbText Then GoTo Delini  'Retain .. at listindex 0
         
    Wend
Delini:
'Rewrites the value with Deleted

                lonStatus = WritePrivateProfileString(lpAppName, IniKey & CStr(I + 0), IniVal, lpFileName)
    
    I = 0 'Set I to zero so the loop starts at the begining before reading
    


End Function

'Get function gets passed values from ini file
Public Function GetIni(ByVal lpAppName As String, ByVal IniKey As String, ByVal lpFileName As String)

Dim sBuff As String * 255
Dim sUser As String
Dim I As Integer

 'Get ini information by looping through the ini file. the CStr(I + 1) keeps track of the
 'position. if the value is not Deleted it adds it to the combobox
 'NOTE: "form1.CboIni.AddItem sUser" you need to use the name of the ComboBox here to add from file
While GetPrivateProfileString(lpAppName, IniKey & CStr(I + 1), "", sBuff, 255, lpFileName)
    sUser = Left(sBuff, InStr(sBuff, Chr(0)) - 1)
    If sUser <> "Deleted" Then Form1.CboIni.AddItem sUser
    I = I + 1

Wend

End Function
