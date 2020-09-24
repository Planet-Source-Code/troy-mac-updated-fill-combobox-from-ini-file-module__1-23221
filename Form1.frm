VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Ini Form"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdSort 
      Caption         =   "Sort List"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.ComboBox CboIni 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   360
      List            =   "Form1.frx":0002
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton cmdAddToList 
      Caption         =   "Add to List"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Written by Troy MacPherson t_macpherson@yahoo.com
'This code can be used for anything by anyone at anytime. I would appreciate credit
'One thing that burns me is the use of the registy as being over used.  Too many programmers  want
'to use it.  I find it better to continue to use ini files for most of my stuff. The registry
'is large enough without adding useless garble to it that never gets removed.  Also this is easier
'for the everyday user to edit.  From a support point of view I would rather have them edit an ini
'file than the registry. A lot of big software companies still use ini files such as.
'Cakwalk, Microsoft, Adobe, Compaq (HP), Norton's Just to name a few. So you are in good company using them... Do a search for
'*.ini and see how many companies still use them.  I came up with 306 ini files hmmm doesn't look like their dead
'to me as some people would have you believe.

'Hope this helps
'USAGE: Create a form form1 Put a Combo box in it called CboIni remove the combo1 from the text field
'       Put a  2 command buttons on the form name them cmdAddToList and CmdSort.  Add the module to your code add copy this
'       text into the (General)

Dim oldCmbText As String
Dim cmbIn As Integer

Private Sub CboIni_Click()
cmbIn = CboIni.ListIndex
End Sub

Private Sub CboIni_KeyDown(KeyCode As Integer, Shift As Integer)
oldCmbText = CboIni.Text
End Sub

Private Sub CboIni_KeyUp(KeyCode As Integer, Shift As Integer)
'Here is the delete code see how I use the oldCmbText from the CboIni_Keydown sub. Normally this
'would have CboIni.Text but that is cleared when you hit delete so I capture it in the CboIni_Keydown first.
'the CboIni_Keydown gets the value of the CboIni.text before it is deleted this is needed otherwise the text = ""


Select Case KeyCode
Case 46 'Delete Key
    CboIni.ListIndex = cmbIn
    
    'Here is where I add the values that will be passed to the functions. Notice I don't have
    'to pass all the values to each function.  As in the GetIni function IniVal is not passed
    'because it is not needed. IniVal is what your asking for from the ini file.  In a function
    'though the number of variables passed much match the variables recieved.
    
        File = App.Path & "\IniMod.ini" '"C:\IniMod.ini" 'Path and file name of ini
        lpAppName = "Customer"  'Section name looks like [Customer]
        IniKey = RTrim("Name") 'Trim spaces out of ini file
        IniVal = oldCmbText 'This adds what ever is in the
        lpFileName = File
        
    'This keeps the ".." as the first entry in the Combobox and assures it wont be deleted
    'It can be removed or changed to what ever you want but make sure to replace it in the ini file
    'and in this if statement
        If oldCmbText = ".." And CboIni.ListIndex = 0 Then
            Exit Sub
        End If
            CboIni.RemoveItem (CboIni.ListIndex)
                IniVal = "Deleted"
    Call Delini(lpAppName, IniKey, IniVal, File, oldCmbText, cmbIn)
    
    CboIni.Clear 'Clear list so you don't get doubles
        Call GetIni(lpAppName, IniKey, lpFileName)
    'Sets the combobox to the first entry
    CboIni.ListIndex = 0

Case 13 'Enter Key
    'This lets you use the enter key to add new entries to the combo list
    Call cmdAddToList_Click
End Select
CmdSort.Caption = "Re-Sort"
End Sub

Private Sub cmdAddToList_Click()
'This will write the new entry that you typed in the combobox to the ini file
File = App.Path & "\IniMod.ini" 'Path and file name of ini
lpAppName = "Customer"  'Section name looks like [Customer]
IniKey = "Name"
IniVal = CboIni.Text 'This adds what ever is in the
lpFileName = File


Call WriteINI(lpAppName, IniKey, IniVal, File)
    CboIni.Clear 'Clear then fill combo with new ini entries NEEDED to keep from getting doubles
Call GetIni(lpAppName, IniKey, lpFileName)

'Sets the combobox to the first entry
CboIni.ListIndex = 0

End Sub

Private Sub CmdSort_Click()
'Bubble sort by Roman Blachman
Dim inArray() As String ' Input array
ReDim inArray(CboIni.ListCount - 1) ' Redim the array To the size of the list count items
If CmdSort.Caption = "Re-Sort" Then CmdSort.Caption = "Sort List"

For ic = 0 To CboIni.ListCount - 1
    inArray(ic) = CboIni.List(ic) ' Put all the values from the list box To the array
Next
cBubbleSort inArray ' Sort the array
CboIni.Clear ' Clear the list


For ic = 0 To UBound(inArray)
    CboIni.AddItem inArray(ic) ' Put the sorted items from the array
Next
CboIni.ListIndex = 0
End Sub

Private Sub Form_Load()
'This is self explanatory it calls the Getini function to fill the Combobox on load
'KeyPreview is for the CboIni_KeyUp CboIni_KeyDown subs. It is needed because I am using
'the keycodes. It allows VB to see what keys are being pressed
KeyPreview = True
'Path and name of ini file this can be changed to whatever u want Just be sure to change it in all
'Subs and functions otherwise you will be reading from one ini file and writing to another
File = App.Path & "\IniMod.ini"
lpAppName = "Customer"
IniKey = "Name"
lpDefault = ""
lpFileName = File

Call GetIni(lpAppName, IniKey, lpFileName)

 'This writes the first entry if there is no ini file.
If CboIni.ListCount = 0 Then
    IniVal = ".."
    Call WriteINI(lpAppName, IniKey, IniVal, lpFileName)
End If
'Call GetIni again to load the file
'Call GetIni(lpAppName, IniKey, lpFileName)
'Sets Combo to 1st entry
    CboIni.ListIndex = 0
    


End Sub
Public Sub cBubbleSort(inputArray As Variant)
'Bubble sort by Roman Blachman
        Dim lDown As Long, lUp As Long
        For lDown = UBound(inputArray) To LBound(inputArray) Step -1
            For lUp = LBound(inputArray) + 1 To lDown
                If inputArray(lUp - 1) > inputArray(lDown) Then SwapValues inputArray(lUp - 1), inputArray(lDown)
            Next lUp
        Next lDown
End Sub


Public Sub SwapValues(firstValue As Variant, secondValue As Variant)
'Bubble sort by Roman Blachman
        Dim tmpValue As Variant
        tmpValue = firstValue
        firstValue = secondValue
        secondValue = tmpValue
End Sub
