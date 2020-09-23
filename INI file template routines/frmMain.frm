VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "clsINI Demo dll"
   ClientHeight    =   8370
   ClientLeft      =   2280
   ClientTop       =   1185
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   11475
   Begin VB.Frame Frame6 
      Caption         =   "Settings section 2"
      Height          =   2415
      Left            =   6120
      TabIndex        =   23
      Top             =   5640
      Width           =   2175
      Begin VB.CommandButton cmdSaveSettings2 
         Caption         =   "Save settings"
         Height          =   375
         Left            =   360
         TabIndex        =   26
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CheckBox chkdeleteInputFiles 
         Caption         =   "Delete input files"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox chkAutoStart 
         Caption         =   "Auto start"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   960
         Width           =   1575
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Modify sections/keys manually"
      Height          =   3495
      Left            =   6600
      TabIndex        =   10
      Top             =   1320
      Width           =   4095
      Begin VB.Frame Frame2 
         Height          =   2295
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   3615
         Begin VB.OptionButton optType 
            Caption         =   "Section"
            Height          =   255
            Index           =   0
            Left            =   2280
            TabIndex        =   29
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optType 
            Caption         =   "Key"
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   28
            Top             =   1080
            Width           =   975
         End
         Begin VB.OptionButton optType 
            Caption         =   "Value"
            Height          =   255
            Index           =   2
            Left            =   2280
            TabIndex        =   27
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox tbValue 
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   285
            Left            =   240
            TabIndex        =   17
            Top             =   1800
            Width           =   1335
         End
         Begin VB.TextBox tbKey 
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   285
            Left            =   240
            TabIndex        =   16
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox tbSection 
            Height          =   285
            Left            =   240
            TabIndex        =   15
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Value"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   20
            Top             =   1560
            Width           =   405
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Key"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   19
            Top             =   840
            Width           =   270
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Section"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   18
            Top             =   120
            Width           =   540
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Rename"
         Height          =   375
         Index           =   1
         Left            =   1560
         TabIndex        =   13
         Top             =   2880
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Delete"
         Height          =   375
         Index           =   2
         Left            =   2880
         TabIndex        =   12
         Top             =   2880
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   2880
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Settings section 1"
      Height          =   2415
      Left            =   600
      TabIndex        =   4
      Top             =   5640
      Width           =   5295
      Begin VB.CommandButton cmdSaveSettings1 
         Caption         =   "Save settings"
         Height          =   375
         Left            =   3480
         TabIndex        =   9
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Frame Frame3 
         Height          =   1335
         Left            =   360
         TabIndex        =   6
         Top             =   840
         Width           =   2655
         Begin VB.OptionButton Option1 
            Caption         =   "Semi automatic"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   21
            Top             =   960
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Manual"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   8
            Top             =   240
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Automatic"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   7
            Top             =   600
            Width           =   1455
         End
      End
      Begin VB.TextBox tbInputDirectory 
         Height          =   285
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Input directory"
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   9480
      TabIndex        =   3
      Top             =   6480
      Width           =   1335
   End
   Begin VB.TextBox tbIni 
      Height          =   4455
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   960
      Width           =   5655
   End
   Begin VB.TextBox txtINIFile 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   5535
   End
   Begin VB.Label lblINIFile 
      Alignment       =   1  'Right Justify
      Caption         =   "INI File:"
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   555
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub chkAutoStart_Click()
   autoStart = chkAutoStart
   cmdSaveSettings2.Enabled = True
End Sub

Private Sub chkdeleteInputFiles_Click()
   deleteInputFiles = chkdeleteInputFiles
   cmdSaveSettings2.Enabled = True
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdSaveSettings1_Click()
   writeSection1
   cmdSaveSettings1.Enabled = False
   showIniFile INI.inifile
End Sub

Private Sub cmdSaveSettings2_Click()
   writeSection2
   cmdSaveSettings2.Enabled = False
   showIniFile INI.inifile
End Sub

Private Sub Command1_Click(Index As Integer)
   Dim newname As String
   
   Select Case Index
   Case 0   ' add
      Select Case optionType
      Case 2         ' add a key/value
         INI.setKeyValue tbKey, tbValue, tbSection
         
      Case 1        ' add a key
         INI.setKey tbKey, tbSection
         
      End Select
      
   Case 1   ' rename
      If InStr(1, INI.getSections, tbSection) = 0 Then
         MsgBox "Section:  """ & tbSection & """  not found!"
      Else
         Select Case optionType
         
         Case 2      ' rename a key (save the new value)
            If InStr(1, INI.getKeysInSection(tbSection), tbKey) = 0 Then
               MsgBox "Key:  """ & tbKey & """ not found in section:  """ & tbSection & """"
            Else
               newname = InputBox("Enter the new value for key: " & tbKey)
               If newname <> "" Then
                  INI.setKeyValue tbKey, tbValue, tbSection
                  tbValue = newname
               End If
            End If
            
         Case 1     ' rename a key
            If InStr(1, INI.getKeysInSection(tbSection), tbKey) = 0 Then
               MsgBox "Key:  """ & tbKey & """ not found in section:  """ & tbSection & """"
            Else
               newname = InputBox("Enter the new name for key: " & tbKey)
               If newname <> "" Then
                  INI.renameKey tbKey, newname, tbSection
                  tbKey = newname
               End If
            End If
         
         Case 0       ' rename a section
            newname = InputBox("Enter the new name for section: " & tbSection)
            If newname <> "" Then
               INI.renameSection newname, tbSection
               tbSection = newname
            End If
            
         End Select
      End If
   Case 2   ' delete
      If InStr(1, INI.getSections, tbSection) = 0 Then
         MsgBox "Section:  """ & tbSection & """  not found!"
      Else
         Select Case optionType
         
         Case 2     ' delete a key/value (erase the value)
            If InStr(1, INI.getKeysInSection(tbSection), tbKey) = 0 Then
               MsgBox "Key:  """ & tbKey & """ not found in section:  """ & tbSection & """"
            Else
               INI.deleteKeyValue tbKey, tbSection
            End If
            
         Case 1     ' delete a key
            If InStr(1, INI.getKeysInSection(tbSection), tbKey) = 0 Then
               MsgBox "Key:  """ & tbKey & """ not found in section:  """ & tbSection & """"
            Else
               INI.deleteKey tbKey, tbSection
            End If
            
         Case 0     ' delete the section
            INI.deleteSection tbSection
            
         End Select
      End If
   End Select
   
   showIniFile INI.inifile
   
End Sub

Private Sub Form_Load()
'
'   The sample application shows how to deal with the inifile at startup of your
'   application.  First it tries to load the values, then it checks if the version
'   as well as the number of keys against hard coded values.  If incorrect a new
'   section will be written.
'   If the inifile is not found at startup a new one will be created automatically
'   with hardcoded default settings.
'   For this all an additional module modIni is supplied with routines:
'      - readSettings
'      - writeSettings
'   In modIni all the global variables are defined and set.
'   So all you have to do in sub Main or in form.load is to call readSettings and then
'   set up the form according to the global values.
'
   
   ' Specify a non-default ini-filename here and everything is taken care of ...
   '
   'INI.inifile = App.Path & "\" & "test.ini"
   
   txtINIFile.Text = INI.inifile
   
   readAllSettings                                ' load all ini-variables
   
                                                  ' setup your application accordingly
   tbInputDirectory = inputDirectory                  ' section1
   Option1(manual) = True
   
   chkdeleteInputFiles.value = deleteInputFiles       ' section2
   chkAutoStart = autoStart
   
   optType(optionType) = True                         ' section3
   tbSection = sectionField
   tbKey = keyField
   tbValue = valueField
   
   cmdSaveSettings1.Enabled = False               ' other init
   cmdSaveSettings2.Enabled = False
   showIniFile INI.inifile
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If cmdSaveSettings1.Enabled Or cmdSaveSettings2.Enabled Then
      If MsgBox("Settings were changed!" & vbNewLine & _
           "Do you want to exit without saving current settings?", _
           vbQuestion + vbYesNo) = vbNo Then
         Cancel = 1           ' do not exit
         Exit Sub
      End If
   End If
   
   Set INI = Nothing          ' exit as usual
   Set frmMain = Nothing
   
End Sub

Private Sub showIniFile(ByVal infile As String)

   Dim sections As String
   Dim section As Variant
   Dim aSections() As String
   
   Dim keys As String
   Dim key As Variant
   Dim aKeys() As String
   
   tbIni = ""
   
   sections = INI.getSections
   aSections = Split(sections, Chr(0))
   For Each section In aSections
      tbIni = tbIni & "[" & section & "]" & vbNewLine
      INI.sectionName = section
      keys = INI.getKeysInSection
      aKeys = Split(keys, Chr(0))
      For Each key In aKeys
         tbIni = tbIni & key & " = " & INI.getKeyValue(key) & vbNewLine
      Next key
      tbIni = tbIni & vbNewLine
   Next section
    
End Sub

Private Sub mnuExit_Click()
   Unload Me
End Sub

Private Sub mnuHelp_Click()
   Load frmHelp
   frmHelp.Show
End Sub

Private Sub Option1_Click(Index As Integer)
   manual = Index
   cmdSaveSettings1.Enabled = True
End Sub

Private Sub optType_Click(Index As Integer)
   
   Select Case Index
   Case 0  ' section level
      tbSection.Enabled = True
      tbSection.BackColor = &H80000005
      
      tbKey.Enabled = False
      tbKey.BackColor = &H80000004
      
      tbValue.Enabled = False
      tbValue.BackColor = &H80000004
      
      Command1(0).Enabled = False
      
   Case 1  ' key level
   
      tbSection.Enabled = True
      tbSection.BackColor = &H80000005
      
      tbKey.Enabled = True
      tbKey.BackColor = &H80000005
      
      tbValue.Enabled = False
      tbValue.BackColor = &H80000004
      
      Command1(0).Enabled = True
      
   Case 2  ' value level
      tbSection.Enabled = True
      tbSection.BackColor = &H80000005
      
      tbKey.Enabled = True
      tbKey.BackColor = &H80000005
      
      tbValue.Enabled = True
      tbValue.BackColor = &H80000005
      
      Command1(0).Enabled = True
   End Select
   optionType = Index
   INI.setKeyValue "optiontype", optionType, "manualInput"
End Sub

Private Sub tbInputDirectory_Change()
   inputDirectory = tbInputDirectory
   cmdSaveSettings1.Enabled = True
End Sub

Private Sub tbKey_Validate(Cancel As Boolean)
   keyField = tbKey
   INI.setKeyValue "keyField", keyField, "manualInput"
End Sub

Private Sub tbSection_Validate(Cancel As Boolean)
   sectionField = tbSection
   INI.setKeyValue "sectionField", sectionField, "manualInput"
End Sub

Private Sub tbValue_Validate(Cancel As Boolean)
   valueField = tbValue
   INI.setKeyValue "valueField", valueField, "manualInput"
End Sub
