VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsINI"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'********************************************************************************
' Author:   Eric Dalquist
'
' email:    ebdalqui@mtu.edu (feel free to email me with questions)
'
' Credit:   Well I found a similar INI class on planet source code about a year
'           ago and search for it so I could give though athor some credit but
'           couldn't find it. So if you think this is based on your code it may
'           well be and I give you credit for it.
'
' Comments: Well the INI API calls are fairly simple and easy to use directly in
'           a program. I felt I wanted something with a bit more power and ease
'           of use. This INI class can be compiled on it's own and included into
'           a project as a dll or you can add the cslINI.cls file to your
'           project so you don't have to keep track of another dll. This class
'           includes all the functions I could think of uses for and actually get
'           to work. One thing that may seem missing is a createSection method.
'           I tried to get this to work and id does kind of if you don't mind a
'           null entry in your section. To do it just take the same idea from
'           the createKey method and replace Key with "" in the API call.
'
' Bugs:     Well I don't know if I would call this a bug but it sure fits under
'           quirks. You might notice in some spots I use "" and in others
'           vbNullString which are supposedly the same thing. When passed to the
'           PrivateProfileString API calls they behave much differently. For
'           example if you create a key with the value "" the key will exist in
'           the INI but have no value. If you create the same key with
'           vbNullString the key will not be written to the file. This is how I
'           ended up implementing the delete methods. Whatever works right?
'
'----------------------------------------------------------------------------------
' Modifications: Eric Kok
'   I have added some functionalities:
'     - Get the number of sections.
'     - Get the number of keys in a section.
'     - Get inifile version.
'     - Get/set defaultSectionName
'     - Get/set sectionname (currentSection) .
'         This will make the supply of the sectionname optional.
'     - Delete all sections
'     - Be able to supply a default value when setting a key/value, so
'         if the key is not found the default will be written to the inifile.
'     - A default section (exe-name) is setup at init.
'     - If an initfile is specified the default sectionname is the initfilename
'         without extension.
'
'   The sample template application shows how to deal with the inifile at startup
'   of your application.  First it tries to load the values, then it checks the
'   version as well as the number of sections/keys against hard coded values.  If
'   incorrect a new section or complete new inifile will be written.
'   This gives the developer the freedom to increase/decrease the number of
'   sections and within the sections the number of keys: the ini-file will always
'   be up to date.
'
'********************************************************************************

'*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!'
' REMEMBER to make the cslINI modual 6 - GlobalMultiUse under instancing in it's '
' properties if compiling this as a standalone DLL! Using the program            '
' cslINIDemo_src changes this setting to private.                                '
'*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!'

Option Explicit

'The two API calls to read and write from the INI file
Private Declare Function GetPrivateProfileString Lib "kernel32" _
     Alias "GetPrivateProfileStringA" ( _
     ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
     ByVal lpDefault As String, ByVal lpReturnedString As String, _
     ByVal nSize As Long, ByVal lpFileName As String) As Long
     
Private Declare Function WritePrivateProfileString& Lib "kernel32" _
     Alias "WritePrivateProfileStringA" ( _
     ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)


Private var_INIfile As String         'The INI file's path
Private var_DefaultRetVal As String   'What to return if no data is found
Private defaultSection As String      'This is the default section (exeName)
Private currentSection As String

Private Sub Class_Initialize()
    'when the object is created we set the default file name
    setDefaultINIFilePath
    'our default for no data is "" (vbNullString)
    var_DefaultRetVal = ""
    defaultSection = App.EXEName
    currentSection = defaultSection
End Sub

Public Sub setDefaultINIFilePath()
    'This sets the INI file to the program's path and exe name,
    'by adding a .ini to the end.
    'Example: the program c:\iniTest\test.exe would have the INI file
    '                     c:\iniTest\test.ini
    'This set up makes it easy for refrence to the file
    
    'If we are running in the root folder app.path will have a \ on the end.
    '
    If Right$(App.Path, 1) = "\" Then
        var_INIfile = App.Path & App.EXEName & ".ini"
    Else 'Otherwise we need to add our own.
        var_INIfile = App.Path & "\" & App.EXEName & ".ini"
    End If
End Sub

Private Sub primeINIFile()
    'Priming the INI file stops the API calls from complaining if the file doesn't exist
    'I'm pretty sure there is a better way to do this but for the sake of
    'ease I just open the INIFile for append (so we don't screw with the data
    'already in it) and then close it right away. If the file was already there nothing
    'happens otherwise the file gets created
    Open var_INIfile For Append As #1
    Close #1
End Sub

Private Function basename(fileSpec) As String
   Dim ptr As Integer
   ptr = InStrRev(fileSpec, "\")
   If ptr > 0 Then
      basename = Right(fileSpec, Len(fileSpec) - ptr)
   Else
      basename = fileSpec
   End If
End Function

'****************************** Start Property Let/Get ******************************
'
Public Property Let inifile(ByVal INIFileIn As String)
    'set the INI path
    '
    Dim ptr As Integer
    var_INIfile = INIFileIn
    defaultSection = basename(var_INIfile)
    ptr = InStrRev(defaultSection, ".")
    If ptr > 0 Then defaultSection = Left(defaultSection, ptr - 1)
    currentSection = defaultSection
End Property

Public Property Get inifile() As String
    'return the INI path
    inifile = var_INIfile
End Property

Public Property Let sectionName(sectionIn As String)
   ' set the section name to a new value
   currentSection = sectionIn
End Property

Public Property Get sectionName() As String
   ' return the current section name
   sectionName = currentSection
End Property

Public Property Let defaultSectionName(ByVal section As String)
    'set the default return value
    defaultSection = section
End Property

Public Property Get defaultSectionName() As String
    'return the default return value
    defaultSectionName = defaultSection
End Property

Public Property Let DefaultReturnValue(ByVal RetVal As String)
    'set the default return value
    var_DefaultRetVal = RetVal
End Property

Public Property Get DefaultReturnValue() As String
    'return the default return value
    DefaultReturnValue = var_DefaultRetVal
End Property
'
'****************************** End Property Let/Get ******************************


'****************************** Start Set Methods ******************************
'
Public Sub setKey(ByVal key As String, _
           Optional ByVal section As String, Optional ByVal INIFileLoc As String)
    Dim useFile As String
    Dim useSection As String
    
   '************************************************************
   'This little block of code is in all the sub and functions
   'that use the INI file. Since there is a file variable that
   'is set by default and can be modified by the user the
   'INIFileLoc parameter is optional. This block of code checks
   'to see if the argument was passed. If it was the UseFile var
   'is set to the passed file if not it is set to the local file
   'variable.
    
    If IsMissing(INIFileLoc) Or INIFileLoc = "" Then
        useFile = var_INIfile
    Else
        useFile = INIFileLoc
    End If
    
    If IsMissing(section) Or section = "" Then
        useSection = currentSection
    Else
        useSection = section
    End If
    '
    '************************************************************
    
    'Call the write profile string but since we only want to create a key but
    'not give it a value we pass "" (vbNullString) for the value
    WritePrivateProfileString useSection, key, "", useFile
End Sub

Public Sub setKeyValue(ByVal key As String, ByVal value As String, _
                Optional ByVal section As String, Optional ByVal INIFileLoc As String)
                
    Dim useFile As String
    Dim useSection As String
    
    If IsMissing(INIFileLoc) Or INIFileLoc = "" Then
        useFile = var_INIfile
    Else
        useFile = INIFileLoc
    End If
    
    If IsMissing(section) Or section = "" Then
        useSection = currentSection
    Else
        useSection = section
    End If
    
    'Call the write profile string with all the parameters
    WritePrivateProfileString useSection, key, value, useFile
End Sub
'
'****************************** End Create Methods ******************************

'****************************** Start Get Methods ******************************
'
Public Function getVersion(Optional ByVal INIFileLoc As String) As String
   '
   ' Assumes an inifile version entry, otherwise writes a default version
   ' in the exeName section
   '
    Dim useFile As String
    Dim version As String
    
    If IsMissing(INIFileLoc) Or INIFileLoc = "" Then
        useFile = var_INIfile
    Else
        useFile = INIFileLoc
    End If
    
    version = getKeyValue("version", "1.0", defaultSection, useFile)
    
    getVersion = version
    
End Function

Public Function getSections(Optional ByVal INIFileLoc As String) As String
    Dim RetVal As String
    Dim KeyLen As Integer
    Dim useFile As String
    
    If IsMissing(INIFileLoc) Or INIFileLoc = "" Then
        useFile = var_INIfile
    Else
        useFile = INIFileLoc
    End If
    
    'prime the INI to make sure it exists
    primeINIFile
    
    'since we could have a lot of sections in one file we are going to use a 1024 char buffer
    RetVal = String$(1024, 0)
    KeyLen = GetPrivateProfileString(vbNullString, vbNullString, var_DefaultRetVal, RetVal, Len(RetVal), useFile)
    
    'if no sections are found then return "" (vbNullString)
    If KeyLen = 0 Then
        getSections = ""
    Else
      'if the retval is > 0 then return the results
      'since we are getting multiple sections but returning them as one string the
      'programer should use the split() function in the returned value with
      'chr$(0) being the delimiter
      getSections = Trim(Left$(RetVal, KeyLen - 1))
      If Right(getSections, 1) = Chr(0) Then
         getSections = Left(getSections, Len(getSections) - 1)
      End If
    End If
End Function

Public Function getNumberOfSections(Optional ByVal INIFileLoc As String) As Integer
   Dim sections As String
   Dim aSections() As String
   Dim useFile As String
   
   If IsMissing(INIFileLoc) Or INIFileLoc = "" Then
        useFile = var_INIfile
   Else
        useFile = INIFileLoc
   End If
    
   sections = getSections(useFile)
   If sections = "" Then
      getNumberOfSections = 0
   Else
      aSections = Split(sections, Chr(0))
      getNumberOfSections = UBound(aSections) + 1
   End If
   
End Function

Public Function getKeysInSection(Optional ByVal section As String, Optional ByVal INIFileLoc As String) As String
    Dim RetVal As String
    Dim KeyLen As Integer
    Dim useFile As String
    Dim useSection As String
    
    If IsMissing(INIFileLoc) Or INIFileLoc = "" Then
        useFile = var_INIfile
    Else
        useFile = INIFileLoc
    End If
    
    If IsMissing(section) Or section = "" Then
        useSection = currentSection
    Else
        useSection = section
    End If
    
    'prime the INI to make sure it exists
    primeINIFile
    
    'since we could have a lot of keys in one section we are going to use a 1024 char buffer
    RetVal = String$(1024, 0)
    KeyLen = GetPrivateProfileString(useSection, vbNullString, var_DefaultRetVal, _
                                     RetVal, Len(RetVal), useFile)
    
    'if no keys are found then return "" (vbNullString)
    If KeyLen = 0 Then
        getKeysInSection = ""
    Else
        'if the retval is > 0 then return the results
        'since we are getting multiple keys but returning them as one string the
        'programer should use the split() function in the returned value with
        'chr$(0) being the delimiter
        getKeysInSection = Left$(RetVal, KeyLen - 1)
    End If
End Function

Public Function getNumberOfKeysInSection(Optional ByVal section As String, Optional ByVal INIFileLoc As String) As Integer

   Dim keys As String
   Dim useFile As String
   Dim useSection As String
   Dim i As Integer, ptr As Integer
    
    If IsMissing(INIFileLoc) Or INIFileLoc = "" Then
        useFile = var_INIfile
    Else
        useFile = INIFileLoc
    End If
    
    If IsMissing(section) Or section = "" Then
        useSection = currentSection
    Else
        useSection = section
    End If
    
    keys = getKeysInSection(useSection, useFile)
    
   If keys = "" Then
      getNumberOfKeysInSection = 0
   Else
      i = 1
      ptr = 1
      ptr = InStr(ptr, keys, Chr(0))
      Do While ptr > 0
         i = i + 1
         ptr = InStr(ptr + 1, keys, Chr(0))
      Loop
      getNumberOfKeysInSection = i
   End If
    
End Function

Public Function getKeyValue(ByVal key As String, Optional ByVal defaultValue As String, _
         Optional ByVal section As String, Optional ByVal INIFileLoc As String) As String
    Dim RetVal As String
    Dim KeyLen As Integer
    Dim useFile As String
    Dim useSection As String
    
    If IsMissing(INIFileLoc) Or INIFileLoc = "" Then
        useFile = var_INIfile
    Else
        useFile = INIFileLoc
    End If
    
    If IsMissing(section) Or section = "" Then
        useSection = currentSection
    Else
        useSection = section
    End If
    
    If IsMissing(defaultValue) Or defaultValue = "" Then
      defaultValue = "<not_set>"
    End If
    
    'call our priming function to be sure the file exists
    primeINIFile
    
    'OK I just choose 255 as a nice number for the max length of data being retrieved
    'I think you can change it and the API calls will still work. It just creates a
    'string of that length in memory for the API call to write the results of it's action
    'to if the key is not found the contents of var_DefaultRetVal are returned
    RetVal = String$(255, 0)
    KeyLen = GetPrivateProfileString(useSection, key, "", RetVal, Len(RetVal), useFile)
    
    ' If the return has 0 length (not found),
    ' return the defaultValue and write that value to the inifile
    '
    If KeyLen = 0 Then
        setKeyValue key, defaultValue, useSection, useFile
        getKeyValue = defaultValue
    Else ' or the key's value
        getKeyValue = Left$(RetVal, KeyLen)
    End If
End Function
'
'****************************** End Get Methods ******************************

'****************************** Start Delete Methods ******************************
'
Public Sub deleteSection(Optional ByVal section As String, _
                         Optional ByVal INIFileLoc As String)
    Dim useFile As String
    Dim useSection As String
    
    If IsMissing(INIFileLoc) Or INIFileLoc = "" Then
        useFile = var_INIfile
    Else
        useFile = INIFileLoc
    End If
    
    If IsMissing(section) Or section = "" Then
        useSection = currentSection
    Else
        useSection = section
    End If
    
    ' To delete a section it's the same thing as a key but we write a
    ' vbNullString to the key name and value
    WritePrivateProfileString useSection, vbNullString, vbNullString, useFile
End Sub

Public Sub deleteAllSections(Optional ByVal INIFileLoc As String)
   Dim useFile As String
   Dim sections As String
   Dim section As Variant
   Dim aSections() As String
   
   If IsMissing(INIFileLoc) Or INIFileLoc = "" Then
       useFile = var_INIfile
   Else
       useFile = INIFileLoc
   End If
   
   sections = getSections
   aSections = Split(sections, Chr(0))
   For Each section In aSections
      If section <> "" Then
         deleteSection section, useFile
      End If
   Next
End Sub

Public Sub deleteKey(ByVal key As String, _
          Optional ByVal section As String, Optional ByVal INIFileLoc As String)
    Dim useFile As String
    Dim useSection As String
    
    If IsMissing(INIFileLoc) Or INIFileLoc = "" Then
        useFile = var_INIfile
    Else
        useFile = INIFileLoc
    End If
    
    If IsMissing(section) Or section = "" Then
        useSection = currentSection
    Else
        useSection = section
    End If
    
    'to delete a key use the actual vbnullstring object ... for some reason
    'it behaves differently than passing the API ""
    WritePrivateProfileString useSection, key, vbNullString, useFile
End Sub

Public Sub deleteKeyValue(ByVal key As String, _
           Optional ByVal section As String, Optional ByVal INIFileLoc As String)
    Dim useFile As String
    Dim useSection As String
    
    If IsMissing(INIFileLoc) Or INIFileLoc = "" Then
        useFile = var_INIfile
    Else
        useFile = INIFileLoc
    End If
    
    If IsMissing(section) Or section = "" Then
        useSection = currentSection
    Else
        useSection = section
    End If
    
    'to delete the key's value just write a null string to it
    WritePrivateProfileString useSection, key, "", useFile
End Sub
'
'****************************** End Delete Methods ******************************

'****************************** Start Rename Methods ******************************
'
Public Sub renameSection(ByVal NewSection As String, _
             Optional ByVal section As String, Optional ByVal INIFileLoc As String)
    Dim KeyList() As String
    Dim KeyCount As Integer
    Dim Index As Integer
    Dim useFile As String
    Dim useSection As String
    
    If IsMissing(INIFileLoc) Or INIFileLoc = "" Then
        useFile = var_INIfile
    Else
        useFile = INIFileLoc
    End If
    
    If IsMissing(section) Or section = "" Then
        useSection = currentSection
    Else
        useSection = section
    End If
    
    'makle sure the file exists
    primeINIFile
    
    'renaming a section works the same way.
    'we get all the key names into an array
    KeyList = Split(getKeysInSection(useSection, useFile), Chr$(0))
    'get the number of keys
    KeyCount = UBound(KeyList)
    
    'for each key in the array
    For Index = 0 To KeyCount
        'create the same key in the new section
        setKeyValue KeyList(Index), getKeyValue(KeyList(Index), "", useSection, useFile), NewSection, useFile
    Next Index
    
    'delete the old section
    deleteSection useSection, useFile
End Sub

Public Sub renameKey(ByVal currentKey As String, ByVal NewKey As String, _
                     Optional ByVal section As String, Optional ByVal INIFileLoc As String)
    Dim useFile As String
    Dim useSection As String
    
    If IsMissing(INIFileLoc) Or INIFileLoc = "" Then
        useFile = var_INIfile
    Else
        useFile = INIFileLoc
    End If
    
    If IsMissing(section) Or section = "" Then
        useSection = currentSection
    Else
        useSection = section
    End If
    
    'make sure the file exists
    primeINIFile
    
    'get the value and create the new key
    setKeyValue NewKey, getKeyValue(currentKey, "", useSection, useFile), useSection, useFile
    'delete the old key
    deleteKey currentKey, useSection, useFile
End Sub
'
'****************************** End Rename Methods ******************************
