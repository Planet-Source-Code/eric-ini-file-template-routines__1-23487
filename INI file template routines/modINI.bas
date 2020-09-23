Attribute VB_Name = "modINI"
Option Explicit

'
' This module is a template you can use to create new applications
' The steps you need to do:
' Declare the public variales for storing the value you want to save and restore
' In readAllSettings:
'   - determine the number of sections
'   - determine the version (start at "1.0")
'   - add/remove lines like: "If Not readSectionXX Then writeSectionXX"
' In writeAllSettings:
'   - add/remove lines like: "writeSectionXX"
' Add/remove routines readSectionXX and writeSectionXX
'
' In your main- or form_load-routine just call readAllSettings
' and setup your application.  Have a look at the sample application form_Load
'

Public INI As New clsINI
Dim iniVersion As String

'-------------------------------------- Declare ini-setting variables here
'
Public inputDirectory As String
Public deleteInputFiles As String
Public autoStart As String
Public manual As String
Public optionType As Integer
Public sectionField As String
Public keyField As String
Public valueField As String

Public Sub readAllSettings()

   Dim numberOfSections As Integer

   numberOfSections = 4                              ' current valid settings
   iniVersion = "1.7"
   
   ' First try to read the values from the inifile
   ' If not in inifile: a defaultvalue is returned AND written to the inifile
   ' If section is out of date then delete and re-write the section
   '
   If Not readDefaultSection Then writeDefaultSection
   If Not readSection1 Then writeSection1
   If Not readSection2 Then writeSection2
   If Not readSection3 Then writeSection3
   
   ' Check if iniFileVersion is up to date:  number of sections, version
   '
   If numberOfSections <> INI.getNumberOfSections Or _
      iniVersion <> INI.getVersion Then
   
      ' Inifilesection is out of date so rewrite all sections
      ' This way you can force a rewrite of the inifile by
      ' incrementing the inifile version.
      '
      INI.deleteAllSections
      writeAllSettings                              ' write all sections again
      
   End If
   
End Sub

Public Sub writeAllSettings()
   
   writeDefaultSection
   writeSection1                                  ' Writing a section, deletes it first
   writeSection2                                  ' so we know we start with a clean version
   writeSection3
   
End Sub

Public Function readDefaultSection() As Boolean
   Dim numberOfSettings As Integer
   
   numberOfSettings = 1
   
   INI.sectionName = INI.defaultSectionName
   
   
   ' return true if number of keys are ok
   '
   If INI.getNumberOfKeysInSection = numberOfSettings Then readDefaultSection = True
End Function

Public Sub writeDefaultSection()

   ' All general settings here
   '
   INI.sectionName = INI.defaultSectionName
   INI.deleteSection                              ' Always cleanup to start with
                                                  ' a fresh/uptodate default section
                                                  
   INI.setKeyValue "version", iniVersion          ' This all there is

End Sub

Public Function readSection1() As Boolean
   Dim numberOfSettings As Integer
   
   numberOfSettings = 2
   
   INI.sectionName = "section-1"
   inputDirectory = INI.getKeyValue("inputDirectory", App.Path)
   manual = INI.getKeyValue("manual", "2")
   
   ' return true if number of keys are ok
   '
   If INI.getNumberOfKeysInSection = numberOfSettings Then readSection1 = True
End Function

Public Function readSection2() As Boolean
   Dim numberOfSettings As Integer
   
   numberOfSettings = 2
   
   INI.sectionName = "section-2"
   deleteInputFiles = INI.getKeyValue("deleteInputFiles", "0")
   autoStart = INI.getKeyValue("autoStart", "0")
   
   ' return true if number of keys are ok
   '
   If INI.getNumberOfKeysInSection = numberOfSettings Then readSection2 = True
End Function

Public Function readSection3() As Boolean
   Dim numberOfSettings As Integer
   
   numberOfSettings = 4
   
   INI.sectionName = "manualInput"
   
   optionType = INI.getKeyValue("optionType", "0")
   sectionField = INI.getKeyValue("sectionField", "newsection")
   keyField = INI.getKeyValue("keyField", "newkey")
   valueField = INI.getKeyValue("valueField", "newvalue")
   
   
   ' return true if number of keys are ok
   '
   If INI.getNumberOfKeysInSection = numberOfSettings Then readSection3 = True
End Function

Public Sub writeSection1()

   ' settings section-1
   '
   INI.sectionName = "section-1"
   INI.deleteSection
   INI.setKeyValue "inputDirectory", inputDirectory
   INI.setKeyValue "manual", manual
   
End Sub

Public Sub writeSection2()

   ' settings section-2
   '
   INI.sectionName = "section-2"
   INI.deleteSection
   INI.setKeyValue "autoStart", autoStart
   INI.setKeyValue "deleteInputFiles", deleteInputFiles
   
End Sub

Public Sub writeSection3()

'Public optionType As Integer
'Public sectionField As String
'Public keyField As String
'Public valueField As String

   ' settings manualInput
   '
   INI.sectionName = "manualInput"
   INI.deleteSection
   INI.setKeyValue "optionType", optionType
   INI.setKeyValue "sectionField", sectionField
   INI.setKeyValue "keyField", keyField
   INI.setKeyValue "valueField", valueField
   
End Sub
