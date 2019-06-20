Attribute VB_Name = "RibbonSetup"
'**************************************************************
'AUTHOR: Chris Newman, TheSpreadsheetGuru
'Instructions on how to use this template can be found at:
'https://www.thespreadsheetguru.com/blog/step-by-step-instructions-create-first-excel-ribbon-vba-addin
'https://www.thespreadsheetguru.com/myfirstaddin-help/
'**************************************************************
Public Const Version = "2.2"

Sub GetVisible(control As IRibbonControl, ByRef MakeVisible)
'PURPOSE: Show/Hide buttons based on how many you need (False = Hide/True = Show)

Select Case control.ID
  Case "GroupA": MakeVisible = True
  Case "aButton01": MakeVisible = True
  Case "aButton02": MakeVisible = True
  Case "aButton03": MakeVisible = True
'  Case "aButton04": MakeVisible = True
'  Case "aButton05": MakeVisible = True
'  Case "aButton06": MakeVisible = True
'  Case "aButton07": MakeVisible = True
'  Case "aButton08": MakeVisible = True
'  Case "aButton09": MakeVisible = True
'  Case "aButton10": MakeVisible = True
  
  Case "GroupB": MakeVisible = True
  Case "bButton01": MakeVisible = True
  Case "bButton02": MakeVisible = True
  Case "bButton03": MakeVisible = True
  Case "bButton04": MakeVisible = True
  Case "bButton05": MakeVisible = True
  Case "bButton06": MakeVisible = True
'  Case "bButton07": MakeVisible = True
'  Case "bButton08": MakeVisible = True
'  Case "bButton09": MakeVisible = True
'  Case "bButton10": MakeVisible = True
  
  Case "GroupC": MakeVisible = True
  Case "cButton01": MakeVisible = True
  Case "cButton02": MakeVisible = True
'  Case "cButton03": MakeVisible = True
'  Case "cButton04": MakeVisible = True
'  Case "cButton05": MakeVisible = True
'  Case "cButton06": MakeVisible = True
'  Case "cButton07": MakeVisible = True
'  Case "cButton08": MakeVisible = True
'  Case "cButton09": MakeVisible = True
'  Case "cButton10": MakeVisible = True
  
  Case "GroupD": MakeVisible = True
  Case "dButton01": MakeVisible = True
'  Case "dButton02": MakeVisible = True
'  Case "dButton03": MakeVisible = True
'  Case "dButton04": MakeVisible = True
'  Case "dButton05": MakeVisible = True
'  Case "dButton06": MakeVisible = True
'  Case "dButton07": MakeVisible = True
'  Case "dButton08": MakeVisible = True
'  Case "dButton09": MakeVisible = True
'  Case "dButton10": MakeVisible = True
'
'  Case "GroupE": MakeVisible = True
'  Case "eButton01": MakeVisible = True
'  Case "eButton02": MakeVisible = True
'  Case "eButton03": MakeVisible = True
'  Case "eButton04": MakeVisible = True
'  Case "eButton05": MakeVisible = True
'  Case "eButton06": MakeVisible = True
'  Case "eButton07": MakeVisible = True
'  Case "eButton08": MakeVisible = True
'  Case "eButton09": MakeVisible = True
'  Case "eButton10": MakeVisible = True
'
'  Case "GroupF": MakeVisible = False
'  Case "fButton01": MakeVisible = True
'  Case "fButton02": MakeVisible = True
'  Case "fButton03": MakeVisible = True
'  Case "fButton04": MakeVisible = True
'  Case "fButton05": MakeVisible = True
'  Case "fButton06": MakeVisible = True
'  Case "fButton07": MakeVisible = True
'  Case "fButton08": MakeVisible = True
'  Case "fButton09": MakeVisible = True
'  Case "fButton10": MakeVisible = True
  
End Select

End Sub

Sub GetLabel(ByVal control As IRibbonControl, ByRef Labeling)
'PURPOSE: Determine the text to go along with your Tab, Groups, and Buttons

Select Case control.ID
  
  Case "CustomTab": Labeling = "PGE186 Tools  v" & Version
  
  Case "GroupA": Labeling = "CAISO Update Tab"
  Case "aButton01": Labeling = "Color Row by Change Type"
  Case "aButton02": Labeling = "   Update Tower Numbers"  ' extra space intentional
  Case "aButton03": Labeling = "Add Change Desc"
'  Case "aButton04": Labeling = "Button"
'  Case "aButton05": Labeling = "Button"
'  Case "aButton06": Labeling = "Button"
'  Case "aButton07": Labeling = "Button"
'  Case "aButton08": Labeling = "Button"
'  Case "aButton09": Labeling = "Button"
'  Case "aButton10": Labeling = "Button"
  
  Case "GroupB": Labeling = "Summary"
  Case "bButton01": Labeling = "Ratings Requested"
  Case "bButton02": Labeling = "Relay Request"
  Case "bButton03": Labeling = "Equipment Added"
  Case "bButton04": Labeling = "Equipment Retired"
  Case "bButton05": Labeling = "Source Docs Used  "  ' extra space intentional
  Case "bButton06": Labeling = "Source Docs Reference"
'  Case "bButton07": Labeling = "Button"
'  Case "bButton08": Labeling = "Button"
'  Case "bButton09": Labeling = "Button"
'  Case "bButton10": Labeling = "Button"
  
  Case "GroupC": Labeling = "Info Requests"
  Case "cButton01": Labeling = "CT"
  Case "cButton02": Labeling = "Field Verification"
'  Case "cButton03": Labeling = "Button"
'  Case "cButton04": Labeling = "Button"
'  Case "cButton05": Labeling = "Button"
'  Case "cButton06": Labeling = "Button"
'  Case "cButton07": Labeling = "Button"
'  Case "cButton08": Labeling = "Button"
'  Case "cButton09": Labeling = "Button"
'  Case "cButton10": Labeling = "Button"
  
  Case "GroupD": Labeling = "PGE186 Tools"
  Case "dButton01": Labeling = "Update Tools"
'  Case "dButton02": Labeling = "Button"
'  Case "dButton03": Labeling = "Button"
'  Case "dButton04": Labeling = "Button"
'  Case "dButton05": Labeling = "Button"
'  Case "dButton06": Labeling = "Button"
'  Case "dButton07": Labeling = "Button"
'  Case "dButton08": Labeling = "Button"
'  Case "dButton09": Labeling = "Button"
'  Case "dButton10": Labeling = "Button"
'
'  Case "GroupE": Labeling = "Update"
'  Case "eButton01": Labeling = "Get Current Add-in"
'  Case "eButton02": Labeling = "Button"
'  Case "eButton03": Labeling = "Button"
'  Case "eButton04": Labeling = "Button"
'  Case "eButton05": Labeling = "Button"
'  Case "eButton06": Labeling = "Button"
'  Case "eButton07": Labeling = "Button"
'  Case "eButton08": Labeling = "Button"
'  Case "eButton09": Labeling = "Button"
'  Case "eButton10": Labeling = "Button"
'
'  Case "GroupF": Labeling = "Group Name"
'  Case "fButton01": Labeling = "Button"
'  Case "fButton02": Labeling = "Button"
'  Case "fButton03": Labeling = "Button"
'  Case "fButton04": Labeling = "Button"
'  Case "fButton05": Labeling = "Button"
'  Case "fButton06": Labeling = "Button"
'  Case "fButton07": Labeling = "Button"
'  Case "fButton08": Labeling = "Button"
'  Case "fButton09": Labeling = "Button"
'  Case "fButton10": Labeling = "Button"
  
End Select
   
End Sub

Sub GetImage(control As IRibbonControl, ByRef RibbonImage)
'PURPOSE: Tell each button which image to load from the Microsoft Icon Library
'TIPS: Image names are case sensitive, if image does not appear in ribbon after re-starting Excel, the image name is incorrect

Select Case control.ID
  
  Case "aButton01": RibbonImage = "ViewBackToColorView"
  Case "aButton02": RibbonImage = "RelationshipsHideTable"
  Case "aButton03": RibbonImage = "FileDocumentInspect"
'  Case "aButton04": RibbonImage = "ObjectPictureFill"
'  Case "aButton05": RibbonImage = "ObjectPictureFill"
'  Case "aButton06": RibbonImage = "ObjectPictureFill"
'  Case "aButton07": RibbonImage = "ObjectPictureFill"
'  Case "aButton08": RibbonImage = "ObjectPictureFill"
'  Case "aButton09": RibbonImage = "ObjectPictureFill"
'  Case "aButton10": RibbonImage = "ObjectPictureFill"
  
  Case "bButton01": RibbonImage = "IndexMarkEntry"
  Case "bButton02": RibbonImage = "MacroArguments"
  Case "bButton03": RibbonImage = "AppointmentColor2"
  Case "bButton04": RibbonImage = "AppointmentColor10"
  Case "bButton05": RibbonImage = "SmartArtAddBullet"
  Case "bButton06": RibbonImage = "AccessListCustomDatasheet"
'  Case "bButton07": RibbonImage = "ObjectPictureFill"
'  Case "bButton08": RibbonImage = "ObjectPictureFill"
'  Case "bButton09": RibbonImage = "ObjectPictureFill"
'  Case "bButton10": RibbonImage = "ObjectPictureFill"
  
  Case "cButton01": RibbonImage = "HighImportance"
  Case "cButton02": RibbonImage = "TentativeAcceptInvitation"
'  Case "cButton03": RibbonImage = "ObjectPictureFill"
'  Case "cButton04": RibbonImage = "ObjectPictureFill"
'  Case "cButton05": RibbonImage = "ObjectPictureFill"
'  Case "cButton06": RibbonImage = "ObjectPictureFill"
'  Case "cButton07": RibbonImage = "ObjectPictureFill"
'  Case "cButton08": RibbonImage = "ObjectPictureFill"
'  Case "cButton09": RibbonImage = "ObjectPictureFill"
'  Case "cButton10": RibbonImage = "ObjectPictureFill"
'
  Case "dButton01": RibbonImage = "ControlsGallery"
'  Case "dButton02": RibbonImage = "ObjectPictureFill"
'  Case "dButton03": RibbonImage = "ObjectPictureFill"
'  Case "dButton04": RibbonImage = "ObjectPictureFill"
'  Case "dButton05": RibbonImage = "ObjectPictureFill"
'  Case "dButton06": RibbonImage = "ObjectPictureFill"
'  Case "dButton07": RibbonImage = "ObjectPictureFill"
'  Case "dButton08": RibbonImage = "ObjectPictureFill"
'  Case "dButton09": RibbonImage = "ObjectPictureFill"
'  Case "dButton10": RibbonImage = "ObjectPictureFill"
'
'  Case "eButton01": RibbonImage = "ControlsGallery"
'  Case "eButton02": RibbonImage = "ObjectPictureFill"
'  Case "eButton03": RibbonImage = "ObjectPictureFill"
'  Case "eButton04": RibbonImage = "ObjectPictureFill"
'  Case "eButton05": RibbonImage = "ObjectPictureFill"
'  Case "eButton06": RibbonImage = "ObjectPictureFill"
'  Case "eButton07": RibbonImage = "ObjectPictureFill"
'  Case "eButton08": RibbonImage = "ObjectPictureFill"
'  Case "eButton09": RibbonImage = "ObjectPictureFill"
'  Case "eButton10": RibbonImage = "ObjectPictureFill"
'
'  Case "fButton01": RibbonImage = "ObjectPictureFill"
'  Case "fButton02": RibbonImage = "ObjectPictureFill"
'  Case "fButton03": RibbonImage = "ObjectPictureFill"
'  Case "fButton04": RibbonImage = "ObjectPictureFill"
'  Case "fButton05": RibbonImage = "ObjectPictureFill"
'  Case "fButton06": RibbonImage = "ObjectPictureFill"
'  Case "fButton07": RibbonImage = "ObjectPictureFill"
'  Case "fButton08": RibbonImage = "ObjectPictureFill"
'  Case "fButton09": RibbonImage = "ObjectPictureFill"
'  Case "fButton10": RibbonImage = "ObjectPictureFill"
  
End Select

End Sub

Sub GetSize(control As IRibbonControl, ByRef Size)
'PURPOSE: Determine if the button size is large or small

Const Large As Integer = 1
Const Small As Integer = 0

Select Case control.ID
    
  Case "aButton01": Size = Large
  Case "aButton02": Size = Large
  Case "aButton03": Size = Large
'  Case "aButton04": Size = Small
'  Case "aButton05": Size = Small
'  Case "aButton06": Size = Small
'  Case "aButton07": Size = Small
'  Case "aButton08": Size = Small
'  Case "aButton09": Size = Small
'  Case "aButton10": Size = Small
  
  Case "bButton01": Size = Large
  Case "bButton02": Size = Large
  Case "bButton03": Size = Large
  Case "bButton04": Size = Large
  Case "bButton05": Size = Large
  Case "bButton06": Size = Large
'  Case "bButton07": Size = Small
'  Case "bButton08": Size = Small
'  Case "bButton09": Size = Small
'  Case "bButton10": Size = Small
  
  Case "cButton01": Size = Small
  Case "cButton02": Size = Small
'  Case "cButton03": Size = Small
'  Case "cButton04": Size = Small
'  Case "cButton05": Size = Small
'  Case "cButton06": Size = Small
'  Case "cButton07": Size = Small
'  Case "cButton08": Size = Small
'  Case "cButton09": Size = Small
'  Case "cButton10": Size = Small
'
  Case "dButton01": Size = Large
'  Case "dButton02": Size = Small
'  Case "dButton03": Size = Small
'  Case "dButton04": Size = Small
'  Case "dButton05": Size = Small
'  Case "dButton06": Size = Small
'  Case "dButton07": Size = Small
'  Case "dButton08": Size = Small
'  Case "dButton09": Size = Small
'  Case "dButton10": Size = Small
'
'  Case "eButton01": Size = Large
'  Case "eButton02": Size = Small
'  Case "eButton03": Size = Small
'  Case "eButton04": Size = Small
'  Case "eButton05": Size = Small
'  Case "eButton06": Size = Small
'  Case "eButton07": Size = Small
'  Case "eButton08": Size = Small
'  Case "eButton09": Size = Small
'  Case "eButton10": Size = Small
'
'  Case "fButton01": Size = Large
'  Case "fButton02": Size = Small
'  Case "fButton03": Size = Small
'  Case "fButton04": Size = Small
'  Case "fButton05": Size = Small
'  Case "fButton06": Size = Small
'  Case "fButton07": Size = Small
'  Case "fButton08": Size = Small
'  Case "fButton09": Size = Small
'  Case "fButton10": Size = Small
  
End Select

End Sub

Sub RunMacro(control As IRibbonControl)
'PURPOSE: Tell each button which macro subroutine to run when clicked

Select Case control.ID
  
  Case "aButton01": Application.Run "Paint"
  Case "aButton02": Application.Run "AddZeroes"
  Case "aButton03": Application.Run "AddChangeDescCode"
'  Case "aButton04": Application.Run "DummyMacro"
'  Case "aButton05": Application.Run "DummyMacro"
'  Case "aButton06": Application.Run "DummyMacro"
'  Case "aButton07": Application.Run "DummyMacro"
'  Case "aButton08": Application.Run "DummyMacro"
'  Case "aButton09": Application.Run "DummyMacro"
'  Case "aButton10": Application.Run "DummyMacro"
  
  Case "bButton01": Application.Run "MakeRatingsRequestedTable"
  Case "bButton02": Application.Run "MakeRelayRequestTable"
  Case "bButton03": Application.Run "MakeEquipAddedTable"
  Case "bButton04": Application.Run "MakeEquipRetiredTable"
  Case "bButton05": Application.Run "MakeSourceDocsUsedTable"
  Case "bButton06": Application.Run "MakeSourceDocsRefTable"
'  Case "bButton07": Application.Run "DummyMacro"
'  Case "bButton08": Application.Run "DummyMacro"
'  Case "bButton09": Application.Run "DummyMacro"
'  Case "bButton10": Application.Run "DummyMacro"
  
  Case "cButton01": Application.Run "MakeInfoRequestCTTable"
  Case "cButton02": Application.Run "MakeInfoRequestTable"
'  Case "cButton03": Application.Run "DummyMacro"
'  Case "cButton04": Application.Run "DummyMacro"
'  Case "cButton05": Application.Run "DummyMacro"
'  Case "cButton06": Application.Run "DummyMacro"
'  Case "cButton07": Application.Run "DummyMacro"
'  Case "cButton08": Application.Run "DummyMacro"
'  Case "cButton09": Application.Run "DummyMacro"
'  Case "cButton10": Application.Run "DummyMacro"
'
  Case "dButton01": Application.Run "CompareAddIns"
'  Case "dButton02": Application.Run "DummyMacro"
'  Case "dButton03": Application.Run "DummyMacro"
'  Case "dButton04": Application.Run "DummyMacro"
'  Case "dButton05": Application.Run "DummyMacro"
'  Case "dButton06": Application.Run "DummyMacro"
'  Case "dButton07": Application.Run "DummyMacro"
'  Case "dButton08": Application.Run "DummyMacro"
'  Case "dButton09": Application.Run "DummyMacro"
'  Case "dButton10": Application.Run "DummyMacro"
'
'  Case "eButton01": Application.Run "DummyMacro"
'  Case "eButton02": Application.Run "DummyMacro"
'  Case "eButton03": Application.Run "DummyMacro"
'  Case "eButton04": Application.Run "DummyMacro"
'  Case "eButton05": Application.Run "DummyMacro"
'  Case "eButton06": Application.Run "DummyMacro"
'  Case "eButton07": Application.Run "DummyMacro"
'  Case "eButton08": Application.Run "DummyMacro"
'  Case "eButton09": Application.Run "DummyMacro"
'  Case "eButton10": Application.Run "DummyMacro"
'
'  Case "fButton01": Application.Run "DummyMacro"
'  Case "fButton02": Application.Run "DummyMacro"
'  Case "fButton03": Application.Run "DummyMacro"
'  Case "fButton04": Application.Run "DummyMacro"
'  Case "fButton05": Application.Run "DummyMacro"
'  Case "fButton06": Application.Run "DummyMacro"
'  Case "fButton07": Application.Run "DummyMacro"
'  Case "fButton08": Application.Run "DummyMacro"
'  Case "fButton09": Application.Run "DummyMacro"
'  Case "fButton10": Application.Run "DummyMacro"

 End Select
    
End Sub

Sub GetScreentip(control As IRibbonControl, ByRef Screentip)
'PURPOSE: Display a specific macro description when the mouse hovers over a button

Select Case control.ID
  
  Case "aButton01": Screentip = "Color Rows based on Type of Change"
  Case "aButton02": Screentip = "Format tower/pole numbers with leading zeroes and red text. Note: Can't undo! Use only on *copy* of row."
  Case "aButton03": Screentip = "Add ChangeIDs based on Change Type and TLS"
'  Case "aButton04": Screentip = "Description"
'  Case "aButton05": Screentip = "Description"
'  Case "aButton06": Screentip = "Description"
'  Case "aButton07": Screentip = "Description"
'  Case "aButton08": Screentip = "Description"
'  Case "aButton09": Screentip = "Description"
'  Case "aButton10": Screentip = "Description"
  
  Case "bButton01": Screentip = "Pull Verification rows into table for copying to Summary Report"
  Case "bButton02": Screentip = "Copy RLY components into table for copying to Summary Report"
  Case "bButton03": Screentip = "Pull 'Create' rows into table for copying to Summary Report"
  Case "bButton04": Screentip = "Pull 'Retire' rows into table for copying to Summary Report"
  Case "bButton05": Screentip = "Copy, split, and sort Comments to summarize documents used"
  Case "bButton06": Screentip = "Copy the needed rows for creating the Source Document Reference Table"
'  Case "bButton07": Screentip = "Description"
'  Case "bButton08": Screentip = "Description"
'  Case "bButton09": Screentip = "Description"
'  Case "bButton10": Screentip = "Description"
  
  Case "cButton01": Screentip = "Make info request rows in 'CT' tab style"
  Case "cButton02": Screentip = "Make info request rows in 'Field Verification' tab style"
'  Case "cButton03": Screentip = "Description"
'  Case "cButton04": Screentip = "Description"
'  Case "cButton05": Screentip = "Description"
'  Case "cButton06": Screentip = "Description"
'  Case "cButton07": Screentip = "Description"
'  Case "cButton08": Screentip = "Description"
'  Case "cButton09": Screentip = "Description"
'  Case "cButton10": Screentip = "Description"
'
  Case "dButton01": Screentip = "Update PGE186 tools to the current version"
'  Case "dButton02": Screentip = "Description"
'  Case "dButton03": Screentip = "Description"
'  Case "dButton04": Screentip = "Description"
'  Case "dButton05": Screentip = "Description"
'  Case "dButton06": Screentip = "Description"
'  Case "dButton07": Screentip = "Description"
'  Case "dButton08": Screentip = "Description"
'  Case "dButton09": Screentip = "Description"
'  Case "dButton10": Screentip = "Description"
'
'  Case "eButton01": Screentip = "Description"
'  Case "eButton02": Screentip = "Description"
'  Case "eButton03": Screentip = "Description"
'  Case "eButton04": Screentip = "Description"
'  Case "eButton05": Screentip = "Description"
'  Case "eButton06": Screentip = "Description"
'  Case "eButton07": Screentip = "Description"
'  Case "eButton08": Screentip = "Description"
'  Case "eButton09": Screentip = "Description"
'  Case "eButton10": Screentip = "Description"
'
'  Case "fButton01": Screentip = "Description"
'  Case "fButton02": Screentip = "Description"
'  Case "fButton03": Screentip = "Description"
'  Case "fButton04": Screentip = "Description"
'  Case "fButton05": Screentip = "Description"
'  Case "fButton06": Screentip = "Description"
'  Case "fButton07": Screentip = "Description"
'  Case "fButton08": Screentip = "Description"
'  Case "fButton09": Screentip = "Description"
'  Case "fButton10": Screentip = "Description"
  
End Select

End Sub





