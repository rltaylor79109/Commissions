Option Explicit

'------------------------------------------------------------------------------'
' Module Name: modMacrosAppSpecific
' Summary: Defines application specific user-facing methods that show up in
'   Alt+F8 because this module does not define the "Option Private Module" directive.
' Date Created: 2026-09-26
' Date Last Modified: 2026-09-26
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Fields
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Constant Fields
'------------------------------------------------------------------------------'

' The name of this module.
Private Const MODULE_NAME As String = "modMacrosAppSpecific"

'------------------------------------------------------------------------------'
' Methods
'------------------------------------------------------------------------------'

'------------------------------------------------------------------------------'
' Public Methods
'------------------------------------------------------------------------------'
'------------------------------------------------------------------------------'
' Summary: Saves the current workbook as a macro free workbook.
' Remarks: Requires reference to "Microsoft Scripting Runtime."
' Date Created: 2026-09-25
' Date Last Modified: 2026-09-26
'------------------------------------------------------------------------------'
Public Sub RunSaveMacroFreeWorkingCopy()
  On Error GoTo Err_Proc
  Const METHOD_NAME As String = "RunSaveMacroFreeWorkingCopy"

  Dim fileFilter As String
  Dim fso As FileSystemObject
  Dim initialFilename As String
  Dim newBaseName As String
  Dim newPath As String
  Dim selectedPath As Variant
  Dim wb As Workbook
  Dim ws As Worksheet
  Dim wsDelColl As Collection
  Dim wsHideColl As Collection

  Set wb = ThisWorkbook
  
  initialFilename = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
  fileFilter = "Standard Excel Files (*.xlsx), *.xlsx"

  selectedPath = Application.GetSaveAsFilename( _
        initialFilename:=initialFilename, _
        fileFilter:=fileFilter, _
        FilterIndex:=1, _
        title:="Select Location and Name to Save Macro Free Version of This Workbook")
  
  If selectedPath = False Then
    MsgBox _
      prompt:="Save cancelled by user.", _
      buttons:=vbInformation, _
      title:="Save Macro Free Working Copy"
    GoTo Exit_Proc
  End If
  
  Set fso = New FileSystemObject
  newPath = fso.GetParentFolderName(selectedPath)
  newBaseName = fso.GetBaseName(selectedPath)
  
  Set wsDelColl = New Collection
  For Each ws In wb.Sheets
    Select Case ws.codeName
      Case _
        SheetCommissions.codeName, _
        SheetDaLineItems.codeName, _
        SheetDropdowns.codeName, _
        SheetMiscVals.codeName, _
        SheetPaymentOutsideOfClosing.codeName, _
        SheetTeamMemberInfo.codeName
      Case Else
        wsDelColl.Add ws
    End Select
  Next ws
  
  Set wsHideColl = New Collection
  wsHideColl.Add SheetDropdowns
  wsHideColl.Add SheetMiscVals
  
  SaveMacroFreeWorkingCopy _
    newPath:=newPath, _
    newBaseName:=newBaseName, _
    wsHideColl:=wsHideColl, _
    wsDelColl:=wsDelColl

Exit_Proc:
  Exit Sub
Err_Proc:
  ShowMethodErrorMsgBox err, MODULE_NAME, METHOD_NAME
  Resume Exit_Proc
End Sub

