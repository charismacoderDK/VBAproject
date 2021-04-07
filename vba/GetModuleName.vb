Sub GetModulesNameOnWorksbooks()

Dim modName As String
Dim wb As Workbook
Dim l As Long

'Set wb = ThisWorkbook
Set wb = Workbooks("PERSONAL.XLSB")

For l = 1 To wb.VBProject.VBComponents.Count

    With wb.VBProject.VBComponents(l)
    'If .Type = 1 Then
    'modName = modName & vbCr & .Name
    'End If
    
    If .Type = 1 Then _
    modName = modName & vbCr & .Name
    End With

Next

MsgBox "Module Names:" & vbCr & modName

Set wb = Nothing

End Sub

Sub GetModulesNameOnPERSONAL()

Dim VBAEditor As VBIDE.VBE
Dim VBProj As VBIDE.VBProject
Dim modName As String
Dim wb As Workbook
Dim l As Long

Set VBAEditor = Application.VBE

Set wb = Workbooks("PERSONAL.XLSB")

Debug.Print wb.Name

Set VBProj = wb.VBAEditor.VBProjects(1)

'Debug.Print VBProj.VBComponents.Count

For l = 1 To VBProj.VBComponents.Count
With VBProj.VBComponents(l)
modName = modName & vbCr & .Name
End With
Next

MsgBox "Module Names:" & vbCr & modName

Set wb = Nothing
Set VBAEditor = Nothing
Set VBProj = Nothing


End Sub

'Set VBCodelist = VBProj.VBComponents.Item(0).CodeModule

'GG = VBCodelist
'Debug.Print GG

Sub break()

If ThisWorkbook.Sheets("Sheet2").Range("B6").Value <> 0 Then _
Range("C6").Value = 0
Range("C5").Value = 0

End Sub
