Attribute VB_Name = "Pages2Word"
' Adds "Import Pages": converts a .pages file with Pages2Word.exe and inserts it
' into the current document (styles come from the current document).
' Requires Pages2Word.exe -> "Propojit s Wordem" to have been clicked once.
Option Explicit

Private Function ConverterPath() As String
    On Error Resume Next
    ConverterPath = CreateObject("WScript.Shell").RegRead("HKCU\Software\Pages2Word\ExePath")
End Function

Public Sub ImportPages()
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim exe As String: exe = ConverterPath()
    If exe = "" Or Not fso.FileExists(exe) Then
        MsgBox "P¯evodnÌk nebyl nalezen." & vbCr & vbCr & _
               "Spusùte Pages2Word.exe a kliknÏte na 'Propojit s Wordem'.", vbExclamation, "Import z Pages"
        Exit Sub
    End If

    Dim dlg As FileDialog: Set dlg = Application.FileDialog(msoFileDialogFilePicker)
    dlg.Title = "Vyberte dokument Pages"
    dlg.AllowMultiSelect = False
    dlg.Filters.Clear
    dlg.Filters.Add "Dokumenty Pages", "*.pages"
    If dlg.Show <> -1 Then Exit Sub

    Dim src As String: src = dlg.SelectedItems(1)
    Dim tmp As String: tmp = fso.BuildPath(fso.GetSpecialFolder(2), "pages2word-import.docx")
    Dim cmd As String
    cmd = """" & exe & """ --output """ & tmp & """ """ & src & """"

    System.Cursor = wdCursorWait
    Dim code As Long: code = CreateObject("WScript.Shell").Run(cmd, 0, True)
    System.Cursor = wdCursorNormal
    If code <> 0 Or Not fso.FileExists(tmp) Then
        MsgBox "Soubor se nepoda¯ilo p¯evÈst.", vbExclamation, "Import z Pages"
        Exit Sub
    End If

    If Documents.Count = 0 Then Documents.Add
    Selection.InsertFile FileName:=tmp, ConfirmConversions:=False
    fso.DeleteFile tmp
End Sub
