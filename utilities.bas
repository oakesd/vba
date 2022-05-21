Option Compare Database
Option Explicit

'for use with CurrentDbC
Private m_db As DAO.Database

'''
'credit to Michael Kaplan, retrieved from MSDN
'https://social.msdn.microsoft.com/Forums/office/en-US/9993d229-8a00-4a59-a796-dfa2dad505bc/cannot-open-any-more-databases?forum=accessdev
'''
Public Property Get CurrentDbC() As DAO.Database

    If (m_db Is Nothing) Then
        Set m_db = CurrentDb
    End If

    Set CurrentDbC = m_db

End Property

Public Function SelectFile()
'@author: Dustin Oakes
'April 2018
'opens a file dialog and returns the path of the selected file as
'a string value.  Returns a zero-length string if no file is selected
On Error GoTo errSelectFile

Dim strPath As String
strPath = ""

'create the file dialog object
Dim objFileDialog As Office.FileDialog
Set objFileDialog = Application.FileDialog(msoFileDialogFilePicker)

'if a file is selected
If objFileDialog.Show Then
    'set the file path to the selected file
    strPath = objFileDialog.SelectedItems(1)
End If

SelectFile = strPath

exitSelectFile:
    If Not objFileDialog Is Nothing Then Set objFileDialog = Nothing
    Exit Function
    
errSelectFile:
    Dim strErrMsg As String
    strErrMsg = "An error occurred while attempting to select the file." & vbCrLf & vbCrLf & _
                "Error: " & Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
                "Please contact your administrator."
    MsgBox strErrMsg
    Resume exitSelectFile

End Function

Public Function ExcelColumnIndexToLetter(columnIndex As Integer)
'@author: Dustin Oakes, 2018
'converts MS Excel column index to alpha label
'@param columnIndex = numerical index of the column in a MS Excel Worksheet
'@return the alpha label that corresponds to columnIndex
On Error GoTo errExcelColumnIndexToLetter

Dim strLabel As String
strLabel = ""

Dim i As Integer
i = columnIndex - 1

Do While i > 25
    strLabel = Chr(65 + (i Mod 26)) & strLabel
    i = i \ 26 - 1
Loop

strLabel = Chr(65 + i) & strLabel

'return the column label as a string
ExcelColumnIndexToLetter = strLabel

exitExcelColumnIndexToLetter:
    Exit Function

errExcelColumnIndexToLetter:
    Dim strMsg As String
    strMsg = "Error: " & Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
             "Please contact your administrator."
    MsgBox strMsg
    Resume exitExcelColumnIndexToLetter

End Function

Public Sub DropTables(arrTables As Variant)
'@author: Dustin Oakes, 2018
On Error GoTo errDropTables

Dim varTable As Variant
Dim strTable As String
Dim strCriteria As String
Dim strSql As String
Dim arrRelations As Variant
Dim i As Integer

For Each varTable In arrTables
    strCriteria = "[Name]=""" & varTable & """ AND [Type] IN (1, 6)"
    If DCount("*", "MSysObjects", strCriteria) > 0 Then
        strTable = varTable
        'remove any relationships involving the table so it can be dropped
        arrRelations = GetRelations(strTable)
        If Not IsNull(arrRelations) Then
            For i = 0 To UBound(arrRelations) - 1
                strSql = "ALTER TABLE " & arrRelations(i, 0) & " DROP CONSTRAINT " & arrRelations(i, 1)
                CurrentDbC.Execute strSql
            Next i
        End If
        
        strSql = "DROP TABLE " & strTable
        CurrentDbC.Execute strSql
    End If
Next

exitDropTables:
    Exit Sub

errDropTables:
    Dim strErrMsg As String
    strErrMsg = "Error: " & Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
                "Please contact your administrator."
    MsgBox strErrMsg
    Resume exitDropTables

End Sub

Public Sub DropLinks()
'@author: Dustin Oakes, 2018
On Error GoTo errDropLinks

Dim tdf As TableDef

CloseAll

For Each tdf In CurrentDbC.TableDefs
    If Left(tdf.Connect, 10) = ";DATABASE=" Then DoCmd.DeleteObject acTable, tdf.NAME
Next tdf

exitDropLinks:
    Exit Sub
    
errDropLinks:
    Dim strErrMsg As String
    strErrMsg = "Error: " & Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
                "Please contact your administrator."
    MsgBox strErrMsg
    Resume exitDropLinks


End Sub

Public Sub CloseAll()
'@author: Dustin Oakes, 2018
On Error GoTo errCloseAll

Dim tdf As TableDef
Dim qry As QueryDef
Dim rpt As Report
Dim frm As Form

For Each tdf In CurrentDbC.TableDefs
    DoCmd.Close acTable, tdf.NAME
Next tdf

For Each qry In CurrentDbC.QueryDefs
    DoCmd.Close acQuery, qry.NAME
Next qry

For Each rpt In Application.Reports
    DoCmd.Close acReport, rpt.NAME
Next rpt

For Each frm In Application.Forms
    DoCmd.Close acForm, frm.NAME
Next frm

exitCloseAll:
    Exit Sub
    
errCloseAll:
     Dim strErrMsg As String
    strErrMsg = "Error: " & Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
                "Please contact your administrator."
    MsgBox strErrMsg
    Resume exitCloseAll

End Sub

Public Sub LinkTables(arrTables As Variant)
'@author: Dustin Oakes, 2018
On Error GoTo errLinkTables

Dim strStartForm As String
'strStartForm = "the form that you want to open automatically"

MsgBox "Select the back-end database file."

Dim strPath As String
strPath = SelectFile

If "" = strPath Then
    MsgBox "No file selected."
    GoTo exitLinkTables
End If

CloseAll
DropLinks
DropTables arrTables

Dim strTable As Variant
For Each strTable In arrTables
    DoCmd.TransferDatabase acLink, "Microsoft Access", strPath, acTable, strTable, strTable
Next

'DoCmd.OpenForm strStartForm

exitLinkTables:
    Exit Sub

errLinkTables:
    Dim strErrMsg As String
    strErrMsg = "Error: " & Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
                "Please contact your administrator."
    MsgBox strErrMsg
    Resume exitLinkTables

End Sub

Public Sub RelinkTables()
'@author: Dustin Oakes
On Error GoTo errRelinkTables

Dim strPath As String

MsgBox "Select the back-end database file."

strPath = SelectFile
If "" < strPath Then
    
    Dim tdf As TableDef
    For Each tdf In CurrentDbC.TableDefs
        If tdf.Connect > "" Then
            tdf.Connect = ";DATABASE=" & strPath
            tdf.RefreshLink
        End If
    Next
    
Else
    MsgBox "No file selected."
End If

exitRelinkTables:
    Exit Sub
    
errRelinkTables:
    Dim strErrMsg As String
    strErrMsg = "Error: " & Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
                "Please contact your administrator."
    MsgBox strErrMsg
    Resume exitRelinkTables

End Sub

Public Sub ListIndexes(strTable As String)
'@author: Dustin Oakes
'@param: strTable = table whose indexes you want to list
'outputs all indexes associated with table, strTable to the immediate window
On Error GoTo errListIndexes

Dim tdf As TableDef
Dim idx As Index

Set tdf = CurrentDbC.TableDefs(strTable)
 
For Each idx In tdf.Indexes
    Debug.Print idx.NAME
Next
 
exitListIndexes:
    On Error Resume Next
    Set tdf = Nothing
    Exit Sub
 
errListIndexes:
    Dim strErrMsg As String
    strErrMsg = "Error: " & Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
                "Please contact your administrator."
    MsgBox strErrMsg
    Resume exitListIndexes
    
End Sub
Public Function GetRelations(strTable As String)
'@author: Dustin Oakes
'returns a 2D array containing object name and relationship name
'for all relationships involving strTable
'returns NULL if no relationships exist
On Error GoTo errGetRelations

Dim rs As DAO.Recordset
Dim arrRelations As Variant
Dim strSql As String
Dim i As Integer

strSql = "SELECT szObject, szRelationship " & _
         "FROM MsysRelationships " & _
         "WHERE szObject = """ & strTable & """ " & _
         "OR szReferencedObject = """ & strTable & """"
         
Set rs = CurrentDbC.OpenRecordset(strSql, 4)

If Not (rs.EOF And rs.BOF) Then
    ReDim arrRelations(rs.RecordCount, 2)
    rs.MoveFirst
    i = 0
    Do Until rs.EOF
        arrRelations(i, 0) = rs.Fields("szObject").Value
        arrRelations(i, 1) = rs.Fields("szRelationship").Value
        i = i + 1
        rs.MoveNext
    Loop
Else
    arrRelations = Null
End If

GetRelations = arrRelations

exitGetRelations:
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Exit Function
    
errGetRelations:
    Dim strErrMsg As String
    strErrMsg = "Error: " & Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
                "Please contact your administrator."
    MsgBox strErrMsg
    Resume exitGetRelations

End Function

Public Sub ListRelationships()
'@author: Dustin Oakes, 2018
On Error GoTo errListRelationships

Dim rel As DAO.Relation
Dim fld As Variant

For Each rel In CurrentDbC.Relations
    Debug.Print rel.NAME
    Debug.Print rel.ForeignTable
    Debug.Print rel.Table
    For Each fld In rel.Fields
        Debug.Print fld.NAME
    Next
Next

exitListRelationships:
    Exit Sub
    
errListRelationships:
    Dim strErrMsg As String
    strErrMsg = "Error: " & Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
                "Please contact your administrator."
    MsgBox strErrMsg
    Resume exitListRelationships

End Sub


