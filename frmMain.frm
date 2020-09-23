VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   Caption         =   "SQL Statement Tester"
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11250
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9165
   ScaleWidth      =   11250
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid ocxResultsGrid 
      Height          =   2895
      Left            =   360
      TabIndex        =   8
      Top             =   5640
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   5106
      _Version        =   393216
      AllowUserResizing=   3
   End
   Begin VB.ListBox lstAvailableTables 
      Appearance      =   0  'Flat
      Height          =   3150
      Left            =   6360
      TabIndex        =   7
      Top             =   1080
      Width           =   4215
   End
   Begin MSComDlg.CommonDialog ocxCommDialog 
      Left            =   840
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "&Execute"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox txtSQLStatement 
      Height          =   3255
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   5895
   End
   Begin VB.CommandButton cmdBrowseSources 
      Caption         =   "..."
      Height          =   255
      Left            =   10320
      TabIndex        =   1
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox txtDataSource 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   9735
   End
   Begin VB.Label lblAvailableTables 
      BackStyle       =   0  'Transparent
      Caption         =   "Tables:"
      Height          =   255
      Left            =   6360
      TabIndex        =   9
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label lblFoundCount 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4560
      TabIndex        =   6
      Top             =   7560
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SQL Statement:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label lblDataSource 
      Caption         =   "Data Source:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowseSources_Click()
'   Any error we get on this particular sub can safely be ignored
On Error Resume Next

'   Open the Windows Browse Dialog using the ocx control painted on the form
'   ocxCommDialog is the Windows Common Dialog control on the form

ocxCommDialog.DialogTitle = "Open Datasource"
ocxCommDialog.Filter = "Microsoft Access|*.mdb"
ocxCommDialog.ShowOpen


'   Retrieve the value from the dialog and put it in the text box
txtDataSource = ocxCommDialog.FileName

'   Using the data source file, fill in the list of available tables
PopulateTableList

End Sub

Private Sub cmdExecute_Click()

'   Run the SQL statement
            '   If it is a select statement, run the select sub...
If UCase(Trim(Left(txtSQLStatement, 6))) = "SELECT" Then
    SQLSelect
Else
'           Otherwise run the Execute Sub
    SQLExecute
End If

End Sub
Sub PopulateTableList()

    Dim dbDataSource As Database
    Dim objTableDef As TableDef
    
    '   Clear out any old entries that may still be in the list
    lstAvailableTables.Clear
    
    '   Open the database file
    Set dbDataSource = OpenDatabase(txtDataSource)
    
    '   Remember...the Database is an object, and it has objects within it.
    '   One of these objects is the TableDef...A TableDef defines a single table.
    '   There is also a collections object of all TableDefs. Since we defined objTableDef
    '   as an object, we can reference it. In doing so, we can get the name of
    '   each table by looping through ALL tables and adding the .name property to the
    '   listbox. In each pass thorugh, the object objTableDef is reset to the next
    '   Table in the TableDefs collection of the source database.
    
    For Each objTableDef In dbDataSource.TableDefs
        lstAvailableTables.AddItem objTableDef.Name
    Next objTableDef
    
End Sub

Sub SQLSelect()

Dim dbDataSource As Database
Dim rsQuery As Recordset

On Error GoTo ERR_SQLSelect

'   Open the database that was entered as the source
Set dbDataSource = OpenDatabase(txtDataSource)

'   Run the entered SQL statement against the data source
'   Since we want to get a recordset back to populate the grid with, use the
'   OpenRecordset method of the database object

Set rsQuery = dbDataSource.OpenRecordset(txtSQLStatement)

'   If any results were returned, populate the data grid
If rsQuery.RecordCount > 0 Then

'   This is sort of tricky...
'   We want to populate the grid with the contents of our newly created recordset, but
'   we also want to do this from other places in the application. Since we already have
'   the recordset populated with lots of data, it would be a waste to do it over again
'   from within the PopulateGrid sub, but we don't want any public (global) recordsets
'   in our application. The solution? Pass the entire recordset as a parameter to the
'   sub. If you look at the sub definition, you will see "PopulateGrid(rsSource As Recordset)"
'   The rsSource will give us our recordset object to pass this data with.

   PopulateGrid rsQuery
Else
    MsgBox "No matching records found."

End If


EXIT_SQLSelect:
Exit Sub


'   This code will only run if there is an error
ERR_SQLSelect:

MsgBox Err.Number & " - " & Err.Description
Resume EXIT_SQLSelect
Resume Next
End Sub

Sub SQLExecute()
Dim dbDataSource As Database
'   This sub will run an action query on the database

'   First, set our reference to our database (i.e. create the database object based on the path)
Set dbDataSource = OpenDatabase(txtDataSource.Text)

'   We could get any number of database related errors on the next step, none of which we can do
'   anything about..
'   This is a method of error handling that lets the code continue, but we immediately check to
'   see if it errored out when it completes the risky action.

On Error Resume Next
'   Run our SQL statement
dbDataSource.Execute (txtSQLStatement.Text)
'   Hopefully it went OK, but if not, show the error.
If Err Then
    MsgBox Err.Description

'   Now that we told what the error is, clear it.
    Err.Clear
Else
    MsgBox "SQL Statement completed successfully. " & dbDataSource.RecordsAffected & " records were modified.", vbInformation
End If



End Sub


Sub PopulateGrid(rsSource As Recordset)
Dim SourceField As Field
Dim intColCount As Long
Dim intRowCount As Long
Dim intFieldCount As Long

'   This is a big one that you may have to chew on for a while...
'   Since I steadfastly refuse to use the data control (yuck!), we are forced to populate this
'   flexgrid manually. This is better because you can control each cell individually. It calls
'   for a lot of code, but it is mostly looping and collections.

'NOTE:  Remember, we already have our data in the rsSource parameter being passed in.

On Error GoTo ERR_PopulateGrid

'   First, populate the Recordset so we can get a valid record count
rsSource.MoveLast
rsSource.MoveFirst
'   Set the # of rows in the grid equal to the number of records + 1
ocxResultsGrid.Rows = rsSource.RecordCount + 1

'   Set the # of Columns in the grid equal to the number of fields + 1
ocxResultsGrid.Cols = rsSource.Fields.Count + 1
        'This process will take a long time, so turn the hourglass pointer on before
        'starting and disble the grid's repaint method (much faster)

MousePointer = vbHourglass

'   By default, Redraw is enabled. That means that it will repaint the screen EACH TIME
'   a cell is populated. On a 100 row x 10 field table, that is 1000 screen refreshes . Who
'   has time for that? So we will turn the Redraw property off until we are done.

ocxResultsGrid.Redraw = False

'       Process each record

'   Set the initial row and column
intRowCount = 1
intColCount = 1

'   Show how many matches we have
lblFoundCount.Caption = rsSource.RecordCount & " matches found."

'First list the field names in the header row
ocxResultsGrid.Row = 0
 For Each SourceField In rsSource.Fields
        '   Set the column
        ocxResultsGrid.Col = intColCount
        '   Since these are headers, make these cells BOLD
        ocxResultsGrid.CellFontBold = True
        
        '   Populate this cell with the name of the field. Notice that we can access
        '   each field in the collection (recordset.fields(number)) without knowing
        '   its name. This is done just like an array ex: rsSource.Fields(2).Name
        '   Also, note the & "" on the end. This prevents a null value fault if the
        '   field is null in the table
        
        ocxResultsGrid.Text = rsSource.Fields(intFieldCount).Name & ""
                
        
        '   Increment the column
        intColCount = intColCount + 1
        intFieldCount = intFieldCount + 1
        '   Get the next field
    Next SourceField

intColCount = 1
intFieldCount = 0
While Not rsSource.EOF
            ' Add data for each field for this row
    ocxResultsGrid.Row = intRowCount
    For Each SourceField In rsSource.Fields
        '   Set the column
        ocxResultsGrid.Col = intColCount
        '   Set the value of the column
        ocxResultsGrid.Text = rsSource.Fields(intFieldCount).Value & ""
        '   Increment the column
        intColCount = intColCount + 1
        intFieldCount = intFieldCount + 1
    Next SourceField
        '   Reset the column count
        intFieldCount = 0
        intColCount = 1
        '   Increment the row count
    intRowCount = intRowCount + 1
    '   Move to the next row in the recordset
    rsSource.MoveNext
    '   Get the next row.
Wend

'   MADE IT!
'   This is the standard way to exit a sub. All exits should be made by passing
'   through this label. See the error handling below for an example.
EXIT_PopulateGrid:

'       Set mouse pointer and redraw back to normal
ocxResultsGrid.Redraw = True
MousePointer = vbNormal

Exit Sub


'   Because of the Exit Sub line above, this will only run if there is an error.

ERR_PopulateGrid:

MsgBox Err.Number & " - " & Err.Description
Resume EXIT_PopulateGrid
Resume Next


End Sub


Private Sub Form_Resize()

'   Do not attempt to resize if window is minimized

If Me.WindowState <> vbMinimized Then
    ' Force form to stay a certain width to prevent errors by resizing too small
    If Me.Width > 3000 Then
    '   Set sizes and locations of controls
        
        txtDataSource.Width = Me.Width * 0.85
        cmdBrowseSources.Left = txtDataSource.Left + txtDataSource.Width + 50
        
        txtSQLStatement.Width = Me.Width / 2
        lstAvailableTables.Left = txtSQLStatement.Left + txtSQLStatement.Width + 400
        lstAvailableTables.Width = Me.Width - lstAvailableTables.Left - 500
        ocxResultsGrid.Width = Me.Width - 700
        ocxResultsGrid.Height = Me.Height - (ocxResultsGrid.Top + 800)
        
        cmdExecute.Left = (Me.Width / 2) - (cmdExecute.Width / 2)
        
        lblFoundCount.Left = (Me.Width / 2) - (lblFoundCount.Width / 2)
        lblFoundCount.Top = ocxResultsGrid.Top + ocxResultsGrid.Height + 150
        lblAvailableTables.Left = lstAvailableTables.Left
    End If
End If

End Sub


Private Sub lstAvailableTables_Click()
On Error GoTo ERR_lstAvailableTables_Click
Dim dbSelectedTable As Database
Dim rsSelectedTable As Recordset

'   This sub will open the selected database and pass the recordset to
'   the PopulateGrid sub.

Set dbSelectedTable = OpenDatabase(txtDataSource)
Set rsSelectedTable = dbSelectedTable.OpenRecordset(lstAvailableTables.Text)
PopulateGrid rsSelectedTable

EXIT_lstAvailableTables_Click:

Exit Sub


ERR_lstAvailableTables_Click:

MsgBox Error
Resume EXIT_lstAvailableTables_Click

End Sub

Private Sub ocxResultsGrid_DblClick()
'   Just a little deal to show the entire contents of a cell and allow you to copy it to the
'   clipboard using <CTL><INSERT> or <CTL> <C>
    InputBox ocxResultsGrid.TextMatrix(0, ocxResultsGrid.Col) & ":" & vbCr & vbCr & vbCr & "(Use <ctl> + <c> to copy to clipboard)", "Zoom", ocxResultsGrid.Text
End Sub

