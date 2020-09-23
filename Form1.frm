VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form1 
   Caption         =   "RDO database example"
   ClientHeight    =   6645
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   ScaleHeight     =   6645
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   5550
      Left            =   2760
      OleObjectBlob   =   "Form1.frx":0015
      TabIndex        =   4
      Top             =   0
      Width           =   7890
   End
   Begin VB.TextBox txtqry 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   6000
      Width           =   9915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Query"
      Height          =   315
      Left            =   9945
      TabIndex        =   2
      Top             =   6000
      Width           =   750
   End
   Begin VB.PictureBox picStatus 
      Align           =   2  'Align Bottom
      BackColor       =   &H00404040&
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   10635
      TabIndex        =   1
      Top             =   6360
      Width           =   10695
   End
   Begin MSRDC.MSRDC MSRDC1 
      Height          =   330
      Left            =   2760
      Negotiate       =   -1  'True
      Top             =   5580
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   582
      _Version        =   393216
      Options         =   0
      CursorDriver    =   1
      BOFAction       =   0
      EOFAction       =   0
      RecordsetType   =   1
      LockType        =   3
      QueryType       =   0
      Prompt          =   3
      Appearance      =   1
      QueryTimeout    =   30
      RowsetSize      =   100
      LoginTimeout    =   15
      KeysetSize      =   0
      MaxRows         =   0
      ErrorThreshold  =   -1
      BatchSize       =   15
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      ReadOnly        =   0   'False
      Appearance      =   -1  'True
      DataSourceName  =   ""
      RecordSource    =   ""
      UserName        =   ""
      Password        =   ""
      Connect         =   ""
      LogMessages     =   ""
      Caption         =   "MSRDC1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00E0E0E0&
      Height          =   5910
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2715
   End
   Begin VB.Menu mnuConnect 
      Caption         =   "&Connect"
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "Settings"
      Begin VB.Menu mnuAllowUpdates 
         Caption         =   "Allow Update"
      End
      Begin VB.Menu mnuAllowAddNew 
         Caption         =   "Allow Add New"
      End
      Begin VB.Menu mnuAllowDelete 
         Caption         =   "Allow Delete"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' define connection and recordset vars globaly
Dim cn As rdo.rdoConnection
Dim en As rdo.rdoEnvironment
Dim tb As rdo.rdoTable

Dim sortorder As String



Private Sub Command1_Click()
On Error GoTo bail

status "Executing query, please wait."
Set MSRDC1.Resultset = cn.OpenResultset(txtqry.Text, rdOpenDynamic, rdConcurRowVer, rdExecDirect)
While MSRDC1.Resultset.StillExecuting
    DoEvents
Wend

MSRDC1.Caption = MSRDC1.Resultset.AbsolutePosition & " of " & MSRDC1.Resultset.RowCount
status ""

Exit Sub
bail:
MsgBox Err.Description & vbCrLf
Err.Clear
End Sub



Private Sub DBGrid1_Click()
If Me.DBGrid1.SelStartCol >= 0 Then

If sortorder & "" = "Asc" Then
     sortorder = "Desc"
Else
    sortorder = "Asc"
End If

status "Sorting by " & DBGrid1.Columns(DBGrid1.SelStartCol).Caption & " " & sortorder

Set MSRDC1.Resultset = cn.OpenResultset(txtqry & " order by " & DBGrid1.Columns(DBGrid1.SelStartCol).Caption & " " & sortorder, rdOpenDynamic, rdConcurRowVer, rdExecDirect)


Do While MSRDC1.Resultset.StillExecuting
    DoEvents: DoEvents: DoEvents
Loop
status ""
End If
End Sub

Private Sub DBGrid1_Error(ByVal DataError As Integer, Response As Integer)
DBGrid1.ReBind
End Sub

Private Sub Form_Resize()
DBGrid1.Width = Form1.ScaleWidth - DBGrid1.Left - 5
Command1.Left = Form1.ScaleWidth - Command1.Width - 5
Command1.Top = Form1.ScaleHeight - 650
txtqry.Width = Command1.Left - 5
txtqry.Top = Command1.Top
List1.Height = Command1.Top - 10
DBGrid1.Height = List1.Height - 375
MSRDC1.Top = List1.Top + List1.Height - MSRDC1.Height
End Sub

Private Sub List1_Click()

txtqry = "select * from " & List1.Text

status "Opening Recordset"
On Error Resume Next
Set MSRDC1.Resultset = cn.OpenResultset("select * from " & List1.Text, rdOpenDynamic, rdConcurRowVer, rdExecDirect)


Do While MSRDC1.Resultset.StillExecuting
    DoEvents: DoEvents: DoEvents
Loop
status ""

MSRDC1.Caption = MSRDC1.Resultset.AbsolutePosition & " of " & MSRDC1.Resultset.RowCount

End Sub


Private Sub status(statusmsg As String)
picStatus.Cls
picStatus.Print statusmsg
End Sub



Private Sub mnuAllowAddNew_Click()
If DBGrid1.AllowAddNew = True Then
    DBGrid1.AllowAddNew = False
    mnuAllowAddNew.Checked = False
Else
    DBGrid1.AllowAddNew = True
    mnuAllowAddNew.Checked = True
End If
End Sub

Private Sub mnuAllowDelete_Click()
If DBGrid1.AllowDelete = True Then
    DBGrid1.AllowDelete = False
    mnuAllowDelete.Checked = False
Else
    DBGrid1.AllowDelete = True
    mnuAllowDelete.Checked = True
End If
End Sub

Private Sub mnuAllowUpdates_Click()
If DBGrid1.AllowUpdate = True Then
    DBGrid1.AllowUpdate = False
    mnuAllowUpdates.Checked = False
Else
    DBGrid1.AllowUpdate = True
    mnuAllowUpdates.Checked = True
End If
End Sub

Private Sub mnuConnect_Click()
On Error GoTo bail

Set MSRDC1.Resultset = Nothing

Set en = rdoEnvironments(0)
Set cn = en.OpenConnection(dsName:="WorkDB", Prompt:=rdDriverPrompt)
List1.Clear
status "Retrieving tables"
Dim tb As rdo.rdoTable
For Each tb In cn.rdoTables
    List1.AddItem tb.Name
Next
status ""

Exit Sub
bail:
MsgBox Err.Description
Err.Clear
End Sub



Private Sub MSRDC1_Reposition()
On Error Resume Next
MSRDC1.Caption = MSRDC1.Resultset.AbsolutePosition & " of " & MSRDC1.Resultset.RowCount
End Sub

