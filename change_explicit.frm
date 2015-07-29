VERSION 5.00
Begin VB.Form change_explicit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "德育管家v0.3 - 加扣分细则"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4860
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   4860
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cancel 
      Caption         =   "关闭"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1763
      TabIndex        =   2
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox explicit 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   113
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   780
      Width           =   4635
   End
   Begin VB.Label Label1 
      Caption         =   "加扣分细则"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1403
      TabIndex        =   1
      Top             =   180
      Width           =   2055
   End
End
Attribute VB_Name = "change_explicit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public conn

Private Sub cancel_Click()
    If explicit.Text = "" Then
        sql = "DELETE FROM " & student_table & " WHERE explicit = ''"
        conn.Execute sql
        sub_score = False
    Else
        sub_score = True
    End If
    Unload Me
End Sub

Private Sub explicit_Change()
    sql = "SELECT COUNT(*) AS date_count FROM " & student_table & " WHERE " & _
        "operate_date = #" & Date & "#"
    db_record = conn.Execute(sql)
    sql = "UPDATE " & student_table & " SET explicit = '" & explicit.Text & "'" & _
        " WHERE index = " & db_record("date_count").Value & " And operate_date = #" & Date & "#"
    conn.Execute sql
End Sub

Private Sub Form_Load()
    Set conn = New ADODB.Connection
    db_drive = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Chr(34) & _
        db_name & table_name & "德育分细则\" & student_table & ".mdb" & Chr(34)
    conn.Open db_drive
    
    sql = "SELECT COUNT(operate_date) AS body_count FROM " & student_table
    db_count = conn.Execute(sql)
    Dim date_index, index
    If Val(db_count("body_count").Value) > 0 Then
        'select max date_index
        sql = "SELECT Max(date_index) AS date_index FROM " & student_table
        db_record = conn.Execute(sql)
        date_index = Val(db_record("date_index").Value)
        sql = "SELECT operate_date FROM " & student_table & " WHERE date_index = " & date_index
        db_record = conn.Execute(sql)
        If db_record("operate_date").Value < Date Then
            date_index = date_index + 1
        End If
        'select max index
        sql = "SELECT COUNT(index) AS record_index FROM " & student_table & _
            " WHERE date_index = " & date_index
        db_record = conn.Execute(sql)
        index = Val(db_record("record_index").Value)
        If index = 0 Then
            sql = "INSERT INTO " & student_table & " VALUES(1," & _
                date_index & ",'" & Date & "'," & change_score & ",'')"
            conn.Execute sql
            Exit Sub
        End If
        sql = "INSERT INTO " & student_table & " VALUES(" & (index + 1) & "," & _
            date_index & ",'" & Date & "'," & change_score & ",'')"
        conn.Execute sql
    Else
        sql = "SELECT COUNT(date_index) AS date_index FROM " & _
            student_table & " WHERE operate_date = #" & Date & "#"
        db_count = conn.Execute(sql)
        sql = "INSERT INTO " & student_table & " VALUES(1," & (db_count("date_index").Value + 1) & _
            ",'" & Date & "'," & change_score & ",'')"
        conn.Execute sql
    End If
End Sub

