VERSION 5.00
Begin VB.Form moral_score_situation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "德育管家v1.0 - 德育分情况"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   6000
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox change_situation 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4680
      TabIndex        =   5
      Top             =   3930
      Width           =   915
   End
   Begin VB.TextBox score_situation 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   2340
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   840
      Width           =   3255
   End
   Begin VB.ComboBox date_show 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1463
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   150
      Width           =   4155
   End
   Begin VB.ListBox date_list 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3660
      ItemData        =   "moral_score_situation.frx":0000
      Left            =   383
      List            =   "moral_score_situation.frx":0002
      TabIndex        =   0
      Top             =   825
      Width           =   1755
   End
   Begin VB.Label Label2 
      Caption         =   "加扣分情况"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2340
      TabIndex        =   4
      Top             =   3960
      Width           =   2235
   End
   Begin VB.Label Label1 
      Caption         =   "日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   443
      TabIndex        =   1
      Top             =   195
      Width           =   855
   End
End
Attribute VB_Name = "moral_score_situation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public connect

Private Sub change_situation_Change()
    change_situation.Text = Val(change_situation.Text)
    sql = "UPDATE " & student_table & " SET change_value = " & change_situation.Text & _
        " WHERE index = " & (date_list.ListIndex + 1) & " And operate_date = #" & Date & "#"
    connect.Execute sql
End Sub

Private Sub date_list_Click()
    sql = "SELECT * FROM " & student_table & " WHERE index = " & (date_list.ListIndex + 1) & _
        " And operate_date = #" & date_show.Text & "#"
    db_record = connect.Execute(sql)
    score_situation.Text = Trim(db_record("explicit").Value)
    change_situation.Text = Val(db_record("change_value").Value)
End Sub

Private Sub date_show_Click()
    Dim date_index, date_count
    sql = "SELECT COUNT(*) AS date_count FROM " & student_table & " WHERE operate_date = #" & _
        date_show.Text & "#"
    date_count = connect.Execute(sql)
    date_list.Clear
    For date_index = 1 To Val(date_count("date_count").Value)
        sql = "SELECT * FROM " & student_table & " WHERE operate_date = #" & date_show.Text & "#" & _
            " And index = " & date_index
        db_record = connect.Execute(sql)
        date_list.AddItem "第" & db_record("index").Value & "条", date_index - 1
    Next
    date_list.ListIndex = 0
    sql = "SELECT * FROM " & student_table & " WHERE operate_date = #" & date_show.Text & "#" & _
        " And index = " & (date_list.ListIndex + 1)
    db_record = connect.Execute(sql)
    score_situation.Text = Trim(db_record("explicit").Value)
    If Val(db_record("change_value").Value) > 0 Then
        change_situation.Text = "+" & db_record("change_value").Value
    ElseIf Val(db_record("change_value").Value) < 0 Then
        change_situation.Text = "-" & db_record("change_value").Value
    End If
End Sub

Private Sub Form_Load()
    Set connect = New ADODB.Connection
    db_single_drive = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Chr(34) & _
        db_name & table_name & "德育分细则\" & student_table & ".mdb" & Chr(34)
    connect.Open db_single_drive
    
    Dim count
    sql = "SELECT COUNT(date_index) AS date_count FROM " & student_table
    count = connect.Execute(sql)
    Dim date_count
    date_count = Val(count("date_count").Value)
    If date_count > 0 Then
        Dim index_date, index_record
        sql = "SELECT Max(date_index) AS date_count FROM " & student_table
        count = connect.Execute(sql)
        date_count = count("date_count").Value
        For index_date = 1 To date_count
            sql = "SELECT operate_date FROM " & student_table & _
                " WHERE date_index = " & index_date
            db_record = connect.Execute(sql)
            date_show.AddItem db_record("operate_date").Value, index_date - 1
        Next
        date_show.ListIndex = 0
        date_list.Clear
        sql = "SELECT Max(index) AS record_count FROM " & student_table & _
            " WHERE operate_date = #" & date_show.Text & "#"
        db_record = connect.Execute(sql)
        date_count = Val(db_record("record_count").Value)
        For index_record = 1 To date_count
            sql = "SELECT * FROM " & student_table & " WHERE index = " & _
                index_record & " And operate_date = #" & date_show.Text & "#"
            db_record = connect.Execute(sql)
            score_situation.Text = Trim(db_record("explicit").Value)
            date_list.AddItem "第" & db_record("index").Value & "条", date_list.ListCount
        Next
        date_list.ListIndex = 0
        sql = "SELECT * FROM " & student_table & " WHERE index = " & _
            Val(Right(date_list.Text, Len(date_list.Text) - 1)) & _
            " And operate_date = #" & date_show.Text & "#"
        db_record = connect.Execute(sql)
        score_situation.Text = Trim(db_record("explicit").Value)
        date_list.ListIndex = 0
    End If
End Sub

Private Sub score_situation_Change()
    sql = "UPDATE " & student_table & " SET explicit = '" & score_situation.Text & "'" & _
        " WHERE index = " & (date_list.ListIndex + 1) & " And operate_date = #" & date_show.Text & "#"
    connect.Execute sql
End Sub
