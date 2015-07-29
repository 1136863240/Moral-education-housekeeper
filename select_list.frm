VERSION 5.00
Begin VB.Form select_list 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "德育管家v0.4 - 列表"
   ClientHeight    =   5355
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5355
   ScaleLeft       =   567
   ScaleMode       =   0  'User
   ScaleWidth      =   5205
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton user_define 
      Caption         =   "自定义"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3435
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
   Begin VB.ListBox students_list 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4260
      Left            =   195
      TabIndex        =   2
      Top             =   960
      Width           =   4815
   End
   Begin VB.ComboBox sort_function 
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
      ItemData        =   "select_list.frx":0000
      Left            =   1200
      List            =   "select_list.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   225
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "排序"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "select_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    sort_function.ListIndex = 0
    students_list.Clear
    db_class_drive = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Chr(34) & _
        db_name & table_name & ".mdb" & Chr(34)
    Dim select_conn As New ADODB.Connection
    Dim record As New ADODB.Recordset
    select_conn.Open db_class_drive
    sql = "SELECT COUNT(*) AS students_count FROM " & table_name
    record.Open sql, select_conn
    Dim students_count
    students_count = Val(record("students_count").Value)
    Dim students_index, item
    students_list.AddItem "   学号     姓名    德育分", 0
    If students_count = 0 Then
        students_list.AddItem "        未查询到任何记录", 1
        user_define.Enabled = False
        sort_function.Enabled = False
        Exit Sub
    End If
    sql = "SELECT id,name,moral_score FROM " & table_name
    If sort_function.ListIndex = 0 Then
        sql = sql & " Order By id"
    Else
        sql = sql & " Order By moral_score"
    End If
    record.Close
    record.Open sql, select_conn
    If record.BOF = False Then
        record.MoveFirst
    End If
    For students_index = 1 To students_count
        item = "    " & record("id").Value & "       " & Trim(record("name").Value) & _
            "       " & record("moral_score").Value
        students_list.AddItem item, students_index
        record.MoveNext
    Next
End Sub

Private Sub Form_Unload(cancel As Integer)
    class_base.Show
    Unload Me
End Sub

Private Sub sort_function_Click()
    Dim select_conn As New ADODB.Connection
    Dim record As New ADODB.Recordset
    Dim all_count, item
    students_list.Clear
    students_list.AddItem "   学号     姓名    德育分", 0
    select_conn.Open db_class_drive
    If record.State = 1 Then
        record.Close
    End If
    sql = "SELECT COUNT(*) AS all_count FROM " & table_name
    record.Open sql, select_conn
    all_count = Val(record("all_count").Value)
    record.Close
    If sort_function.ListIndex = 0 Then
        sql = "SELECT id,name,moral_score FROM " & table_name & " Order By id"
        record.Open sql, select_conn
        Dim id_index
        For id_index = 1 To all_count
            item = "    " & record("id").Value & "       " & Trim(record("name").Value) & _
                "       " & record("moral_score").Value
            students_list.AddItem item, id_index
            record.MoveNext
        Next
    Else
        sql = "SELECT id,name,moral_score FROM " & table_name & " Order By moral_score"
        record.Open sql, select_conn
        For id_index = 1 To all_count
            item = "    " & record("id").Value & "       " & Trim(record("name").Value) & _
                "       " & record("moral_score").Value
            students_list.AddItem item, id_index
            record.MoveNext
        Next
    End If
End Sub

Private Sub students_list_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If students_list.ListIndex = 0 Then
        students_list.ListIndex = 1
    End If
End Sub

Private Sub user_define_Click()
    Dim select_conn As New ADODB.Connection
    Dim record As New ADODB.Recordset
    select_conn.Open db_class_drive
    sort_index = sort_function.ListIndex
    user_define_data.Show vbModal, Me
    If is_select = False Then
        Exit Sub
    End If
    students_list.Clear
    students_list.AddItem "   学号     姓名    德育分", 0
    Dim item
    If select_id = True And _
        select_score = False Then
        sql = "SELECT COUNT(*) AS id_count FROM " & table_name & " WHERE id >= " & _
            range_id_least & " And id <= " & range_id_most
        record.Open sql, select_conn
        Dim id_count
        id_count = Val(record("id_count").Value)
        If Val(record("id_count").Value) = 0 Then
            students_list.AddItem "        未查询到任何记录", 1
            Exit Sub
        End If
        id_count = Val(record("id_count").Value)
        record.Close
        sql = "SELECT id,name,moral_score FROM " & table_name & " WHERE id >= " & _
            range_id_least & " And id <= " & range_id_most
        If sort_as_id = True Then
            sql = sql & " Order By id"
            record.Open sql, select_conn
            Dim count_index
            For count_index = 1 To id_count
                item = "    " & record("id").Value & "       " & Trim(record("name").Value) & _
                    "       " & record("moral_score").Value
                students_list.AddItem item, count_index
                record.MoveNext
            Next
        Else
            sql = sql & " Order By moral_score"
            record.Open sql, select_conn
            For count_index = 1 To id_count
                item = "    " & record("id").Value & "       " & Trim(record("name").Value) & _
                    "       " & record("moral_score").Value
                students_list.AddItem item, count_index
                record.MoveNext
            Next
        End If
    ElseIf select_id = True And _
        select_score = True Then
        sql = "SELECT COUNT(*) AS id_count FROM " & table_name & " WHERE id >= " & _
            range_id_least & " And id <= " & range_id_most & " And moral_score >= " & _
            range_score_least & " And moral_score <= " & range_score_most
        record.Open sql, select_conn
        If Val(record("id_count").Value) = 0 Then
            students_list.AddItem "        未查询到任何记录", 1
            Exit Sub
        End If
        id_count = Val(record("id_count").Value)
        record.Close
        sql = "SELECT id,name,moral_score FROM " & table_name & " WHERE id >= " & _
            range_id_least & " And id <= " & range_id_most & " And moral_score >= " & _
            range_score_least & " And moral_score <= " & range_score_most
        If sort_as_id = True Then
            sql = sql & " Order By id"
            record.Open sql, select_conn
            For count_index = 1 To id_count
                item = "    " & record("id").Value & "       " & Trim(record("name").Value) & _
                    "       " & record("moral_score").Value
                students_list.AddItem item, count_index
                record.MoveNext
            Next
        Else
            sql = sql & " Order By moral_score"
            record.Open sql, select_conn
            For count_index = 1 To id_count
                item = "    " & record("id").Value & "       " & Trim(record("name").Value) & _
                    "       " & record("moral_score").Value
                students_list.AddItem item, count_index
                record.MoveNext
            Next
        End If
    ElseIf select_id = False And _
        select_score = True Then
        sql = "SELECT COUNT(*) AS count_moral_score FROM " & table_name & " WHERE moral_score >= " & _
            range_score_least & " And moral_score <= " & range_score_most
        record.Open sql, select_conn
        If Val(record("count_moral_score").Value) = 0 Then
            students_list.AddItem "        未查询到任何记录", 1
            Exit Sub
        End If
        id_count = Val(record("count_moral_score").Value)
        record.Close
        sql = "SELECT id,name,moral_score FROM " & table_name & " WHERE moral_score >= " & _
            range_score_least & " And moral_score <= " & range_score_most
        If sort_as_id = True Then
            sql = sql & " Order By id"
            record.Open sql, select_conn
            For count_index = 1 To id_count
                item = "    " & record("id").Value & "       " & Trim(record("name").Value) & _
                    "       " & record("moral_score").Value
                students_list.AddItem item, count_index
                record.MoveNext
            Next
        Else
            sql = sql & " Order By moral_score"
            record.Open sql, select_conn
            For count_index = 1 To id_count
                item = "    " & record("id").Value & "       " & Trim(record("name").Value) & _
                    "       " & record("moral_score").Value
                students_list.AddItem item, count_index
                record.MoveNext
            Next
        End If
    Else
        If record.State = 1 Then
            record.Close
        End If
        sql = "SELECT COUNT(*) AS all_count FROM " & table_name
        record.Open sql, select_conn
        id_count = Val(record("all_count").Value)
        record.Close
        sql = "SELECT id,name,moral_score FROM " & table_name
        If sort_as_id = True Then
            sql = sql & " Order By id"
            record.Open sql, select_conn
            For count_index = 1 To id_count
                item = "    " & record("id").Value & "       " & Trim(record("name").Value) & _
                    "       " & record("moral_score").Value
                students_list.AddItem item, count_index
                record.MoveNext
            Next
            sort_function.ListIndex = 0
        Else
            sql = sql & " Order By moral_score"
            record.Open sql, select_conn
            For count_index = 1 To id_count
                item = "    " & record("id").Value & "       " & Trim(record("name").Value) & _
                    "       " & record("moral_score").Value
                students_list.AddItem item, count_index
                record.MoveNext
            Next
            sort_function.ListIndex = 1
        End If
    End If
End Sub
