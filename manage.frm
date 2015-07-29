VERSION 5.00
Begin VB.Form manage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "德育管家v0.3 - 管理"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5790
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   5790
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton moral_score_check 
      Caption         =   "德育分情况"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3159
      TabIndex        =   17
      Top             =   1876
      Width           =   1665
   End
   Begin VB.CommandButton change_moral_score 
      Caption         =   "确定修改"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3780
      TabIndex        =   13
      Top             =   2551
      Width           =   1485
   End
   Begin VB.CommandButton up 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3159
      TabIndex        =   12
      Top             =   2626
      Width           =   420
   End
   Begin VB.TextBox change_value 
      Height          =   390
      Left            =   2316
      TabIndex        =   11
      Top             =   2611
      Width           =   645
   End
   Begin VB.CommandButton down 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1698
      TabIndex        =   10
      Top             =   2626
      Width           =   420
   End
   Begin VB.TextBox moral_score 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1620
      TabIndex        =   9
      Top             =   1936
      Width           =   1260
   End
   Begin VB.TextBox strName 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3705
      TabIndex        =   7
      Top             =   1246
      Width           =   1260
   End
   Begin VB.TextBox id 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1320
      TabIndex        =   5
      Top             =   1246
      Width           =   1260
   End
   Begin VB.CommandButton btn_next 
      Caption         =   "下一个"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3818
      TabIndex        =   3
      Top             =   240
      Width           =   1290
   End
   Begin VB.TextBox page 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2198
      TabIndex        =   1
      Top             =   255
      Width           =   540
   End
   Begin VB.CommandButton btn_previous 
      Caption         =   "上一个"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   803
      TabIndex        =   0
      Top             =   240
      Width           =   1290
   End
   Begin VB.Label Label1 
      Caption         =   "加扣分"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   525
      TabIndex        =   16
      Top             =   2641
      Width           =   990
   End
   Begin VB.Label grade 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1620
      TabIndex        =   15
      Top             =   3391
      Width           =   1035
   End
   Begin VB.Label Label5 
      Caption         =   "等级"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   525
      TabIndex        =   14
      Top             =   3376
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "德育分"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   525
      TabIndex        =   8
      Top             =   1966
      Width           =   990
   End
   Begin VB.Label Label3 
      Caption         =   "姓名"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2910
      TabIndex        =   6
      Top             =   1291
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "学号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   525
      TabIndex        =   4
      Top             =   1291
      Width           =   615
   End
   Begin VB.Label for_count 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2850
      TabIndex        =   2
      Top             =   330
      Width           =   765
   End
End
Attribute VB_Name = "manage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public isSave

Private Sub btn_next_Click()
    'check is changed
    If Val(page.Text) = 0 Then
        id.Enabled = True
        strName.Enabled = True
        moral_score.Enabled = True
        change_value.Enabled = True
        grade.Enabled = True
        btn_next.Caption = "添加"
        down.Enabled = True
        up.Enabled = True
        change_moral_score.Enabled = True
        moral_score_check.Enabled = True
        sql = "INSERT INTO " & table_name & " VALUES(" & (page.Text + 1) & "," & (page.Text + 1) & _
            ",'',75,'" & db_name & table_name & "德育分细则\" & strName.Text & ".mdb" & "')"
        db_conn.Execute sql
    ElseIf Val(page.Text) = Val(Mid(for_count.Caption, 3)) Then
        If id.Text = "" Or _
            Trim(strName.Text) = "" Or _
            moral_score.Text = "" Then
            MsgBox "学号、姓名或德育分不能为空", vbOKOnly + vbExclamation, "提示"
            Exit Sub
        End If
        sql = "INSERT INTO " & table_name & " VALUES(" & (page.Text + 1) & "," & (page.Text + 1) & _
            ",'',75,'" & db_name & table_name & "德育分细则\" & strName.Text & ".mdb" & "')"
        db_conn.Execute sql
    ElseIf Val(page.Text) = Val(Mid(for_count.Caption, 3)) - 1 Then
        If id.Text = "" Or _
            Trim(strName.Text) = "" Or _
            moral_score.Text = "" Then
            MsgBox "学号、姓名或德育分不能为空", vbOKOnly + vbExclamation, "提示"
            Exit Sub
        End If
        btn_previous.Enabled = True
        btn_next.Caption = "添加"
        sql = "SELECT COUNT(*) AS row FROM " & table_name
        db_count = db_conn.Execute(sql)
    End If
    page.Text = Val(page.Text) + 1
    sql = "SELECT COUNT(*) AS row FROM " & table_name
    db_count = db_conn.Execute(sql)
    for_count.Caption = "/ " & db_count("row").Value
    sql = "SELECT * FROM " & table_name & " WHERE index = " & page.Text
    db_record = db_conn.Execute(sql)
    id.Text = db_record("id").Value
    strName.Text = Trim(db_record("name").Value)
    moral_score.Text = db_record("moral_score").Value
    change_value.Text = "0"
    If Val(page.Text) > 0 Then
        btn_previous.Enabled = True
    End If
End Sub

Private Sub btn_previous_Click()
    page.Text = Val(page.Text) - 1
    sql = "SELECT * FROM " & table_name & " WHERE index = " & page.Text
    If Val(page.Text) = 1 Then
        btn_previous.Enabled = False
    End If
    db_record = db_conn.Execute(sql)
    btn_next.Caption = "下一条"
    id.Text = db_record("id").Value
    strName.Text = Trim(db_record("name").Value)
    moral_score.Text = db_record("moral_score").Value
    change_value.Text = "0"
End Sub

Private Sub change_moral_score_Click()
    If Val(change_value.Text) = 0 Then
        MsgBox "修改值为零，德育分未修改", vbOKOnly + vbExclamation, "提示"
    Else
        If strName.Text = "" Then
            MsgBox "姓名不能为空！", vbOKOnly + vbExclamation, "提示"
            Exit Sub
        End If
        change_score = change_value.Text
        student_table = strName.Text
        If Dir(db_name & table_name & "德育分细则\" & strName.Text & ".mdb") = "" Then
            If Dir(db_name & table_name & "德育分细则", vbDirectory) = "" Then
                MkDir db_name & table_name & "德育分细则"
            End If
            db_drive = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Chr(34) & _
                db_name & table_name & "德育分细则\" & strName.Text & ".mdb" & Chr(34)
            Set catalog = Nothing
            Set db_table = Nothing
            Set catalog = New adox.catalog
            Set db_table = New adox.Table
            catalog.Create db_drive
            catalog.ActiveConnection = db_drive
            db_table.Name = strName.Text
            db_table.Columns.Append "index", adox.DataTypeEnum.adInteger
            db_table.Columns.Append "date_index", adox.DataTypeEnum.adInteger
            db_table.Columns.Append "operate_date", adox.DataTypeEnum.adDate
            db_table.Columns.Append "change_value", adox.DataTypeEnum.adInteger
            db_table.Columns.Append "explicit", adox.DataTypeEnum.adWChar
            catalog.Tables.Append db_table
        End If
        Set db_conn_explicit = New ADODB.Connection
        db_drive = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Chr(34) & _
            db_name & table_name & "德育分细则\" & student_table & ".mdb" & Chr(34)
        db_conn_explicit.Open db_drive
        sql = "SELECT COUNT(*) AS explicit_count FROM " & student_table & _
            " WHERE operate_date = #" & Date & "#"
        db_record = db_conn_explicit.Execute(sql)
        If Val(db_record("explicit_count").Value) > 0 Then
            db_index = db_record("explicit_count").Value
        Else
            db_index = 0
        End If
        change_score = change_value.Text
        change_explicit.Show vbModal, Me
        If sub_score = True Then
            moral_score.Text = Val(moral_score.Text) + Val(change_value.Text)
        End If
        change_value.Text = "0"
    End If
End Sub

Private Sub change_value_Change()
    If change_value.Text = "" Then
        change_value.Text = "0"
    Else
        change_value.Text = Val(change_value.Text)
    End If
End Sub

Private Sub down_Click()
    change_value.Text = Val(change_value.Text) - 1
End Sub

Private Sub Form_Unload(cancel As Integer)
    If Val(page.Text) > 0 Then
        sql = "SELECT * FROM " & table_name & " WHERE index = " & Mid(for_count.Caption, 3)
        db_record = db_conn.Execute(sql)
        If db_record("id").Value = "" Or _
            Trim(db_record("name").Value) = "" Or _
            db_record("moral_score").Value = "" Then
            sql = "DELETE FROM " & table_name & " WHERE index = " & Mid(for_count.Caption, 3)
            db_conn.Execute sql
            db_conn.Close
            End
        End If
        If id.Text = "" Or _
            strName.Text = "" Or _
            moral_score.Text = "" Then
            sql = "DELETE FROM " & table_name & " WHERE index = " & page.Text
            db_conn.Execute sql
            db_conn.Close
            End
        End If
    Else
        db_conn.Close
        End
    End If
    db_conn.Close
    End
End Sub

Private Sub Form_Load()
    Me.Caption = "德育分管理器v0.3 - " & table_name & "管理"
    On Error GoTo OperatorError
    Set db_conn = New ADODB.Connection
    db_conn.Open db_drive, "Admin"
    Set db_record = New ADODB.Recordset
    Set db_count = New ADODB.Recordset
    sql = "SELECT COUNT(*) AS row FROM " & table_name
    db_count = db_conn.Execute(sql)
    for_count.Caption = "/ " & db_count("row").Value
    If Val(db_count("row").Value) > 0 Then
        page.Text = "1"
        If page.Text = Mid(for_count.Caption, 3) Then
            btn_next.Caption = "添加"
        ElseIf Val(page.Text) > Val(Mid(for_count.Caption, 3)) Then
            btn_next.Caption = "下一条"
        End If
        sql = "SELECT * FROM " & table_name & " WHERE index = " & page.Text
        db_record = db_conn.Execute(sql)
        id.Text = db_record("id").Value
        strName.Text = Trim(db_record("name").Value)
        moral_score.Text = db_record("moral_score").Value
        change_value.Text = "0"
        btn_previous.Enabled = False
    Else
        page.Text = "0"
        id.Enabled = False
        strName.Enabled = False
        moral_score.Enabled = False
        change_value.Enabled = False
        grade.Enabled = False
        btn_previous.Enabled = False
        btn_next.Caption = "添加"
        down.Enabled = False
        up.Enabled = False
        change_moral_score.Enabled = False
        moral_score_check.Enabled = False
    End If
    Exit Sub
OperatorError:
    MsgBox "操作出错，错误信息：" & Err.Description, vbOKOnly + vbExclamation, "提示"
    End
End Sub

Private Sub id_Change()
    If id.Text = "" Then
        id.Text = "1"
    Else
        sql = "UPDATE " & table_name & " SET id = " & id.Text & " WHERE index = " & page.Text
        db_conn.Execute sql
    End If
End Sub

Private Sub moral_score_Change()
    moral_score.Text = Val(moral_score.Text)
    If Val(moral_score.Text) < 60 Then
        grade.Caption = "不及格"
    ElseIf Val(moral_score.Text) >= 60 And _
        Val(moral_score.Text) < 75 Then
        grade.Caption = "及格"
    ElseIf Val(moral_score.Text) >= 75 And _
        Val(moral_score.Text) < 85 Then
        grade.Caption = "良好"
    Else
        grade.Caption = "优秀"
    End If
End Sub

Private Sub moral_score_check_Click()
    student_table = strName.Text
    If Dir(db_name & table_name & "德育分细则\" & student_table & ".mdb") = "" Then
        MsgBox "无任何修改相关记录", vbOKOnly + vbExclamation, "提示"
        Exit Sub
    End If
    Dim connect_manager, drive
    Set connect_manager = New ADODB.Connection
    drive = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Chr(34) & _
        db_name & table_name & "德育分细则\" & student_table & ".mdb" & Chr(34)
    connect_manager.Open drive
    
    sql = "SELECT COUNT(*) AS body_count FROM " & student_table
    db_record = connect_manager.Execute(sql)
    If Val(db_record("body_count").Value) = 0 Then
        MsgBox "无任何修改相关记录", vbOKOnly + vbExclamation, "提示"
        Exit Sub
    End If
    If strName.Text = "" Then
        MsgBox "姓名不能为空", vbOKOnly + vbExclamation, "提示"
    Else
        moral_score_situation.Show vbModal, Me
    End If
End Sub

Private Sub page_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(page.Text) > Val(Mid(for_count.Caption, 3)) And _
            Val(Mid(for_count.Caption, 3)) > 0 Then
            page.Text = Mid(for_count.Caption, 3)
        ElseIf Val(page.Text) = 0 Then
            Exit Sub
        End If
        sql = "SELECT * FROM " & table_name & " WHERE index = " & page.Text
        db_record = db_conn.Execute(sql)
        id.Text = db_record("id").Value
        strName.Text = Trim(db_record("name").Value)
        moral_score.Text = db_record("moral_score").Value
        change_value.Text = "0"
    End If
End Sub

Private Sub strName_Change()
    sql = "UPDATE " & table_name & " SET name = '" & strName.Text & _
        "',explicit = '" & db_name & table_name & "德育分细则\" & _
        strName.Text & ".mdb" & "' WHERE index = " & page.Text
    db_conn.Execute sql
End Sub

Private Sub up_Click()
    change_value.Text = Val(change_value.Text) + 1
End Sub
