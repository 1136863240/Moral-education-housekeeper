VERSION 5.00
Begin VB.Form manage 
   Caption         =   "德育分管理器v0.1 - 管理"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5790
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   5790
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton change_moral_score 
      Caption         =   "确定修改"
      Height          =   495
      Left            =   4665
      TabIndex        =   15
      Top             =   1560
      Width           =   885
   End
   Begin VB.CommandButton up 
      Caption         =   "+"
      Height          =   345
      Left            =   4155
      TabIndex        =   14
      Top             =   1635
      Width           =   420
   End
   Begin VB.TextBox change_value 
      Height          =   390
      Left            =   3495
      TabIndex        =   13
      Top             =   1605
      Width           =   525
   End
   Begin VB.CommandButton down 
      Caption         =   "-"
      Height          =   345
      Left            =   2955
      TabIndex        =   12
      Top             =   1635
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
      Left            =   1560
      TabIndex        =   11
      Top             =   1605
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
      Left            =   3645
      TabIndex        =   9
      Top             =   945
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
      Left            =   1260
      TabIndex        =   7
      Top             =   945
      Width           =   1260
   End
   Begin VB.CommandButton btn_return 
      Caption         =   "还原"
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
      Left            =   3165
      TabIndex        =   5
      Top             =   2925
      Width           =   1290
   End
   Begin VB.CommandButton btn_save 
      Caption         =   "保存"
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
      Left            =   1230
      TabIndex        =   4
      Top             =   2925
      Width           =   1290
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
      Left            =   3758
      TabIndex        =   3
      Top             =   180
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
      Height          =   450
      Left            =   2138
      TabIndex        =   1
      Top             =   195
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
      Left            =   743
      TabIndex        =   0
      Top             =   180
      Width           =   1290
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
      Left            =   1560
      TabIndex        =   17
      Top             =   2295
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
      Height          =   375
      Left            =   465
      TabIndex        =   16
      Top             =   2295
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
      Height          =   375
      Left            =   465
      TabIndex        =   10
      Top             =   1650
      Width           =   975
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
      Height          =   375
      Left            =   2850
      TabIndex        =   8
      Top             =   990
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
      Height          =   375
      Left            =   465
      TabIndex        =   6
      Top             =   990
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
      Left            =   2783
      TabIndex        =   2
      Top             =   270
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
    If Val(Page.Text) = 0 Then
        sql = "INSERT INTO " & table_name & " VALUES(" & (db_count("row").Value + 1) & "," & _
            (db_count("row").Value + 1) & ",'',75)"
        db_conn.Execute sql
        id.Enabled = True
        strName.Enabled = True
        moral_score.Enabled = True
        change_value.Enabled = True
        grade.Enabled = True
        btn_next.Caption = "添加"
        btn_save.Enabled = True
        btn_return.Enabled = True
        down.Enabled = True
        up.Enabled = True
        change_value.Text = "0"
        Page.Text = "1"
        for_count.Caption = "/ 1"
        sql = "SELECT * FROM " & table_name & " WHERE index = " & Page.Text
        db_record = db_conn.Execute(sql)
        id.Text = db_record("id").Value
        strName.Text = Trim(db_record("name").Value)
        moral_score.Text = db_record("moral_score").Value
        change_value.Text = "0"
    ElseIf Val(Page.Text) = Val(Mid(for_count.Caption, 3)) Then
        sql = "SELECT * FROM " & table_name & " WHERE index = " & Page.Text
        db_record = db_conn.Execute(sql)
        If id.Text <> db_record("id").Value Or _
            strName.Text <> Trim(db_record("name").Value) Or _
            moral_score.Text <> db_record("moral_score").Value Then
            isSave = MsgBox("记录已修改，是否保存记录？" & vbCrLf & Chr(34) & "是" & Chr(34) & _
                "为保存并向下一条记录，" & Chr(34) & "否" & Chr(34) & "为还原并" & _
                "向下一条记录", vbYesNo + vbExclamation, "温馨提示")
            If isSave = vbYes Then
                sql = "UPDATE " & table_name & " SET id = " & id.Text & ", name = '" & strName.Text & _
                    "', moral_score = " & moral_score.Text & " WHERE index = " & Page.Text
                db_conn.Execute (sql)
            Else
                id.Text = db_record("id").Value
                strName.Text = Trim(db_record("name").Value)
                moral_score.Text = db_record("moral_score").Value
                change_value.Text = "0"
                If id.Text = "" Or _
                    Trim(strName.Text) = "" Or _
                    moral_score.Text = "" Then
                    MsgBox "学号、姓名或德育分不能为空", vbOKOnly + vbExclamation, "提示"
                    Exit Sub
                End If
            End If
        End If
        If id.Text = "" Or _
            Trim(strName.Text) = "" Or _
            moral_score.Text = "" Then
            MsgBox "学号、姓名或德育分不能为空", vbOKOnly + vbExclamation, "提示"
            Exit Sub
        End If
        Page.Text = Val(Page.Text) + 1
        for_count.Caption = "/ " & Page.Text
        sql = "INSERT INTO " & table_name & " VALUES(" & Page.Text & _
            "," & Page.Text & ",'',75)"
        db_conn.Execute sql
        sql = "SELECT * FROM " & table_name & " WHERE index = " & Page.Text
        db_record = db_conn.Execute(sql)
        id.Text = db_record("id").Value
        strName.Text = Trim(db_record("name").Value)
        moral_score.Text = db_record("moral_score").Value
        change_value.Text = "0"
        If Val(Page.Text) > 0 Then
            btn_previous.Enabled = True
        End If
    ElseIf Val(Page.Text) = Val(Mid(for_count.Caption, 3)) - 1 Then
        If id.Text = "" Or _
            Trim(strName.Text) = "" Or _
            moral_score.Text = "" Then
            MsgBox "学号、姓名或德育分不能为空", vbOKOnly + vbExclamation, "提示"
            Exit Sub
        End If
        If db_count("row").Value > 0 Then
            If Val(id.Text) <> db_record("id").Value Or _
                Trim(strName.Text) <> Trim(db_record("name").Value) Or _
                Val(moral_score.Text) <> db_record("moral_score").Value Then
                isSave = MsgBox("记录已修改，是否保存记录？" & vbCrLf & Chr(34) & "是" & Chr(34) & _
                    "为保存并向下一条记录，" & Chr(34) & "否" & Chr(34) & "为还原并" & _
                    "向下一条记录", vbYesNo + vbExclamation, "温馨提示")
                If isSave = vbYes Then
                    sql = "UPDATE " & table_name & " SET id = " & id.Text & ", name = '" & strName.Text & _
                        "', moral_score = " & moral_score.Text & " WHERE index = " & Page.Text
                    db_conn.Execute (sql)
                Else
                    id.Text = db_record("id").Value
                    strName.Text = Trim(db_record("name").Value)
                    moral_score.Text = db_record("moral_score").Value
                    change_value.Text = "0"
                    If id.Text = "" Or _
                        Trim(strName.Text) = "" Or _
                        moral_score.Text = "" Then
                        MsgBox "学号、姓名或德育分不能为空", vbOKOnly + vbExclamation, "提示"
                        Exit Sub
                    End If
                End If
            End If
        End If
        btn_previous.Enabled = True
        btn_next.Caption = "添加"
        sql = "SELECT COUNT(*) AS row FROM " & table_name
        db_count = db_conn.Execute(sql)
        Page.Text = Val(Page.Text) + 1
        sql = "SELECT * FROM " & table_name & " WHERE index = " & Page.Text
        db_record = db_conn.Execute(sql)
        id.Text = db_record("id").Value
        strName.Text = Trim(db_record("name").Value)
        moral_score.Text = db_record("moral_score").Value
        change_value.Text = "0"
    Else
        If id.Text = "" Or _
            Trim(strName.Text) = "" Or _
            moral_score.Text = "" Then
            MsgBox "学号、姓名或德育分不能为空", vbOKOnly + vbExclamation, "提示"
            Exit Sub
        Else
            'first check note
            sql = "SELECT * FROM " & table_name & " WHERE index = " & Page.Text
            db_record = db_conn.Execute(sql)
            'when database's count is not empty
            If db_count("row").Value > 0 Then
                If Val(id.Text) <> db_record("id").Value Or _
                    Trim(strName.Text) <> Trim(db_record("name").Value) Or _
                    Val(moral_score.Text) <> db_record("moral_score").Value Then
                    isSave = MsgBox("记录已修改，是否保存记录？" & vbCrLf & Chr(34) & "是" & Chr(34) & _
                        "为保存并向下一条记录，" & Chr(34) & "否" & Chr(34) & "为还原并" & _
                        "向下一条记录", vbYesNo + vbExclamation, "温馨提示")
                    If isSave = vbYes Then
                        sql = "UPDATE " & table_name & " SET id = " & id.Text & ", name = '" & strName.Text & _
                            "', moral_score = " & moral_score.Text & " WHERE index = " & Page.Text
                        db_conn.Execute (sql)
                    Else
                        id.Text = db_record("id").Value
                        strName.Text = Trim(db_record("name").Value)
                        moral_score.Text = db_record("moral_score").Value
                        change_value.Text = "0"
                        If id.Text = "" Or _
                            Trim(strName.Text) = "" Or _
                            moral_score.Text = "" Then
                            MsgBox "学号、姓名或德育分不能为空", vbOKOnly + vbExclamation, "提示"
                            Exit Sub
                        End If
                    End If
                    'finish check
                    'select next
                    Page.Text = Val(Page.Text) + 1
                    sql = "SELECT * FROM " & table_name & " WHERE index = " & Page.Text
                    db_record = db_conn.Execute(sql)
                    id.Text = db_record("id").Value
                    strName.Text = Trim(db_record("name").Value)
                    moral_score.Text = db_record("moral_score").Value
                    change_value.Text = "0"
                Else
                    Page.Text = Val(Page.Text) + 1
                    sql = "SELECT * FROM " & table_name & " WHERE index = " & Page.Text
                    db_record = db_conn.Execute(sql)
                    id.Text = db_record("id").Value
                    strName.Text = Trim(db_record("name").Value)
                    moral_score.Text = db_record("moral_score").Value
                    change_value.Text = "0"
                    btn_previous.Enabled = True
                End If
            'or ...
            Else
                Page.Text = Val(Page.Text) + 1
                sql = "SELECT * FROM " & table_name & " WHERE index = " & Page.Text
                db_record = db_conn.Execute(sql)
                id.Text = db_record("id").Value
                strName.Text = Trim(db_record("name").Value)
                moral_score.Text = db_record("moral_score").Value
                change_value.Text = "0"
                btn_previous.Enabled = True
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
            End If
        End If
    End If
End Sub

Private Sub btn_previous_Click()
    'check is changed
    sql = "SELECT * FROM " & table_name & " WHERE index = " & Page.Text
    db_record = db_conn.Execute(sql)
    If id.Text <> db_record("id").Value Or _
        Trim(strName.Text) <> Trim(db_record("name").Value) Or _
        moral_score.Text <> db_record("moral_score").Value Then
        isSave = MsgBox("记录已修改，是否保存记录？" & vbCrLf & Chr(34) & "是" & Chr(34) & _
            "为保存并向上一条记录，" & Chr(34) & "否" & Chr(34) & "为还原并" & _
            "向上一条记录", vbYesNo + vbExclamation, "温馨提示")
        If isSave = vbYes Then
            sql = "UPDATE " & table_name & " SET id = " & id.Text & ", name = '" & strName.Text & _
                "', moral_score = " & moral_score.Text & " WHERE index = " & Page.Text
            db_conn.Execute sql
        End If
    End If
    'go into previous note
    If Val(Page.Text) = 2 Then
        Page.Text = Val(Page.Text) - 1
        btn_previous.Enabled = False
        sql = "SELECT * FROM " & table_name & " WHERE index = 1"
        db_record = db_conn.Execute(sql)
        id.Text = db_record("id").Value
        strName.Text = Trim(db_record("name").Value)
        moral_score.Text = db_record("moral_score").Value
        change_value.Text = "0"
    ElseIf Val(Page.Text) > 2 Then
        Page.Text = Val(Page.Text) - 1
        sql = "SELECT * FROM " & table_name & " WHERE index = " & Page.Text
        db_record = db_conn.Execute(sql)
        id.Text = db_record("id").Value
        strName.Text = Trim(db_record("name").Value)
        moral_score.Text = db_record("moral_score").Value
        change_value.Text = "0"
    End If
        btn_next.Caption = "下一条"
End Sub

Private Sub btn_return_Click()
    Dim isReturn
    isReturn = MsgBox("确定要还原记录吗？" & vbCrLf & Chr(34) & "是" & Chr(34) & _
        "为还原记录，" & Chr(34) & "否" & Chr(34) & "为不还原", vbYesNo + vbExclamation, "温馨提示")
    If isReturn = vbYes Then
        sql = "SELECT * FROM " & table_name & " WHERE index = " & Page.Text
        db_record = db_conn.Execute(sql)
        If db_count("row").Value > 0 Then
            id.Text = db_record("id").Value
            strName.Text = Trim(db_record("name").Value)
            moral_score.Text = db_record("moral_score").Value
            change_value.Text = "0"
        Else
            MsgBox "数据库中无相关内容，无法还原", vbOKOnly + vbExclamation, "提示"
        End If
    End If
End Sub

Private Sub btn_save_Click()
    On Error GoTo Error
    If id.Text = "" Or _
        Trim(strName.Text) = "" Or _
        moral_score.Text = "" Then
            MsgBox "学号、姓名或德育分不能为空", vbOKOnly + vbExclamation, "提示"
            Exit Sub
    End If
    sql = "SELECT COUNT(*) AS row FROM " & table_name
    db_count = db_conn.Execute(sql)
    If Val(db_count("row").Value) > 0 And Val(db_count("row").Value) >= Val(Page.Text) Then
        sql = "UPDATE " & table_name & " SET id = " & id.Text & ", name = '" & strName.Text & _
            "', moral_score = " & moral_score.Text & " WHERE index = " & Page.Text
        db_conn.Execute (sql)
    Else
        sql = "INSERT INTO " & table_name & " VALUES(" & Page.Text & "," & id.Text & ",'" & _
            strName.Text & "'," & moral_score.Text & ")"
        MsgBox sql
        db_conn.Execute (sql)
    End If
    sql = "SELECT * FROM " & table_name & " WHERE index = " & Page.Text
    db_record = db_conn.Execute(sql)
    If db_record("id").Value = id.Text And _
        Trim(db_record("name").Value) = Trim(strName.Text) And _
        db_record("moral_score").Value = moral_score.Text Then
        MsgBox "保存成功", vbOKOnly + vbExclamation, "提示"
    Else
        MsgBox "保存失败", vbOKOnly + vbCritical, "错误"
    End If
    If db_record("moral_score").Value < 60 Then
        grade.Caption = "不及格"
    ElseIf db_record("moral_score").Value >= 60 And _
        db_record("moral_score").Value < 75 Then
        grade.Caption = "及格"
    ElseIf db_record("moral_score").Value >= 75 And _
        db_record("moral_score").Value < 85 Then
        grade.Caption = "良好"
    Else
        grade.Caption = "优秀"
    End If
    Exit Sub
Error:
    MsgBox "出现错误，错误信息：" & Err.Description, vbOKOnly, "出错"
    Exit Sub
End Sub

Private Sub change_moral_score_Click()
    moral_score.Text = Val(moral_score.Text) + Val(change_value.Text)
    If db_record("moral_score").Value < 60 Then
        grade.Caption = "不及格"
    ElseIf db_record("moral_score").Value >= 60 And _
        db_record("moral_score").Value < 75 Then
        grade.Caption = "及格"
    ElseIf db_record("moral_score").Value >= 75 And _
        db_record("moral_score").Value < 85 Then
        grade.Caption = "良好"
    Else
        grade.Caption = "优秀"
    End If
End Sub

Private Sub change_value_Change()
    If change_value.Text = "" Then
        change_value.Text = "0"
    End If
End Sub

Private Sub down_Click()
    change_value.Text = Val(change_value.Text) - 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Val(Page.Text) > 0 Then
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
            sql = "DELETE FROM " & table_name & " WHERE index = " & Page.Text
            db_conn.Execute sql
            db_conn.Close
            End
        End If
        If db_count("row").Value > 0 Then
            If id.Text <> db_record("id").Value Or _
                Trim(strName.Text) <> Trim(db_record("name").Value) Or _
                moral_score.Text <> db_record("moral_score").Value Then
                isSave = MsgBox("记录已修改，是否保存记录？" & vbCrLf & Chr(34) & "是" & Chr(34) & _
                    "为保存并退出，" & Chr(34) & "否" & Chr(34) & "为还原并" & "退出", vbYesNo + _
                    vbExclamation, "温馨提示")
                If isSave = vbYes Then
                    sql = "UPDATE " & table_name & " SET id = " & id.Text & ", name = '" & strName.Text & _
                        "', moral_score = " & moral_score.Text & " WHERE index = " & Page.Text
                    db_conn.Execute sql
                End If
            End If
        Else
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
    On Error GoTo OperatorError
    Set db_conn = New ADODB.Connection
    db_conn.Open db_drive, "Admin"
    Set db_record = New ADODB.Recordset
    Set db_count = New ADODB.Recordset
    sql = "SELECT COUNT(*) AS row FROM " & table_name
    db_count = db_conn.Execute(sql)
    for_count.Caption = "/ " & db_count("row").Value
    If Val(db_count("row").Value) > 0 Then
        Page.Text = "1"
        If Page.Text = Mid(for_count.Caption, 3) Then
            btn_next.Caption = "添加"
        ElseIf Val(Page.Text) > Val(Mid(for_count.Caption, 3)) Then
            btn_next.Caption = "下一条"
        End If
        sql = "SELECT * FROM " & table_name & " WHERE index = " & Page.Text
        db_record = db_conn.Execute(sql)
        id.Text = db_record("id").Value
        strName.Text = Trim(db_record("name").Value)
        moral_score.Text = db_record("moral_score").Value
        change_value.Text = "0"
        btn_previous.Enabled = False
        If db_record("moral_score").Value < 60 Then
            grade.Caption = "不及格"
        ElseIf db_record("moral_score").Value >= 60 And _
            db_record("moral_score").Value < 75 Then
            grade.Caption = "及格"
        ElseIf db_record("moral_score").Value >= 75 And _
            db_record("moral_score").Value < 85 Then
            grade.Caption = "良好"
        Else
            grade.Caption = "优秀"
        End If
    Else
        Page.Text = "0"
        id.Enabled = False
        strName.Enabled = False
        moral_score.Enabled = False
        change_value.Enabled = False
        grade.Enabled = False
        btn_previous.Enabled = False
        btn_next.Caption = "添加"
        btn_save.Enabled = False
        btn_return.Enabled = False
        down.Enabled = False
        up.Enabled = False
    End If
    Exit Sub
OperatorError:
    MsgBox "操作出错，错误信息：" & Err.Description, vbOKOnly + vbExclamation, "提示"
    db_conn.Close
    End
End Sub

Private Sub moral_score_Change()
    If moral_score.Text = "" Then
        moral_score.Text = "0"
    End If
    Dim index As Integer
    For index = 1 To Len(moral_score.Text)
        If Asc(Mid(moral_score.Text, index, 1)) < Asc("0") Or _
            Asc(Mid(moral_score.Text, index, 1)) > Asc("9") Then
            MsgBox "德育分不允许出现非数字", vbOKOnly + vbExclamation, "提示"
        Else
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
        End If
    Next
End Sub

Private Sub up_Click()
    change_value.Text = Val(change_value.Text) + 1
End Sub
