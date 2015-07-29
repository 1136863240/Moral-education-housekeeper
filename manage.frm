VERSION 5.00
Begin VB.Form manage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "德育管家v1.0 - 德育分查看"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   6060
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   6060
   StartUpPosition =   2  '屏幕中心
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
      Left            =   1508
      TabIndex        =   7
      Top             =   1207
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
      Left            =   4388
      TabIndex        =   5
      Top             =   337
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
      Left            =   1508
      TabIndex        =   3
      Top             =   337
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
      Left            =   4313
      TabIndex        =   1
      Top             =   2025
      Width           =   1290
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
      Left            =   458
      TabIndex        =   0
      Top             =   2025
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
      Left            =   4508
      TabIndex        =   9
      Top             =   1252
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
      Left            =   3533
      TabIndex        =   8
      Top             =   1252
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
      Left            =   413
      TabIndex        =   6
      Top             =   1252
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
      Left            =   3533
      TabIndex        =   4
      Top             =   382
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
      Left            =   413
      TabIndex        =   2
      Top             =   382
      Width           =   615
   End
   Begin VB.Menu back_to_class_base 
      Caption         =   "返回"
      Index           =   0
   End
   Begin VB.Menu mode 
      Caption         =   "模式"
      Index           =   10
      Begin VB.Menu mode_change 
         Caption         =   "修改模式"
         Index           =   11
      End
      Begin VB.Menu mode_see 
         Caption         =   "查看模式"
         Checked         =   -1  'True
         Index           =   12
      End
   End
   Begin VB.Menu select_data 
      Caption         =   "筛选"
      Index           =   20
      Begin VB.Menu condition_for_select 
         Caption         =   "条件筛选"
         Index           =   21
      End
      Begin VB.Menu see_all_for_select 
         Caption         =   "查看所有"
         Index           =   22
      End
   End
End
Attribute VB_Name = "manage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public isSave

Private Sub back_to_class_base_Click(Index As Integer)
    Unload Me
    class_base.Show
End Sub

Private Sub btn_next_Click()
    'check is changed
    If Val(page.Text) = 0 Then
        id.Enabled = True
        strName.Enabled = True
        moral_score.Enabled = True
        change_value.Enabled = True
        grade.Enabled = True
        btn_next.Enabled = False
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
        strName.Enabled = True
        strName.SetFocus
    ElseIf Val(page.Text) = Val(Mid(for_count.Caption, 3)) - 1 Then
        If id.Text = "" Or _
            Trim(strName.Text) = "" Or _
            moral_score.Text = "" Then
            MsgBox "学号、姓名或德育分不能为空", vbOKOnly + vbExclamation, "提示"
            Exit Sub
        End If
        btn_previous.Enabled = True
        strName.Enabled = False
        btn_next.Enabled = False
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
    If id.Text = "" Or _
        Trim(strName.Text) = "" Or _
        moral_score.Text = "" Then
        MsgBox "学号、姓名或德育分不能为空", vbOKOnly + vbExclamation, "提示"
        Exit Sub
    End If
    page.Text = Val(page.Text) - 1
    sql = "SELECT * FROM " & table_name & " WHERE index = " & page.Text
    If Val(page.Text) = 1 Then
        btn_previous.Enabled = False
    End If
    db_record = db_conn.Execute(sql)
    btn_next.Caption = "下一个"
    id.Text = db_record("id").Value
    strName.Text = Trim(db_record("name").Value)
    moral_score.Text = db_record("moral_score").Value
    change_value.Text = "0"
    strName.Enabled = False
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
        End If
    End If
    db_conn.Close
End Sub

Private Sub Form_Load()
    Me.Caption = "德育分管理器v1.0 - " & table_name & "德育分查看"
    Set db_conn = New ADODB.Connection
    db_conn.Open db_class_drive, "Admin"
    Set db_record = New ADODB.Recordset
    Set db_count = New ADODB.Recordset
    sql = "SELECT COUNT(*) AS row FROM " & table_name
    db_count = db_conn.Execute(sql)
    If Val(db_count("row").Value) > 0 Then
        page.Text = "1"
        If page.Text = Mid(for_count.Caption, 3) Then
            btn_next.Enabled = False
        ElseIf Val(page.Text) > Val(Mid(for_count.Caption, 3)) Then
            btn_next.Caption = "下一个"
        End If
        sql = "SELECT * FROM " & table_name & " WHERE index = " & page.Text
        db_record = db_conn.Execute(sql)
        id.Text = db_record("id").Value
        strName.Text = Trim(db_record("name").Value)
        moral_score.Text = db_record("moral_score").Value
        btn_previous.Enabled = False
        strName.Enabled = False
    Else
        page.Text = "0"
        id.Enabled = False
        strName.Enabled = False
        moral_score.Enabled = False
        change_value.Enabled = False
        grade.Enabled = False
        btn_previous.Enabled = False
        btn_next.Enabled = False
        down.Enabled = False
        up.Enabled = False
        change_moral_score.Enabled = False
        moral_score_check.Enabled = False
    End If
    Exit Sub
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

Private Sub id_KeyPress(KeyAscii As Integer)
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
        If Val(page.Text) = Val(Mid(for_count.Caption, 3)) And _
            Val(page.Text) > 1 Then
            btn_next.Enabled = False
            btn_previous.Enabled = True
        ElseIf Val(page.Text) = Val(Mid(for_count.Caption, 3)) And _
            Val(page.Text) = 1 Then
            btn_next.Enabled = False
            btn_previous.Enabled = False
        ElseIf Val(page.Text) < Val(Mid(for_count.Caption, 3)) And _
            Val(page.Text) = 1 Then
            btn_next.Caption = "下一个"
            btn_previous.Enabled = False
        Else
            btn_next.Caption = "下一个"
            btn_previous.Enabled = True
        End If
    End If
End Sub

Private Sub moral_score_Change()
    moral_score.Text = Val(moral_score.Text)
    sql = "UPDATE " & table_name & " SET moral_score = " & moral_score.Text & _
        " WHERE index = " & page.Text
    db_conn.Execute sql
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

Private Sub strName_Change()
    sql = "UPDATE " & table_name & " SET name = '" & strName.Text & _
        "',explicit = '" & db_name & table_name & "德育分细则\" & _
        strName.Text & ".mdb" & "' WHERE index = " & page.Text
    db_conn.Execute sql
End Sub
