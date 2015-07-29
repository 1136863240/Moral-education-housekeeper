VERSION 5.00
Begin VB.Form login 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "德育管家v1.0 - 请输入班级"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   4740
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox login_user 
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
      IMEMode         =   3  'DISABLE
      Left            =   1635
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   1072
      Width           =   2415
   End
   Begin VB.TextBox Password 
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
      IMEMode         =   3  'DISABLE
      Left            =   1635
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1710
      Width           =   2415
   End
   Begin VB.ComboBox Class_Grade 
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
      ItemData        =   "login.frx":0000
      Left            =   1635
      List            =   "login.frx":0002
      TabIndex        =   0
      Top             =   435
      Width           =   2415
   End
   Begin VB.CommandButton exit 
      Cancel          =   -1  'True
      Caption         =   "退出"
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
      Left            =   2603
      TabIndex        =   3
      Top             =   2535
      Width           =   1455
   End
   Begin VB.CommandButton login 
      Caption         =   "确定"
      Default         =   -1  'True
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
      Left            =   683
      TabIndex        =   2
      Top             =   2535
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "身份"
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
      Left            =   686
      TabIndex        =   7
      Top             =   1114
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "密码"
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
      Left            =   690
      TabIndex        =   5
      Top             =   1740
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "班级"
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
      Left            =   686
      TabIndex        =   4
      Top             =   488
      Width           =   615
   End
   Begin VB.Menu about_software_link 
      Caption         =   "关于软件(&A)..."
      NegotiatePosition=   1  'Left
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub about_software_link_Click()
    about_software.Show vbModal, Me
End Sub

Private Sub login_Click()
    For strIndex = 1 To Len(Class_Grade.Text)
        If Mid(Class_Grade.Text, strIndex, 1) = " " Then
            MsgBox "班级名称中不能出现空格，请检查后重试", vbOKOnly + vbExclamation, "提示"
            Exit Sub
        End If
    Next

    If Class_Grade.Text = "" Then
        MsgBox "班级名称不能为空", vbOKOnly + vbExclamation, "温馨提示"
    Else
        'create a drive string
        db_password = Password.Text
        table_name = Class_Grade.Text
        If db_password = "" Then
            If Dir(db_name & Class_Grade.Text & ".mdb") = "" Then
                Dim isPassword
                isPassword = MsgBox("密码为空，安全性较低，是否继续？", _
                    vbYesNo + vbExclamation, "温馨提示")
                If isPassword = vbYes Then
                    db_class_drive = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Chr(34) & _
                        db_name & table_name & ".mdb" & Chr(34)
                    If Dir(db_name & Class_Grade.Text & ".mdb") = "" Then
                        Set catalog = New adox.catalog
                        Set db_table = New adox.Table
                        On Error GoTo DatabaseError
                        catalog.Create db_class_drive
                        catalog.ActiveConnection = db_class_drive
                        table_name = Class_Grade.Text
                        db_table.Name = table_name
                        db_table.Columns.Append "index", adox.DataTypeEnum.adInteger, 10
                        db_table.Columns.Append "id", adox.DataTypeEnum.adInteger, 10
                        db_table.Columns.Append "name", adox.DataTypeEnum.adWChar, 10
                        db_table.Columns.Append "moral_score", adox.DataTypeEnum.adInteger
                        db_table.Columns.Append "explicit", adox.DataTypeEnum.adWChar, 200
                        catalog.Tables.Append db_table
                        If Dir(db_name & table_name & "德育分细则\", vbDirectory) = "" Then
                            MkDir db_name & table_name & "德育分细则\"
                        End If
                        class_base.Show
                        Unload Me
                    Else
                        db_password = Password.Text
                        db_single_drive = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Chr(34) & _
                            db_name & Class_Grade.Text & ".mdb" & Chr(34)
                        Set db_conn = New ADODB.Connection
                        On Error Resume Next
                        db_conn.Open db_single_drive
                        If db_conn.State = adStateClosed Then
                            MsgBox "密码有误！", vbOKOnly + vbCritical, "错误"
                            Exit Sub
                        Else
                            db_conn.Close
                            class_base.Show
                            Unload Me
                        End If
                    End If
                End If
            Else
                db_class_drive = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Chr(34) & _
                    db_name & table_name & ".mdb" & Chr(34)
                If Dir(db_name & Class_Grade.Text & ".mdb") = "" Then
                    Set catalog = New adox.catalog
                    Set db_table = New adox.Table
                    On Error GoTo DatabaseError
                    catalog.Create db_class_drive
                    catalog.ActiveConnection = db_class_drive
                    table_name = Class_Grade.Text
                    db_table.Name = table_name
                    db_table.Columns.Append "index", adox.DataTypeEnum.adInteger
                    db_table.Columns.Append "id", adox.DataTypeEnum.adInteger
                    db_table.Columns.Append "name", adox.DataTypeEnum.adWChar
                    db_table.Columns.Append "moral_score", adox.DataTypeEnum.adInteger
                    db_table.Columns.Append "explicit", adox.DataTypeEnum.adWChar, 200
                    catalog.Tables.Append db_table
                    manage.Show
                    Unload Me
                Else
                    db_password = Password.Text
                    db_single_drive = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Chr(34) & _
                        db_name & Class_Grade.Text & ".mdb" & Chr(34)
                    Set db_conn = New ADODB.Connection
                    On Error Resume Next
                    db_conn.Open db_single_drive
                    If db_conn.State = adStateClosed Then
                        MsgBox "密码有误！", vbOKOnly + vbCritical, "错误"
                        Exit Sub
                    Else
                        db_conn.Close
                        class_base.Show
                        Unload Me
                    End If
                End If
            End If
        Else
            db_single_drive = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Chr(34) & _
                db_name & Class_Grade.Text & ".mdb" & Chr(34) & ";Jet OLEDB:Database" & _
                " Password=" & db_password
            'if no exist database and table
            'create them
            table_name = Class_Grade.Text
            If Dir(db_name & Class_Grade.Text & ".mdb") = "" Then
                Set catalog = New adox.catalog
                Set db_table = New adox.Table
                On Error GoTo DatabaseError
                catalog.Create db_single_drive
                catalog.ActiveConnection = db_single_drive
                table_name = Class_Grade.Text
                db_table.Name = table_name
                db_table.Columns.Append "index", adox.DataTypeEnum.adInteger
                db_table.Columns.Append "id", adox.DataTypeEnum.adInteger
                db_table.Columns.Append "name", adox.DataTypeEnum.adWChar
                db_table.Columns.Append "moral_score", adox.DataTypeEnum.adInteger
                db_table.Columns.Append "explicit", adox.DataTypeEnum.adWChar
                catalog.Tables.Append db_table
                class_base.Show
                Unload Me
            Else
                db_password = Password.Text
                db_single_drive = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Chr(34) & _
                    db_name & Class_Grade.Text & ".mdb" & Chr(34) & ";Jet OLEDB:Database" & _
                    " Password=" & db_password
                Set db_conn = New ADODB.Connection
                On Error Resume Next
                db_conn.Open db_single_drive
                If db_conn.State = adStateClosed Then
                    MsgBox "密码有误！", vbOKOnly + vbCritical, "错误"
                    Exit Sub
                Else
                    db_conn.Close
                    class_base.Show
                    Unload Me
                End If
            End If
        End If
    End If
    Exit Sub
DatabaseError:
    MsgBox "数据库操作失败，错误信息：" & Err.Description, vbOKOnly + vbExclamation, "提示"
    Exit Sub
End Sub

Private Sub exit_Click()
    End
End Sub

Private Sub Form_Load()
    MsgBox d_en("123456")
    MsgBox d_de(d_en("123456"))
    On Error GoTo Error
    db_name = App.Path
    If Right(db_name, 1) <> "\" Then 'if it is child directory
        If Dir(db_name & "\db", vbDirectory) = "" Then 'if \db directory is not exist
            'create a directory, name is \db
            MkDir db_name & "\db"
        End If
        db_name = db_name & "\db\"
    Else
        If Dir(db_name & "db", vbDirectory) = "" Then
            'same as
            MkDir db_name & "db"
        End If
        db_name = db_name & "db\"
    End If
    Dim file
    Set found_file = CreateObject("Scripting.FileSystemObject")
    Set folder = found_file.GetFolder(db_name)
    Set file_count = folder.Files
    Class_Grade.Clear
    For Each file In file_count
        Dim file_name
        If Right(file.Name, 4) <> ".mdb" Then
            GoTo GoNextLoop
        End If
        file_name = Left(file.Name, InStr(1, file.Name, ".mdb") - 1)
        Class_Grade.AddItem file_name, Class_Grade.ListCount
GoNextLoop:
    Next
    If Class_Grade.ListCount > 0 Then
        Class_Grade.ListIndex = 0
    End If
    Exit Sub
Error:
    MsgBox "出现错误，错误信息：" & Err.Description, vbOKOnly + vbCritical, "错误"
End Sub
