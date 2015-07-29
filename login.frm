VERSION 5.00
Begin VB.Form login 
   Caption         =   "德育分管理器v0.1 - 请输入班级"
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   4740
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
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
      Top             =   1448
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
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
      Top             =   1448
      Width           =   1455
   End
   Begin VB.TextBox Class_Grade 
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
      Left            =   1639
      TabIndex        =   1
      Top             =   428
      Width           =   2415
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
      TabIndex        =   0
      Top             =   488
      Width           =   615
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    db_name = App.Path
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
        db_drive = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='"
        If Right(db_name, 1) <> "\" Then 'if it is child directory
            If Dir(db_name & "\db", vbDirectory) = "" Then 'if \db directory is not exist
                'create a directory, name is \db
                MkDir (db_name & "\db")
            End If
            db_name = db_name & "\db\"
        Else
            If Dir(db_name & "db", vbDirectory) = "" Then
                'same as
                MkDir (db_name & "db")
            End If
            db_name = db_name & "db\"
        End If
        db_drive = db_drive & db_name & Class_Grade.Text & ".mdb'"
        'if no exist database and table
        'create them
        If Dir(db_name & Class_Grade.Text & ".mdb") = "" Then
            Set catalog = New ADOX.catalog
            Set db_table = New ADOX.Table
            catalog.Create (db_drive)
            catalog.ActiveConnection = db_drive
            On Error GoTo DatabaseError
            table_name = Class_Grade.Text
            db_table.Name = table_name
            db_table.Columns.Append "index", ADOX.DataTypeEnum.adInteger
            db_table.Columns.Append "id", ADOX.DataTypeEnum.adInteger
            db_table.Columns.Append "name", ADOX.DataTypeEnum.adWChar
            db_table.Columns.Append "moral_score", ADOX.DataTypeEnum.adInteger
            catalog.Tables.Append db_table
        End If
        table_name = Class_Grade.Text
        Me.Hide
        manage.Show
        Exit Sub
    End If
DatabaseError:
    MsgBox "数据库操作失败，错误信息：" & Err.Description, vbOKOnly + vbExclamation, "提示"
    Exit Sub
End Sub

Private Sub Command2_Click()
End
End Sub

