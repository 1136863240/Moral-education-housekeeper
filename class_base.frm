VERSION 5.00
Begin VB.Form class_base 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "德育管家v0.4 - 班级基本情况"
   ClientHeight    =   4470
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   6525
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton change_class 
      Caption         =   "切换班级"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4815
      TabIndex        =   8
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton change_single 
      Caption         =   "单个查看"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2595
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton look_all 
      Caption         =   "查看所有"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   375
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label prompt 
      Caption         =   "←单击显示所有最高者"
      Height          =   195
      Left            =   3780
      TabIndex        =   13
      Top             =   3840
      Width           =   1875
   End
   Begin VB.Label Label6 
      Caption         =   "人数"
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
      Left            =   480
      TabIndex        =   12
      Top             =   1710
      Width           =   975
   End
   Begin VB.Label count_student 
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
      Left            =   1680
      TabIndex        =   11
      Top             =   1710
      Width           =   1755
   End
   Begin VB.Label max_score_student 
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
      Left            =   2160
      TabIndex        =   10
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "最高者"
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
      Left            =   480
      TabIndex        =   9
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label max_score 
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
      Left            =   2160
      TabIndex        =   5
      Top             =   3030
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "最高分"
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
      Left            =   480
      TabIndex        =   4
      Top             =   3030
      Width           =   1335
   End
   Begin VB.Label sqr_score 
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
      Left            =   2160
      TabIndex        =   3
      Top             =   2340
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "平均分"
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
      Left            =   480
      TabIndex        =   2
      Top             =   2340
      Width           =   1335
   End
   Begin VB.Label class_name 
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
      Left            =   1680
      TabIndex        =   1
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "班级"
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
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   975
   End
End
Attribute VB_Name = "class_base"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub change_class_Click()
    db_conn.Close
    login.Show
    Unload Me
End Sub

Private Sub change_single_Click()
    manage.Show
    Unload Me
End Sub

Private Sub Form_Load()
    Set db_conn = New ADODB.Connection
    db_conn.Open db_class_drive
    class_name.Caption = table_name
    sql = "SELECT COUNT(*) AS students_count FROM " & table_name
    db_record = db_conn.Execute(sql)
    Dim students_count
    students_count = Val(db_record("students_count").Value)
    If students_count = 0 Then
        sqr_score.Caption = "0"
        max_score.Caption = "0"
        count_student.Caption = "0"
        max_score_student.Caption = "无"
        prompt.Visible = False
        Exit Sub
    End If
    count_student.Caption = students_count
    sql = "SELECT Avg(moral_score) AS sqr_score FROM " & table_name
    db_record = db_conn.Execute(sql)
    Dim sqr
    sqr = Val(db_record("sqr_score").Value)
    sqr_score.Caption = sqr
    sql = "SELECT Max(moral_score) AS max_score FROM " & table_name
    db_record = db_conn.Execute(sql)
    Dim max
    max = Val(db_record("max_score").Value)
    max_score.Caption = max
    sql = "SELECT COUNT(*) AS max_count FROM " & table_name & " WHERE moral_score = " & max
    db_record = db_conn.Execute(sql)
    Dim max_count
    max_count = Val(db_record("max_count").Value)
    If max_count > 1 Then
        max_score_student.Caption = max_count & "人"
        prompt.Visible = True
    Else
        sql = "SELECT name FROM " & table_name & " WHERE moral_score = " & max
        db_record = db_conn.Execute(sql)
        max_score_student.Caption = db_record("name").Value
        prompt.Visible = False
    End If
End Sub

Private Sub look_all_Click()
    select_list.Show
    Unload Me
End Sub

Private Sub max_score_student_Click()
    If db_conn.State = adStateClosed Then
        db_conn.Open db_drive
    End If
    sql = "SELECT Max(moral_score) AS max_score FROM " & table_name
    db_record = db_conn.Execute(sql)
    Dim max
    max = Val(db_record("max_score").Value)
    sql = "SELECT COUNT(*) AS max_count FROM " & table_name & " WHERE moral_score = " & max
    db_record = db_conn.Execute(sql)
    Dim max_count
    max_count = Val(db_record("max_count").Value)
    If max_count > 1 Then
        max_students.Show vbModal, Me
    End If
End Sub
