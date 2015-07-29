VERSION 5.00
Begin VB.Form select_range 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "德育管家v0.4 - 范围选择"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4005
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   4005
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cancel 
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2227
      TabIndex        =   5
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton okay 
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
      Height          =   555
      Left            =   300
      TabIndex        =   4
      Top             =   1440
      Width           =   1455
   End
   Begin VB.ComboBox tallest_value 
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
      Left            =   1905
      TabIndex        =   3
      Top             =   840
      Width           =   1635
   End
   Begin VB.ComboBox least_value 
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
      Left            =   1905
      TabIndex        =   1
      Top             =   240
      Width           =   1635
   End
   Begin VB.Label Label2 
      Caption         =   "最大值"
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
      Left            =   465
      TabIndex        =   2
      Top             =   893
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "最小值"
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
      Left            =   465
      TabIndex        =   0
      Top             =   293
      Width           =   975
   End
End
Attribute VB_Name = "select_range"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cancel_Click()
    isChange = False
    Unload Me
End Sub

Private Sub Form_Load()
    least_value.Clear
    tallest_value.Clear
    Select Case select_range_index
    Case 0
        Caption = "德育管家v0.4 - 学号范围选择"
        least_value.AddItem "1", 0
        least_value.AddItem "自定义<...>", 1
        least_value.ListIndex = 0
        tallest_value.AddItem "最大值", 0
        tallest_value.AddItem "自定义<...>", 1
        tallest_value.ListIndex = 0
    Case 1
        Caption = "德育管家v0.4 - 德育分范围选择"
        least_value.AddItem "最小值", 0
        least_value.AddItem "自定义<...>", 1
        least_value.ListIndex = 0
        tallest_value.AddItem "最大值", 0
        tallest_value.AddItem "自定义<...>", 1
        tallest_value.ListIndex = 0
    End Select
End Sub

Private Sub least_value_Click()
    If select_range_index = 1 Then
        If least_value.ListIndex = 0 Then
            If least_value.Text = "最小值" Then
                sql = "SELECT Min(moral_score) AS min_score FROM " & table_name
                If db_conn.State = adStateClosed Then
                    db_conn.Open db_class_drive
                End If
                db_record = db_conn.Execute(sql)
                least_value.Text = Trim(Str(db_record("min_score").Value))
            End If
        ElseIf least_value.ListIndex = 1 Then
            least_value.Text = ""
        End If
    End If
End Sub

Private Sub okay_Click()
    If least_value.Text = "" Or _
        tallest_value.Text = "" Then
        MsgBox "最小值或最大值不能为空", vbOKOnly + vbCritical, "提示"
        Exit Sub
    End If
    If select_range_index = 0 Or _
        select_range_index = 1 Then
        If Val(least_value.Text) > Val(tallest_value.Text) Then
            MsgBox "最小值不能大于最大值", vbOKOnly + vbCritical, "提示"
            Exit Sub
        End If
    Else
        If least_value.ListIndex = -1 Or _
            tallest_value.ListIndex = -1 Then
            MsgBox "最小值或最大值不合法", vbOKOnly + vbCritical, "提示"
            Exit Sub
        ElseIf least_value.ListIndex > tallest_value.ListIndex Then
            MsgBox "最小值不能大于最大值", vbOKOnly + vbCritical, "提示"
            Exit Sub
        End If
    End If
    isChange = True
    If select_range_index = 0 Then
        range_id_least = Val(least_value.Text)
        range_id_most = Val(tallest_value.Text)
    ElseIf select_range_index = 1 Then
        range_score_least = Val(least_value.Text)
        range_score_most = Val(tallest_value.Text)
    End If
    Unload Me
End Sub

Private Sub tallest_value_Click()
    If select_range_index = 0 Then
        If tallest_value.ListIndex = 0 Then
            If tallest_value.Text = "最大值" Then
                sql = "SELECT Max(id) AS max_id FROM " & table_name
                If db_conn.State = adStateClosed Then
                    db_conn.Open db_class_drive
                End If
                db_record = db_conn.Execute(sql)
                tallest_value.Text = Trim(Str(db_record("max_id").Value))
            End If
        ElseIf tallest_value.ListIndex = 1 Then
            tallest_value.Text = ""
        End If
    End If
    If select_range_index = 1 Then
        If tallest_value.ListIndex = 0 Then
            If tallest_value.Text = "最大值" Then
                sql = "SELECT Max(moral_score) AS max_score FROM " & table_name
                If db_conn.State = adStateClosed Then
                    db_conn.Open db_class_drive
                End If
                db_record = db_conn.Execute(sql)
                tallest_value.Text = Trim(Str(db_record("max_score").Value))
            End If
        ElseIf tallest_value.ListIndex = 1 Then
            tallest_value.Text = ""
        End If
    End If
End Sub
