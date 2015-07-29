VERSION 5.00
Begin VB.Form user_define_data 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "德育管家v1.0 - 自定义筛选"
   ClientHeight    =   2745
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4560
   StartUpPosition =   2  '屏幕中心
   Begin VB.OptionButton as_id 
      Caption         =   "按学号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   600
      Width           =   1095
   End
   Begin VB.OptionButton as_moral_score 
      Caption         =   "按德育分"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cancel 
      Caption         =   "取消"
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
      Left            =   2633
      TabIndex        =   5
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton okay 
      Caption         =   "确定"
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
      Left            =   593
      TabIndex        =   4
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CheckBox moral_score_range 
      Caption         =   "德育分范围"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2393
      TabIndex        =   3
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CheckBox id_range 
      Caption         =   "学号范围"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   593
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label select 
      Caption         =   "筛选条件"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label sort 
      Caption         =   "排列方式"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "user_define_data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancel_Click()
    is_select = False
    Unload Me
End Sub

Private Sub Form_Load()
    If sort_index = 0 Then
        as_id.Value = True
        as_moral_score.Value = False
    Else
        as_id.Value = False
        as_moral_score.Value = True
    End If
    If select_id = True Then
        id_range.Value = 1
    Else
        id_range.Value = 0
    End If
    If select_score = True Then
        moral_score_range.Value = 1
    Else
        moral_score_range.Value = 0
    End If
End Sub

Private Sub id_range_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If id_range.Value = 0 Then
        Exit Sub
    End If
    select_range_index = 0
    select_range.Show vbModal, Me
    If isChange = True Then
        id_range.Value = 1
    Else
        id_range.Value = 0
    End If
End Sub

Private Sub moral_score_range_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If moral_score_range.Value = 0 Then
        Exit Sub
    End If
    select_range_index = 1
    select_range.Show vbModal, Me
    If isChange = True Then
        moral_score_range.Value = 1
    Else
        moral_score_range.Value = 0
    End If
End Sub

Private Sub okay_Click()
    'sort as ...
    If as_id.Value = True Then
        sort_as_id = True
        is_select = True
    Else
        sort_as_id = False
    End If
    If as_moral_score.Value = True Then
        sort_as_moral_score = True
        is_select = True
    Else
        sort_as_moral_score = False
    End If
    
    'select range
    If id_range.Value = 1 Then
        select_id = True
    Else
        select_id = False
    End If
    If moral_score_range.Value = 1 Then
        select_score = True
    Else
        select_score = False
    End If
    Unload Me
End Sub
