VERSION 5.00
Begin VB.Form max_students 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "德育管家v0.4 - 德育分最高者"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   4620
   StartUpPosition =   2  '屏幕中心
   Begin VB.ListBox max_students_list 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4515
   End
End
Attribute VB_Name = "max_students"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    max_students_list.Clear
    Dim conn As New ADODB.Connection
    Dim record As New ADODB.Recordset
    conn.Open db_class_drive
    sql = "SELECT Max(moral_score) AS max_moral_score FROM " & table_name
    record.Open sql, conn
    Dim max
    max = Val(record("max_moral_score").Value)
    sql = "SELECT COUNT(*) AS max_students_count FROM " & table_name & _
        " WHERE moral_score = " & max
    record.Close
    record.Open sql, conn
    Dim max_students_count, max_students_index, item
    item = "   学号     姓名    德育分"
    max_students_list.AddItem item, 0
    max_students_count = Val(record("max_students_count").Value)
    sql = "SELECT id,name,moral_score FROM " & table_name & _
        " WHERE moral_score = " & max
    record.Close
    record.Open sql, conn
    For max_students_index = 1 To max_students_count
        item = "    " & record("id").Value & "       " & record("name").Value & _
            "" & record("moral_score").Value
        max_students_list.AddItem item, max_students_index
        record.MoveNext
    Next
End Sub
