VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "文件时间修改器"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5100
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   5100
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command5 
      Caption         =   "本地时间"
      Height          =   375
      Left            =   1360
      TabIndex        =   20
      Top             =   3000
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4800
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "退出软件"
      Height          =   375
      Left            =   3840
      TabIndex        =   19
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "关于软件"
      Height          =   375
      Left            =   2600
      TabIndex        =   18
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "确认修改"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "更改信息(每一个时间参数用空格隔开，如：2008 8 8)"
      ForeColor       =   &H000040C0&
      Height          =   1095
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   4815
      Begin VB.TextBox Text4 
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   960
         TabIndex        =   16
         Top             =   720
         Width           =   3735
      End
      Begin VB.TextBox Text3 
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   960
         TabIndex        =   15
         Top             =   480
         Width           =   3735
      End
      Begin VB.TextBox Text2 
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   960
         TabIndex        =   14
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "访问时间："
         ForeColor       =   &H000040C0&
         Height          =   180
         Index           =   5
         Left            =   120
         TabIndex        =   13
         Top             =   765
         Width           =   900
      End
      Begin VB.Label Label2 
         Caption         =   "修改时间："
         ForeColor       =   &H000040C0&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   488
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "创建时间："
         ForeColor       =   &H000040C0&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   248
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "原有信息"
      ForeColor       =   &H000040C0&
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   4815
      Begin VB.Label Label5 
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   960
         TabIndex        =   9
         Top             =   720
         Width           =   3735
      End
      Begin VB.Label Label4 
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label Label3 
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   960
         TabIndex        =   7
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "访问时间："
         ForeColor       =   &H000040C0&
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   900
      End
      Begin VB.Label Label2 
         Caption         =   "修改时间："
         ForeColor       =   &H000040C0&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "创建时间："
         ForeColor       =   &H000040C0&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "浏览(&V)"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   1320
      TabIndex        =   1
      Top             =   172
      Width           =   2655
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "要修改的文件"
      ForeColor       =   &H000040C0&
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   217
      Width           =   1080
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type FILETIME '结构体声明
dwLowDateTime As Long
dwHighDateTime As Long
End Type
Private Type SYSTEMTIME
wYear As Integer
wMonth As Integer
wDayOfWeek As Integer
wDay As Integer
wHour As Integer
wMinute As Integer
wSecond As Integer
wMilliseconds As Integer
End Type
Private Const GENERIC_WRITE = &H40000000 '常数声明
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2

'API声明
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Boolean
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Sub Command1_Click()

CommonDialog1.DialogTitle = "打开所有文件"
CommonDialog1.Filter = "所有文件|*.*"
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
Text1.Text = CommonDialog1.FileName
Call shuxin
End If
End Sub

Private Sub Command2_Click()
Dim lngHandle As Long
Dim udtFileTime As FILETIME
Dim udtLocalTime As FILETIME
Dim xiutime As FILETIME
Dim fangtime As FILETIME
Dim udtSystemTime As SYSTEMTIME
Dim fenxi() As String
If Text1.Text = "" Then
MsgBox "请选择要修改的文件！", vbCritical, "错误提示"
Text1.SetFocus
Exit Sub
End If
If Dir(Text1.Text, vbNormal Or vbHidden Or vbSystem Or vbReadOnly) = "" Then
MsgBox "说选的文件不存在！", vbCritical, "错误提示"
Text1.SetFocus
Exit Sub
End If
lngHandle = CreateFile(Text1.Text, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
fenxi = Split(Text2.Text, " ")
If UBound(fenxi) < 5 Then
MsgBox "时间参数有错误请重新设置，参考：2008 8 8 8 8 8", vbCritical, "参数错误"
Exit Sub
End If
udtSystemTime.wYear = Val(fenxi(0)) '设年
udtSystemTime.wMonth = Val(fenxi(1)) '月"
udtSystemTime.wDay = Val(fenxi(2)) '日
udtSystemTime.wDayOfWeek = 0 '周
udtSystemTime.wHour = Val(fenxi(3)) '时
udtSystemTime.wMinute = Val(fenxi(4)) '分
udtSystemTime.wSecond = Val(fenxi(5)) '秒
udtSystemTime.wMilliseconds = 0 '毫秒
' 转换时间格式 ,不知道微软为什么要搞得这么麻烦
SystemTimeToFileTime udtSystemTime, udtLocalTime
' 再转换
LocalFileTimeToFileTime udtLocalTime, udtFileTime
fenxi = Split(Text3.Text, " ")
If UBound(fenxi) < 5 Then
MsgBox "时间参数有错误请重新设置，参考：2008 8 8 8 8 8", vbCritical, "参数错误"
Exit Sub
End If
udtSystemTime.wYear = Val(fenxi(0)) '设年
udtSystemTime.wMonth = Val(fenxi(1)) '月"
udtSystemTime.wDay = Val(fenxi(2)) '日
udtSystemTime.wDayOfWeek = 0 '周
udtSystemTime.wHour = Val(fenxi(3)) '时
udtSystemTime.wMinute = Val(fenxi(4)) '分
udtSystemTime.wSecond = Val(fenxi(5)) '秒
udtSystemTime.wMilliseconds = 0 '毫秒
' 转换时间格式 ,不知道微软为什么要搞得这么麻烦
SystemTimeToFileTime udtSystemTime, udtLocalTime
' 再转换
LocalFileTimeToFileTime udtLocalTime, xiutime
fenxi = Split(Text4.Text, " ")
If UBound(fenxi) < 5 Then
MsgBox "时间参数有错误请重新设置，参考：2008 8 8 8 8 8", vbCritical, "参数错误"
Exit Sub
End If
udtSystemTime.wYear = Val(fenxi(0)) '设年
udtSystemTime.wMonth = Val(fenxi(1)) '月"
udtSystemTime.wDay = Val(fenxi(2)) '日
udtSystemTime.wDayOfWeek = 0 '周
udtSystemTime.wHour = Val(fenxi(3)) '时
udtSystemTime.wMinute = Val(fenxi(4)) '分
udtSystemTime.wSecond = Val(fenxi(5)) '秒
udtSystemTime.wMilliseconds = 0 '毫秒
' 转换时间格式 ,不知道微软为什么要搞得这么麻烦
SystemTimeToFileTime udtSystemTime, udtLocalTime
' 再转换
LocalFileTimeToFileTime udtLocalTime, fangtime

If SetFileTime(lngHandle, udtFileTime, fangtime, xiutime) = 1 Then
MsgBox "文件时间修改成功！", vbInformation, "操作提示"
Call shuxin
Else
If MsgBox("文件时间修改失败，文件可能只读，是否修改文件的属性并重试？", vbInformation + vbYesNo, "错误提示") = vbYes Then
SetAttr Text1.Text, vbNormal
Call shuxin
Call Command2_Click
End If
End If
End Sub

Private Sub Command3_Click()
MsgBox "软件制作：心翔zsd（小翟个人工作室）" + vbCrLf + "软件作者：翟士丹" + vbCrLf + "联系邮箱：6520186zsd@163.com", vbInformation, "关于软件"
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Command5_Click()
Dim time
time = Now
time = Replace(time, "-", " ")
time = Replace(time, ":", " ")
Text2.Text = time
End Sub

Private Sub Text2_Change()
Text3.Text = Text2.Text
Text4.Text = Text2.Text
End Sub

Private Sub Text3_Change()
Text4.Text = Text3.Text
End Sub
Sub shuxin()
Dim lngHandle As Long
Dim chuangtime As FILETIME
Dim xiutime As FILETIME
Dim fangtime As FILETIME
Dim gettime As Boolean
Dim filesys As SYSTEMTIME
Dim gett As Long
lngHandle = CreateFile(Text1.Text, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
getime = GetFileTime(lngHandle, chuangtime, fangtime, xiutime)
If gettime = ture Then
gett = FileTimeToLocalFileTime(chuangtime, chuangtime)
gett = FileTimeToSystemTime(chuangtime, filesys)
Label3.Caption = filesys.wYear & "年" & filesys.wMonth & "月" & filesys.wDay & "日, " & filesys.wHour & ":" & filesys.wMinute & ":" & filesys.wSecond
Text2.Text = filesys.wYear & " " & filesys.wMonth & " " & filesys.wDay & " " & filesys.wHour & " " & filesys.wMinute & " " & filesys.wSecond

gett = FileTimeToLocalFileTime(xiutime, xiutime)
gett = FileTimeToSystemTime(xiutime, filesys)
Label4.Caption = filesys.wYear & "年" & filesys.wMonth & "月" & filesys.wDay & "日, " & filesys.wHour & ":" & filesys.wMinute & ":" & filesys.wSecond
Text3.Text = filesys.wYear & " " & filesys.wMonth & " " & filesys.wDay & " " & filesys.wHour & " " & filesys.wMinute & " " & filesys.wSecond

gett = FileTimeToLocalFileTime(fangtime, fangtime)
gett = FileTimeToSystemTime(fangtime, filesys)
Label5.Caption = filesys.wYear & "年" & filesys.wMonth & "月" & filesys.wDay & "日, " & filesys.wHour & ":" & filesys.wMinute & ":" & filesys.wSecond
Text4.Text = filesys.wYear & " " & filesys.wMonth & " " & filesys.wDay & " " & filesys.wHour & " " & filesys.wMinute & " " & filesys.wSecond

Else

End If

End Sub
