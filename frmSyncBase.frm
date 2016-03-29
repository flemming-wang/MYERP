VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmSyncBase 
   Caption         =   "基础档案同步"
   ClientHeight    =   9495
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19470
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   19470
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   10695
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   10695
   End
   Begin VB.Frame fraLog 
      Caption         =   "同步日志"
      Height          =   2295
      Left            =   6360
      TabIndex        =   2
      Top             =   6960
      Width           =   12735
      Begin VB.CommandButton cmdLogTimeE 
         Caption         =   "..."
         Height          =   285
         Left            =   9600
         TabIndex        =   32
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdLogTimeS 
         Caption         =   "..."
         Height          =   285
         Left            =   7440
         TabIndex        =   31
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtLogTimeE 
         Height          =   285
         Left            =   8160
         TabIndex        =   30
         Top             =   360
         Width           =   1500
      End
      Begin VB.TextBox txtLogTimeS 
         Height          =   285
         Left            =   6000
         TabIndex        =   28
         Top             =   360
         Width           =   1500
      End
      Begin VB.TextBox txtLogBase 
         Height          =   285
         Left            =   3480
         TabIndex        =   26
         Top             =   360
         Width           =   1500
      End
      Begin VB.TextBox txtLogPlan 
         Height          =   285
         Left            =   960
         TabIndex        =   24
         Top             =   360
         Width           =   1500
      End
      Begin VSFlex8Ctl.VSFlexGrid gridLog 
         Height          =   1095
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   10935
         _cx             =   19288
         _cy             =   1931
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.CommandButton cmdLogQuery 
         Caption         =   "查询"
         Height          =   285
         Left            =   10080
         TabIndex        =   33
         Top             =   360
         Width           =   800
      End
      Begin VB.Label lblLogTimeTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   195
         Left            =   7920
         TabIndex        =   29
         Top             =   360
         Width           =   180
      End
      Begin VB.Label lblLogTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "同步时间"
         Height          =   195
         Left            =   5160
         TabIndex        =   27
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lblLogBase 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "档案编码"
         Height          =   195
         Left            =   2640
         TabIndex        =   25
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lblLogPlan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "方案编码"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.Frame fraBase 
      Caption         =   "同步数据"
      Height          =   6135
      Left            =   6360
      TabIndex        =   1
      Top             =   720
      Width           =   12735
      Begin VB.CommandButton cmdSynchronize 
         Caption         =   "同步"
         Height          =   285
         Left            =   11040
         TabIndex        =   34
         Top             =   360
         Width           =   800
      End
      Begin VB.CommandButton cmdBaseTimeE 
         Caption         =   "..."
         Height          =   285
         Left            =   9600
         TabIndex        =   22
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdBaseTimeS 
         Caption         =   "..."
         Height          =   285
         Left            =   7440
         TabIndex        =   21
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdBaseQuery 
         Caption         =   "查询"
         Height          =   285
         Left            =   10080
         TabIndex        =   20
         Top             =   360
         Width           =   800
      End
      Begin VB.TextBox txtBaseTimeE 
         Height          =   285
         Left            =   8160
         TabIndex        =   19
         Top             =   360
         Width           =   1500
      End
      Begin VB.TextBox txtBaseTimeS 
         Height          =   285
         Left            =   6000
         TabIndex        =   17
         Top             =   360
         Width           =   1500
      End
      Begin VB.TextBox txtBaseName 
         Height          =   285
         Left            =   3480
         TabIndex        =   15
         Top             =   360
         Width           =   1500
      End
      Begin VB.TextBox txtBaseCode 
         Height          =   285
         Left            =   960
         TabIndex        =   13
         Top             =   360
         Width           =   1500
      End
      Begin VSFlex8Ctl.VSFlexGrid gridBase 
         Height          =   3855
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   10935
         _cx             =   19288
         _cy             =   6800
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblBaseTimeTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   195
         Left            =   7920
         TabIndex        =   18
         Top             =   360
         Width           =   180
      End
      Begin VB.Label lblBaseTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "创建时间"
         Height          =   195
         Left            =   5160
         TabIndex        =   16
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lblBaseName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "档案名称"
         Height          =   195
         Left            =   2640
         TabIndex        =   14
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lblBaseCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "档案编码"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.Frame fraPlan 
      Caption         =   "同步方案"
      Height          =   8535
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6135
      Begin VB.CommandButton cmdPlanQuery 
         Caption         =   "查询"
         Height          =   285
         Left            =   4440
         TabIndex        =   11
         Top             =   360
         Width           =   800
      End
      Begin VB.TextBox txtPlanName 
         Height          =   285
         Left            =   3120
         TabIndex        =   10
         Top             =   360
         Width           =   1200
      End
      Begin VB.TextBox txtPlanCode 
         Height          =   285
         Left            =   960
         TabIndex        =   8
         Top             =   360
         Width           =   1200
      End
      Begin VSFlex8Ctl.VSFlexGrid gridPlan 
         Height          =   6255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   5895
         _cx             =   10398
         _cy             =   11033
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblPlanName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "方案名称"
         Height          =   195
         Left            =   2280
         TabIndex        =   9
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lblPlanCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "方案编码"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmSyncBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iSelBaseDataCnt As Integer  '选择的档案数
Dim bSynchroning As Boolean     '正在同步数据标识
Dim m_cFrmID As String          '窗体代码

Private Sub cmdBaseQuery_Click()
    InitBaseGrid
    FillBaseGrid
End Sub

Private Sub cmdBaseTimeE_Click()
    Dim objClendar As Object
    Set objClendar = CreateObject("CalendarAPP.ICaleCom")
    txtBaseTimeE.Text = Format(objClendar.Calendar(txtBaseTimeE.hWnd), "YYYY-MM-DD")
    Set objClendar = Nothing
End Sub

Private Sub cmdBaseTimeS_Click()
    Dim objClendar As Object
    Set objClendar = CreateObject("CalendarAPP.ICaleCom")
    txtBaseTimeS.Text = Format(objClendar.Calendar(txtBaseTimeS.hWnd), "YYYY-MM-DD")
    Set objClendar = Nothing
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdLogQuery_Click()
    InitLogGrid
    FillLogGrid
End Sub

Private Sub cmdLogTimeE_Click()
    Dim objClendar As Object
    Set objClendar = CreateObject("CalendarAPP.ICaleCom")
    txtLogTimeE.Text = Format(objClendar.Calendar(txtLogTimeE.hWnd), "YYYY-MM-DD")
    Set objClendar = Nothing
End Sub

Private Sub cmdLogTimeS_Click()
    Dim objClendar As Object
    Set objClendar = CreateObject("CalendarAPP.ICaleCom")
    txtLogTimeS.Text = Format(objClendar.Calendar(txtLogTimeS.hWnd), "YYYY-MM-DD")
    Set objClendar = Nothing
End Sub

Private Sub cmdPlanQuery_Click()
    InitPlanGrid
    FillPlanGrid
End Sub

Private Sub cmdSynchronize_Click()
    If iSelBaseDataCnt > 0 And gridPlan.RowSel > 0 Then
        bSynchroning = True
    Else
        Exit Sub
    End If
'
    Dim i As Integer
    Dim sql As String
    Dim err As String
    Dim rs As ADODB.Recordset

    g_cComputer = ComputerName
    TabBaseSet = Replace("TempSyncData_" & g_cComputer & "_TempBaseSet", "-", "_")
    TabBaseData = Replace("TempSyncData_" & g_cComputer & "_TempBaseData", "-", "_")
    TabBaseLog = Replace("TempSyncData_" & g_cComputer & "_TempBaseLog", "-", "_")

    DropTempTbl

    sql = ""
    sql = " CREATE TABLE " & TabBaseSet & " ( planid int) "
    sql = sql & " CREATE TABLE " & TabBaseData & " ( basecode nvarchar(50)) "
    
    '创建临时表
    If ExtSql(sql) Then
        sql = ""
        
        '每次只能选择一个方案
        i = gridPlan.RowSel

        sql = sql & "insert into " & TabBaseSet & "(planid) "
        sql = sql & " values ("
        If gridPlan.ColIndex("ID") > -1 Then
            sql = sql & "'" & gridPlan.TextMatrix(i, gridPlan.ColIndex("ID")) & " '"
        Else
            sql = sql & "''"
        End If
        sql = sql & ") " & vbCrLf

        For i = 1 To gridBase.Rows - 1
            If gridBase.Cell(flexcpChecked, i, 0) = flexChecked Then
                sql = sql & "insert into " & TabBaseData & "(basecode) "
                sql = sql & " values ("
                If gridBase.ColIndex("basecode") > -1 Then
                    sql = sql & "'" & gridBase.TextMatrix(i, gridBase.ColIndex("basecode")) & " '"
                Else
                    sql = sql & "''"
                End If
                sql = sql & ") " & vbCrLf
            End If
        Next

        If Not IsBlank(sql) Then
            If ExtSql(sql) Then
                sql = ""
                err = ""

                Call ExtBaseSyncPro(err)
                DoEvents
                If Not IsBlank(err) Then
                    MsgBox err, vbOKOnly, "基础档案同步"
                Else
                    MsgBox "同步成功", vbOKOnly, "基础档案同步"
                    InitLogGrid
                    ClearLogCondition
                    txtLogPlan.Text = gridPlan.TextMatrix(gridPlan.RowSel, gridPlan.ColIndex("cPlanCode"))
                    FillLogGrid True
                End If
            End If
        End If
    End If

    bSynchroning = False
End Sub

Private Sub Form_Load()
    Me.WindowState = vbMaximized
    InitCondition
    InitGrid
    FillGrid
End Sub

Private Sub InitCondition()
    ClearPlanCondition
    ClearBaseCondition
    ClearLogCondition
End Sub

Private Sub ClearPlanCondition()
    txtPlanCode.Text = ""
    txtPlanName.Text = ""
End Sub

Private Sub ClearBaseCondition()
    txtBaseCode.Text = ""
    txtBaseName.Text = ""
    
    '当月第一天
    txtBaseTimeS.Text = Format(DateSerial(Year(Now), Month(Now), 1), "YYYY-MM-DD")
    
    '下个月第一天减1天
    If Month(Now) = 12 Then
        txtBaseTimeE.Text = Format(DateSerial(Year(Now), Month(Now), 31), "YYYY-MM-DD")
    Else
        txtBaseTimeE.Text = Format(DateAdd("D", -1, DateSerial(Year(Now), Month(Now) + 1, 1)), "YYYY-MM-DD")
    End If
End Sub

Private Sub ClearLogCondition()
    txtLogPlan.Text = ""
    txtLogPlan.Text = ""
    
    '当月第一天
    txtLogTimeS.Text = Format(DateSerial(Year(Now), Month(Now), 1), "YYYY-MM-DD")
    
    '下个月第一天减1天
    If Month(Now) = 12 Then
        txtLogTimeE.Text = Format(DateSerial(Year(Now), Month(Now), 31), "YYYY-MM-DD")
    Else
        txtLogTimeE.Text = Format(DateAdd("D", -1, DateSerial(Year(Now), Month(Now) + 1, 1)), "YYYY-MM-DD")
    End If
End Sub

Private Sub InitGrid()
    InitPlanGrid
    InitBaseGrid
    InitLogGrid
End Sub

Private Sub InitPlanGrid()
    Dim i As Integer
    With gridPlan
        .AllowBigSelection = False  '不能点击左上角
        .AllowSelection = False     '不能多选
        .AllowUserResizing = flexResizeColumns  '可拖动列宽
        .Clear
        Set .DataSource = Nothing
        
        .Rows = 1
        .Cols = 10
        .TextMatrix(0, 0) = "行号"
        .TextMatrix(0, 1) = "方案标识"
        .TextMatrix(0, 2) = "方案编码"
        .TextMatrix(0, 3) = "方案名称"
        .TextMatrix(0, 4) = "源表账套"
        .TextMatrix(0, 5) = "目的账套"
        .TextMatrix(0, 6) = "源表名"
        .TextMatrix(0, 7) = "目的表名"
        .TextMatrix(0, 8) = "源大类"
        .TextMatrix(0, 9) = "目的大类"
        .ColWidth(0) = 500
        .ColWidth(1) = 0
        .ColWidth(2) = 1500
        .ColWidth(3) = 3000
        .ColWidth(4) = 0
        .ColWidth(5) = 0
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        
        'ID,cPlanCode,cPlanName,cAccID,cAccIDP,cTabName,cTabNameP,cType,cTypeP
        .ColKey(0) = "rowno"
        .ColKey(1) = "ID"
        .ColKey(2) = "cPlanCode"
        .ColKey(3) = "cPlanName"
        .ColKey(4) = "cAccID"
        .ColKey(5) = "cAccIDP"
        .ColKey(6) = "cTabName"
        .ColKey(7) = "cTabNameP"
        .ColKey(8) = "cType"
        .ColKey(9) = "cTypeP"
        
        '行号居中
        .ColAlignment(0) = flexAlignCenterCenter
        
        For i = 0 To .Cols - 1
            '默认都设置成文本类型
            .ColDataType(i) = flexcpText
        Next
        
    End With
End Sub

Private Sub InitBaseGrid()
    Dim i As Integer
    With gridBase
        .Clear
        Set .DataSource = Nothing
        .AutoResize = True  '自适应列宽
        .AllowUserResizing = flexResizeColumns  '可拖动列宽
        
        .Rows = 1
        '选择框、行号居中
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        
        For i = 0 To .Cols - 1
            '默认都设置成文本类型
            .ColDataType(i) = flexcpText
        Next
        
        .Cell(flexcpChecked, 0, 0) = flexNoCheckbox
        iSelBaseDataCnt = 0
    End With
End Sub

Private Sub InitLogGrid()
    Dim i As Integer
    With gridLog
        .Clear
        Set .DataSource = Nothing
        .Rows = 1
        .Cols = 11
        .TextMatrix(0, 0) = "行号"
        .TextMatrix(0, 1) = "方案编码"
        .TextMatrix(0, 2) = "源账套"
        .TextMatrix(0, 3) = "源ID"
        .TextMatrix(0, 4) = "源编码"
        .TextMatrix(0, 5) = "源时间戳"
        .TextMatrix(0, 6) = "目的账套"
        .TextMatrix(0, 7) = "目的ID"
        .TextMatrix(0, 8) = "目的编码"
        .TextMatrix(0, 9) = "目的时间戳"
        .TextMatrix(0, 10) = "生成日期"
        .ColWidth(0) = 500
        .ColWidth(1) = 1500
        .ColWidth(2) = 1500
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColWidth(5) = 1500
        .ColWidth(6) = 1500
        .ColWidth(7) = 1000
        .ColWidth(8) = 1000
        .ColWidth(9) = 1500
        .ColWidth(10) = 2000
        
        '行号居中
        .ColAlignment(0) = flexAlignCenterCenter
        
        For i = 0 To .Cols - 1
            '默认都设置成文本类型
            .ColDataType(i) = flexcpText
        Next
        
    End With
End Sub

Private Sub FillGrid()
    FillPlanGrid
    FillBaseGrid
End Sub

Private Sub FillPlanGrid(Optional bClearCondition As Boolean = False)
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim sPlanCode As String
    Dim sPlanName As String
    
    If bClearCondition Then ClearPlanCondition
    
    sql = "select ID,cPlanCode,cPlanName,cAccID,cAccIDP,cTabName,cTabNameP,cType,cTypeP from JT_BaseSet where 1=1 "
    
    '方案编码
    sPlanCode = txtPlanCode.Text
    If Not IsBlank(sPlanCode) Then
        sql = sql & " and cPlanCode like '%" & sPlanCode & "%'"
    End If
    
    '方案名称
    sPlanName = txtPlanName.Text
    If Not IsBlank(sPlanName) Then
        sql = sql & " and cPlanName like '%" & sPlanName & "%'"
    End If
    
    With gridPlan
        Set rs = QueryExt(sql)
        If Not rs.BOF And Not rs.EOF Then
            Do While Not rs.EOF
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = .Rows - 1
                .TextMatrix(.Rows - 1, 1) = GetNoNullValue(rs!ID)
                .TextMatrix(.Rows - 1, 2) = GetNoNullValue(rs!cPlanCode)
                .TextMatrix(.Rows - 1, 3) = GetNoNullValue(rs!cPlanName)
                .TextMatrix(.Rows - 1, 4) = GetNoNullValue(rs!cAccID)
                .TextMatrix(.Rows - 1, 5) = GetNoNullValue(rs!cAccIDP)
                .TextMatrix(.Rows - 1, 6) = GetNoNullValue(rs!cTabName)
                .TextMatrix(.Rows - 1, 7) = GetNoNullValue(rs!cTabNameP)
                .TextMatrix(.Rows - 1, 8) = GetNoNullValue(rs!cType)
                .TextMatrix(.Rows - 1, 9) = GetNoNullValue(rs!cTypeP)
                rs.MoveNext
            Loop
            '默认选择第一行
            gridPlan.RowSel = 1
            InitBaseGrid
            FillBaseGrid (True)
            InitLogGrid
            ClearLogCondition
            txtLogPlan.Text = gridPlan.TextMatrix(gridPlan.RowSel, gridPlan.ColIndex("cPlanCode"))
            'FillLogGrid
        End If
    End With
End Sub

Private Sub FillBaseGrid(Optional bClearCondition As Boolean = False)
    Dim i As Integer
    Dim sql As String
    Dim sBaseCode As String
    Dim sBaseName As String
    Dim dateFrom As String
    Dim dateTo As String
    Dim sSrcType As String
    Dim sSrcTabName As String
    Dim sDesTabName As String
    Dim sPlanCode As String     '方案编号
    Dim sSrcAccId As String     '源账套
    Dim sSrcAccIdP As String    '目的账套
    
    If bClearCondition Then ClearBaseCondition
    
    sql = ""
    sBaseCode = txtBaseCode.Text
    sBaseName = txtBaseName.Text
    dateFrom = txtBaseTimeS.Text
    dateTo = txtBaseTimeE.Text
    
    If Not IsBlank(dateFrom) And Not IsDate(dateFrom) Then
        MsgBox "开始时间格式错误！", vbOKOnly, "基础档案同步"
        Exit Sub
    End If
    
    If Not IsBlank(dateTo) And Not IsDate(dateTo) Then
        MsgBox "结束时间格式错误！", vbOKOnly, "基础档案同步"
        Exit Sub
    End If
    
    i = gridPlan.RowSel
    If i < 1 Then Exit Sub  '未选择方案
    sSrcType = gridPlan.TextMatrix(i, gridPlan.ColIndex("cType"))
    sSrcTabName = gridPlan.TextMatrix(i, gridPlan.ColIndex("cTabName"))
    sDesTabName = gridPlan.TextMatrix(i, gridPlan.ColIndex("cTabNameP"))
    sPlanCode = gridPlan.TextMatrix(i, gridPlan.ColIndex("cPlanCode"))
    sSrcAccId = gridPlan.TextMatrix(i, gridPlan.ColIndex("cAccId"))
    sSrcAccIdP = gridPlan.TextMatrix(i, gridPlan.ColIndex("cAccIdP"))
    
    With gridBase

        Dim rs As ADODB.Recordset
        Select Case LCase(sSrcTabName & "-" & sDesTabName)
            Case "b_department-department"
                sql = "select '' as 行号,a.cdepcode as basecode,a.* from B_Department a where 1=1" 'dDepBeginDate between '" & dateFrom & "' and '" & dateTo & "'"
                If Not IsBlank(sBaseCode) Then sql = sql & " and cdepcode like '%" & sBaseCode & "%'"
                If Not IsBlank(sBaseName) Then sql = sql & " and (cdepname like '%" & sBaseName & "%' or cdepfullname like '%" & sBaseName & "%')"
                
                '过滤掉目标账套已存在，且时间戳没变的几率
                sql = sql & "  and (not exists (select top 1 * from " & sSrcAccIdP & ".dbo." & sDesTabName & " b where b.cdepcode = a.cdepcode)"
                sql = sql & "       or"
                sql = sql & "       exists (select top 1 * from " & sSrcAccId & ".dbo.JT_BaseLog c where c.ccode = a.cdepcode and isnull(c.myufts,'') <> convert(nchar,convert(money,a.myufts),2) and c.cplancode='" & sPlanCode & "'and c.cAccID = '" & sSrcAccId & "' and c.cAccIDP = '" & sSrcAccIdP & "')"
                sql = sql & "       )"

            Case "b_warehouse-warehouse"
                sql = "select '' as 行号,a.cwhcode as basecode,a.* from B_Warehouse a where 1=1" 'dModifyDate between '" & dateFrom & "' and '" & dateTo & "'"
                If Not IsBlank(sBaseCode) Then sql = sql & " and cwhcode like '%" & sBaseCode & "%'"
                If Not IsBlank(sBaseName) Then sql = sql & " and cwhname like '%" & sBaseName & "%'"
                
                sql = sql & "  and (not exists (select top 1 * from " & sSrcAccIdP & ".dbo." & sDesTabName & " b where b.cwhcode = a.cwhcode)"
                sql = sql & "       or"
                sql = sql & "       exists (select top 1 * from " & sSrcAccId & ".dbo.JT_BaseLog c where c.ccode = a.cwhcode and isnull(c.myufts,'') <> convert(nchar,convert(money,a.myufts),2) and c.cplancode='" & sPlanCode & "'and c.cAccID = '" & sSrcAccId & "' and c.cAccIDP = '" & sSrcAccIdP & "')"
                sql = sql & "       )"
            
            Case "b_customer-customer"
                sql = "select '' as 行号,a.ccuscode as basecode,a.* from B_Customer a where 1=1" 'dCusCreateDatetime between '" & dateFrom & "' and '" & dateTo & "'"
                If Not IsBlank(sBaseCode) Then sql = sql & " and ccuscode like '%" & sBaseCode & "%'"
                If Not IsBlank(sBaseName) Then sql = sql & " and (ccusname like '%" & sBaseName & "%' or ccusabbname like '%" & sBaseName & "%')"
                '客户分类 符合赫夫曼编码规则** *** ****
                If Not IsBlank(sSrcType) Then sql = sql & "  and ccccode like '" & sSrcType & "%'"
                
                sql = sql & "  and (not exists (select top 1 * from " & sSrcAccIdP & ".dbo." & sDesTabName & " b where b.ccuscode = a.ccuscode)"
                sql = sql & "       or"
                sql = sql & "       exists (select top 1 * from " & sSrcAccId & ".dbo.JT_BaseLog c where c.ccode = a.ccuscode and isnull(c.myufts,'') <> convert(nchar,convert(money,a.myufts),2) and c.cplancode='" & sPlanCode & "'and c.cAccID = '" & sSrcAccId & "' and c.cAccIDP = '" & sSrcAccIdP & "')"
                sql = sql & "       )"
                
            Case "b_vendor-vendor"
                sql = "select '' as 行号,a.cvencode as basecode,a.* from B_Vendor a where 1=1" 'dVenCreateDatetime between '" & dateFrom & "' and '" & dateTo & "'"
                If Not IsBlank(sBaseCode) Then sql = sql & " and cvencode like '%" & sBaseCode & "%'"
                If Not IsBlank(sBaseName) Then sql = sql & " and (cvenname like '%" & sBaseName & "%' or cvenabbname like '%" & sBaseName & "%')"
                '供应商分类 符合赫夫曼编码规则** *** ****
                If Not IsBlank(sSrcType) Then sql = sql & "  and cvccode like '" & sSrcType & "%'"
                
                sql = sql & "  and (not exists (select top 1 * from " & sSrcAccIdP & ".dbo." & sDesTabName & " b where b.cvencode = a.cvencode)"
                sql = sql & "       or"
                sql = sql & "       exists (select top 1 * from " & sSrcAccId & ".dbo.JT_BaseLog c where c.ccode = a.cvencode and isnull(c.myufts,'') <> convert(nchar,convert(money,a.myufts),2) and c.cplancode='" & sPlanCode & "'and c.cAccID = '" & sSrcAccId & "' and c.cAccIDP = '" & sSrcAccIdP & "')"
                sql = sql & "       )"
                
            Case "b_inventory-inventory"
                sql = "select '' as 行号,a.cinvcode as basecode,a.* from b_inventory a where 1=1" 'dsdate between '" & dateFrom & "' and '" & dateTo & "'"
                If Not IsBlank(sBaseCode) Then sql = sql & " and cinvcode like '%" & sBaseCode & "%'"
                If Not IsBlank(sBaseName) Then sql = sql & " and cinvname like '%" & sBaseName & "%'"
                '客户分类 符合赫夫曼编码规则
                If Not IsBlank(sSrcType) Then sql = sql & "  and cinvccode like '" & sSrcType & "%'"
                
                sql = sql & "  and (not exists (select top 1 * from " & sSrcAccIdP & ".dbo." & sDesTabName & " b where b.cinvcode = a.cinvcode)"
                sql = sql & "       or"
                sql = sql & "       exists (select top 1 * from " & sSrcAccId & ".dbo.JT_BaseLog c where c.ccode = a.cinvcode and isnull(c.myufts,'') <> convert(nchar,convert(money,a.myufts),2) and c.cplancode='" & sPlanCode & "'and c.cAccID = '" & sSrcAccId & "' and c.cAccIDP = '" & sSrcAccIdP & "')"
                sql = sql & "       )"
            
            Case ""
                'no Plan do nothing
        End Select
        If Not IsBlank(sql) Then
            '五大档案之外的数据，直接绑定rs到DataSource
            Set rs = QueryExt(sql)
            Set .DataSource = rs

            If .Rows > 1 Then
                .Cell(flexcpChecked, 0, 0) = flexChecked
                For i = 1 To .Rows - 1
                    .Cell(flexcpChecked, i, 0) = flexChecked
                    .TextMatrix(i, 1) = i
                Next
            End If
            
            .ColWidth(0) = 250
            .ColWidth(1) = 500
            .ColWidth(2) = 0 '隐藏一个basecode列，统一名称，方便取数
            
            '选择框、行号居中
            .ColAlignment(0) = flexAlignCenterCenter
            .ColAlignment(1) = flexAlignCenterCenter

            '默认都选中
            iSelBaseDataCnt = .Rows - 1
        End If
    End With
End Sub

'bShowUpdateOnly 是否只显示本次同步日志
Private Sub FillLogGrid(Optional bShowUpdateOnly As Boolean = False)
    Dim rs As ADODB.Recordset
    
    Dim sql As String
    Dim sPlanCode As String
    Dim sBaseCode As String
    Dim dateFrom As String
    Dim dateTo As String
    
    sql = ""
    If bShowUpdateOnly Then
        '显示本次同步日志，从临时表查
        sql = "select * from " & TabBaseLog & " where 1=1 "
    Else
        sPlanCode = txtLogPlan.Text
        sBaseCode = txtLogBase.Text
        dateFrom = txtLogTimeS.Text
        dateTo = txtLogTimeE.Text
        
        If Not IsBlank(dateFrom) And Not IsDate(dateFrom) Then
            MsgBox "开始时间格式错误！", vbOKOnly, "基础档案同步"
            Exit Sub
        End If
        
        If Not IsBlank(dateTo) And Not IsDate(dateTo) Then
            MsgBox "结束时间格式错误！", vbOKOnly, "基础档案同步"
            Exit Sub
        End If
        
        sql = "select * from JT_BaseLog where 1=1 "
        
        If Not IsBlank(sPlanCode) Then sql = sql & " and cplancode like '%" & sPlanCode & "%'"
        If Not IsBlank(sBaseCode) Then sql = sql & " and (ccode like '%" & sBaseCode & "%' or ccodep like '%" & sBaseCode & "%')"
        
        If Not IsBlank(dateFrom) And Not IsBlank(dateTo) Then
            sql = sql & " and dcdate between '" & dateFrom & "' and '" & dateTo & "'"
        ElseIf IsBlank(dateFrom) Then
            sql = sql & " and dcdate >= '" & dateFrom & "'"
        ElseIf IsBlank(dateTo) Then
            sql = sql & " and dcdate <= '" & dateTo & "'"
        End If
    End If
    
    Set rs = QueryExt(sql)
    'Set gridLog.DataSource = rs
    
    With gridLog
        If Not rs.BOF And Not rs.EOF Then
            Do While Not rs.EOF
                .Rows = .Rows + 1

                .TextMatrix(.Rows - 1, 0) = .Rows - 1
                .TextMatrix(.Rows - 1, 1) = GetNoNullValue(rs!cPlanCode)
                .TextMatrix(.Rows - 1, 2) = GetNoNullValue(rs!cAccID)
                .TextMatrix(.Rows - 1, 3) = GetNoNullValue(rs!iID)
                .TextMatrix(.Rows - 1, 4) = GetNoNullValue(rs!cCode)
                .TextMatrix(.Rows - 1, 5) = GetNoNullValue(rs!myufts)
                .TextMatrix(.Rows - 1, 6) = GetNoNullValue(rs!cAccIDP)
                .TextMatrix(.Rows - 1, 7) = GetNoNullValue(rs!IDP)
                .TextMatrix(.Rows - 1, 8) = GetNoNullValue(rs!cCodeP)
                .TextMatrix(.Rows - 1, 9) = GetNoNullValue(rs!myuftsP)
                .TextMatrix(.Rows - 1, 10) = GetNoNullValue(rs!dCDate)

                rs.MoveNext
            Loop
        End If
    End With
End Sub

Private Sub Form_Resize()
    '按钮
    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If
    picToolbar.Top = 100
    picToolbar.Left = 100
    picToolbar.Width = Me.ScaleWidth - 200
    picToolbar.Height = Me.ScaleHeight * 4 / 100 '占屏幕高度的4%
    
    
    '同步方案
    fraPlan.Left = picToolbar.Left
    fraPlan.Top = picToolbar.Top + picToolbar.Height
    fraPlan.Width = Me.ScaleWidth * 3 / 10  '占屏幕宽度的25%
    fraPlan.Height = Me.ScaleHeight * 95 / 100  '占屏幕高度的95%
    
    lblPlanCode.Left = 200
    lblPlanCode.Top = 400
    txtPlanCode.Left = lblPlanCode.Left + lblPlanCode.Width + 100
    txtPlanCode.Top = lblPlanCode.Top

    lblPlanName.Left = txtPlanCode.Left + txtPlanCode.Width + 100
    lblPlanName.Top = txtPlanCode.Top
    txtPlanName.Left = lblPlanName.Left + lblPlanName.Width + 100
    txtPlanName.Top = lblPlanName.Top
    
    cmdPlanQuery.Left = txtPlanName.Left + txtPlanName.Width + 100
    cmdPlanQuery.Top = txtPlanName.Top
    
    gridPlan.Top = lblPlanCode.Top + lblPlanCode.Height + 200
    gridPlan.Left = lblPlanCode.Left
    gridPlan.Width = fraPlan.Width - 400
    gridPlan.Height = fraPlan.Height - gridPlan.Top - 200
    
    
    '同步数据
    fraBase.Left = fraPlan.Left + fraPlan.Width + 100
    fraBase.Top = fraPlan.Top
    fraBase.Width = Me.ScaleWidth * 7 / 10 - 300 '占屏幕宽度的75%
    fraBase.Height = fraPlan.Height * 3 / 5 '右侧高度的60%
    
    lblBaseCode.Left = 200
    lblBaseCode.Top = 400
    txtBaseCode.Left = lblPlanCode.Left + lblPlanCode.Width + 100
    txtBaseCode.Top = lblPlanCode.Top

    lblBaseName.Left = txtBaseCode.Left + txtBaseCode.Width + 100
    lblBaseName.Top = txtBaseCode.Top
    txtBaseName.Left = lblBaseName.Left + lblBaseName.Width + 100
    txtBaseName.Top = lblBaseName.Top
    
    lblBaseTime.Left = txtBaseName.Left + txtBaseName.Width + 100
    lblBaseTime.Top = txtBaseName.Top
    
    txtBaseTimeS.Left = lblBaseTime.Left + lblBaseTime.Width + 100
    txtBaseTimeS.Top = lblBaseTime.Top
    
    cmdBaseTimeS.Left = txtBaseTimeS.Left + txtBaseTimeS.Width - cmdBaseTimeS.Width
    cmdBaseTimeS.Top = txtBaseTimeS.Top
    cmdBaseTimeS.Height = txtBaseTimeS.Height
    
    lblBaseTimeTo.Left = cmdBaseTimeS.Left + cmdBaseTimeS.Width + 100
    lblBaseTimeTo.Top = cmdBaseTimeS.Top
    
    txtBaseTimeE.Left = lblBaseTimeTo.Left + lblBaseTimeTo.Width + 100
    txtBaseTimeE.Top = lblBaseTimeTo.Top
    
    cmdBaseTimeE.Left = txtBaseTimeE.Left + txtBaseTimeE.Width - cmdBaseTimeE.Width
    cmdBaseTimeE.Top = txtBaseTimeE.Top
    cmdBaseTimeE.Height = txtBaseTimeE.Height
    
    cmdBaseQuery.Left = cmdBaseTimeE.Left + cmdBaseTimeE.Width + 100
    cmdBaseQuery.Top = cmdBaseTimeE.Top
    
    cmdSynchronize.Left = cmdBaseQuery.Left + cmdBaseQuery.Width + 100
    cmdSynchronize.Top = cmdBaseQuery.Top
    
    gridBase.Top = lblBaseCode.Top + lblBaseCode.Height + 200
    gridBase.Left = lblBaseCode.Left
    gridBase.Width = fraBase.Width - 400
    gridBase.Height = fraBase.Height - gridBase.Top - 200
    
    
    '同步历史
    fraLog.Left = fraPlan.Left + fraPlan.Width + 100
    fraLog.Top = fraBase.Top + fraBase.Height + 100
    fraLog.Width = fraBase.Width
    fraLog.Height = fraPlan.Height * 2 / 5 - 100 '右侧高度的40%
    
    lblLogPlan.Left = 200
    lblLogPlan.Top = 400
    txtLogPlan.Left = lblLogPlan.Left + lblLogPlan.Width + 100
    txtLogPlan.Top = lblLogPlan.Top

    lblLogBase.Left = txtLogPlan.Left + txtLogPlan.Width + 100
    lblLogBase.Top = txtLogPlan.Top
    txtLogBase.Left = lblLogBase.Left + lblLogBase.Width + 100
    txtLogBase.Top = lblLogBase.Top
    
    lblLogTime.Left = txtLogBase.Left + txtLogBase.Width + 100
    lblLogTime.Top = txtLogBase.Top
    
    txtLogTimeS.Left = lblLogTime.Left + lblLogTime.Width + 100
    txtLogTimeS.Top = lblLogTime.Top
    
    cmdLogTimeS.Left = txtLogTimeS.Left + txtLogTimeS.Width - cmdLogTimeS.Width
    cmdLogTimeS.Top = txtLogTimeS.Top
    cmdLogTimeS.Height = txtLogTimeS.Height
    
    lblLogTimeTo.Left = cmdLogTimeS.Left + cmdLogTimeS.Width + 100
    lblLogTimeTo.Top = cmdLogTimeS.Top
    
    txtLogTimeE.Left = lblLogTimeTo.Left + lblLogTimeTo.Width + 100
    txtLogTimeE.Top = lblLogTimeTo.Top
    
    cmdLogTimeE.Left = txtLogTimeE.Left + txtLogTimeE.Width - cmdLogTimeE.Width
    cmdLogTimeE.Top = txtLogTimeE.Top
    cmdLogTimeE.Height = txtLogTimeE.Height
    
    cmdLogQuery.Left = cmdLogTimeE.Left + cmdLogTimeE.Width + 100
    cmdLogQuery.Top = cmdLogTimeE.Top
    
    gridLog.Top = lblLogPlan.Top + lblLogPlan.Height + 200
    gridLog.Left = lblLogPlan.Left
    gridLog.Width = fraLog.Width - 400
    gridLog.Height = fraLog.Height - gridLog.Top - 200
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If bSynchroning Then
        MsgBox "正在同步数据，请稍后！", vbOKOnly, "数据同步"
        Cancel = True
        bSynchroning = False
    Else
        DropTempTbl
    End If
End Sub

Private Sub gridBase_Click()
    Dim i, j, k As Integer
    Dim bChecked As Boolean
    
    With gridBase
        If .Rows = 1 Then Exit Sub
        
        i = .MouseRow
        j = .MouseCol
    
        If i = -1 Or j = -1 Or i > .Rows - 1 Or j > .Cols - 1 Then Exit Sub
    
        bChecked = IIf(.Cell(flexcpChecked, i, j) = flexChecked, True, False)
        
        If j = 0 Then
            '全选/全消
            If i = 0 Then
                For k = 1 To .Rows - 1
                    .Cell(flexcpChecked, k, 0) = IIf(bChecked, flexUnchecked, flexChecked)
                Next
                .Cell(flexcpChecked, 0, 0) = IIf(bChecked, flexUnchecked, flexChecked)
                iSelBaseDataCnt = IIf(bChecked, 0, .Rows - 1)
            Else
                '单选
                .Cell(flexcpChecked, i, 0) = IIf(bChecked, flexUnchecked, flexChecked)
                iSelBaseDataCnt = IIf(bChecked, iSelBaseDataCnt - 1, iSelBaseDataCnt + 1)
                If iSelBaseDataCnt = .Rows - 1 Then
                    .Cell(flexcpChecked, 0, 0) = flexChecked
                Else
                    .Cell(flexcpChecked, 0, 0) = flexUnchecked
                End If
            End If
        End If
    End With
End Sub

Public Sub ExitForm(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub

Private Sub gridPlan_DblClick()
    Dim iSel As Integer
    If gridPlan.Rows = 1 Then Exit Sub
    InitBaseGrid
    FillBaseGrid (True)
    InitLogGrid
    ClearLogCondition
    txtLogPlan.Text = gridPlan.TextMatrix(gridPlan.RowSel, gridPlan.ColIndex("cPlanCode"))
    'FillLogGrid
End Sub

Public Property Get FrmID() As String
    FrmID = m_cFrmID
End Property

Public Property Let FrmID(ByVal RHS As String)
     m_cFrmID = RHS
End Property

