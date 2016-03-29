VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmSyncBusiness 
   Caption         =   "业务数据同步"
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
      Begin VB.TextBox txtLogBusiness 
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
      Begin VB.Label lblLogBusiness 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单据编号"
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
   Begin VB.Frame fraBusiness 
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
      Begin VB.CommandButton cmdBusinessTimeE 
         Caption         =   "..."
         Height          =   285
         Left            =   9600
         TabIndex        =   22
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdBusinessTimeS 
         Caption         =   "..."
         Height          =   285
         Left            =   7440
         TabIndex        =   21
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdBusinessQuery 
         Caption         =   "查询"
         Height          =   285
         Left            =   10080
         TabIndex        =   20
         Top             =   360
         Width           =   800
      End
      Begin VB.TextBox txtBusinessTimeE 
         Height          =   285
         Left            =   8160
         TabIndex        =   19
         Top             =   360
         Width           =   1500
      End
      Begin VB.TextBox txtBusinessTimeS 
         Height          =   285
         Left            =   6000
         TabIndex        =   17
         Top             =   360
         Width           =   1500
      End
      Begin VB.TextBox txtBusinessName 
         Height          =   285
         Left            =   3480
         TabIndex        =   15
         Top             =   360
         Width           =   1500
      End
      Begin VB.TextBox txtBusinessCode 
         Height          =   285
         Left            =   960
         TabIndex        =   13
         Top             =   360
         Width           =   1500
      End
      Begin VSFlex8Ctl.VSFlexGrid gridBusiness 
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
      Begin VB.Label lblBusinessTimeTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   195
         Left            =   7920
         TabIndex        =   18
         Top             =   360
         Width           =   180
      End
      Begin VB.Label lblBusinessTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "创建时间"
         Height          =   195
         Left            =   5160
         TabIndex        =   16
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lblBusinessName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "制单人"
         Height          =   195
         Left            =   2760
         TabIndex        =   14
         Top             =   360
         Width           =   540
      End
      Begin VB.Label lblBusinessCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单据编号"
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
Attribute VB_Name = "frmSyncBusiness"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iSelBusinessDataCnt As Integer  '选择的单据数
Dim bSynchroning As Boolean         '正在同步数据标识
Dim m_cFrmID As String              '窗体代码

Private Sub cmdBusinessQuery_Click()
    InitBusinessGrid
    FillBusinessGrid
End Sub

Private Sub cmdBusinessTimeE_Click()
    Dim objClendar As Object
    Set objClendar = CreateObject("CalendarAPP.ICaleCom")
    txtBusinessTimeE.Text = Format(objClendar.Calendar(txtBusinessTimeE.hWnd), "YYYY-MM-DD")
    Set objClendar = Nothing
End Sub

Private Sub cmdBusinessTimeS_Click()
    Dim objClendar As Object
    Set objClendar = CreateObject("CalendarAPP.ICaleCom")
    txtBusinessTimeS.Text = Format(objClendar.Calendar(txtBusinessTimeS.hWnd), "YYYY-MM-DD")
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
    If iSelBusinessDataCnt > 0 And gridPlan.RowSel > 0 Then
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
    TabBusinessSet = Replace("TempSyncData_" & g_cComputer & "_TempBusinessSet", "-", "_")
    TabBusinessData = Replace("TempSyncData_" & g_cComputer & "_TempBusinessData", "-", "_")
    TabBusinessLog = Replace("TempSyncData_" & g_cComputer & "_TempBusinessLog", "-", "_")

    DropTempTbl

    sql = ""
    sql = " CREATE TABLE " & TabBusinessSet & " ( planid int) "
    sql = sql & " CREATE TABLE " & TabBusinessData & " ( businessid bigint) "
    
    '创建临时表
    If ExtSql(sql) Then
        sql = ""
        
        '每次只能选择一个方案
        i = gridPlan.RowSel

        sql = sql & "insert into " & TabBusinessSet & "(planid) "
        sql = sql & " values ("
        If gridPlan.ColIndex("ID") > -1 Then
            sql = sql & "'" & gridPlan.TextMatrix(i, gridPlan.ColIndex("ID")) & " '"
        Else
            sql = sql & "''"
        End If
        sql = sql & ") " & vbCrLf

        For i = 1 To gridBusiness.Rows - 1
            If gridBusiness.Cell(flexcpChecked, i, 0) = flexChecked Then
                sql = sql & "insert into " & TabBusinessData & "(businessid) "
                sql = sql & " values ("
                If gridBusiness.ColIndex("businessid") > -1 Then
                    sql = sql & "'" & gridBusiness.TextMatrix(i, gridBusiness.ColIndex("businessid")) & " '"
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

                Call ExtBusinessSyncPro(err)
                DoEvents
                If Not IsBlank(err) Then
                    MsgBox err, vbOKOnly, "业务单据同步"
                Else
                    MsgBox "同步成功", vbOKOnly, "业务单据同步"
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
    ClearBusinessCondition
    ClearLogCondition
End Sub

Private Sub ClearPlanCondition()
    txtPlanCode.Text = ""
    txtPlanName.Text = ""
End Sub

Private Sub ClearBusinessCondition()
    txtBusinessCode.Text = ""
    txtBusinessName.Text = ""
    
    '当月第一天
    txtBusinessTimeS.Text = Format(DateSerial(Year(Now), Month(Now), 1), "YYYY-MM-DD")
    
    '下个月第一天减1天
    If Month(Now) = 12 Then
        txtBusinessTimeE.Text = Format(DateSerial(Year(Now), Month(Now), 31), "YYYY-MM-DD")
    Else
        txtBusinessTimeE.Text = Format(DateAdd("D", -1, DateSerial(Year(Now), Month(Now) + 1, 1)), "YYYY-MM-DD")
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
    InitBusinessGrid
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
        .Cols = 14
        .TextMatrix(0, 0) = "行号"
        .TextMatrix(0, 1) = "方案标识"
        .TextMatrix(0, 2) = "方案编码"
        .TextMatrix(0, 3) = "方案名称"
        .TextMatrix(0, 4) = "源表名"
        .TextMatrix(0, 5) = "目的表名"
        .TextMatrix(0, 6) = "源单据类型"
        .TextMatrix(0, 7) = "源业务类型"
        .TextMatrix(0, 8) = "源部门"
        .TextMatrix(0, 9) = "源仓库"
        .TextMatrix(0, 10) = "源客商"
        .TextMatrix(0, 11) = "源模板"
        .TextMatrix(0, 12) = "源账套"
        .TextMatrix(0, 13) = "目的账套"
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
        .ColWidth(10) = 0
        .ColWidth(11) = 0
        .ColWidth(12) = 0
        .ColWidth(13) = 0
        
        'ID,cPlanCode,cPlanName,TabName,cTabNameP,cVouchType,cBusType,cDepCode,cWhCode,cVenCode,VT_ID
        .ColKey(0) = "rowno"
        .ColKey(1) = "ID"
        .ColKey(2) = "cPlanCode"
        .ColKey(3) = "cPlanName"
        .ColKey(4) = "cTabName"
        .ColKey(5) = "cTabNameP"
        .ColKey(6) = "cVouchType"
        .ColKey(7) = "cBusType"
        .ColKey(8) = "cDepCode"
        .ColKey(9) = "cWhCode"
        .ColKey(10) = "cVenCode"
        .ColKey(11) = "VT_ID"
        .ColKey(12) = "cAccId"
        .ColKey(13) = "cAccIdP"
        
        '行号居中
        .ColAlignment(0) = flexAlignCenterCenter
        
        For i = 0 To .Cols - 1
            '默认都设置成文本类型
            .ColDataType(i) = flexcpText
        Next
        
    End With
End Sub

Private Sub InitBusinessGrid()
    Dim i As Integer
    With gridBusiness
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
        iSelBusinessDataCnt = 0
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
    FillBusinessGrid
End Sub

Private Sub FillPlanGrid(Optional bClearCondition As Boolean = False)
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim sPlanCode As String
    Dim sPlanName As String
    
    If bClearCondition Then ClearPlanCondition
    
    sql = "select ID,cPlanCode,cPlanName,cAccID,cAccIDP,cTabName,cTabNameP,cVouchType,cBusType,cDepCode,cWhCode,cVenCode,VT_ID from JT_BusinessSet where 1=1 "
    
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
    
    Select Case UCase(m_cFrmID)
    
        Case "JT0202"
            sql = sql & " and cTabName = 'rdrecord32'"
        Case "JT0203"
            sql = sql & " and cTabName = 'ExpenseVouch'"
        Case "JT0204"
            sql = sql & " and cTabName = 'SalePayVouch'"
        Case "JT0205"
            sql = sql & " and cTabName = 'rdrecord01'"
        Case "JT0206"
            sql = sql & " and cTabName = 'ap_closebill'"
        Case Else
            'show all
    End Select
    
    With gridPlan
        Set rs = QueryExt(sql)
        If Not rs.BOF And Not rs.EOF Then
            Do While Not rs.EOF
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = .Rows - 1
                .TextMatrix(.Rows - 1, 1) = GetNoNullValue(rs!ID)
                .TextMatrix(.Rows - 1, 2) = GetNoNullValue(rs!cPlanCode)
                .TextMatrix(.Rows - 1, 3) = GetNoNullValue(rs!cPlanName)
                .TextMatrix(.Rows - 1, 4) = GetNoNullValue(rs!cTabName)
                .TextMatrix(.Rows - 1, 5) = GetNoNullValue(rs!cTabNameP)
                .TextMatrix(.Rows - 1, 6) = GetNoNullValue(rs!cVouchType)
                .TextMatrix(.Rows - 1, 7) = GetNoNullValue(rs!cBusType)
                .TextMatrix(.Rows - 1, 8) = GetNoNullValue(rs!cDepCode)
                .TextMatrix(.Rows - 1, 9) = GetNoNullValue(rs!cWhCode)
                .TextMatrix(.Rows - 1, 10) = GetNoNullValue(rs!cVenCode)
                .TextMatrix(.Rows - 1, 11) = GetNoNullValue(rs!VT_ID)
                .TextMatrix(.Rows - 1, 12) = GetNoNullValue(rs!cAccId)
                .TextMatrix(.Rows - 1, 13) = GetNoNullValue(rs!cAccIdP)
                rs.MoveNext
            Loop
            '默认选择第一行
            gridPlan.RowSel = 1
            InitBusinessGrid
            FillBusinessGrid (True)
            InitLogGrid
            ClearLogCondition
            txtLogPlan.Text = gridPlan.TextMatrix(gridPlan.RowSel, gridPlan.ColIndex("cPlanCode"))
            'FillLogGrid
        End If
    End With
End Sub

Private Sub FillBusinessGrid(Optional bClearCondition As Boolean = False)
    Dim i As Integer
    Dim sql As String
    Dim sBusinessCode As String '单号
    Dim sBusinessName As String '制单人
    Dim dateFrom As String      '开始日期
    Dim dateTo As String        '结束日期
    Dim sSrcAccId As String     '源账套
    Dim sSrcTabName As String   '源表名
    Dim sSrcAccIdP As String    '目的账套
    Dim sDesTabName As String   '目的表名
    Dim sSrcVouchType As String '单据类型
    Dim sSrcBusType As String   '业务类型
    Dim sSrcDepCode As String   '部门
    Dim sSrcWhCode As String    '仓库
    Dim sSrcVenCode As String   '客商
    Dim sSrcVT_ID As String     '模板
    
    If bClearCondition Then ClearBusinessCondition
    
    sql = ""
    sBusinessCode = txtBusinessCode.Text
    sBusinessName = txtBusinessName.Text
    dateFrom = txtBusinessTimeS.Text
    dateTo = txtBusinessTimeE.Text
    
    If Not IsBlank(dateFrom) And Not IsDate(dateFrom) Then
        MsgBox "开始时间格式错误！", vbOKOnly, "业务单据同步"
        Exit Sub
    End If
    
    If Not IsBlank(dateTo) And Not IsDate(dateTo) Then
        MsgBox "结束时间格式错误！", vbOKOnly, "业务单据同步"
        Exit Sub
    End If
    
    i = gridPlan.RowSel
    If i < 1 Then Exit Sub  '未选择数据
    sSrcAccId = gridPlan.TextMatrix(i, gridPlan.ColIndex("cAccId"))
    sSrcTabName = gridPlan.TextMatrix(i, gridPlan.ColIndex("cTabName"))
    sSrcAccIdP = gridPlan.TextMatrix(i, gridPlan.ColIndex("cAccIdP"))
    sDesTabName = gridPlan.TextMatrix(i, gridPlan.ColIndex("cTabNameP"))
    sSrcVouchType = gridPlan.TextMatrix(i, gridPlan.ColIndex("cVouchType"))
    sSrcBusType = gridPlan.TextMatrix(i, gridPlan.ColIndex("cBusType"))
    sSrcDepCode = gridPlan.TextMatrix(i, gridPlan.ColIndex("cDepCode"))
    sSrcWhCode = gridPlan.TextMatrix(i, gridPlan.ColIndex("cWhCode"))
    sSrcVenCode = gridPlan.TextMatrix(i, gridPlan.ColIndex("cVenCode"))
    sSrcVT_ID = gridPlan.TextMatrix(i, gridPlan.ColIndex("VT_ID"))
    
    With gridBusiness

        Dim rs As ADODB.Recordset
        Select Case LCase(sSrcTabName & "-" & sDesTabName)
            Case "rdrecord32-rdrecord32"
            
                sql = "select '' as 行号,a.id as businessid,a.* from rdrecord32 a where ddate between '" & dateFrom & "' and '" & dateTo & "'"
                If Not IsBlank(sBusinessCode) Then sql = sql & " and ccode like '%" & sBusinessCode & "%'"
                If Not IsBlank(sBusinessName) Then sql = sql & " and cmaker like '%" & sBusinessName & "%'"
                If Not IsBlank(sSrcBusType) Then sql = sql & "  and cbustype = '" & sSrcBusType & "'"
                If Not IsBlank(sSrcDepCode) Then sql = sql & "  and cdepcode = '" & sSrcDepCode & "'"
                If Not IsBlank(sSrcWhCode) Then sql = sql & "  and cwhcode = '" & sSrcWhCode & "'"
                If Not IsBlank(sSrcVenCode) Then sql = sql & "  and ccuscode = '" & sSrcVenCode & "'"
                If Not IsBlank(sSrcVT_ID) Then sql = sql & "  and vt_id = '" & sSrcVT_ID & "'"
                
                '过滤掉目标账套已存在，且时间戳没变的几率
                sql = sql & "  and (not exists (select top 1 * from " & sSrcAccIdP & ".dbo." & sDesTabName & " b where b.cdefine11 = '" & sSrcAccId & "' and b.cdefine12 = a.id)"
                sql = sql & "       or"
                sql = sql & "       exists (select top 1 * from " & sSrcAccIdP & ".dbo." & sDesTabName & " c where c.cdefine11 = '" & sSrcAccId & "' and c.cdefine12 = a.id and isnull(c.cdefine14,'') <> convert(nchar,convert(money,a.ufts),2))"
                sql = sql & "       )"
                
            Case "expensevouch-expensevouch"
            
                '代垫费用单
                sql = "select '' as 行号,a.id as businessid,a.* from expensevouch a where ddate between '" & dateFrom & "' and '" & dateTo & "'"
                If Not IsBlank(sBusinessCode) Then sql = sql & " and cEVCode like '%" & sBusinessCode & "%'"
                If Not IsBlank(sBusinessName) Then sql = sql & " and cmaker like '%" & sBusinessName & "%'"
                If Not IsBlank(sSrcDepCode) Then sql = sql & "  and cdepcode = '" & sSrcDepCode & "'"
                If Not IsBlank(sSrcVenCode) Then sql = sql & "  and ccuscode = '" & sSrcVenCode & "'"
                If Not IsBlank(sSrcVT_ID) Then sql = sql & "  and iVTid = '" & sSrcVT_ID & "'"
                
                sql = sql & "  and (not exists (select top 1 * from " & sSrcAccIdP & ".dbo." & sDesTabName & " b where b.cdefine11 = '" & sSrcAccId & "' and b.cdefine12 = a.id)"
                sql = sql & "       or"
                sql = sql & "       exists (select top 1 * from " & sSrcAccIdP & ".dbo." & sDesTabName & " c where c.cdefine11 = '" & sSrcAccId & "' and c.cdefine12 = a.id and isnull(c.cdefine14,'') <> convert(nchar,convert(money,a.ufts),2))"
                sql = sql & "       )"
                
             Case "salepayvouch-salepayvouch"
             
                '费用支出单
                sql = "select '' as 行号,a.id as businessid,a.* from salepayvouch a where ddate between '" & dateFrom & "' and '" & dateTo & "'"
                If Not IsBlank(sBusinessCode) Then sql = sql & " and cSPVCode like '%" & sBusinessCode & "%'"
                If Not IsBlank(sBusinessName) Then sql = sql & " and cmaker like '%" & sBusinessName & "%'"
                If Not IsBlank(sSrcDepCode) Then sql = sql & "  and cdepcode = '" & sSrcDepCode & "'"
                If Not IsBlank(sSrcVenCode) Then sql = sql & "  and ccuscode = '" & sSrcVenCode & "'"
                If Not IsBlank(sSrcVT_ID) Then sql = sql & "  and iVTid = '" & sSrcVT_ID & "'"
                
                sql = sql & "  and (not exists (select top 1 * from " & sSrcAccIdP & ".dbo." & sDesTabName & " b where b.cdefine11 = '" & sSrcAccId & "' and b.cdefine12 = a.id)"
                sql = sql & "       or"
                sql = sql & "       exists (select top 1 * from " & sSrcAccIdP & ".dbo." & sDesTabName & " c where c.cdefine11 = '" & sSrcAccId & "' and c.cdefine12 = a.id and isnull(c.cdefine14,'') <> convert(nchar,convert(money,a.ufts),2))"
                sql = sql & "       )"
                
             Case "ap_closebill-ap_closebill"
             
                If sSrcVouchType = "48" Then
                
                    '收款单
                    sql = "select '' as 行号,a.iID as businessid,a.* from ap_closebill a where cvouchtype='48' and dVouchDate between '" & dateFrom & "' and '" & dateTo & "'"
                    If Not IsBlank(sBusinessCode) Then sql = sql & " and cVouchID like '%" & sBusinessCode & "%'"
                    If Not IsBlank(sBusinessName) Then sql = sql & " and cOperator like '%" & sBusinessName & "%'"
                    If Not IsBlank(sSrcBusType) Then sql = sql & "  and iBusType = '" & sSrcBusType & "'"
                    If Not IsBlank(sSrcDepCode) Then sql = sql & "  and cdeptcode = '" & sSrcDepCode & "'"
                    If Not IsBlank(sSrcVenCode) Then sql = sql & "  and cdwcode = '" & sSrcVenCode & "'"
                    If Not IsBlank(sSrcVT_ID) Then sql = sql & "  and vt_id = '" & sSrcVT_ID & "'"
                    
                ElseIf sSrcVouchType = "49" Then
                
                    '付款单
                    sql = "select '' as 行号,a.iID as businessid,a.* from ap_closebill a where cvouchtype='49' and dVouchDate between '" & dateFrom & "' and '" & dateTo & "'"
                    If Not IsBlank(sBusinessCode) Then sql = sql & " and cVouchID like '%" & sBusinessCode & "%'"
                    If Not IsBlank(sBusinessName) Then sql = sql & " and cOperator like '%" & sBusinessName & "%'"
                    If Not IsBlank(sSrcBusType) Then sql = sql & "  and iBusType = '" & sSrcBusType & "'"
                    If Not IsBlank(sSrcDepCode) Then sql = sql & "  and cdeptcode = '" & sSrcDepCode & "'"
                    If Not IsBlank(sSrcVenCode) Then sql = sql & "  and cdwcode = '" & sSrcVenCode & "'"
                    If Not IsBlank(sSrcVT_ID) Then sql = sql & "  and vt_id = '" & sSrcVT_ID & "'"
                End If
                
                sql = sql & "  and (not exists (select top 1 * from " & sSrcAccIdP & ".dbo." & sDesTabName & " b where b.cdefine11 = '" & sSrcAccId & "' and b.cdefine12 = a.iid)"
                sql = sql & "       or"
                sql = sql & "       exists (select top 1 * from " & sSrcAccIdP & ".dbo." & sDesTabName & " c where c.cdefine11 = '" & sSrcAccId & "' and c.cdefine12 = a.iid and isnull(c.cdefine14,'') <> convert(nchar,convert(money,a.ufts),2))"
                sql = sql & "       )"
             Case ""
                'no Plan do nothing
        End Select
        If Not IsBlank(sql) Then
            '五大单据之外的数据，直接绑定rs到DataSource
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
            .ColWidth(2) = 0 '隐藏一个Businesscode列，统一名称，方便取数
            
            '选择框、行号居中
            .ColAlignment(0) = flexAlignCenterCenter
            .ColAlignment(1) = flexAlignCenterCenter

            '默认都选中
            iSelBusinessDataCnt = .Rows - 1
        End If
    End With
End Sub

'bShowUpdateOnly 是否只显示本次同步日志
Private Sub FillLogGrid(Optional bShowUpdateOnly As Boolean = False)
    Dim rs As ADODB.Recordset
    
    Dim sql As String
    Dim sPlanCode As String
    Dim sBusinessCode As String
    Dim dateFrom As String
    Dim dateTo As String
    
    sql = ""
    If bShowUpdateOnly Then
        '显示本次同步日志
        sql = "select * from " & TabBusinessLog & " where 1=1 "
    Else
        sPlanCode = txtLogPlan.Text
        sBusinessCode = txtLogBusiness.Text
        dateFrom = txtLogTimeS.Text
        dateTo = txtLogTimeE.Text
        
        If Not IsBlank(dateFrom) And Not IsDate(dateFrom) Then
            MsgBox "开始时间格式错误！", vbOKOnly, "业务单据同步"
            Exit Sub
        End If
        
        If Not IsBlank(dateTo) And Not IsDate(dateTo) Then
            MsgBox "结束时间格式错误！", vbOKOnly, "业务单据同步"
            Exit Sub
        End If
        
        sql = "select * from JT_BusinessLog where 1=1 "
        
        If Not IsBlank(sPlanCode) Then sql = sql & " and cplancode like '%" & sPlanCode & "%'"
        If Not IsBlank(sBusinessCode) Then sql = sql & " and (ccode like '%" & sBusinessCode & "%' or ccodep like '%" & sBusinessCode & "%')"
        
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
    fraBusiness.Left = fraPlan.Left + fraPlan.Width + 100
    fraBusiness.Top = fraPlan.Top
    fraBusiness.Width = Me.ScaleWidth * 7 / 10 - 300 '占屏幕宽度的75%
    fraBusiness.Height = fraPlan.Height * 3 / 5 '右侧高度的60%
    
    lblBusinessCode.Left = 200
    lblBusinessCode.Top = 400
    txtBusinessCode.Left = lblPlanCode.Left + lblPlanCode.Width + 100
    txtBusinessCode.Top = lblPlanCode.Top

    lblBusinessName.Left = txtBusinessCode.Left + txtBusinessCode.Width + 100
    lblBusinessName.Top = txtBusinessCode.Top
    txtBusinessName.Left = lblBusinessName.Left + lblBusinessName.Width + 100
    txtBusinessName.Top = lblBusinessName.Top
    
    lblBusinessTime.Left = txtBusinessName.Left + txtBusinessName.Width + 100
    lblBusinessTime.Top = txtBusinessName.Top
    
    txtBusinessTimeS.Left = lblBusinessTime.Left + lblBusinessTime.Width + 100
    txtBusinessTimeS.Top = lblBusinessTime.Top
    
    cmdBusinessTimeS.Left = txtBusinessTimeS.Left + txtBusinessTimeS.Width - cmdBusinessTimeS.Width
    cmdBusinessTimeS.Top = txtBusinessTimeS.Top
    cmdBusinessTimeS.Height = txtBusinessTimeS.Height
    
    lblBusinessTimeTo.Left = cmdBusinessTimeS.Left + cmdBusinessTimeS.Width + 100
    lblBusinessTimeTo.Top = cmdBusinessTimeS.Top
    
    txtBusinessTimeE.Left = lblBusinessTimeTo.Left + lblBusinessTimeTo.Width + 100
    txtBusinessTimeE.Top = lblBusinessTimeTo.Top
    
    cmdBusinessTimeE.Left = txtBusinessTimeE.Left + txtBusinessTimeE.Width - cmdBusinessTimeE.Width
    cmdBusinessTimeE.Top = txtBusinessTimeE.Top
    cmdBusinessTimeE.Height = txtBusinessTimeE.Height
    
    cmdBusinessQuery.Left = cmdBusinessTimeE.Left + cmdBusinessTimeE.Width + 100
    cmdBusinessQuery.Top = cmdBusinessTimeE.Top
    
    cmdSynchronize.Left = cmdBusinessQuery.Left + cmdBusinessQuery.Width + 100
    cmdSynchronize.Top = cmdBusinessQuery.Top
    
    gridBusiness.Top = lblBusinessCode.Top + lblBusinessCode.Height + 200
    gridBusiness.Left = lblBusinessCode.Left
    gridBusiness.Width = fraBusiness.Width - 400
    gridBusiness.Height = fraBusiness.Height - gridBusiness.Top - 200
    
    
    '同步历史
    fraLog.Left = fraPlan.Left + fraPlan.Width + 100
    fraLog.Top = fraBusiness.Top + fraBusiness.Height + 100
    fraLog.Width = fraBusiness.Width
    fraLog.Height = fraPlan.Height * 2 / 5 - 100 '右侧高度的40%
    
    lblLogPlan.Left = 200
    lblLogPlan.Top = 400
    txtLogPlan.Left = lblLogPlan.Left + lblLogPlan.Width + 100
    txtLogPlan.Top = lblLogPlan.Top

    lblLogBusiness.Left = txtLogPlan.Left + txtLogPlan.Width + 100
    lblLogBusiness.Top = txtLogPlan.Top
    txtLogBusiness.Left = lblLogBusiness.Left + lblLogBusiness.Width + 100
    txtLogBusiness.Top = lblLogBusiness.Top
    
    lblLogTime.Left = txtLogBusiness.Left + txtLogBusiness.Width + 100
    lblLogTime.Top = txtLogBusiness.Top
    
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

Private Sub gridBusiness_Click()
    Dim i, j, k As Integer
    Dim bChecked As Boolean
    
    With gridBusiness
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
                iSelBusinessDataCnt = IIf(bChecked, 0, .Rows - 1)
            Else
                '单选
                .Cell(flexcpChecked, i, 0) = IIf(bChecked, flexUnchecked, flexChecked)
                iSelBusinessDataCnt = IIf(bChecked, iSelBusinessDataCnt - 1, iSelBusinessDataCnt + 1)
                If iSelBusinessDataCnt = .Rows - 1 Then
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
    InitBusinessGrid
    FillBusinessGrid (True)
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

