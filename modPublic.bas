Attribute VB_Name = "modPublic"

Public MY_Login As SYS_Login.MY_Login

Public G_ConnString As String
Public DBconnSys As New ADODB.Connection     '登录时连接Master数据库
Public DBconnData As New ADODB.Connection     '登录时连接Master数据库
Public DBUFconnSys As New ADODB.Connection     '登录时连接Master数据库
Public DBUFconnData As New ADODB.Connection     '登录时连接Master数据库
Public DBconnPrice As New ADODB.Connection


Public TabBaseSet As String    '待同步方案临时表
Public TabBaseData As String   '待同步数据临时表
Public TabBaseLog As String    '本次同步历史记录表

Public TabBusinessSet As String  '待同步方案临时表
Public TabBusinessData As String '待同步数据临时表
Public TabBusinessLog As String  '本次同步历史记录表

'取机器名
Public g_cComputer As String
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Sub Main()
    On Error Resume Next
        
    If DBconnData.State <> 0 Then Exit Sub
    
    If DBconnSys.State <> 0 Then DBconnSys.Close
    G_ConnString = MY_Login.MYSystemADODb
    DBconnSys.ConnectionString = G_ConnString
    DBconnSys.CursorLocation = adUseClient
    DBconnSys.ConnectionTimeout = 30
    DBconnSys.CommandTimeout = 0
    DBconnSys.Open
    
    If DBconnData.State <> 0 Then DBconnData.Close
    G_ConnString = MY_Login.MYDataADODb
    DBconnData.ConnectionString = G_ConnString
    DBconnData.CursorLocation = adUseClient
    DBconnData.ConnectionTimeout = 30
    DBconnData.CommandTimeout = 0
    DBconnData.Open

End Sub

'只用来执行Insert，update，delete语句
Public Function ExtSql(tmpSql As String) As Boolean
    On Error GoTo errHandler
    
    Dim cmd As New ADODB.Command
    Set cmd.ActiveConnection = DBconnData
    cmd.CommandText = tmpSql
    cmd.Execute
    Set cmd = Nothing
    'Db_DisConnect
    
    ExtSql = True
    Exit Function
    
errHandler:
    ExtSql = False
    MsgBox err.Description, vbOKOnly, "基础档案同步"
    VBA.err.Clear
    Exit Function
End Function

'查询的 程序
Public Function QueryExt(tmpSql As String) As ADODB.Recordset
    Dim rst As New ADODB.Recordset
    Set rst.ActiveConnection = DBconnData
    rst.CursorType = adOpenDynamic
    rst.LockType = adLockOptimistic
    rst.Open tmpSql
    Set QueryExt = rst
End Function

'执行档案同步存储过程
Public Function ExtBaseSyncPro(ByRef errmsg As String) As Boolean
On Error GoTo errHandler
    Dim cmd As New ADODB.Command
    
    With cmd
        .ActiveConnection = DBconnData
        .CommandText = "JT_Proc_SyncBase"
        .CommandType = adCmdStoredProc
        .Prepared = False
        .Parameters.Append .CreateParameter("TabBaseSet", adVarWChar, adParamInput, 256, TabBaseSet)
        .Parameters.Append .CreateParameter("TabBaseData", adVarWChar, adParamInput, 256, TabBaseData)
        .Parameters.Append .CreateParameter("TabBaseLog", adVarWChar, adParamInput, 256, TabBaseLog)
        .Parameters.Append .CreateParameter("ret", adVarWChar, adParamOutput, 256)
        .Execute
        errmsg = .Parameters.Item("ret")
    End With
    If Not IsBlank(errmsg) Then
        ExtBaseSyncPro = False
    Else
        ExtBaseSyncPro = True
    End If
    
ExitFunction:
    Set cmd = Nothing
    Exit Function
    
errHandler:
    ExtBaseSyncPro = False
    errmsg = err.Description
    VBA.err.Clear
    
    GoTo ExitFunction
End Function

'执行业务数据同步存储过程
Public Function ExtBusinessSyncPro(ByRef errmsg As String) As Boolean
On Error GoTo errHandler
    Dim cmd As New ADODB.Command
    
    With cmd
        .ActiveConnection = DBconnData
        .CommandText = "JT_Proc_SyncBusiness"
        .CommandType = adCmdStoredProc
        .Prepared = False
        .Parameters.Append .CreateParameter("TabBusinessSet", adVarWChar, adParamInput, 256, TabBusinessSet)
        .Parameters.Append .CreateParameter("TabBusinessData", adVarWChar, adParamInput, 256, TabBusinessData)
        .Parameters.Append .CreateParameter("TabBusinessLog", adVarWChar, adParamInput, 256, TabBusinessLog)
        .Parameters.Append .CreateParameter("ret", adVarWChar, adParamOutput, 256)
        .Execute
        errmsg = .Parameters.Item("ret")
    End With
    If Not IsBlank(errmsg) Then
        ExtBusinessSyncPro = False
    Else
        ExtBusinessSyncPro = True
    End If
    
ExitFunction:
    Set cmd = Nothing
    Exit Function
    
errHandler:
    ExtBusinessSyncPro = False
    errmsg = err.Description
    VBA.err.Clear
    
    GoTo ExitFunction
End Function

Public Function ComputerName()

    Dim strTemp As String * 255
    GetComputerName strTemp, 255
    ComputerName = Left(strTemp, InStr(1, strTemp, vbNullChar) - 1) & "_" & GetSessionID()
    
End Function

Public Function GetSessionID() As String
    On Error Resume Next
    
    Dim o As Object
    GetSessionID = ""
    Set o = CreateObject("TermMisc.Terminal")
    If Not (o Is Nothing) Then
        GetSessionID = o.GetSessionID()
    Else
        Debug.Print "Fail To Create the TermMisc.Terminal Object in module temptablemager"
    End If
    Set o = Nothing
    
    VBA.err.Clear
End Function


Public Function IsBlank(ByVal strString As String)

    '替换换行符
    strString = Replace(strString, vbCr, "")
    strString = Replace(strString, vbLf, "")
    strString = Replace(strString, vbCrLf, "")
    If Len(strString) = 0 Then
       IsBlank = True
    Else
       IsBlank = False
    End If
    
End Function

Public Function GetNoNullValue(ByVal vTarget As Variant, Optional ByVal vReplace As Variant = "") As Variant

    If IsNull(vTarget) Then
        GetNoNullValue = vReplace
    Else
        GetNoNullValue = vTarget
    End If
    
End Function

Public Sub DropTempTbl()
   
    On Error Resume Next
    
    If DBconnData Is Nothing Then
        Exit Sub
    End If
    
    If DBconnData.State <> adStateOpen Then
        Exit Sub
    End If
    
    Dim strSql As String
    strSql = ""
    If Not IsBlank(TabBaseSet) Then
        strSql = " IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES" & _
           " WHERE TABLE_NAME =  '" & TabBaseSet & " ' And TABLE_TYPE= 'BASE TABLE')" & vbCrLf & _
           " DROP Table " & TabBaseSet & vbCrLf
    End If
    If Not IsBlank(TabBaseData) Then
        strSql = strSql & " IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES" & _
           " WHERE TABLE_NAME =  '" & TabBaseData & " ' And TABLE_TYPE= 'BASE TABLE')" & vbCrLf & _
           " DROP Table " & TabBaseData & vbCrLf
    End If
    If Not IsBlank(TabBaseLog) Then
        strSql = strSql & " IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES" & _
           " WHERE TABLE_NAME =  '" & TabBaseLog & " ' And TABLE_TYPE= 'BASE TABLE')" & vbCrLf & _
           " DROP Table " & TabBaseLog & vbCrLf
    End If
    If Not IsBlank(TabBusinessSet) Then
        strSql = strSql & " IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES" & _
           " WHERE TABLE_NAME =  '" & TabBusinessSet & " ' And TABLE_TYPE= 'BASE TABLE')" & vbCrLf & _
           " DROP Table " & TabBusinessSet & vbCrLf
    End If
    If Not IsBlank(TabBusinessData) Then
        strSql = strSql & " IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES" & _
           " WHERE TABLE_NAME =  '" & TabBusinessData & " ' And TABLE_TYPE= 'BASE TABLE')" & vbCrLf & _
           " DROP Table " & TabBusinessData & vbCrLf
    End If
    If Not IsBlank(TabBusinessLog) Then
        strSql = strSql & " IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES" & _
           " WHERE TABLE_NAME =  '" & TabBusinessLog & " ' And TABLE_TYPE= 'BASE TABLE')" & vbCrLf & _
           " DROP Table " & TabBusinessLog & vbCrLf
    End If
    If Not IsBlank(strSql) Then DBconnData.Execute strSql
    
    VBA.err.Clear
End Sub
