VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCallin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
 

Dim PauthCode As String
Dim cModCode As String

Dim strSql As String
Dim rst As New ADODB.Recordset

Public Function CallFunction(ByVal cMenuId As String, ByVal FrmID As String, ByVal cAuthId As String, ByVal cCmdLine As String) As Object
    Dim frmTemp As Form
    
    Call Main
     
    cMenuId = GetFrmType(cMenuId, FrmID)
    Select Case cMenuId
        Case "JT0201"
'            If GetVersion(cMenuId) <> "2.0.6.6" Then
'                MsgBox "您当前版本太低，请进OA打补丁。"
'                Exit Function
'            End If
'            If GetAuth(FrmID) = False Then
'                Exit Function
'            End If
            
            Set frmTemp = New frmSyncBase
            frmTemp.FrmID = FrmID
            frmTemp.Show
        Case "JT0202", "JT0203", "JT0204", "JT0205", "JT0206", "JT0209"
        
            Set frmTemp = New frmSyncBusiness
            frmTemp.FrmID = FrmID
            frmTemp.Show
            
        Case Else     '系统原有的菜单
            MsgBox "此功能配置不正确，或者功能没有完善"
             
            
'            MsgBox "没有此权利", vbInformation, "提示"
    End Select
End Function
 
Public Function IintLoadFrm(LoadFrmLogin As Object, cMenuId As String, ByVal FrmID As String)

    Set MY_Login = New SYS_Login.MY_Login
    If MY_Login.cAcc_Id <> LoadFrmLogin.cAcc_Id Then
        If DBconnSys.State <> 0 Then DBconnSys.Close
        If DBconnData.State <> 0 Then DBconnData.Close
        If DBUFconnSys.State <> 0 Then DBUFconnSys.Close
        If DBUFconnData.State <> 0 Then DBUFconnData.Close
        If DBconnPrice.State <> 0 Then DBconnPrice.Close
    End If
    
    Set MY_Login = LoadFrmLogin
    
    Call CallFunction(cMenuId, FrmID, "", "")
End Function
 
Private Function GetFrmType(cMenuId As String, ByRef FrmID As String) As String
    Dim strSql As String
    Dim RsV As New ADODB.Recordset
    
    If FrmID = "" Then
        strSql = "select isnull(FrmID,'') from Sys_Menu where MenuCode = '" & cMenuId & "'"
        RsV.Open strSql, DBconnSys, adOpenStatic, adLockReadOnly
        If RsV.RecordCount = 0 Then
            GetFrmType = cMenuId
            RsV.Close
            Exit Function
        Else
            If RsV.Fields(0) <> "" Then
                FrmID = RsV.Fields(0)
            Else
                GetFrmType = cMenuId
                RsV.Close
                Exit Function
            End If
        End If
        RsV.Close
    End If
    
    strSql = "select isnull(FrmType,0) from Sys_FrmProperty where frmid = '" & FrmID & "'"
    RsV.Open strSql, DBconnSys, adOpenStatic, adLockReadOnly
    If RsV.RecordCount = 0 Then
        GetFrmType = cMenuId
    Else
        If RsV.Fields(0) > 0 Then
            GetFrmType = RsV.Fields(0)
        Else
            GetFrmType = cMenuId
        End If
    End If
    RsV.Close
   
    
End Function

