Attribute VB_Name = "mdl_Creo_pub"
Public asyncConnection As pfcls.IpfcAsyncConnection
Public Casync As New pfcls.CCpfcAsyncConnection
Public creoSsn As pfcls.IpfcBaseSession
Public Sub Conn2creo()
    Set asyncConnection = Casync.Connect(Null, Null, Null, Null)
    Set creoSsn = asyncConnection.session
End Sub
Public Sub DisconnCreo()
    asyncConnection.End
End Sub
Public Sub editParam(mdl As pfcls.IpfcModel, pName As String, vstr As String)
Dim po As IpfcParameterOwner
Set po = mdl
Dim pv As pfcls.IpfcParamValue
Set pv = po.getParam(pName).GetScaledValue
pv.StringValue = vstr
Call po.getParam(pName).SetScaledValue(pv, Null)
End Sub
Public Function getParam(mdl As pfcls.IpfcModel, pName As String)
Dim po As IpfcParameterOwner
Set po = mdl
Dim pv As pfcls.IpfcParamValue
Set pv = po.getParam(pName).GetScaledValue
getParam = pv.StringValue
End Function
