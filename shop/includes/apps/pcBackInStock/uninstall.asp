<!--#include file="common.asp"-->
<%
'response.Write("This simulates an error uninstalling the application.")
'response.End()

'// Disable
call pcs_UpdateFeatureStatusByCode(pcv_strThisFeatureCode, 0)

'// Clear Table
call pcs_Update("", "", "")

'// Delete Table (if possible)
query="DELETE FROM pcWebServiceBackInStock;"
set rs = server.CreateObject("ADODB.RecordSet")
set rs = conntemp.execute(query)
set rs = nothing

'// Drop Table (if possible)
query="DROP TABLE pcWebServiceBackInStock;"
set rs = server.CreateObject("ADODB.RecordSet")
set rs = conntemp.execute(query)
set rs = nothing

if err.number <> 0 then
    Err.Description=""
    err.number=0
end if

query="SELECT pcPCWS_Uid, pcPCWS_AuthToken, pcPCWS_Username, pcPCWS_Password FROM pcWebServiceSettings;"
Set rs=connTemp.execute(query)
If Not rs.eof Then
    pcv_strUid = rs("pcPCWS_Uid")
    pcv_AuthToken = rs("pcPCWS_AuthToken")  
    pcv_strUsername = rs("pcPCWS_Username")  
    pcv_strPassword = enDeCrypt(rs("pcPCWS_Password"), scCrypPass)          
End If
Set rs=nothing

'// Remove
call pcs_RemoveFeatureByCode(pcv_strThisFeatureCode) 
call pcs_RemoveWidgetByCode("PrdBackInStock")
call pcs_RemoveWidgetByCode("PrdBackInStockModal")
call pcs_RemoveEventHookByCode("PrdBackInStock")
call pcs_RemoveEventHookByCode("BackInStockCPanelJS")
call pcs_RemoveEventHookByCode("BackInStockMenu")

'// Update Global Settings
call pcs_GenGlobalWebServiceSettings()  

response.Write("success")
response.End()
%>