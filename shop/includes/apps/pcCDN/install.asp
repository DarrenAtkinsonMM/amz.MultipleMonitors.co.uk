<!--#include file="common.asp"-->
<%
pcv_strDomain = getDomainFromURL(scStoreURL)            

if not TableExists("pcWebServiceCDN") then

    query="CREATE TABLE pcWebServiceCDN ("
    query=query&"pcCDN_Id [INT] IDENTITY (1,1) PRIMARY KEY NOT NULL ,"
    query=query&"pcCDN_Domain [nvarchar] (400) NULL ,"
    query=query&"pcCDN_Distribution [nvarchar] (250) NULL "
    query=query&");"

    set rs = server.CreateObject("ADODB.RecordSet")
    set rs = conntemp.execute(query)
    set rs = nothing

end if

if err.number <> 0 then
    Err.Description=""
    err.number=0
    Session("pcAdminInstallMsg") = "Error installing database scripts."
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

'// Add 
call pcs_AddFeatureByCode(pcv_strThisFeatureCode) 
  
'// Provision 
call pcs_provision(pcv_strDomain, pcv_strUid, scCrypPass, pcv_strThisFeatureCode)

'// Update Global Settings
call pcs_GenGlobalWebServiceSettings()  

'// Install App Include
call pcs_GenGlobalAppInclude()  


response.Write("success")
response.End()
%>