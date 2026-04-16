<!--#include file="common.asp"-->
<%    
if not TableExists("pcWebServiceCartStack") then

    query="CREATE TABLE pcWebServiceCartStack ("
    query=query&"pcCS_Id [INT] IDENTITY (1,1) PRIMARY KEY NOT NULL ,"
    query=query&"pcCS_Name [nvarchar] (400) NULL ,"
    query=query&"pcCS_Email [nvarchar] (250) NULL ,"
    query=query&"pcCS_Url [nvarchar] (250) NULL ,"
    query=query&"pcCS_AccountId [nvarchar] (250) NULL ,"
    query=query&"pcCS_SiteId [nvarchar] (250) NULL ,"
    query=query&"pcCS_APIKey [nvarchar] (250) NULL "
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
call pcs_provision(getDomainFromURL(scStoreURL), pcv_strUid, scCrypPass, pcv_strThisFeatureCode)

'// Update Global Settings
call pcs_GenGlobalWebServiceSettings()  

'// Install App Include
call pcs_GenGlobalAppInclude()  


response.Write("success")
response.End()
%>