<!--#include file="common.asp"-->
<%
pcv_strUrl = getDomainFromURL(scStoreURL)  


'// Install Hook - Check For Fraud
Call pcs_AddEventHook("CheckForFraud", "Advanced Fraud", "Execute", "../includes/apps/pcFraud/interface.asp", "CheckForFraud", "PrePayment")

'// Install Hook - Display Fraud Details
Call pcs_AddEventHook("DisplayFraudDetails", "Advanced Fraud", "Execute", "../includes/apps/pcFraud/interface.asp", "DisplayFraudDetails", "CPanelTabPaymentInfo")

'// Create table pcWebServiceFraud
if not TableExists("pcWebServiceFraud") then
    
    query="CREATE TABLE pcWebServiceFraud ("
    query=query&"pcPay_FA_Id [INT] IDENTITY (1,1) PRIMARY KEY NOT NULL,"
    query=query&"pcPay_FA_Active [tinyint] NOT NULL DEFAULT(0),"
    query=query&"pcPay_FA_LicenseKey [nvarchar] (50),"
    query=query&"pcPay_FA_RiskScore [decimal](5,2) NOT NULL DEFAULT(0),"
    query=query&"pcPay_FA_OrderStatus [tinyint] NOT NULL DEFAULT(0),"
    query=query&"pcPay_FA_SendShipping [tinyint] NOT NULL DEFAULT(0),"
    query=query&"pcPay_FA_SendEmail [tinyint] NOT NULL DEFAULT(0),"
    query=query&"pcPay_FA_SendPhone [tinyint] NOT NULL DEFAULT(0),"
    query=query&"pcPay_FA_RiskScoreEmail [decimal](5,2) NOT NULL DEFAULT(0),"
    query=query&"pcPay_FA_Emails [nvarchar] (200),"
    query=query&"pcPay_FA_RiskScoreLock [decimal](5,2) NOT NULL DEFAULT(0)"
    query=query&");"
    
    set rs = server.CreateObject("ADODB.RecordSet")
    set rs = conntemp.execute(query)
    set rs = nothing

    query="SELECT pcPay_FA_Id FROM pcWebServiceFraud;"
    set rs=conntemp.execute(query)
    if rs.eof then
        query="INSERT INTO pcWebServiceFraud (pcPay_FA_Active) VALUES (0);"
        conntemp.execute(query)
    end if
    set rs = nothing

    call AlterTableSQL("orders", "ADD", "faAccountId", "[nvarchar](50)", 0, "", "0")
    call AlterTableSQL("orders", "ADD", "faRiskScore", "[decimal](5,2)", 1, "0", "1")

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
call pcs_provision(pcv_strUrl, pcv_strUid, scCrypPass, pcv_strThisFeatureCode)

'// Update Global Settings
call pcs_GenGlobalWebServiceSettings()  

'// Install App Include
call pcs_GenGlobalAppInclude()  


response.Write("success")
response.End()
%>