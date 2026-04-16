<!--#include file="common.asp"-->
<%
pcv_strDomain = getDomainFromURL(scStoreURL)


'// Install Widget - Minimal Widget
Call pcs_InstallWidget("PrdBackInStock", "Back In Stock Minimal", "Execute", "../includes/apps/pcBackInStock/interface.asp", "BackInStockWidget")

'// Install Widget - Modal Widget
Call pcs_InstallWidget("PrdBackInStockModal", "Back In Stock Modal", "Execute", "../includes/apps/pcBackInStock/interface.asp", "BackInStockWidgetModal")

'// Install Hook - Add to Wait List
Call pcs_AddEventHook("PrdBackInStock", "Back In Stock Wait List", "Execute", "../includes/apps/pcBackInStock/interface.asp", "PrdBackInStockWaitList", "InStockEvent")

'// Install Hook - Add JS to Control Panel for Auto Send
Call pcs_AddEventHook("BackInStockCPanelJS", "Back In Stock JS", "Execute", "../includes/apps/pcBackInStock/interface.asp", "BackInStockCPanelJS", "CPanelFooterJS")

'// Install Hook - Add to Wait List
Call pcs_AddEventHook("BackInStockMenu", "Back In Stock Alert", "Execute", "../includes/apps/pcBackInStock/interface.asp", "BackInStockMenu", "PreCPanelMenu")


'// pcWebServiceBackInStock
if not TableExists("pcWebServiceBackInStock") then

    query="CREATE TABLE pcWebServiceBackInStock ("
    query=query&"pcBIS_Id [INT] IDENTITY (1,1) PRIMARY KEY NOT NULL, "
    query=query&"pcBIS_Msg [nvarchar](MAX) NULL,"
    query=query&"pcBIS_Subject [nvarchar](250) NULL,"
    query=query&"pcBIS_FromEmail [nvarchar](250) NULL,"
    query=query&"pcBIS_FromName [nvarchar](250) NULL,"
    query=query&"pcBIS_Auto [INT] DEFAULT(0) NULL,"
    query=query&"pcBIS_ButtonText [nvarchar](150) NULL"
    query=query&");"
    set rs = server.CreateObject("ADODB.RecordSet")
    set rs = conntemp.execute(query)
    set rs = Nothing
    
end if

'// pcBIS_ListEmails
if not TableExists("pcBIS_ListEmails") then
    query="CREATE TABLE pcBIS_ListEmails ("
    query=query&"BackInStockUID [INT] IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"
    query=query&"idProduct [INT] DEFAULT(0) NULL,"
    query=query&"Guid [nvarchar](250) NULL,"
    query=query&"Email [nvarchar](250) NULL,"
    query=query&"ParentProductID [INT] DEFAULT(0) NULL,"
    query=query&"AddedTime [DateTime] NULL,"
    query=query&"Sent [INT] DEFAULT(0) NULL,"
    query=query&"SentTime [DateTime] NULL,"
    query=query&"Quantity [INT] DEFAULT(0) NULL"
    query=query&");"
		set rs = server.CreateObject("ADODB.RecordSet")
    set rs = conntemp.execute(query)
    set rs = Nothing
end if

'// pcBIS_WaitList
if not TableExists("pcBIS_WaitList") then
    query="CREATE TABLE pcBIS_WaitList ("
    query=query&"pcBIS_ID [INT] IDENTITY (1, 1) PRIMARY KEY NOT NULL,"
    query=query&"idProduct [INT] DEFAULT(0) NULL"
    query=query&");"
    set rs=server.CreateObject("ADODB.RecordSet")
    set rs=conntemp.execute(query)
    set rs=nothing
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
