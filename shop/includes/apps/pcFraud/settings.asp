<%
On Error Resume Next

pcv_strAdminPrefix="1"

If (session("admin") = 0 OR session("admin") = 1 OR session("admin") = "") _
	OR _
	((instr(session("PmAdmin"),"*")=0 And instr(session("PmAdmin"),"19")=0)) _
	OR _
	(len(session("CUID"))=0) _
	OR _
	(session("admin." & pcf_getAdminToken()) <> Session.SessionID) Then
	
	call closeDb()
    response.write("You do not have enough permissions to access the selected page.")
    response.End
    
end if

Private Function pcf_getAdminToken()
	pcv_strLocalAddress = Request.ServerVariables("LOCAL_ADDR") 
	pcv_strLocalSessionID = Session.SessionID
	pcv_strAdminToken = pcv_strLocalAddress & "." & pcv_strLocalSessionID
	pcf_getAdminToken = pcv_strAdminToken
End Function

'// START: Update Settings
If request("updateSettings")<>"" Then
	
    pcv_intFeatureOn = getUserInput(request("FeatureOn"), 1)
    If len(pcv_intFeatureOn)=0 Then
        pcv_intFeatureOn = 0
    Else
        pcv_intFeatureOn = 1
    End If

	pcPay_FA_Active=Request.Form("pcPay_FA_Active")
	pcPay_FA_LicenseKey=Request.Form("pcPay_FA_LicenseKey")
	pcPay_FA_RiskScore=Request.Form("pcPay_FA_RiskScore")
	if pcPay_FA_RiskScore="" then
		pcPay_FA_RiskScore = 0
	end if
	pcPay_FA_SendShipping=Request.Form("pcPay_FA_SendShipping")
	if pcPay_FA_SendShipping="" then
		pcPay_FA_SendShipping = 0
	end if
	pcPay_FA_SendEmail=Request.Form("pcPay_FA_SendEmail")
	if pcPay_FA_SendEmail="" then
		pcPay_FA_SendEmail = 0
	end if
	pcPay_FA_SendPhone=Request.Form("pcPay_FA_SendPhone")
	if pcPay_FA_SendPhone="" then
		pcPay_FA_SendPhone = 0
	end if
	pcPay_FA_RiskScoreEmail=Request.Form("pcPay_FA_RiskScoreEmail")
	if pcPay_FA_SendPhone="" then
		pcPay_FA_SendPhone = 0
	end if
	pcPay_FA_Emails=Request.Form("pcPay_FA_Emails")
	pcPay_FA_RiskScoreLock=Request.Form("pcPay_FA_RiskScoreLock")
    
    '// Debug...
    'response.Write("pcPay_FA_Active:  " & pcPay_FA_Active & "<br />")
    'response.Write("pcPay_FA_LicenseKey:  " & pcPay_FA_LicenseKey & "<br />")
    'response.Write("pcPay_FA_RiskScore:  " & pcPay_FA_RiskScore & "<br />")
    'response.Write("pcPay_FA_SendShipping:  " & pcPay_FA_SendShipping & "<br />")
    'response.Write("pcPay_FA_SendEmail:  " & pcPay_FA_SendEmail & "<br />")
    'response.Write("pcPay_FA_SendPhone:  " & pcPay_FA_SendPhone & "<br />")
    'response.Write("pcPay_FA_RiskScoreEmail:  " & pcPay_FA_RiskScoreEmail & "<br />")
    'response.Write("pcPay_FA_Emails:  " & pcPay_FA_Emails & "<br />")
    'response.Write("pcPay_FA_RiskScoreLock:  " & pcPay_FA_RiskScoreLock & "<br />")

	
    query="UPDATE pcWebServiceFraud SET pcPay_FA_RiskScore=" & pcPay_FA_RiskScore & ", pcPay_FA_SendShipping=" & pcPay_FA_SendShipping & ", pcPay_FA_SendEmail=" & pcPay_FA_SendEmail & ", pcPay_FA_SendPhone=" & pcPay_FA_SendPhone & ", pcPay_FA_RiskScoreEmail=" & pcPay_FA_RiskScoreEmail & ", pcPay_FA_Emails='" & pcPay_FA_Emails & "', pcPay_FA_RiskScoreLock=" & pcPay_FA_RiskScoreLock & " "
    
    'response.Write("query:  " & query & "<br />")
    
    set rs=Server.CreateObject("ADODB.Recordset")
    set rs=connTemp.execute(query)
    set rs=nothing
    
    'response.End()

    call pcs_UpdateFeatureStatusByCode(pcv_strThisFeatureCode, pcv_intFeatureOn)

    call pcs_GenGlobalWebServiceSettings()

	msg = "Settings saved successfully!"
    
End If
'// END: Update Settings


'// START: Unsubscribe
If request("unsubscribe")<>"" Then

    '// Unsubscribe
    pcv_boolIsUnsubscribed = True
    
    If pcv_boolIsUnsubscribed Then
    
        '// Turn Feature Off
        call pcs_UpdateFeatureStatusByCode(pcv_strThisFeatureCode, 0)
        
        '// Disable Service
        call pcs_UpdateFeature(pcv_strThisFeatureCode, 0)
        
        call pcs_GenGlobalWebServiceSettings()
    
        msg = "Unsubscribed successfully!"
    
    Else
    
        msg = "We couldn't unsubscribe you at the moment. Please call or try again later."
    
    End If
  
End If
'// END: Unsubscribe
%>

<%
'// START: Page Load
pcv_intIsActive = pcf_IsFeatureActiveByCode(pcv_strThisFeatureCode)
pcv_intIsEnabled = pcf_GetFeatureStatusByCode(pcv_strThisFeatureCode)

query = "SELECT * FROM pcWebServiceFraud"
set rs = server.CreateObject("ADODB.RecordSet")
set rs = conntemp.execute(query)
If Not rs.Eof Then
    pcPay_FA_LicenseKey=rs("pcPay_FA_LicenseKey")
    pcPay_FA_Active=rs("pcPay_FA_Active")
    pcPay_FA_RiskScore=rs("pcPay_FA_RiskScore")
    pcPay_FA_OrderStatus=rs("pcPay_FA_OrderStatus")
    pcPay_FA_SendShipping=rs("pcPay_FA_SendShipping")
    pcPay_FA_SendEmail=rs("pcPay_FA_SendEmail")
    pcPay_FA_SendPhone=rs("pcPay_FA_SendPhone")
    pcPay_FA_RiskScoreEmail=rs("pcPay_FA_RiskScoreEmail")
    pcPay_FA_Emails=rs("pcPay_FA_Emails")
    pcPay_FA_RiskScoreLock=rs("pcPay_FA_RiskScoreLock")
End If
set rs = Nothing
'// END: Page Load
%>
<h3 class="pcHeader">Fraud Alert Settings</h3>

<div class="bs-callout bs-callout-info">
    <h4>Read Me</h4>    
    <p>
        Before you turn this feature on we recommend reading the <a target="_blank" href="https://productcart.desk.com/customer/portal/articles/2509280-fraud-alert">user guide</a>.     
        
    </p>
</div>

<!--
<% if msg<>"" then %>
    <div class="pcCPmessage"><%=msg%></div>
<% end if %>
-->

<form name="form1" method="post" action="pcws_Settings.asp?fc=<%=pcv_strThisFeatureCode %>" class="pcForms">

    <div class="form-group">
        <div class="col-sm-12">
            <div class="app-toggle">
                <div class="onoffswitch">
                    <input type="checkbox" name="FeatureOn" class="onoffswitch-checkbox" id="myonoffswitch2" <% if pcv_intIsEnabled=1 then%>checked<% end if %>>
                    <label class="onoffswitch-label" for="myonoffswitch2">
                        <span class="onoffswitch-inner"></span>
                        <span class="onoffswitch-switch"></span>
                    </label>
                </div>
            </div>
        </div>
    </div>

    <div class="form-group">
        <label for="pcPay_FA_RiskScore">Max Risk Score: </label>
        <input type="text" class="form-control" id="pcPay_FA_RiskScore" name="pcPay_FA_RiskScore" value="<%=pcPay_FA_RiskScore%>" size="1">
        <span class="help-block">This field contains the risk score, from 0.01 to 99. For example, a score of 20 indicates a 20% chance that a transaction is fraudulent. </span>
    </div> 
    
    <div class="checkbox">
        <label>
            <input type="checkbox" id="pcPay_FA_SendShipping" name="pcPay_FA_SendShipping" value="1" <% if pcPay_FA_SendShipping = 1 then %>checked<%end if%>> Send order's shipping address. 
        </label>
    </div>
    
    <div class="checkbox">
        <label>
            <input type="checkbox" name="pcPay_FA_SendEmail" value="1" <% if pcPay_FA_SendEmail = 1 then %>checked<%end if%>> Send customer's email. 
        </label>
    </div>
    
    <div class="checkbox">
        <label>
            <input type="checkbox" name="pcPay_FA_SendPhone" value="1" <% if pcPay_FA_SendPhone = 1 then %>checked<%end if%>> Send customer's phone number. 
        </label>
    </div>

    <!--
    <div class="form-group">
        <label for="pcPay_FA_RiskScoreEmail"></label>
        If risk score exceeds &nbsp;<input type="text" id="pcPay_FA_RiskScoreEmail" name="pcPay_FA_RiskScoreEmail" size="1" value="<%=pcPay_FA_RiskScoreEmail%>" />&nbsp; then email the following address &nbsp;<input type="text" name="pcPay_FA_Emails" value="<%=pcPay_FA_Emails%>" />
    </div> 
    -->
    
    <input type="hidden" id="pcPay_FA_RiskScoreEmail" name="pcPay_FA_RiskScoreEmail" value="99" />
    <input type="hidden" id="pcPay_FA_Emails" name="pcPay_FA_Emails" value="na" />
    
    <div class="form-group">
        <label for="pcPay_FA_RiskScoreLock"></label>
        Lock the account if risk score exceeds &nbsp;<input type="text" id="pcPay_FA_RiskScoreLock" name="pcPay_FA_RiskScoreLock" size="1" value="<%=pcPay_FA_RiskScoreLock%>" />
    </div> 

    <div class="form-group">
        <div class="col-sm-12">
            <hr />
            
            <input type="submit" name="updateSettings" value="<%=dictLanguageCP.Item(Session("language")&"_cpCommon_107")%>" class="btn btn-primary">
            
            &nbsp;
            
            <input type="button" data-toggle="modal" data-target="#myModal" name="uninstall" value="Uninstall" class="btn btn-danger">


  
            <!-- Modal -->
            <div class="modal fade" id="myModal" tabindex="-1" role="dialog" aria-labelledby="myModalLabel">
              <div class="modal-dialog" role="document">
                <div class="modal-content">
                  <div class="modal-header">
                    <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
                    <h4 class="modal-title" id="myModalLabel">Please Confirm</h4>
                  </div>
                  <div class="modal-body">
                    <div ng-bind-html="error"></div>
                    Are you sure you want to uninstall this app?
                  </div>
                  <div class="modal-footer">
                    <button type="button" class="btn btn-default" data-dismiss="modal">No</button>
                    <button type="button" class="btn btn-primary" data-ng-click="Uninstall('/MyApps/<%=pcv_strUid %>', 'pcFraud');">Yes</button>
                  </div>
                </div>
              </div>
            </div>
            
            
        </div>
    </div>

</form>