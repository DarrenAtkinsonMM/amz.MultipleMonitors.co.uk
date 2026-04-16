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
'// END: Page Load
%>
<h3 class="pcHeader">CartStack Settings</h3>

<div class="bs-callout bs-callout-info">
    <h4>Read Me</h4>    
    <p>
        Before you turn this feature on we recommend reading the <a target="_blank" href="https://productcart.desk.com/customer/portal/articles/2509281-cartstack">user guide</a>.
       
    </p>
</div>

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
                    <button type="button" class="btn btn-primary" data-ng-click="Uninstall('/MyApps/<%=pcv_strUid %>', 'pcCartStack');">Yes</button>
                  </div>
                </div>
              </div>
            </div>
            
            
        </div>
    </div>

</form>