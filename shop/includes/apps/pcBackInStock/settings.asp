<%
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

	pcv_strMsg = request("nmMsg") '// getUserInput2(request("nmMsg"),0)    
	'pcv_strMsg = replace(pcv_strMsg, vbcrlf, "<br>")
    pcv_strMsg = replace(pcv_strMsg, "'", "''")
    pcv_intAuto = getUserInput(request("nmAuto"),0)    
	pcv_strBText = getUserInput(request("nmBText"),0)    
    pcv_strSubject = getUserInput(request("nmSubject"),0)
    pcv_strFromName = getUserInput(request("nmFromName"),0)
    pcv_strFromEmail = getUserInput(request("nmFromEmail"),0)

    query="SELECT [pcBIS_Msg] FROM pcWebServiceBackInStock"
    Set rs2 = server.CreateObject("ADODB.RecordSet")
    Set rs2 = connTemp.execute(query)
    If Not rs2.Eof Then    
        query="UPDATE pcWebServiceBackInStock SET "
        query = query & "pcBIS_Msg='" & pcv_strMsg & "', "
        query = query & "pcBIS_Auto='" & pcv_intAuto & "', "
        query = query & "pcBIS_Subject='" & pcv_strSubject & "', "
        query = query & "pcBIS_FromName='" & pcv_strFromName & "', "
        query = query & "pcBIS_FromEmail='" & pcv_strFromEmail & "', "
        query = query & "pcBIS_ButtonText='" & pcv_strBText & "' "
        'response.Write(query)
        'response.End()
        Set rs4 = server.CreateObject("ADODB.RecordSet")
        Set rs4 = connTemp.execute(query)
        Set rs4 = Nothing     
    Else    
        query="INSERT INTO pcWebServiceBackInStock ([pcBIS_Msg], [pcBIS_Auto], [pcBIS_Subject], [pcBIS_FromName], [pcBIS_FromEmail], [pcBIS_ButtonText]) VALUES ('" & pcv_strMsg & "', '" & pcv_intAuto & "', '" & pcv_strSubject & "', '" & pcv_strFromName & "', '" & pcv_strFromEmail & "', '" & pcv_strBText & "');"
        Set rs4 = server.CreateObject("ADODB.RecordSet")
        Set rs4 = connTemp.execute(query)
        Set rs4 = Nothing    
    End If
    Set rs2 = Nothing 


    '// Turn ON/ OFF
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

query="SELECT * FROM pcWebServiceBackInStock"
Set rs2 = server.CreateObject("ADODB.RecordSet")
Set rs2 = connTemp.execute(query)
If Not rs2.Eof Then    
    pcv_strMsg = rs2("pcBIS_Msg")
    pcv_strAuto = rs2("pcBIS_Auto")
    pcv_strButtonText = rs2("pcBIS_ButtonText")
    pcv_strSubject = rs2("pcBIS_Subject")
    pcv_strFromEmail = rs2("pcBIS_FromEmail")
    pcv_strFromName = rs2("pcBIS_FromName")
End If
Set rs2 = Nothing 

nmSubject = pcv_strSubject 
nmFromName = pcv_strFromName 
nmFromEmail = pcv_strFromEmail 

If len(nmFromName)=0 Or IsNull(nmFromName) Then
    nmFromName = scCompanyName
End If
If len(nmFromEmail)=0 Or IsNull(nmFromEmail) Then
    nmFromEmail = scFrmEmail
End If
If len(nmSubject)=0 Or IsNull(nmSubject) Then
    nmSubject = "Back in Stock {productname}"
End If

nmMsg = replace(pcv_strMsg, "<br>", vbcrlf)
nmMsg = replace(nmMsg, "<br/>", vbcrlf)
nmMsg = replace(nmMsg, "<br />", vbcrlf)
if nmMsg="" then
    nmMsg="{firstname} {lastname}," & vbcrlf & "The product {productname} ({productsku}) is back in stock." & vbcrlf & "We currently have {quantity_available} units in stock. Please follow the link below to purchase it:" & vbcrlf & "{product_url}" & vbcrlf & vbcrlf & "Best regards," & vbcrlf & "{storename}"
end if

nmAuto=pcv_strAuto
if nmAuto="" then
    nmAuto=0
end if

nmBText=pcv_strButtonText
if nmBText="" then
    nmBText="Notify In-Stock"
end if

pcv_intIsActive = pcf_IsFeatureActiveByCode(pcv_strThisFeatureCode)
pcv_intIsEnabled = pcf_GetFeatureStatusByCode(pcv_strThisFeatureCode)
'// END: Page Load
%>
<h3 class="pcHeader">Settings</h3>

<% ' START show message, if any %>
	<!--include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

<div class="bs-callout bs-callout-info">
    <h4>Read Me</h4>    
    <p>
        Before you turn this feature on we recommend reading the <a target="_blank" href="https://productcart.desk.com/customer/portal/articles/">user guide</a>.     
        
    </p>
</div>

<script type=text/javascript>
function checkFormA(tmpForm)
{
    if (tmpForm.nmMsg.value=="")
    {
        alert("Please enter a value for 'E-mail Template' field");
        tmpForm.nmMsg.focus();
        return(false);
    }
    if (tmpForm.nmBText.value == false)
    {
        alert("Please enter a value for 'Button Text' field");
        tmpForm.nmBText.focus();
        return(false);
    }
    return(true);
}
</script>
<form name="form1" method="post" action="pcws_Settings.asp?fc=<%=pcv_strThisFeatureCode %>" class="pcForms" onSubmit="javascript: return(checkFormA(this));" autocomplete="off">

    <h2>Manage</h2>
    
   <div class="form-group">
        <div class="col-sm-12">
            
            
            <a href="<%=scStoreURL%>/store/<%=scAdminFolderName%>/nmReports.asp" target="_blank" class="btn btn-default">Reports</a>

            &nbsp;
            
            <a href="<%=scStoreURL%>/store/<%=scAdminFolderName%>/nmSendManually.asp" target="_blank" class="btn btn-default">Manual Send</a>

            
            <hr />
        </div>
    </div>
    
    <h2>Settings</h2>
    
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
        <div class="col-sm-4">
            <label for="nmFromName">From Name</label>
            <input class="form-control" type="text" name="nmFromName" size="250" value="<%=nmFromName%>">
            <br />
        </div>
    </div>
    
    <div class="form-group">
        <div class="col-sm-4">
            <label for="nmFromEmail">From Email</label>
            <input class="form-control" type="text" name="nmFromEmail" size="30" value="<%=nmFromEmail%>">
            <br />
        </div>
    </div>

    <div class="form-group">
        <div class="col-sm-4">
            <label for="nmSubject">Subject</label>
            <input class="form-control" type="text" name="nmSubject" size="30" value="<%=nmSubject%>">
            <br />
        </div>
    </div>
 
    <div class="form-group">
        <div class="col-sm-12">
            <label for="nmMsg">E-mail Template</label>
            <textarea id="nmMsg" name="nmMsg" rows="8" class="htmleditor"><%=nmMsg%></textarea>
        </div>
    </div>

    <div class="form-group">
        <div class="col-sm-12">
            <div class="radio">
                    <label>
                        <input type="radio" name="nmAuto" value="1"  <%if nmAuto="1" then%>checked<%end if%> class="clearBorder">
                        Automatically send emails
                    </label>
            </div>
            
            <div class="radio">
                    <label>
                        <input type="radio" name="nmAuto" value="0" <%if nmAuto<>"1" then%>checked<%end if%> class="clearBorder">
                        Notify me to manually send the emails
                    </label>
            </div>
        </div>
    </div>
    
    <div class="form-group">
        <div class="col-sm-4">
            <label for="nmMsg">Button Text</label>
            <input class="form-control" type="text" name="nmBText" size="30" value="<%=nmBText%>">
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
                    <button type="button" class="btn btn-primary" data-ng-click="Uninstall('/MyApps/<%=pcv_strUid %>', 'pcBackInStock');">Yes</button>
                  </div>
                </div>
              </div>
            </div>
            
            
        </div>
    </div>

</form>