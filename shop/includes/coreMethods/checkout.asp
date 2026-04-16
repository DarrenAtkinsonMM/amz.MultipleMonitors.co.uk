<%
'// Check Logged In AJAX Calls
Public Sub pcs_CheckLoggedIn

    If Session("idCustomer")=0 Or Session("idCustomer")="" Then
        Response.Clear
        Response.Write "SECURITY"
        Response.End
    End If

End Sub



'// Advanced Security
Public Sub advancedSecurity()

    If scSecurity=1 Then

        Session("store_userlogin")="1"
        session("store_adminre")="1"
        
        If (scUserLogin=1 Or scUserReg=1) And (scUseImgs=1) Then 
            %>
            <div class="pcSpacer"></div>
            <div id="pcCAPTCHA">
			<%if scCaptchaType="1" then
				call pcs_genReCaptcha()
			else%>
                <!--#include file="../../CAPTCHA/CAPTCHA_form_inc.asp" -->
			<%end if%>
            </div>
        <% Else %>
            <div id="show_security"></div>
        <% End If %>

    <% Else %>

        <div id="show_security"></div>

    <% 
    End If
    
End Sub
%>



<%
Public Function showBillingAddressTypeArea(pcv_NOShippingAtAll, pcv_AlwAltShipAddress, scComResShipAddress)

    result = 0
    
    If ((pcv_NOShippingAtAll="1") And (pcv_AlwAltShipAddress="1")) Then

        If scComResShipAddress = "0" Then        
            result = 1
        End If

    End If 
    
    showBillingAddressTypeArea = result

End Function



Public Function showShippingAddressTypeArea(pcv_NOShippingAtAll, pcv_AlwAltShipAddress, scComResShipAddress)

    result = 0

    If (pcv_NOShippingAtAll="1") Or (pcv_NOShippingAtAll="2" And pcv_AlwAltShipAddress="2") Then 
       
        If scComResShipAddress = "0" Then         
            result = 1
        End If
        
    End If
    
    showShippingAddressTypeArea = result

End Function



Public Sub specialCustomerFields()

    tmpCustCFList=""
    pcSFCustFieldsExist=""

    query="SELECT pcCField_ID, pcCField_Name, pcCField_FieldType, pcCField_Value, pcCField_Length, pcCField_Maximum, pcCField_Required, pcCField_PricingCategories, pcCField_ShowOnReg, pcCField_ShowOnCheckout,'',pcCField_Description,0 FROM pcCustomerFields ORDER BY pcCField_Order ASC, pcCField_Name ASC;"
    set rs=server.CreateObject("ADODB.RecordSet")
    set rs=connTemp.execute(query)
    if not rs.eof then
        pcSFCustFieldsExist="YES"
        tmpCustCFList=rs.GetRows()
    end if
    set rs=nothing

    if pcSFCustFieldsExist="YES" AND Session("idCustomer")<>0 then
        pcArr=tmpCustCFList
    
        For k=0 to ubound(pcArr,2)
    
            pcArr(10,k)=""
   
            query="SELECT pcCFV_Value FROM pcCustomerFieldsValues WHERE idcustomer=" & Session("idCustomer") & " AND pcCField_ID=" & pcArr(0,k) & ";"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=connTemp.execute(query)
            if not rs.eof then
                pcArr(10,k)=rs("pcCFV_Value")
            end if
            set rs=nothing
 
        Next
        
        tmpCustCFList=pcArr
        
    end if '// if pcSFCustFieldsExist="YES" AND Session("idCustomer")<>0 then
    
    
    if pcSFCustFieldsExist="YES" then
        pcArr=tmpCustCFList
    
        For k=0 to ubound(pcArr,2)						
            pcv_ShowField=0
            if pcArr(9,k)="1" then
                pcv_ShowField=1
            end if
    
            if (pcv_ShowField=1) AND (pcArr(7,k)="1") then
    
                if session("idCustomer")>"0" then

                    query="SELECT pcCustFieldsPricingCats.idcustomerCategory FROM pcCustFieldsPricingCats INNER JOIN Customers ON (pcCustFieldsPricingCats.pcCField_ID=" & pcArr(0,k) & " AND pcCustFieldsPricingCats.idCustomerCategory=customers.idCustomerCategory) WHERE customers.idcustomer=" & session("idCustomer")
                    set rs=Server.CreateObject("ADODB.Recordset")
                    set rs=conntemp.execute(query)												
                    if NOT rs.eof then
                        pcv_ShowField=1
                    else
                        pcv_ShowField=0
                    end if
                    set rs=nothing

                else '// if session("idCustomer")>"0" then
                
                    pcv_ShowField=0
                    
                end if
            
            end if '// if (pcv_ShowField=1) AND (pcArr(7,k)="1") then
            
            pcArr(12,k)=pcv_ShowField
            
        Next
        tmpCustCFList=pcArr
        
    end if '// if pcSFCustFieldsExist="YES" then
    
    
    if pcSFCustFieldsExist="YES" then
        pcArr=tmpCustCFList
    
        For k=0 to ubound(pcArr,2)
    
            pcv_ShowField=pcArr(12,k)
    
            if pcv_ShowField=1 then 
                %>

                <%
                'get fields
                cfID 		= pcArr(0,k)
                cfName 	= pcArr(1,k)
                cfType 	= pcArr(2,k)
                cfValue = pcArr(3,k)
                cfLen  	= pcArr(4,k)
                cfMax  	= pcArr(5,k)
                cfReq 	= pcArr(6,k)
                
                'get value
                cfChecked = ""
                if pcArr(10,k)<>"" then
                    cfValue = pcArr(10,k)	
                    cfChecked = "checked"														
                end if
                
                'if cfValue="" then
                '	cfValue = "1"
                'end if
                
                'get required
                cfClass = ""
                if cfReq="1" then
                    cfClass = "required"
                end if
                %>

                <% if cfType="1" then %> 
                                           
                    <div class="form-group">                            
                        <div class="checkbox">                                
                        <label for="custfield<%=cfID%>"><%=pcArr(1,k)%><% If cfClass = "required" Then %><div class="pcRequiredIcon"><img src="<%=pcv_strRequiredIcon%>" alt="Required"></div><% End If %>
                            <input type="checkbox" name="custfield<%=cfID%>" id="custfield<%=cfID%>" value="<%= cfValue %>" <%= cfChecked %> class="<%= cfClass %> clearBorder">                                
                        </label>
                        <%if trim(pcArr(11,k))<>"" then%>                                
                            <span class="help-text"><%=pcArr(11,k)%></span>                               
                        <%end if%>                                
                        </div>                           
                    </div>
                
                <% else %>
                
                    <div class="form-group">                            
                         <label for="custfield<%=cfID%>"><%=pcArr(1,k)%>:<% If cfClass = "required" Then %><div class="pcRequiredIcon"><img src="<%=pcv_strRequiredIcon%>" alt="Required"></div><% End If %></label>
                         <input type="text" name="custfield<%=cfID%>" id="custfield<%=cfID%>" value="<%= cfValue %>"  <%if cfMax>"0" then%>maxlength="<%=cfMax%>" <%end if%> class="<%= cfClass %> form-control">
                        <%if trim(pcArr(11,k))<>"" then%>                                
                            <span class="help-block"><%=pcArr(11,k)%></span></span>                                
                        <%end if%>                                
                    </div>
                    
                <% end if %>

            <%
            end if '// if pcv_ShowField=1 then

       Next
    end if

End Sub
%>


<% 'Referrer Field %>
<%
Public Sub referrerFields()

    If (Session("idCustomer")>"0") Then

        query="SELECT IDRefer FROM Customers WHERE idCustomer=" & Session("idCustomer") & ";"
        Set rs = Server.CreateObject("ADODB.Recordset")
        Set rs = connTemp.execute(query)
        If Not rs.Eof Then
            Session("pcSFIDrefer") = rs("IDRefer")                    
        End If
        Set rs = Nothing

    End If
    
    If ((Session("idCustomer")=0) Or ((Session("idCustomer")>"0") And (Session("CustomerGuest")<>"0"))) And (RefNewCheckout="1") Then 
        %>
        <div id="opcReferrer" class="form-group">
            <label for="IDRefer"><%=ReferLabel%><% If ViewRefer="1" Then %><div class="pcRequiredIcon"><img src="<%=pcv_strRequiredIcon%>" alt="Required"></div><% End If %></label>
            <select name="IDRefer" id="IDRefer" class="form-control <% if ViewRefer="1" then %>required<% end if %>">
                    <option value="" <%if Session("pcSFIDrefer")="" then%>selected<%end if%>></option>
                    <%
                    query="Select idrefer, [name] From Referrer Where removed=0 Order By SortOrder;"
                    Set rs = Server.CreateObject("ADODB.Recordset")
                    Set rs = connTemp.execute(query)
                    Do While Not rs.Eof
                        intIdrefer = rs("idrefer")
                        strName = rs("name") 
                        %>
                        <option value="<%=intIdrefer%>" <%if Session("pcSFIDrefer")=trim(intIdrefer) then%>selected<%end if%>><%=strName%></option>
                        <% 
                        rs.movenext
                    Loop
                    Set rs = Nothing 
                    %>
            </select>
        </div>
        <% 
    End If 
    
End Sub 
%>



<% 'Newsletter Area %>
<% 
Public Sub newsletterFields()

    '// If newsletter is enabled, show it for new customer and when existing customers edit their account
    If (AllowNews="1") And (NewsCheckout="1") Then
    %>
    <div id="pcNewsletter" class="checkbox">
        <label for="CRecvNews"><%=NewsLabel%><input type="checkbox" value="1" name="CRecvNews" <% If pcIntRecvNews="1" Then %>checked<% End If %> class="clearBorder" /></label>
    </div>
    <% 
    End If

End Sub 
 %>



<% 'Terms Area %>
<% 
Public Sub termsAndConditions()
    
    '// Terms & Conditions Agreement
    pcv_AgreedToTerms=0
    
    If scTermsShown=1 Then
        pcv_AgreedToTerms=1
    Else
        If (pcAgreeTerms="0" Or pcAgreeTerms="") Then
            pcv_AgreedToTerms=1
        End If
    End if
    
    If scTerms=1 AND pcv_AgreedToTerms=1 then
        %>
        <div class="opcRow checkbox" id="AgreeArea">
            
            <script type=text/javascript>
                var pcCustomerTermsAgreed=0;
            </script>
    
            <%
            Session("pcCustomerTermsAgreed")="0"

            query="SELECT pcStoreSettings_TermsLabel, pcStoreSettings_TermsCopy FROM pcStoreSettings;"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
    
            pcStrTermsLabel=rs("pcStoreSettings_TermsLabel")    
            pcStrTermsCopy=rs("pcStoreSettings_TermsCopy")
            
            if trim(pcStrTermsCopy)<>"" and not isNull(pcStrTermsCopy) then
                pcStrTermsCopy=replace(pcStrTermsCopy, CHR(10),"<br>")
                pcStrTermsCopy=replace(pcStrTermsCopy, "&lt;","<")
                pcStrTermsCopy=replace(pcStrTermsCopy, "&gt;",">")
            end if
    
            set rs=nothing

            If len(pcStrTermsLabel)>0 Then
                pcv_strTermsLabel = pcStrTermsLabel & " " & "<a href=""javascript:;"" id=""ViewTerms"">" & dictLanguage.Item(Session("language")&"_opc_50") & "</a>"
            Else
                pcv_strTermsLabel = "<a href=""javascript:;"" id=""ViewTerms"">" & dictLanguage.Item(Session("language")&"_opc_50") & "</a>"
            End If
            %>
            <label for="AgreeTerms"> 
                <input type="checkbox" value="1" id="AgreeTerms" name="AgreeTerms" class="clearBorder" />&nbsp;<%=pcv_strTermsLabel %></a>
            </label>

             <div id="TermsDialog" class="modal fade">
              <div class="modal-dialog">
                <div class="modal-content">
                  <div class="modal-header">
                    <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                    <h4 class="modal-title"><%=pcStrTermsLabel%></h4>
                  </div>
                  <div id="TermsMsg" class="modal-body">
                    <p><%=pcStrTermsCopy%></p>
                  </div>
                  <div class="modal-footer">
                    <button type="button" class="btn btn-default" data-dismiss="modal">Close</button>
                  </div>
                </div>
              </div>
            </div>

        </div>
        
    <% Else %>
    
        <script type=text/javascript>
            var pcCustomerTermsAgreed=1;
        </script>
    
        <%
        Session("pcCustomerTermsAgreed")="1"					
    End If                 
    %>
    
    <%
    'SB S
    If pcIsSubscription Then

        pcv_strRegAgree = 0
    
        for f=1 to pcCartIndex
        
            pcSubscriptionId = pcCartArray(f,38)
    
            if pcSubscriptionId<>"0" then

                '// Check Package Level
                query= "SELECT SB_Agree, SB_AgreeText FROM SB_Packages WHERE SB_PackageID ="&pcSubscriptionId&" AND  SB_Agree=1" 
                set rss=server.CreateObject("ADODB.RecordSet")
                set rss=connTemp.execute(query)	
                If not rss.eof Then
                    session("pcCartSession")=pcCartArray		
                    session("pcIsRegAgree") = true
                    pSubAgreeText=rss("SB_AgreeText")
                    pcv_strRegAgree=1								 
                Else
                    '// Check Global
                    if scSBRegAgree="1" then
                        session("pcCartSession")=pcCartArray		
                        session("pcIsRegAgree") = true
                        pSubAgreeText=scSBAgreeText
                        pcv_strRegAgree=1
                    end if
                End if 
                set rss = nothing

            end if
        Next
        
    End if
    
    If pcv_strRegAgree=1 Then
    %>

        <div class="opcRow checkbox" id="sb_AgreeArea">
            
            <script type=text/javascript>
                var pcCustomerRegAgreed=0;
            </script>

            <%
            Session("pcCustomerRegAgreed")="0"
      
            pcStrTermsLabel_SB=scSBLang1
      
            if trim(pSubAgreeText)<>"" and not isNull(pSubAgreeText) then
                pSubAgreeText=replace(pSubAgreeText, CHR(10),"<br>")
                pSubAgreeText=replace(pSubAgreeText, "&lt;","<")
                pSubAgreeText=replace(pSubAgreeText, "&gt;",">")
            end if
            %>
            <label for="sb_AgreeTerms">
                <%=pcStrTermsLabel_SB%>
                <a href="javascript:;" id="sb_ViewTerms"><%=scSBLang4%></a>.<input type="checkbox" value="1" id="sb_AgreeTerms" name="sb_AgreeTerms" class="clearBorder" />
            </label>
            <div id="sb_TermsDialog" class="modal fade">
                <div class="modal-dialog">
                    <div class="modal-content">
                        <div class="modal-header">
                            <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                            <h4 class="modal-title"><%=pcStrTermsLabel_SB%></h4>
                        </div>
                        <div id="TermsMsg" class="modal-body">
                            <p><%=pSubAgreeText%></p>
                        </div>
                        <div class="modal-footer">
                            <button type="button" class="btn btn-default" data-dismiss="modal">Close</button>
                        </div>
                    </div>
                </div>
            </div>

        </div>
        
    <% Else %>
    
        <script type=text/javascript>
            var pcCustomerRegAgreed=1;
        </script>
    
        <%
        Session("pcCustomerRegAgreed")="1"				
    End If  				 			
    'SB E
    
End Sub              
%>