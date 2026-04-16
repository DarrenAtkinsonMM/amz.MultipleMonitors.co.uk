<%
pIdProduct = session("idProductRedirect")

BackInStockAddUrl = pcv_marketURL & "api/backinstock/add"
BackInStockRmvUrl = pcv_marketURL & "api/backinstock/remove"
BackInStockSendUrl = pcv_marketURL & "api/backinstock/send"

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Methods
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub pcs_BISCPanelJS()
    %>
    <script type=text/javascript>
        $pc(document).ready(function()
        {
            function pcf_BackInStockSendEmails() {
                    $pc.ajax({
                        type: "POST",
                        url: pcRootUrl + "/includes/apps/pcBackInStock/autosend.asp",
                        timeout: 6000,
                        global: false,
                        success: function(data, textStatus){}
                    });
            }            
            // pcf_BackInStockSendEmails();
        });	
    </script>
    <%
End Sub



Public Sub pcs_BISMenu()
    On Error Resume Next
    Dim query, rs

    If (session("PmAdmin")="19") Then

        If scNM_Auto="0" Then
            query="SELECT idProduct FROM pcBIS_WaitList;"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=connTemp.execute(query)
            if not rs.eof then
            	msg = msg & "<li>There are pending <strong>product stock notification emails</strong> awaiting sending. <a href='nmSendManually.asp'>Send now &gt;&gt;</a></li>"
            end if
            set rs=nothing
        End If
    End If
    Session("msg") = msg

End Sub



Public Sub pcs_BIS_StorefrontJS()
    %>
    <script type="text/javascript">
        //$pc(document).ready(function()
        //{
    
            function isValidEmailAddress(emailAddress) {
                var pattern = /^([a-z\d!#$%&'*+\-\/=?^_`{|}~\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]+(\.[a-z\d!#$%&'*+\-\/=?^_`{|}~\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]+)*|"((([ \t]*\r\n)?[ \t]+)?([\x01-\x08\x0b\x0c\x0e-\x1f\x7f\x21\x23-\x5b\x5d-\x7e\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]|\\[\x01-\x09\x0b\x0c\x0d-\x7f\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]))*(([ \t]*\r\n)?[ \t]+)?")@(([a-z\d\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]|[a-z\d\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF][a-z\d\-._~\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]*[a-z\d\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])\.)+([a-z\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]|[a-z\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF][a-z\d\-._~\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]*[a-z\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])\.?$/i;
                return pattern.test(emailAddress);
            };        
            function setCookie(cname, cvalue, exdays) {
                var d = new Date();
                d.setTime(d.getTime() + (exdays*24*60*60*1000));
                var expires = "expires="+d.toUTCString();
                document.cookie = cname + "=" + cvalue + "; " + expires + "; path=/";
            }        
            function sendBackInStock() {
                <% If pcv_Apparel="1" Then %>
                    if (MyAccept==0)
                    {
                        alert("<%=dictLanguage.Item(Session("language")&"_viewPrd_spmsg2")%>\n");
                        return(false);
                    }
                <% End If %>
                if(!isValidEmailAddress($pc("#nmEmail").val()))
                {
                    alert("<%=dictLanguage.Item(Session("language")&"_Custmoda_16")%>\n");
                    return(false);
                }
                $.ajax({
                  url: pcRootUrl + "/includes/apps/pcBackInStock/notifications.asp",
                  type:"POST",
                  headers: { 
                    "App" : "NetSource ProductCart"
                  },
                  data: "action=add&nmEmail=" + $pc("#nmEmail").val() + "&idproduct=" + document.additem.idproduct.value + "&quantity=" + document.additem.quantity.value,
                  success: function (data) {
                    var tmpStr1=data.split("||");
                    if (tmpStr1[0]=="SUCCESS")
                    {
                        $pc("#nmResult").html("<div class='pcSuccessMessage'>" + tmpStr1[1] + "</div>");
                        $pc("#nmResult").show();
                        $pc("#nmArea").hide();
                    }
                    else
                    {
                        $pc("#nmResult").html("<div class='pcErrorMessage'>" + tmpStr1[1] + "</div>");
                        $pc("#nmResult").show();
                        $pc("#nmArea").show();
                    }
                    $('#bis_modal').modal('hide');
                  }
                });
            }        
            function rmvBackInStock(tmpGUID) {           
                $.ajax({
                  url: pcRootUrl + "/includes/apps/pcBackInStock/notifications.asp",
                  type: "POST",
                  headers: {
                    "App" : "NetSource ProductCart"
                  },
                  data: "action=rmv&nmGUID=" + tmpGUID + "&idproduct=" + document.additem.idproduct.value,
                  success: function (data) {
                    var tmpStr1=data.split("||");
                    if (tmpStr1[0]=="SUCCESS")
                    {
                        $pc("#nmResult").html("<div class='pcSuccessMessage'>" + tmpStr1[1] + "</div>");
                        $pc("#nmResult").show();
                        $pc("#nmArea").hide();
                    }
                    else
                    {
                        $pc("#nmResult").html("<div class='pcErrorMessage'>" + tmpStr1[1] + "</div>");
                        $pc("#nmResult").show();
                        $pc("#nmArea").show();
                    }                    
                  }
                });
            }
        
        //});	
    </script>
    <%
End Sub



Public Sub pcs_PrdBackInStock(pcv_strWidgetTemplate)

    Dim query, rs
    Dim nmTurnOn, nmMsg, nmAuto, nmBText

    nmMsg=""
    nmAuto=0
    nmBText=""

    nmTurnOn=scNM_IsEnabled      
    nmMsg=scNM_Msg
    nmAuto=scNM_Auto
    if IsNull(nmAuto) OR nmAuto="" then
        nmAuto=0
    end if
    nmBText=scNM_ButtonText
    if nmBText="" then
        nmBText="Email When Available"
    end if

    If session("idcustomer")<>"" And session("idcustomer")<>"0" Then
        query="SELECT name,lastName,email FROM customers WHERE idCustomer=" & session("idcustomer")
        set rs=conntemp.execute(query)
        if not rs.eof then
            Session("pcSFFromEmail") = rs("email")
        end if
        set rs=nothing
    End If

    If (nmTurnOn=1) And (pcf_CheckShowNM=1) Then
        pcs_BIS_StorefrontJS()
        %>
        <div id="nmArea" class="pcShowAddToCart">
            <%
            hadNM=0
            If Request.Cookies("BackInStockPrdID" & pIdProduct)<>"" Then
                pcv_strGuid = Request.Cookies("BackInStockPrdID" & pIdProduct)
                pcv_strGuid = getUserInput(pcv_strGuid, 0)
                hadNM="1"
            End If
            
            If hadNM="1" Then
                %>
                <div class="pcSuccessMessage">        
                    <%=dictLanguage.Item(Session("language")&"_BackInStock_1a") %>
                    &nbsp;&nbsp;
                    <input name="nmButton" value="<%=dictLanguage.Item(Session("language")&"_css_cancel") %>" class="btn btn-default" onclick="javascript:rmvBackInStock('<%=pcv_strGuid %>');" type="button">
                    
                </div>
            <% Else %>
            
                <% If pcv_strWidgetTemplate = "minimal" Then %>
                    <!--#include file="widgets/minimal.asp"-->
                <% Else %>
                    <!--#include file="widgets/modal.asp"-->
                <% End If %>
                
            <% End If %>
        </div>
        <div id="nmResult" style="display:none"></div>
        <%
    End If

End Sub


Public Sub pcs_AddWaitList()
    Dim query,rs
    
    tmpID = session("idProductRedirect")
    
    query="SELECT Products.idProduct FROM Products INNER JOIN pcBIS_ListEmails ON Products.idProduct=pcBIS_ListEmails.idproduct WHERE pcBIS_ListEmails.idProduct=" & tmpID & " AND pcBIS_ListEmails.Sent=0 AND Products.stock>0;"

    set rs=connTemp.execute(query)
    if not rs.eof then
        set rs=nothing
        query="DELETE FROM pcBIS_WaitList WHERE idProduct=" & tmpID & ";"
        set rs=connTemp.execute(query)
        set rs=nothing
        query="INSERT INTO pcBIS_WaitList (idProduct) VALUES (" & tmpID & ");"
        set rs=connTemp.execute(query)
        set rs=nothing
    end if
    set rs=nothing

End Sub


Public Sub pcs_AddWaitListParent(tmpID)
    Dim query,rs,pcArr,i,intCount
    
    query="SELECT Products.idProduct FROM Products INNER JOIN pcBIS_ListEmails ON Products.idProduct=pcBIS_ListEmails.idproduct WHERE pcBIS_ListEmails.ParentProductID=" & tmpID & " AND pcBIS_ListEmails.Sent=0 AND Products.stock>0;"
    set rs=connTemp.execute(query)
    if not rs.eof then
        pcArr=rs.getRows()
        intCount=ubound(pcArr,2)
        set rs=nothing
        For i=0 to intCount
        query="DELETE FROM pcBIS_WaitList WHERE idProduct=" & pcArr(0,i) & ";"
        set rs=connTemp.execute(query)
        set rs=nothing
        query="INSERT INTO pcBIS_WaitList (idProduct) VALUES (" & pcArr(0,i) & ");"
        set rs=connTemp.execute(query)
        set rs=nothing
        Next
    end if
    set rs=nothing

End Sub


Public Sub pcs_RmvWaitList(tmpID)
    Dim query,rs
    
    query="DELETE FROM pcBIS_WaitList WHERE idProduct=" & tmpID & ";"
    set rs=connTemp.execute(query)
    set rs=nothing

End Sub

Public Function pcf_CheckShowNM()
    Dim tmpShow

    tmpShow=1

	if pFormQuantity="-1" and NotForSaleOverride(session("customerCategory"))=0 then
	    tmpShow=1
    else 
		If scorderlevel = "0" OR pcf_WholesaleCustomerAllowed Then
			if pcf_OutStockPurchaseAllow then
				If ((pserviceSpec<>0) AND ((pnoprices>0) OR (pPrice=0) OR (scConfigPurchaseOnly=1))) or ((iBTOQuoteSubmitOnly=1) and (pserviceSpec<>0)) then 
					tmpShow=0
				else 
					tmpShow=0
				end if 
			end if
		end if
		If (not pcf_OutStockPurchaseAllow) OR (scorderlevel = "2") OR ((pcf_WholesaleCustomerAllowed or scorderlevel = "1") and session("customerType")<>"1") then
			tmpShow=1
		End if
	end if
    
    'response.Write(pcf_OutStockPurchaseAllow & "<br/>")
    'response.Write(scorderlevel & "<br/>")
    'response.Write(pcf_WholesaleCustomerAllowed & "<br/>")
    'response.Write(session("customerType") & "<br/>")
    
    '// If customer logged in, and not showing, then double check cookies were not deleted...
    If len(Session("pcSFFromEmail"))>0 Then
        pcv_boolCookieIsDeleted = False
        'pcv_strGuid = "28563FE2-3754-4797-834A-BCF053801C50" '// pcf_getGuidBIS(Session("pcSFFromEmail"), pIdProduct)
        If len(pcv_strGuid)>0 Then
            pcv_boolCookieIsDeleted = True
        End If
        If pcv_boolCookieIsDeleted Then
            Response.Cookies("BackInStockPrdID" & pIdProduct) = pcv_strGuid
            Response.Cookies("BackInStockPrdID" & pIdProduct).Expires = Date()+365  
            tmpShow=1
        End If
	End If
    pcf_CheckShowNM = tmpShow
    
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Methods
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>