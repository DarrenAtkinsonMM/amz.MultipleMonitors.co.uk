<%
' ------------------------------------------------------
'Start SDBA - Notify Drop-Shipping
' ------------------------------------------------------
If scShipNotifySeparate="1" And pcCartIndex>1 Then
    tmp_showmsg=0
    
    For f=1 To pcCartIndex
        tmp_idproduct=pcCartArray(f,0)        
        
        query = "SELECT pcProd_IsDropShipped FROM products WHERE idproduct=" & tmp_idproduct & " AND pcProd_IsDropShipped=1;"
        Set rs = connTemp.execute(query)
        If err.number<>0 Then
            call LogErrorToDatabase()
            set rs=nothing
            call closedb()
            response.redirect "techErr.asp?err="&pcStrCustRefID
        End If
        If Not rs.eof Then
            tmp_showmsg=1
            exit for
        End If
        Set rs = nothing
        
    Next

    If tmp_showmsg = 1 Then
        %>
        <div class="pcAttention">
            <img src="<%=pcf_getImagePath("images","sds_boxes.gif")%>" alt="<%response.write ship_dictLanguage.Item(Session("language")&"_dropshipping_msg")%>">
            <%response.write ship_dictLanguage.Item(Session("language")&"_dropshipping_msg")%>
        </div>
        <%
    End If
End If

' ------------------------------------------------------
' END SDBA - Notify Drop-Shipping
' ------------------------------------------------------

' ------------------------------------------------------
' Start Cross Selling - Notify Accessory Added
' ------------------------------------------------------
If Session("cs_Accessory") <> "" Then 
    cs_Msg = replace(dictLanguage.Item(Session("language")&"_showcart_28"),"<main product name>", Session("cs_Accessory")) %>
    <div class="pcAttention">
        <%=cs_Msg %>
    </div>
    <% 
    Session("cs_Accessory") = ""
End If
' ------------------------------------------------------
' End Cross Selling - End Accessory Added
' ------------------------------------------------------
%>