<%
Function pcf_AddToCartQV(IdProduct)
Dim query,rs
	IF scQuickBuy = "1" THEN '// Feature is inactive
			pcf_AddToCartQV=False
			Exit Function
			
	ELSE '// Feature is active
	
		If IdProduct<>"" AND NOT isNULL(IdProduct) Then
	
			pcf_AddToCartQV=True

			pcv_intSkipDetailsPage=0
			query="SELECT pcProd_SkipDetailsPage FROM Products WHERE idProduct=" & IdProduct & ";"
			set rs=connTemp.execute(query)
			if not rs.eof then
				pcv_intSkipDetailsPage=rs("pcProd_SkipDetailsPage")
				if isNull(pcv_intSkipDetailsPage) or pcv_intSkipDetailsPage="" then
					pcv_intSkipDetailsPage=0
				end if
			end if
			set rs=nothing
			
			If pcv_intSkipDetailsPage=1 Then
				pcf_AddToCartQV=False
				Exit Function
			End If
			
			If NOT (scorderlevel = "0" OR pcf_WholesaleCustomerAllowed) Then '// if [everyone] OR [wholesale w/ wholesale only turned on]
				pcf_AddToCartQV=False
				Exit Function
			End If
	
	
			If NOT pcf_OutStockPurchaseAllow Then  
				pcf_AddToCartQV=False
				Exit Function
			End If
	
	
			If pFormQuantity="-1" and NotForSaleOverride(session("customerCategory"))=0 Then
				pcf_AddToCartQV=False
				Exit Function
			End If			
			
		Else  '// If IdProduct<>"" AND NOT isNULL(IdProduct) Then
			
			pcf_AddToCartQV=False
			Exit Function
			
		End If  '// If IdProduct<>"" AND NOT isNULL(IdProduct) Then
		
	END IF '// Feature is active
	
End Function

Function pcf_QuickViewBtn(pIdProduct)
    If session("Mobile")=""  Then
    %>
    <div class="pcQuickView">
        <a href="javascript:CallQuickView(<%=pIdProduct%>);" class="pcQuickViewTrigger btnHover">
            <img src="images/qv-button-hover.png" alt="Quick View" class="btnHorverImg" onmousedown="this.src='images/qv-button-down.png'" onmouseout="this.src='images/qv-button-hover.png'"/>
        </a>
    </div>
    <%
    End If
End Function
%>
