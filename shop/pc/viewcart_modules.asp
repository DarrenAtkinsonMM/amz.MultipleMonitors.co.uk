<%
Public Sub pcs_ShowCartWeight()
	' ------------------------------------------------------
	' START - Show shopping cart content weight
	' ------------------------------------------------------
    %>
    <div class="pcTotalCartWeight" data-ng-if="Evaluate(shoppingcart.showCartWeight)">
        <div class="pcSpacer"></div>
        <b><%response.write ship_dictLanguage.Item(Session("language")&"_viewCart_a")%></b>
        <span data-ng-show="!IsEmpty(shoppingcart.kilos)">
            {{shoppingcart.kilos}}                
            <span data-ng-if="!IsEmpty(shoppingcart.weightG)">{{shoppingcart.weightG}}</span>
        </span>
        <span data-ng-show="!IsEmpty(shoppingcart.pounds)">
            {{shoppingcart.pounds}}
            <span data-ng-if="!IsEmpty(shoppingcart.weightOZ)">{{shoppingcart.weightOZ}}</span>
        </span>
    </div>
    <% 
	' ------------------------------------------------------
	' END - Show shopping cart content weight
	' ------------------------------------------------------    
End Sub
%>


<%
Public Sub pcs_EstimateShipping() 
	' ------------------------------------------------------
	' START - Show estimated shipping charges link
	' ------------------------------------------------------    
    Dim iShipService
    iShipService=0    
    
    query="SELECT * FROM shipService WHERE serviceActive=-1;"
    set rs=server.CreateObject("ADODB.RecordSet")
    set rs=conntemp.execute(query)
    if rs.eof then
        iShipService = 1
    end if
    set rs=nothing

    If iShipService=0 Then
        If (scShowEstimateLink="-1" Or request("show")="1") Then
            %>
            <div data-ng-show="shoppingcart.weight>0">
            	<h2><%=ship_dictLanguage.Item(Session("language")&"_viewCart_b")%></h2>
				<div class="row" data-ng-init="showEstShip()" id="pcEstimateShippingArea"></div>
            </div>
            <%
        End If
    End If
	' ------------------------------------------------------
	' END - Show estimated shipping charges link
	' ------------------------------------------------------    
End Sub
%>


<%
Public Sub pcs_ShowPromoMessage()
	' ------------------------------------------------------
	' START - Promotions
	' ------------------------------------------------------
	If PromoMsgStr<>"" Then
        %>
        <div class="pcPromoMessage">
            <span class="pcLargerText"><%=dictLanguage.Item(Session("language")&"_showcart_29")%></span>
            <ul><%=PromoMsgStr%></ul>
        </div>        
        <%
	End If
	' ------------------------------------------------------
	' END - Promotions
	' ------------------------------------------------------
End Sub
%>


<%
Public Sub pcs_ShowGiftWrapOverview()
	' ------------------------------------------------------
	' START - Show Gift Wrapping Overview
	' ------------------------------------------------------	
	If (session("Cust_GW")="1") And (pcv_GWDetails<>"") And (session("Cust_GWText")="1") Then 
        %>
        <div class="pcGiftWrapMessage"><%=pcf_FixHTMLContentPaths(pcv_GWDetails)%></div>
        <% 
    End If	
	' ------------------------------------------------------
	' END - Show Gift Wrapping Overview
	' ------------------------------------------------------
    %>
<%
End Sub
%>


<%
Public Sub pcs_ShowCrossSelling()
	' ------------------------------------------------------
	' START - Cross selling
	' ------------------------------------------------------
    Dim scCS, cs_showprod, cs_showcart, cs_showimage, crossSellText, pcv_strCSQuery, pcv_strHaveResults, pcv_intProductCount, pcArray_CSRelations

    '// Get Cross Sell Settings - Sitewide 
    query = "SELECT cs_status,cs_showprod,cs_showcart,cs_showimage,crossSellText,cs_CartViewCnt, cs_showNFS FROM crossSelldata WHERE id=1;"
    Set rs = server.CreateObject("ADODB.RecordSet")
    Set rs = conntemp.execute(query)
    if err.number<>0 then
        call LogErrorToDatabase()
        set rs=nothing
        call closedb()
        response.redirect "techErr.asp?err="&pcStrCustRefID
    end if		
    If Not rs.Eof Then
        scCS = rs("cs_status")
        cs_showprod = rs("cs_showprod")
        cs_showcart = rs("cs_showcart")
        cs_showimage = rs("cs_showimage")
        crossSellText = rs("crossSellText")
        cs_ViewCnt = rs("cs_CartViewCnt")
        cs_showNFS = rs("cs_showNFS")
    End If
    Set rs = Nothing
  
    '// Do Not Display if CS is turned "Off"
    If scCS=-1 And cs_showcart="-1" Then
    
        '// Check if there are items for cross sell in the database
        pcv_strCSQuery = ""
        cs_Source=1
        pcv_cs_headerflag=0
        tmp_PList=""
        pcv_strCSQuery = pcv_strCSQuery & "cs_relationships.idproduct=0"  
        tCnt=0
        
        For f = pcCartIndex To 1 Step -1
		
			pidproduct = pcf_ProductIdFromArray(pcCartArray, f)

            If inStr(","& tmp_PList &",",","& pidproduct &",")=0 Then  
            
                If tCnt=0 Then
                    tmp_PList=tmp_PList & pidproduct 				
                Else
                    tmp_PList=tmp_PList & "," & pidproduct  
                End If				
            
                '// Build Cross Sell Relationship List
                pcv_strCSQuery = pcv_strCSQuery & " OR cs_relationships.idproduct="& pidproduct
                
            End If '// Check existing IDProduct
            tCnt = tCnt + 1
            
        Next '// Move to next product in the cart
        tCnt = 0
        
        If len(tmp_PList)>0 Then
            pcv_strCSunavailable = "(cs_relationships.idrelation NOT IN ("& tmp_PList &")) AND " 
        Else
            pcv_strCSunavailable = ""
        End If

        query="SELECT cs_relationships.idproduct, cs_relationships.idrelation, cs_relationships.cs_type, cs_relationships.discount, cs_relationships.ispercent,cs_relationships.isRequired, products.servicespec, products.price, products.description FROM cs_relationships INNER JOIN products ON cs_relationships.idrelation=products.idProduct WHERE ("& pcv_strCSunavailable &"("& pcv_strCSQuery &") AND ((products.active)=-1) AND ((products.removed)=0)) ORDER BY cs_relationships.num,cs_relationships.idrelation;"
        Set rs = server.createobject("adodb.recordset")
        set rs = conntemp.execute(query)	
        If err.number<>0 Then
            call LogErrorToDatabase()
            Set rs = Nothing
            call closedb()
            response.redirect "techErr.asp?err="&pcStrCustRefID
        End If
        pcv_strHaveResults = 0
        If Not rs.Eof Then
            pcArray_CSRelations = rs.getRows()
            pcv_intProductCount = UBound(pcArray_CSRelations,2)+1
            pcv_strHaveResults=1
        End If
        Set rs = Nothing		

        tCnt = Cint(0)	
        pcsFilterOverRide = "1"

        If pcv_strHaveResults=1 Then	
        
            '// Start: viewPrd
            cs_pCnt=Cint(0)
            cs_pOptCnt=Cint(0)
            cs_pAddtoCart=Cint(0)
            pcv_intCategoryActive=2	'// set bundle group to inactive
            pcv_intAccessoryActive=2 '// set accessories group to inactive
            cs_count=Cint(0)
            session("listcross")=""
            
            Do While ( (tCnt < pcv_intProductCount) And (tCnt < cs_ViewCnt))				
                
                pidrelation = pcArray_CSRelations(1,tCnt) '// rs("idrelation")
                pcsType = pcArray_CSRelations(2,tCnt) '// rs("cs_type")			
                pDiscount = pcArray_CSRelations(3,tCnt) '// rs("discount")
                cs_pserviceSpec = pcArray_CSRelations(6,tCnt)				
                pcArray_CSRelations(8,tCnt) = 1
                
                If (pcsType="Accessory") Or ((pcsType="Bundle") And (pDiscount>0)) Then

                    '// CHECK IF BUNDLES GROUP HAS AT LEAST ONE PRODUCT FROM AN ACTIVE CATEGORY		
                    '// CHECK IF ACCESSORIES GROUP HAS AT LEAST ONE PRODUCT FROM AN ACTIVE CATEGORY  						
                    If Session("customerType")=1 Then
                        pcv_strCSTemp=""
                    Else
                        pcv_strCSTemp=" AND pccats_RetailHide<>1 "
                    End If
                                                        
                    query="SELECT categories_products.idProduct "
                    query=query+"FROM categories_products " 
                    query=query+"INNER JOIN categories "
                    query=query+"ON categories_products.idCategory = categories.idCategory "
                    query=query+"WHERE categories_products.idProduct="& pidrelation &" AND iBTOhide=0 " & pcv_strCSTemp & " "
                    query=query+"ORDER BY priority, categoryDesc ASC;"	
                    Set rsCheckCategory = server.CreateObject("ADODB.RecordSet")
                    Set rsCheckCategory = conntemp.execute(query)									
                    If Not rsCheckCategory.Eof Then
                        If pcsType="Accessory" Then
                            pcv_intAccessoryActive = 1
                        End If
                        If pcsType="Bundle" Then							
                            pcv_intCategoryActive = 1
                        End If	
                    Else
                        session("listcross")=session("listcross") & "," & pidrelation					
                    End If	
                    Set rsCheckCategory = Nothing
 
                End If '// If (pcsType="Bundle") AND (pDiscount>0) Then	

                pcv_intOptionsExist = 0
                
                '// CHECK FOR REQUIRED OPTIONS							
                pcv_intOptionsExist = pcf_CheckForReqOptions(pidrelation) '// check options function (1=YES, 2=NO)			


                '// CHECK FOR REQUIRED INPUT FIELDS
                If pcv_intOptionsExist = 2 Then
                    pcv_intOptionsExist = pcf_CheckForReqInputFields(pidrelation)
                End If				


                '// VALIDATE
                If (cs_pserviceSpec=true) Or (pcv_intOptionsExist = 1) Then
                    If pcsType <> "Accessory" Then
                        cs_pOptCnt = cs_pOptCnt + 1
                    End If
                    pcArray_CSRelations(8,tCnt) = 0					
                End If	
                If pcsType <> "Accessory" Then
                    cs_pCnt = cs_pCnt + 1 
                End If
                tCnt = tCnt + 1	
                            
            Loop  '// Do While ( (tCnt < pcv_intProductCount) And (tCnt < cs_ViewCnt))
        
            If pcv_intAccessoryActive=1 Then
                
                cs_DisplayCheckBox=0
                cs_Bundle=0

                If pcv_cs_headerflag=0 Then
                
                    '// Only display header once
                    pcv_cs_headerflag=1 
                    %>
                    <div class="pcSectionTitle">
                        <%=crossSellText%>
                    </div>
                    <% 
                End If %>

                <% If cs_showImage="-1" Then %>
                <!--#include file="cs_img.asp"-->
                <% Else %>
                <!--#include file="cs.asp"-->
                <% End If %>

                <div class="pcSpacer"></div>
            <%		
            End If '// If pcv_intAccessoryActive=1 Then				
                        
        End If '// If pcv_strHaveResults=1 Then	
            
    End If '// If scCS=-1 And cs_showcart="-1" Then
    
    session("listcross")=""
	' ------------------------------------------------------
	' END - Cross selling
	' ------------------------------------------------------
End Sub
%>