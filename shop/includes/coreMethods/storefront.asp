<%
Public Function pcf_TotalReviewCount(pIDProduct)
    Dim rsReviewCount
    query = "SELECT COUNT(*) as ct FROM pcReviews WHERE pcRev_IDProduct=" & pIDProduct & " AND pcRev_Active=1"
    Set rsReviewCount = server.CreateObject("ADODB.RecordSet")
    Set rsReviewCount = connTemp.execute(query)
    intCount=0
    If Not rsReviewCount.Eof Then
        intCount = clng(rsReviewCount("ct"))
    End If
    Set rsReviewCount = Nothing
    pcf_TotalReviewCount = intCount
End Function

Public Sub pcf_do301Redirect(url)
    Response.Status = "301 Moved Permanently" 
    Response.AddHeader "Location", url
    Response.End
End Sub

Public Sub safeExecute(path)
    on error resume next
    server.Execute(path)
    err.clear
End Sub

Public Function pcf_FormatPhoneNumber(phoneNum)

    phoneNum = replace(phoneNum, "(", "")
    phoneNum = replace(phoneNum, ")", "")
    phoneNum = replace(phoneNum, "-", "")
    phoneNum = replace(phoneNum, ".", "")
    phoneNum = replace(phoneNum, " ", "")
    If len(phoneNum)=10 Then
        phoneNum = "(" & Mid(phoneNum, 1, 3) & ")" & Mid(phoneNum, 4, 3) & "-" & Mid(phoneNum, 7, 4)
    Else
        phoneNum = scCompanyPhoneNumber
    End If
    pcf_FormatPhoneNumber = phoneNum
    
End Function

Public Function GenerateSearchURL(pIdCategory)

    '*******************************
    ' Generate Base Navigation Url
    '*******************************
    baseNavUrl = pcStrPageName
    baseNavUrl = baseNavUrl & "?incSale=" & incSale 
    baseNavUrl = baseNavUrl & "&IDSale=" & tmpIDSale 
    baseNavUrl = baseNavUrl & "&ProdSort=" & ProdSort
    baseNavUrl = baseNavUrl & "&PageStyle=" & pcPageStyle
    baseNavUrl = baseNavUrl & "&customfield=" & pcustomfield
    baseNavUrl = baseNavUrl & "&SearchValues=" & pCValues
    baseNavUrl = baseNavUrl & "&exact=" & intExact
    baseNavUrl = baseNavUrl & "&keyword=" & tKeywords
    baseNavUrl = baseNavUrl & "&priceFrom=" & pPriceFrom
    baseNavUrl = baseNavUrl & "&priceUntil=" & pPriceUntil
    If Len(pIdCategory)>0 Then
        baseNavUrl = baseNavUrl & "&idCategory=" & pIdCategory
    End If
    baseNavUrl = baseNavUrl & "&IdSupplier=" & IdSupplier
    baseNavUrl = baseNavUrl & "&withStock=" & pWithStock
    baseNavUrl = baseNavUrl & "&IDBrand=" & IDBrand
    baseNavUrl = baseNavUrl & "&SKU=" & pSearchSKU
    baseNavUrl = baseNavUrl & "&order=" & strORD
    baseNavUrl = baseNavUrl & pcv_strCSFieldQuery
    
    GenerateSearchURL = baseNavUrl

End Function


Public Sub storeSSLRedirect(pcStrIntSSLPage)

    If scSSL="1" And scIntSSLPage=pcStrIntSSLPage Then
      If (Request.ServerVariables("HTTPS") = "off") Then
          Dim xredir__, xqstr__
          xredir__ = "https://" & Request.ServerVariables("SERVER_NAME") & _
          Request.ServerVariables("SCRIPT_NAME")
          xqstr__ = Request.ServerVariables("QUERY_STRING")
          If xqstr__ <> "" Then xredir__ = xredir__ & "?" & xqstr__
          Response.redirect xredir__
      End If
    End If
    
End Sub



Public Sub storeURLRedirect()

    '// Redirect to maintain consistent URL. 
		'// If enabled, will redirect to the domain configured in the store constants. (e.g. shop.mystore.com)
    intDoRedirect = scConURL ' 0 = NO; 1 = YES
    
    IF intDoRedirect = 1 THEN
        strOrigDomain = Request.ServerVariables("HTTP_HOST")
        strPath = Request.ServerVariables("URL")
        strQueryString = Request.ServerVariables("QUERY_STRING")
        strHttps = ucase(Request.ServerVariables("HTTPS"))
        strNewDomain = getDomainFromURL(scStoreURL)

        '// Clean up and concatenate
        if trim(strQueryString)<>"" then
            strQueryString = "?" & strQueryString
        end if
        if strHttps="ON" then
            strURLPrefix = "https://"
            strURL = strURLPrefix & strNewDomain & strPath & strQueryString
        else
            strURLPrefix = "http://"
            strURL = strURLPrefix & strNewDomain & strPath & strQueryString
        end if
        If len(strURL)>0 Then
            If instr(strURL,"404;")>0 Then
                strURL = "/" & Right(strURL,Len(strURL)-instr(strURL,":80")-3)			
                if strHttps="ON" then                    
                    strURL = strURLPrefix & strNewDomain & left(strURL,len(strURL))
                else                    
                    strURL = strURLPrefix & strNewDomain & left(strURL,len(strURL))
                end if
            end if
        End If
				
        '// Redirect to store URL      
				if lcase(strOrigDomain) <> lcase(strNewDomain) then
						Response.Status="301 Moved Permanently" 
						Response.AddHeader "Location", strURL
				end if
    END IF

End Sub



Public Function getFirstCategoryID(pIdProduct, pIdCategory)

	'// If category ID doesn't exist, get the first category that the product has been assigned to, filtering out hidden categories
	if pIdCategory=0 or trim(pIdCategory)="" then

        query="SELECT categories_products.idCategory FROM categories_products INNER JOIN categories ON categories_products.idCategory = categories.idCategory WHERE categories_products.idProduct="& pIdProduct &" AND categories.iBTOhide<>1 AND categories.pccats_RetailHide<>1"
		set rsCat=server.CreateObject("ADODB.RecordSet")
		set rsCat=conntemp.execute(query)
			if err.number<>0 then
				call LogErrorToDatabase()
				set rsCat=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
		if not rsCat.EOF then
			pIdCategoryTemp=rsCat("idCategory")
		else
			pIdCategoryTemp=1
		end if
		set rsCat=nothing
 
	end if
    
    getFirstCategoryID = pIdCategoryTemp
    
End Function

Public Function pcf_getStoreMsg(msg)
  msgStr = ""

  select case msg
  case 1
    msgStr = dictLanguage.Item(Session("language")&"_showcart_1")&"<br><br><a href=default.asp>"&GetButtonLink("continueshop")&"</a>"    
  case 2  
    msgStr = dictLanguage.Item(Session("language")&"_forgotpassworderror") &"<br><br><a href=""javascript:history.go(-1)"">"&GetButtonLink("back")&"</a>"
  case 3
    msgStr = dictLanguage.Item(Session("language")&"_advSrcb_2") &"<br><br><a href=search.asp>"&GetButtonLink("back")&"</a>"
  case 4
    msgStr = dictLanguage.Item(Session("language")&"_advSrcb_2") &"<br><br><a href=search.asp>"&GetButtonLink("back")&"</a>"
  case 5
    msgStr = dictLanguage.Item(Session("language")&"_advSrcb_2") &"<br><br><a href=search.asp>"&GetButtonLink("back")&"</a>"
  case 6
    msgStr = dictLanguage.Item(Session("language")&"_advSrcb_2") &"<br><br><a href=search.asp>"&GetButtonLink("back")&"</a>"
  case 7
    msgStr = dictLanguage.Item(Session("language")&"_advSrcb_2") &"<br><br><a href=search.asp>"&GetButtonLink("back")&"</a>"
  case 8
    dstr=replace(scStoreMsg,"''","~~~")
    dstr=replace(dstr,"'","""")
    dstr=replace(dstr,"~~~","'")
    msgStr = dstr
  case 9
    msgStr = dictLanguage.Item(Session("language")&"_checkout_1")&"<br><br><a href=default.asp>"&GetButtonLink("continueshop")&"</a>"    
  case 10
    msgStr = dictLanguage.Item(Session("language")&"_msg_202")
  case 11
    msgStr = dictLanguage.Item(Session("language")&"_CustviewPastD_16")
  case 12
    msgStr = dictLanguage.Item(Session("language")&"_msg_12")
  case 13
    msgStr = dictLanguage.Item(Session("language")&"_msg_13")
  case 14
    msgStr = dictLanguage.Item(Session("language")&"_chooseShpmnt_1") & "<br><br><a href=""javascript:history.go(-1)"">"&GetButtonLink("back")&"</a>" 
  case 15
    msgStr = dictLanguage.Item(Session("language")&"_chooseShpmnt_2") &"<br><br><a href='checkout.asp?cmode=2'>"&GetButtonLink("back")&"</a>"
  case 16
    msgStr = dictLanguage.Item(Session("language")&"_cRec_1")
  case 17
    msgStr = dictLanguage.Item(Session("language")&"_cRec_2") &"<br><br><a href=""javascript:history.go(-1)"">"&GetButtonLink("back")&"</a>"
  case 18
    msgStr = "" 
  case 19
    msgStr = dictLanguage.Item(Session("language")&"_cRemv") 
  case 20
    msgStr = dictLanguage.Item(Session("language")&"_advSrcb_2") &"<br><br><a href=search.asp>"&GetButtonLink("back")&"</a>"       
  case 21
    msgStr = dictLanguage.Item(Session("language")&"_advSrcb_2") &"<br><br><a href=search.asp>"&GetButtonLink("back")&"</a>" 
  case 22
    msgStr = dictLanguage.Item(Session("language")&"_Custmodb_2")
  case 23
    msgStr = dictLanguage.Item(Session("language")&"_CustPastAdd_2")  
  case 24
    msgStr = dictLanguage.Item(Session("language")&"_CustPastAdd_3")
  case 25
    msgStr = dictLanguage.Item(Session("language")&"_CustPastAdd_4")
  case 26
    msgStr = dictLanguage.Item(Session("language")&"_CustPastAdd_1")
  case 27
    msgStr = dictLanguage.Item(Session("language")&"_additem_6")
  case 28
    msgStr = dictLanguage.Item(Session("language")&"_CustRegb_3")
  case 29
    msgStr = dictLanguage.Item(Session("language")&"_CustRegb_1")
  case 30
    msgStr = dictLanguage.Item(Session("language")&"_CustRegb_2")
  case 31
    dstr=replace(scStoreMsg,"''","~~~")
    dstr=replace(dstr,"'","""")
    dstr=replace(dstr,"~~~","'")
    msgStr = dstr
  case 32
    msgStr = dictLanguage.Item(Session("language")&"_Custvb_1") &"<br><br><a href=Checkout.asp?cmode=1&redirectUrl=" & Server.URLEncode(session("redirectUrlLI")) & ">"&GetButtonLink("back")&"</a>"    
  case 33
    msgStr = dictLanguage.Item(Session("language")&"_Custvb_2") &"<br><br><a href=Checkout.asp?cmode=1&redirectUrl=" & Server.URLEncode(session("redirectUrlLI")) & ">"&GetButtonLink("back")&"</a>" 
  case 34
    msgStr = dictLanguage.Item(Session("language")&"_CustviewPast_1") &"<br><br><a href=""CustPref.asp"">"&GetButtonLink("back")&"</a>"
  case 35
    msgStr = dictLanguage.Item(Session("language")&"_CustviewPastD_1")
  case 36
    msgStr = dictLanguage.Item(Session("language")&"_Custwl_1")
  case 37
    msgStr = dictLanguage.Item(Session("language")&"_CustwlRmv_1")
  case 38
    msgStr = dictLanguage.Item(Session("language")&"_msg_38")
  case 39
    msgStr = dictLanguage.Item(Session("language")&"_instPrd_B") & "<br><br><a href=""javascript:history.go(-1)"">"&GetButtonLink("back")&"</a>"   
  case 40
    msgStr = dictLanguage.Item(Session("language")&"_instPrd_C")& "<br><br><a href=""javascript:history.go(-1)"">"&GetButtonLink("back")&"</a>"     
  case 41
    msgStr = dictLanguage.Item(Session("language")&"_instPrd_D")    
  case 42
    msgStr = dictLanguage.Item(Session("language")&"_instPrd_E")
  case 43
    msgStr = dictLanguage.Item(Session("language")&"_instPrd_E")  
  case 44
    msgStr = dictLanguage.Item(Session("language")&"_instPrd_E")
  case 45
    msgStr = dictLanguage.Item(Session("language")&"_instPrd_E")
  case 46
    msgStr = dictLanguage.Item(Session("language")&"_instPrd_E")
  case 47
    msgStr = dictLanguage.Item(Session("language")&"_instPrd_A")
  case 48
    msgStr = ""     
  case 49
    msgStr = dictLanguage.Item(Session("language")&"_instPrd_C")
  case 50
    msgStr = dictLanguage.Item(Session("language")&"_instPrd_B") & "<br><br><a href=""javascript:history.go(-1)"">"&GetButtonLink("back")&"</a>"  
  case 51
    msgStr = dictLanguage.Item(Session("language")&"_instPrd_C")& "<br><br><a href=""javascript:history.go(-1)"">"&GetButtonLink("back")&"</a>" 
  case 52
    msgStr = dictLanguage.Item(Session("language")&"_instPrd_D")
  case 53
    msgStr = "" 
  case 54
  msgStr = dictLanguage.Item(Session("language")&"_login_2") &"<br><br><a href=checkout.asp>"&GetButtonLink("back")&"</a>" 
  case 55
    msgStr = dictLanguage.Item(Session("language")&"_login_2")&"<br><br><a href=checkout.asp>"&GetButtonLink("back")&"</a>"  
  case 56
    msgStr = dictLanguage.Item(Session("language")&"_login_3")&"<br><br><a href=checkout.asp>"&GetButtonLink("back")&"</a>"    
  case 57
    msgStr = ""  
  case 58
  msgStr = dictLanguage.Item(Session("language")&"_instPrd_C")
  case 59
    dstr=replace(scStoreMsg,"''","~~~")
    dstr=replace(dstr,"'","""")
    dstr=replace(dstr,"~~~","'")
    msgStr = dstr
  case 60
    msgStr = dictLanguage.Item(Session("language")&"_mainIndex_1")
  case 61
    msgStr = dictLanguage.Item(Session("language")&"_NewCust_1")&"<br><br><a href=default.asp>"&GetButtonLink("continueshop")&"</a>"    
  case 62
    msgStr = dictLanguage.Item(Session("language")&"_orderverify_1")
  case 63
    msgStr = dictLanguage.Item(Session("language")&"_orderverify_2")
  case 64
    msgStr = dictLanguage.Item(Session("language")&"_paymntb_c_1")
  case 65
    msgStr = dictLanguage.Item(Session("language")&"_paymntb_c_2")
  case 66
    msgStr = dictLanguage.Item(Session("language")&"_paymntb_o_1")
  case 67
    msgStr = dictLanguage.Item(Session("language")&"_paymntb_o_6") & "<br><br><a href=""javascript:history.go(-1)"">"&GetButtonLink("back")&"</a>" 
  case 68
    msgStr = dictLanguage.Item(Session("language")&"_paymntb_o_5") & "<br><br><a href=""javascript:history.go(-1)"">"&GetButtonLink("back")&"</a>" 
  case 69
    msgStr = dictLanguage.Item(Session("language")&"_paymntb_o_8") & "<br><br><a href=""javascript:history.go(-1)"">"&GetButtonLink("back")&"</a>" 
  case 70
    msgStr = dictLanguage.Item(Session("language")&"_paymntb_o_7") & "<br><br><a href=""javascript:history.go(-1)"">"&GetButtonLink("back")&"</a>"    
  case 71
    msgStr = dictLanguage.Item(Session("language")&"_paymntb_o_4")&"<br><br><a href=default.asp>"&GetButtonLink("continueshop")&"</a>"
  case 72
    msgStr = dictLanguage.Item(Session("language")&"_paymntb_o_4")
  case 73
    msgStr = dictLanguage.Item(Session("language")&"_msg_73")
  case 74
    msgStr = dictLanguage.Item(Session("language")&"_viewPrd_1")
  case 75
    msgStr = dictLanguage.Item(Session("language")&"_msg_75")&"<br><br><a href=""javascript:history.go(-1)"">"&GetButtonLink("back")&"</a>"
  case 76
    msgStr = dictLanguage.Item(Session("language")&"_msg_76")
  case 77
    msgStr = dictLanguage.Item(Session("language")&"_orderverify_5")&"<br><br><a href=""javascript:history.go(-1)"">"&GetButtonLink("back")&"</a>"
  case 78
    msgStr = dictLanguage.Item(Session("language")&"_advSrcb_1") &"<br><br><a href=search.asp>"&GetButtonLink("back")&"</a>"
  case 79
    msgStr = dictLanguage.Item(Session("language")&"_advSrcb_2") &"<br><br><a href=search.asp>"&GetButtonLink("back")&"</a>"    
  case 80
    msgStr = ""
  case 81
    msgStr = ""
  case 82
    msgStr = dictLanguage.Item(Session("language")&"_updOrdStats_2")&"<br><br><a href=default.asp>"&GetButtonLink("continueshop")&"</a>"    
  case 83
    dstr=replace(scStoreMsg,"''","~~~")
    'dstr=replace(dstr,"'","""")
    dstr=replace(dstr,"~~~","'")
    dstr=replace(dstr, "&lt;BR&gt;", "<br>")
    msgStr = dstr
  case 84
    dstr=replace(scStoreMsg,"''","~~~")
    'dstr=replace(dstr,"'","""")
    dstr=replace(dstr,"~~~","'")
    dstr=replace(dstr, "&lt;BR&gt;", "<br>")
    msgStr = dstr
  case 85
    msgStr = dictLanguage.Item(Session("language")&"_viewCat_P_1")
  case 86
    msgStr = dictLanguage.Item(Session("language")&"_viewCat_P_6")
  case 87
    msgStr = dictLanguage.Item(Session("language")&"_viewCat_P_1")
  case 88
    msgStr = dictLanguage.Item(Session("language")&"_viewPrd_2")
  case 89
    msgStr = dictLanguage.Item(Session("language")&"_viewSpc_1")
  case 90
    msgStr = dictLanguage.Item(Session("language")&"_advSrcb_2") &"<br><br><a href=""viewBrands.asp"">"&GetButtonLink("back")&"</a>"
  case 91
    msgStr = dictLanguage.Item(Session("language")&"_AffLogin_10") &"<br><br><a href=AffiliateLogin.asp?redirectUrl=" & Server.URLEncode(session("redirectUrlLI")) & ">"&GetButtonLink("back")&"</a>"    
  case 92
    msgStr = dictLanguage.Item(Session("language")&"_AffLogin_11") &"<br><br><a href=AffiliateLogin.asp?redirectUrl=" & Server.URLEncode(session("redirectUrlLI")) & ">"&GetButtonLink("back")&"</a>" 
  case 93
    msgStr = dictLanguage.Item(Session("language")&"_viewNewArrivals_1")
  case 94
    msgStr = dictLanguage.Item(Session("language")&"_viewBestSellers_1")
  case 95
    msgStr = dictLanguage.Item(Session("language")&"_viewPrd_62")
        
  'GGG Add-on start
  case 96
    msgStr = dictLanguage.Item(Session("language")&"_msg_4") &"<br><br><a href=""javascript:history.go(-1)"">"&GetButtonLink("back")&"</a>"
  case 97
    msgStr = dictLanguage.Item(Session("language")&"_msg_5") &"<br><br><a href=""javascript:history.go(-1)"">"&GetButtonLink("back")&"</a>"
  case 98
    msgStr = dictLanguage.Item(Session("language")&"_msg_6") &"<br><br><a href=""javascript:history.go(-1)"">"&GetButtonLink("back")&"</a>"
  case 99
    msgStr = dictLanguage.Item(Session("language")&"_msg_7") &"<br><br><a href=""javascript:history.go(-1)"">"&GetButtonLink("back")&"</a>"
  case 100
    msgStr = dictLanguage.Item(Session("language")&"_msg_8") &"<br><br><a href=""javascript:history.go(-1)"">"&GetButtonLink("back")&"</a>"
  case 101
    msgStr = dictLanguage.Item(Session("language")&"_msg_9") &"<br><br><a href=""javascript:history.go(-1)"">"&GetButtonLink("back")&"</a>"
  case 102
    msgStr = dictLanguage.Item(Session("language")&"_msg_10") &"<br><br><a href=""javascript:history.go(-1)"">"&GetButtonLink("back")&"</a>"
  'GGG Add-on end
      
  case 130
    msgStr = bto_dictLanguage.Item(Session("language")&"_configurePrd_19") &"<br><br><a href=""javascript:history.go(-1)"">"&GetButtonLink("back")&"</a>"
  case 131
    msgStr = dictLanguage.Item(Session("language")&"_checkout_13") &"<br><br><a href=viewCart.asp>"&GetButtonLink("back")&"</a>"
  case 132
    msgStr = dictLanguage.Item(Session("language")&"_alert_12") &"<br><br><a href=""javascript:history.go(-1)"">"&GetButtonLink("back")&"</a>"
  case 133
    msgStr = dictLanguage.Item(Session("language")&"_alert_13") &"<br><br><a href=""repeatorder.asp?idOrder=" & request("idorder") & "&OrderRepeat=haveto"" class=""pcButton pcButtonContinue""><img src="""& rslayout("submit") & """ alt="""&dictLanguage.Item(Session("language")&"_css_submit")&"""><span class=""pcButtonText"">"&dictLanguage.Item(Session("language")&"_css_submit")&"</span></a>&nbsp;<a href=""javascript:history.go(-1)"">"&GetButtonLink("back")&"</a>"
  case 134
    msgStr = dictLanguage.Item(Session("language")&"_alert_14") &"<br><br><a href=""javascript:history.go(-1)"">"&GetButtonLink("back")&"</a>"
  case 135
    msgStr = dictLanguage.Item(Session("language")&"_alert_15") &"<br><br><a href=""addsavedprdstocart.asp?OrderRepeat=haveto"" class=""pcButton pcButtonContinue""><img src="""& rslayout("submit") & """ alt="""&dictLanguage.Item(Session("language")&"_css_submit")&"""><span class=""pcButtonText"">"&dictLanguage.Item(Session("language")&"_css_submit")&"</span></a>&nbsp;<a href=""javascript:history.go(-1)"">"&GetButtonLink("back")&"</a>"
  case 200
    msgStr = dictLanguage.Item(Session("language")&"_techErr_2")
  case 201
    msgStr = dictLanguage.Item(Session("language")&"_msg_201")
  case 202
    msgStr = dictLanguage.Item(Session("language")&"_sdsLogin_8") &"<br><br><a href=sds_Login.asp?redirectUrl=" & Server.URLEncode(session("redirectUrlLI")) & ">"&GetButtonLink("back")&"</a>"
  case 203  
    msgStr = dictLanguage.Item(Session("language")&"_sds_forgotpassworderror") &"<br><br><a href=""javascript:history.go(-1)"">"&GetButtonLink("back")&"</a>"
        
  case 204
    ' Error Description: Quantity being ordered is greater than quantity in stock
    ' Set local variables and clear session variables
    pDescription = session("pcErrStrPrdDesc")
    session("pcErrStrPrdDesc") = ""
    pStock = session("pcErrIntStock")
    session("pcErrIntStock") = Cint(0)
    msgStr = dictLanguage.Item(Session("language")&"_instPrd_2")&pDescription&dictLanguage.Item(Session("language")&"_instPrd_3")&pStock&dictLanguage.Item(Session("language")&"_instPrd_4")&"<br><br><a href=""javascript:history.go(-1)"">"&GetButtonLink("back")&"</a>"
        
  case 205
    ' Error Description: Wholesale minimum not met, so customer cannot checkout
    msgStr = dictLanguage.Item(Session("language")&"_checkout_2")& scCurSign &money(scWholesaleMinPurchase)&dictLanguage.Item(Session("language")&"_techErr_3") & "<BR><BR><a href='viewCart.asp'>"&dictLanguage.Item(Session("language")&"_mainIndex_5")&"</a>"
        
  case 206
    ' Error Description: Retail minimum not met, so customer cannot checkout
    msgStr = dictLanguage.Item(Session("language")&"_checkout_2")& scCurSign &money(scMinPurchase) & "<BR><BR><a href='viewCart.asp'>"&dictLanguage.Item(Session("language")&"_mainIndex_5")&"</a>"
        
  case 207
    ' Error Description: The product ID could not be retrieved
    msgStr = dictLanguage.Item(Session("language")&"_PrdError_1")&"<br /><br/><a href=default.asp>"&GetButtonLink("continueshop")&"</a>"
  case 208  
    msgStr = dictLanguage.Item(Session("language")&"_PayPal_2") &"<br><br><a href=""javascript:history.go(-1)"">"&GetButtonLink("back")&"</a>"
  case 209  
    msgStr = dictLanguage.Item(Session("language")&"_PayPal_3") &"<br><br><a href=""javascript:history.go(-1)"">"&GetButtonLink("back")&"</a>"
        
  case 210
    ' Error Description: generic error due to invalid product or other ID
    msgStr = dictLanguage.Item(Session("language")&"_msg_210")
        
  case 211
    ' Error Description: Your session is invalid
    msgStr = dictLanguage.Item(Session("language")&"_validateform_9")&"<br /><br/><a href=default.asp>"&GetButtonLink("continueshop")&"</a>"
  case 212
    ' Error Description: This browser does not accept cookies.
    msgStr = dictLanguage.Item(Session("language")&"_PrdError_2")&"<br /><br/><a href=default.asp>"&GetButtonLink("continueshop")&"</a>"
        
  '// ProductCart v4
  case 300
    ' Content page cannot be accessed
    msgStr = dictLanguage.Item(Session("language")&"_viewPages_1") &"<br><br><a href=""javascript:history.go(-1)"">"&GetButtonLink("back")&"</a>"
  case 301
    ' Content page cannot be accessed
    msgStr = dictLanguage.Item(Session("language")&"_ShowRecentRev_2")
  case 302
    ' No brands available
    msgStr = dictLanguage.Item(Session("language")&"_msg_302")
  case 303
    ' No subbrands available
    msgStr = dictLanguage.Item(Session("language")&"_msg_303")
  case 304
    ' Did not setup header and footer properly
    dstr=replace(scStoreMsg,"''","~~~")
    dstr=replace(dstr,"~~~","'")
    dstr=replace(dstr, "&lt;BR&gt;", "<br>")
    msgStr = dstr
  case 305
    ' Subscription product in cart
    msgStr = scSBLang5
  case 306
    ' Subscription product not allowed
    msgStr = scSBLang6
  case 307
    ' No payment methods available
    msgStr = dictLanguage.Item(Session("language")&"_EIG_17")
  case 308
    ' Duplicate Order Detected
    msgStr = dictLanguage.Item(Session("language")&"_OPC_Alert_01")
  case 309
    ' BTO items are low of stock
    msgStr = dictLanguage.Item(Session("language")&"_instConfQty_1")
  case 310
    'Suspended Account Checkout
    msgStr = dictLanguage.Item(Session("language")&"_opc_checkorv_3")
  case 311
    'Amazon Payment Errors
    msgStr = dictLanguage.Item(Session("language")&"_AmazonPay_7")
  case 312
    'Amazon Payment Errors
    msgStr = dictLanguage.Item(Session("language")&"_AmazonPay_8") & session("amzError")
	
  case 313
    msgStr = dictLanguage.Item(Session("language")&"_newpass_11") & session("pcSFLockMinutes") & dictLanguage.Item(Session("language")&"_newpass_11a") &"<br><br><a href=checkout.asp>"&GetButtonLink("back")&"</a>"
	session("pcSFLockMinutes")=""
	
  case 314
    msgStr = dictLanguage.Item(Session("language")&"_newpass_12") & session("pcSFLockMinutes") & dictLanguage.Item(Session("language")&"_newpass_12a") &"<br><br><a href=checkout.asp>"&GetButtonLink("back")&"</a>"
	session("pcSFLockMinutes")=""
	
  case 315
    msgStr = dictLanguage.Item(Session("language")&"_newpass_13") & "<a href='checkout.asp?cmode=2&fmode=1'>" & dictLanguage.Item(Session("language")&"_newpass_14a") & "</a>" & dictLanguage.Item(Session("language")&"_newpass_14b")
	session("pcSFLockMinutes")=""
	
  case 316
    msgStr = dictLanguage.Item(Session("language")&"_alert_20")
  case 317
    msgStr = dictLanguage.Item(Session("language")&"_alert_21")
	
  case 318
    msgStr = dictLanguage.Item(Session("language")&"_viewPages_3")
	
  case 319
    ' Error Description: Quantity being ordered is greater than desired Gift Registry quantity
    ' Set local variables and clear session variables
    pDescription = session("pcErrStrPrdDesc")
    session("pcErrStrPrdDesc") = ""
    pRemain = session("pcErrIntRemain")
    session("pcErrIntRemain") = Cint(0)
    msgStr = dictLanguage.Item(Session("language")&"_instPrd_2")&pDescription&dictLanguage.Item(Session("language")&"_instPrd_5")&pRemain&dictLanguage.Item(Session("language")&"_instPrd_6")&"<br><br><a href=""javascript:history.go(-1)"">"&GetButtonLink("back")&"</a>"
	
  end select 

  pcf_getStoreMsg = msgStr
End Function



Function pcf_GetCurrentPage()

    if scATCEnabled="1" then
    
        Dim originalurl, rooturl, homepageurl
        
        pcv_URLPrefix = scStoreURL & "/" & scPcFolder
        pcv_URLPrefix = replace(pcv_URLPrefix,"//","/")
        pcv_URLPrefix = replace(pcv_URLPrefix,"http:/","http://")
        pcv_URLPrefix = replace(pcv_URLPrefix,"https:/","https://")
        
        rooturl = pcv_URLPrefix&"/pc/"
        if scURLredirect = "" then
            homepageurl = pcv_URLPrefix&"/pc/home.asp"
        else
            homepageurl = scURLredirect
        end if
        
        atc_Debug = 0	' 1 = ON
        
        ' --------------------------------------------------------------------
         
        originalurl 	= lcase(Request.ServerVariables("HTTP_REFERER"))
        If len(originalurl)=0 Then
            if scSeoURLs = 0 then
                originalurl	= rooturl & "viewPrd.asp"
            else
                originalurl	= rooturl & "viewcart.asp"
            end if
        End If
        atc_idProduct 	= getuserinput(request("idproduct"),0)
        
        ' --------------------------------------------------------------------
        ' // Debugging
        ' response.write "homepageurl=" & homepageurl & "<br>"
        ' response.write "rooturl=" & rooturl & "<br>"
        ' response.write "originalurl=" & originalurl & "<br>"
        ' Exit Function()
        ' --------------------------------------------------------------------
        
        if originalurl = rooturl then
            originalurl = homepageurl & "?idproduct=" & atc_idproduct 
            if InStr(originalurl,"atc=")= 0 then
                if InStr(originalurl,"?") then
                    originalurl = originalurl & "&atc=1"
                else
                    originalurl = originalurl & "?atc=1"			
                end if
                if atc_debug = 1 then originalurl = originalurl & "&home=1"			
            end if
            pcf_GetCurrentPage = originalurl
            Exit Function
        
        elseif originalurl = homepageurl then
            originalurl = homepageurl & "?idproduct=" & atc_idproduct 
            if InStr(originalurl,"atc=")= 0 then
                if InStr(originalurl,"?") then
                    originalurl = originalurl & "&atc=1"
                else
                    originalurl = originalurl & "?atc=1"			
                end if
                if atc_debug = 1 then originalurl = originalurl & "&home=2"			
            end if
            pcf_GetCurrentPage = originalurl
            Exit Function
        
        elseif InStr(originalurl,"showbestsellers.asp") <> 0  then
            originalurl = rooturl2 & "showbestsellers.asp?idproduct=" & atc_idproduct 
            if InStr(originalurl,"atc=")= 0 then
                if InStr(originalurl,"?") then
                    originalurl = originalurl & "&atc=1"
                else
                    originalurl = originalurl & "?atc=1"			
                end if
                if atc_debug = 1 then originalurl = originalurl & "&bestsellers"			
            end if
            pcf_GetCurrentPage = originalurl
            Exit Function
        
        elseif InStr(originalurl,"showfeatured.asp") <> 0  then
            originalurl = rooturl2 & "showfeatured.asp?idproduct=" & atc_idproduct 
            if InStr(originalurl,"atc=")= 0 then
                if InStr(originalurl,"?") then
                    originalurl = originalurl & "&atc=1"
                else
                    originalurl = originalurl & "?atc=1"			
                end if
                if atc_debug = 1 then originalurl = originalurl & "&featured"			
            end if
            pcf_GetCurrentPage = originalurl
            Exit Function
        
        elseif InStr(originalurl,"shownewarrivals.asp") <> 0  then
            originalurl = rooturl2 & "shownewarrivals.asp?idproduct=" & atc_idproduct 
            if InStr(originalurl,"atc=")= 0 then
                if InStr(originalurl,"?") then
                    originalurl = originalurl & "&atc=1"
                else
                    originalurl = originalurl & "?atc=1"			
                end if
                if atc_debug = 1 then originalurl = originalurl & "&new"			
            end if
            pcf_GetCurrentPage = originalurl
            Exit Function
        
        elseif InStr(originalurl,"showspecials.asp") <> 0  then
            originalurl = rooturl2 & "showspecials.asp?idproduct=" & atc_idproduct 
            if InStr(originalurl,"atc=")= 0 then
                if InStr(originalurl,"?") then
                    originalurl = originalurl & "&atc=1"
                else
                    originalurl = originalurl & "?atc=1"			
                end if
                if atc_debug = 1 then originalurl = originalurl & "&specials"			
            end if
            pcf_GetCurrentPage = originalurl
            Exit Function
        
        elseif InStr(originalurl,"showsearchresults.asp") <> 0  then
        
            if InStr(originalurl,"atc=")= 0 then
                ' on the first pass, all we have to do is add the flag and the product id
                if InStr(originalurl,"?") then
                    originalurl = originalurl & "&atc=1&idproduct=" & atc_idproduct
                else
                    originalurl = originalurl & "?atc=1&idproduct=" & atc_idproduct
                end if
                if atc_debug = 1 then originalurl = originalurl & "&searchresults=1"			
                pcf_GetCurrentPage = originalurl
                Exit Function
            else
                ' on subsequent passes, first we have to strip the flag and product id, then add them back
            
                startpos 	  		= InStr(originalurl,"atc")			' where atc begins
                searchurl			= left(originalurl,startpos-1)
                originalurl = searchurl 
                originalurl = originalurl & "&atc=1"
                originalurl = originalurl & "&idproduct=" & atc_idproduct
                if atc_debug = 1 then originalurl = originalurl & "&searchresults=2"			 
                pcf_GetCurrentPage = originalurl
                Exit Function
            end if
        
        elseif InStr(originalurl,"viewcategories.asp") <> 0  then
        
            if InStr(originalurl,"idcategory")=0 then
                originalurl=originalurl & "?idcategory=1"
            end if
            
            lenourl	  			= len(originalurl)							' length of referrer URL string
            startpos 	  		= InStr(originalurl,"idcategory")			' where idcategory begins
            midstring 			= mid(originalurl, startpos, lenourl-1)
            eqpos 				= instr(midstring, "=")						' where the "=" following idcategory is located
            lenvalue 			= lenourl - eqpos
            if lenvalue > 0 then 
                beyondeq		= mid(midstring, eqpos+1, lenvalue)
            end if
            ampersandpos	  		= instr(1,beyondeq,"&") ' is there another variable beyond the category?
            if ampersandpos > 0 then
                category		  	= left(beyondeq,ampersandpos-1) 
            else
                category		  	= beyondeq
            end if 
      
            if lcase(InStr(originalurl,"sfid")) = 0  then ' normal category page call        
                originalurl = rooturl 
                if category>1 then
                originalurl = originalurl & "viewCategories.asp?idcategory=" & category 
                originalurl = originalurl & "&idproduct=" & atc_idproduct 
                originalurl = originalurl & "&atc=1"
                else
                    originalurl = originalurl & "viewCategories.asp?atc=1"
                end if
            else ' custom search fields
                if InStr(originalurl,"atc=1") = 0  then
                    originalurl = originalurl & "&atc=1"
                end if 
            end if
        
        ElseIf InStr(originalurl,"viewprd.asp") <> 0  Then
        
            If InStr(originalurl,"atc=")= 0 Then
                if InStr(originalurl,"?") then
                    originalurl = originalurl & "&atc=1"
                else
                    originalurl = originalurl & "?atc=1"			
                end if
                if atc_debug = 1 then originalurl = originalurl & "&viewprd=1"			 
            End If
            pcf_GetCurrentPage = originalurl
            Exit Function
        
        elseif request("pCnt") > 0  then        
            if InStr(originalurl,"atc=")= 0 then
                if InStr(originalurl,"?") then
                    originalurl = originalurl & "&atc=1"
                else
                    originalurl = originalurl & "?atc=1"			
                end if
            end if
            pcf_GetCurrentPage = originalurl
            Exit Function
        end if
        
        
        
        if InStr(originalurl,".htm")<> 0 then
        
            If InStr(originalurl,"atc=")= 0 Then
                ' on the first pass, all we have to do is add the flag and the product id
                if InStr(originalurl,"?") then
                    originalurl = originalurl & "&atc=1&idproduct=" & atc_idproduct
                else
                    originalurl = originalurl & "?atc=1&idproduct=" & atc_idproduct
                end if
                if atc_debug = 1 then originalurl = originalurl & "&alternate=1"			 
                pcf_GetCurrentPage = originalurl
                Exit Function
        
            Else
                ' on subsequent passes, first we have to strip the flag and product id, then add them back
                startpos    = InStr(originalurl,"?atc")			' where atc begins
                searchurl   = Left(originalurl,startpos - 1)
                originalurl = searchurl
                originalurl = originalurl & "?atc=1"
                originalurl = originalurl & "&idproduct=" & atc_idproduct
                if atc_debug = 1 then originalurl = originalurl & "&alternate=2"			 
        
                If InStr(originalurl,"atc=")= 0 Then
                    ' on the first pass, all we have to do is add the flag and the product id
                    if InStr(originalurl,"?") then
                        originalurl = originalurl & "&atc=1&idproduct=" & atc_idproduct
                    else
                        originalurl = originalurl & "?atc=1&idproduct=" & atc_idproduct
                    end if
                    pcf_GetCurrentPage = originalurl
                    Exit Function
                Else
                    ' on subsequent passes, first we have to strip the flag and product id, then add them back
                    startpos    = InStr(originalurl,"?atc")			' where atc begins
                    searchurl   = Left(originalurl,startpos - 1)
                    originalurl = searchurl
                    originalurl = originalurl & "?atc=1"
                    originalurl = originalurl & "&idproduct=" & atc_idproduct
                    if atc_debug = 1 then originalurl = originalurl & "&alternate=3"		
                    pcf_GetCurrentPage = originalurl
                    Exit Function
                End If
        
            End If
        
            If InStr(originalurl,"idproduct") = 0 Then
                originalurl = originalurl & "&idproduct=" & atc_idproduct
            End If
        
            pcf_GetCurrentPage = originalurl
            Exit Function
        End If
    
        If InStr(originalurl,"idproduct") = 0 Then
            if InStr(originalurl,"atc=")= 0 then
                if InStr(originalurl,"?") then
                    originalurl = originalurl & "&atc=1"
                else
                    originalurl = originalurl & "?atc=1"			
                end if			
            end if
            originalurl = originalurl & "&idproduct=" & atc_idproduct
        End If
        
        pcf_GetCurrentPage = originalurl
    End if 
    
        
End Function

'// Test if the provided string has visible content
Function pcf_HasHTMLContent(Content)
	pcf_HasHTMLContent = Len(Trim(Replace(Replace(Content&"", "<br />", ""), vbCrLf, ""))) > 0
End Function

'// Use if loading from a page inside pc/ directory
Function pcf_FixHTMLContentPaths(Content)
	newContent = Content

	'// Fix paths
	newContent = Replace(newContent, "../pc/", "")

	'// Fix YouTube and Vimeo URLs if we're using HTTPS
	If Request.ServerVariables("HTTPS") = "on" Then
		newContent = Replace(newContent, "http://www.youtube.com/", "https://www.youtube.com/")
		newContent = Replace(newContent, "http://player.vimeo.com/", "https://player.vimeo.com/")
	End If

	pcf_FixHTMLContentPaths = newContent
End Function

%>