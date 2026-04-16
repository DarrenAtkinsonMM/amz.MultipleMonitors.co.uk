<%
Public Sub pcs_SaveThemeToSettings(pcStrThemeFolder)

    Dim strtext1
    
	pcStrFolder = "../includes"

    strtext1 = ""
    
    strtext1 = strtext1 & "<" & Chr(37) & vbNewLine 
    
    strtext1 = strtext1 & "private const scThemePath = """ & pcStrThemeFolder & """" & vbCrLf

    strtext1 = strtext1 & Chr(37) & ">"
    
    'response.Write(strtext1)
    'response.End()

    call pcs_SaveUTF8(pcStrFolder & "\themesettings.asp", pcStrFolder & "\themesettings.asp", strtext1)

End Sub


Public Function pcf_displayThemeName(handle)

    handle = replace(handle, "_", " ")
    handle = replace(handle, "-", " ")
    handle = replace(handle, ".", " ")
    pcf_displayThemeName = handle

End Function


Public Sub pcs_IndexThemeFolder()
    Dim ThemePath

    If PPD="1" Then
        ThemePath = Server.MapPath("/" & scPcFolder & "/pc/theme/") & "/"
    Else
        ThemePath = Server.MapPath("../pc/theme/") & "/"
    End If
    
    Dim ThemeFS, ThemeDir, ThemeId, ThemeName, ThemeStatus
    Set ThemeFS = Server.CreateObject("Scripting.FileSystemObject")
    Set ThemeDir = ThemeFS.GetFolder(ThemePath)
    

    If ThemeDir.SubFolders.Count > 0 Then

        For Each Folder in ThemeDir.SubFolders
        
            If Folder.name<>"v47_upgrade" And Folder.name<>"_common" Then
            
                pcs_SaveTheme(Folder.name)
                
                pcv_strThemeList = pcv_strThemeList & "'" & Folder.name & "'" & ","

            End If
            
        Next
        
        pcv_strThemeList = left(pcv_strThemeList,len(pcv_strThemeList)-1)
        
    End If
    
    query = "DELETE FROM pcThemes WHERE pcThemes_Name Not IN (" & pcv_strThemeList & ")"
    Set rs = server.CreateObject("ADODB.RecordSet")
    Set rs = conntemp.execute(query)
    Set rs = Nothing
    
    
    Set ThemeFS = Nothing
    Set ThemeDir = Nothing
    
End Sub


Public Sub pcs_SaveTheme(handle)

    query="SELECT [pcThemes_Name] FROM pcThemes WHERE [pcThemes_Name]='" & handle & "'"
    Set rs2 = server.CreateObject("ADODB.RecordSet")
    Set rs2 = connTemp.execute(query)
    If Not rs2.Eof Then    
        call pcs_UpdateTheme(handle)    
    Else    
        call pcs_AddTheme(handle)    
    End If
    Set rs2 = Nothing 

End Sub


Public Sub pcs_AddTheme(handle)

    query="INSERT INTO pcThemes ([pcThemes_Name], [pcThemes_Active]) VALUES ('" & handle & "', 0);"
    Set rs2 = server.CreateObject("ADODB.RecordSet")
    Set rs2 = connTemp.execute(query)
    Set rs2 = Nothing 

End Sub


Public Sub pcs_UpdateTheme(handle)

    query="UPDATE pcThemes SET "
    query = query & "pcThemes_Name='" & handle & "' "
    query = query & "WHERE [pcThemes_Name]='" & handle & "'"
    Set rs2 = server.CreateObject("ADODB.RecordSet")
    Set rs2 = connTemp.execute(query)
    Set rs2 = Nothing  

End Sub

Function DisplayZone(tmpZone)
    Select Case trim(tmpZone)
        Case "CatTree":	pcs_CategoryTree
        Case "PrdName": pcs_ProductName
        Case "PrdSKU": pcs_ShowSKU
        Case "PrdRate": pcs_ShowRating
        Case "PrdW": pcs_DisplayWeight
        Case "PrdBrand": pcs_ShowBrand
        Case "PrdStock": pcs_UnitsStock
        Case "PrdDesc": pcs_ProductDescription
        Case "PrdLDesc": pcs_LongProductDescription
        Case "PrdConfig": pcs_BTOConfiguration
        Case "PrdSearch": pcs_CustomSearchFields
        Case "PrdRP": pcs_RewardPoints
        Case "PrdPrice": pcs_ProductPrices
        Case "PrdImg": pcs_ProductImage
                    pcs_AdditionalImages
        Case "PrdOpt": 
            if pcf_VerifyShowOptions then
                pcs_OptionsN
            end if
        Case "PrdInput":
            if pcf_VerifyShowOptions then
                pcs_OptionsX							
            end if
        Case "PrdQDisc": pcs_QtyDiscounts
        Case "PrdCS": pcs_RequiredCrossSelling
                    pcs_CrossSellingDiscounts
                    pcs_CrossSellingAccessories
        Case "PrdRev":%>
            <!--#include file="../../pc/prv_increviews.asp"-->
        <%Case "PrdSB": pcs_SubscriptionProduct
        Case "PrdOSM": pcs_OutStockMessage
        Case "PrdPromo": pcs_ProductPromotionMsg
        Case "PrdNoShip", "PrdFreeShip": pcs_NoShippingText
        Case "PrdBOM": pcs_DisplayBOMsg
        Case "PrdWL": pcs_WishList
        Case "PrdAT": if scAddThisDisplay<>0 then pcs_AddThis
        Case "PrdBtns": pcs_NextButtons
        Case "PrdATC":
            if pFormQuantity="-1" and NotForSaleOverride(session("customerCategory"))=0 then
                if pEmailText<>"" then 
                    response.write "<div class=pcShowProductNFS>" 
                    response.write pEmailText '// reason why it's not for sale
                    response.write "</div>" 
                end if
                tmp_showQty="1"
                if pcv_lngMinimumQty>0 then
                    tmpMinQty=pcv_lngMinimumQty
                else
                    tmpMinQty=1
                end if%>
                <input type="hidden" name="quantity" value="<%=tmpMinQty%>">
            <%else 
                If scorderlevel = "0" OR pcf_WholesaleCustomerAllowed Then
                    if pcf_OutStockPurchaseAllow then
                        If ((pserviceSpec<>0) AND ((pnoprices>0) OR (pPrice=0) OR (scConfigPurchaseOnly=1))) or ((iBTOQuoteSubmitOnly=1) and (pserviceSpec<>0)) then
                            pcs_CustomizeButton
                        else
                            pcs_OptionsXTab
                            pcs_AddtoCart
                        end if 
                    end if
                end if
                If (not pcf_OutStockPurchaseAllow) OR (scorderlevel = "2") OR ((pcf_WholesaleCustomerAllowed or scorderlevel = "1") and session("customerType")<>"1") then%>
                    <input type="hidden" name="quantity" value="1">
                <input type="hidden" name="idproduct" value="<%=pidProduct%>">
                <%End if
            end if
        Case Else	
            Call pcs_addWidget(tmpZone)
    End Select
End Function
%>