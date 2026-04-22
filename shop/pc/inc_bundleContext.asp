<%
' ==============================================================
' inc_bundleContext.asp
' 2026 redesign - bundle end-page shared context.
'
' Included by viewPrd-<machine>-bundle-v2.asp after common.asp.
' Validates sid/mid/cid querystrings, queries products for live
' stand + monitor rows, derives monitor count, maps to bundle
' discount, and exposes VBScript variables used throughout the
' bundle page. Also provides mmEmitBundleHiddenInputs() which
' writes the hidden <input> fields instPrd.asp consumes to
' process the add-to-cart as a bundle (pCnt=3 + idproduct2/3).
'
' Prerequisites - the including page must define BEFORE the
' include directive:
'   Const MM_PRODUCT_ID = <pc idProduct>
'   Const MM_VAT_RATE   = 1.2
' ==============================================================

' ------------------------------------------------------------
' 1. Parse + validate querystrings
' ------------------------------------------------------------
Dim mmBunSid, mmBunMid, mmBunCid

If Request.QueryString("sid") = "" Then
  Response.Redirect "/bundles/"
End If
If Not IsNumeric(Request.QueryString("sid")) Then
  Response.Redirect "/404.html"
End If
mmBunSid = CLng(Request.QueryString("sid"))

If Request.QueryString("mid") = "" Then
  Response.Redirect "/bundles/?sid=" & mmBunSid
End If
If Not IsNumeric(Request.QueryString("mid")) Then
  Response.Redirect "/404.html"
End If
mmBunMid = CLng(Request.QueryString("mid"))

If Request.QueryString("cid") = "" Then
  Response.Redirect "/bundles/?sid=" & mmBunSid & "&mid=" & mmBunMid
End If
If Not IsNumeric(Request.QueryString("cid")) Then
  Response.Redirect "/404.html"
End If
mmBunCid = CLng(Request.QueryString("cid"))

' The including page sets MM_PRODUCT_ID to its PC. If cid does
' not match, the URL is for the wrong bundle page - send the
' customer back to the builder with the stand+monitor preserved.
If mmBunCid <> MM_PRODUCT_ID Then
  Response.Redirect "/bundles/?sid=" & mmBunSid & "&mid=" & mmBunMid
End If

' ------------------------------------------------------------
' 2. Stand + monitor DB lookup (live prices + images)
' ------------------------------------------------------------
Dim mmBunStandName, mmBunStandPriceInc, mmBunStandImg, mmBunStandSku
Dim mmBunMonName,   mmBunMonPriceInc,   mmBunMonImg,   mmBunMonSku
mmBunStandName = "" : mmBunStandPriceInc = 0 : mmBunStandImg = "" : mmBunStandSku = ""
mmBunMonName   = "" : mmBunMonPriceInc   = 0 : mmBunMonImg   = "" : mmBunMonSku   = ""

Dim mmBunSql, mmBunRs

mmBunSql = "SELECT description, sku, price, smallImageUrl " & _
           "FROM products " & _
           "WHERE idProduct = " & mmBunSid & _
           "  AND active = -1 AND removed = 0"
Set mmBunRs = connTemp.Execute(mmBunSql)
If Not mmBunRs.EOF Then
  mmBunStandName     = mmBunRs("description") & ""
  mmBunStandSku      = mmBunRs("sku") & ""
  mmBunStandPriceInc = CDbl(mmBunRs("price"))
  mmBunStandImg      = mmBunRs("smallImageUrl") & ""
End If
mmBunRs.Close : Set mmBunRs = Nothing

If mmBunStandName = "" Then
  Response.Redirect "/bundles/"
End If

mmBunSql = "SELECT description, sku, price, smallImageUrl " & _
           "FROM products " & _
           "WHERE idProduct = " & mmBunMid & _
           "  AND active = -1 AND removed = 0"
Set mmBunRs = connTemp.Execute(mmBunSql)
If Not mmBunRs.EOF Then
  mmBunMonName     = mmBunRs("description") & ""
  mmBunMonSku      = mmBunRs("sku") & ""
  mmBunMonPriceInc = CDbl(mmBunRs("price"))
  mmBunMonImg      = mmBunRs("smallImageUrl") & ""
End If
mmBunRs.Close : Set mmBunRs = Nothing

If mmBunMonName = "" Then
  Response.Redirect "/bundles/?sid=" & mmBunSid
End If

' ------------------------------------------------------------
' 3. Derive monitor count from the stand name
'    (matches funBundlesCalcs in viewPrdCode.asp)
' ------------------------------------------------------------
Dim mmBunMonCount
mmBunMonCount = 1
If InStr(mmBunStandName, "Dual")   > 0 Then mmBunMonCount = 2
If InStr(mmBunStandName, "Triple") > 0 Then mmBunMonCount = 3
If InStr(mmBunStandName, "Quad")   > 0 Then mmBunMonCount = 4
If InStr(mmBunStandName, "Five")   > 0 Then mmBunMonCount = 5
If InStr(mmBunStandName, "Six")    > 0 Then mmBunMonCount = 6
If InStr(mmBunStandName, "Eight")  > 0 Then mmBunMonCount = 8

' ------------------------------------------------------------
' 4. Bundle discount by monitor count
'    (matches inc_headerDAJS.asp numBunDiscount)
' ------------------------------------------------------------
Dim mmBunDiscount
Select Case mmBunMonCount
  Case 2, 3 : mmBunDiscount = 25
  Case 4, 5 : mmBunDiscount = 50
  Case 6, 8 : mmBunDiscount = 100
  Case Else : mmBunDiscount = 0
End Select

' ------------------------------------------------------------
' 5. Derived totals - ex-VAT values used throughout the page
' ------------------------------------------------------------
Dim mmBunStandPriceEx, mmBunMonPriceEx
Dim mmBunMonSubtotalEx, mmBunMonSubtotalInc
mmBunStandPriceEx   = mmBunStandPriceInc / MM_VAT_RATE
mmBunMonPriceEx     = mmBunMonPriceInc   / MM_VAT_RATE
mmBunMonSubtotalEx  = mmBunMonPriceEx    * mmBunMonCount
mmBunMonSubtotalInc = mmBunMonPriceInc   * mmBunMonCount

' ------------------------------------------------------------
' 6. Cleaned-up display names (strip ProductCart admin prefixes)
' ------------------------------------------------------------
Dim mmBunStandDispName, mmBunMonDispName
mmBunStandDispName = Replace(mmBunStandName, "Synergy ", "")
mmBunStandDispName = Replace(mmBunStandDispName, "Monitor ", "")
mmBunMonDispName   = mmBunMonName

' ------------------------------------------------------------
' 7. Image URLs with fallback
' ------------------------------------------------------------
Dim mmBunStandImgSrc, mmBunMonImgSrc
If mmBunStandImg <> "" Then
  mmBunStandImgSrc = "/shop/pc/catalog/" & mmBunStandImg
Else
  mmBunStandImgSrc = "/shop/pc/catalog/no_image.gif"
End If
If mmBunMonImg <> "" Then
  mmBunMonImgSrc = "/shop/pc/catalog/" & mmBunMonImg
Else
  mmBunMonImgSrc = "/shop/pc/catalog/no_image.gif"
End If

' ------------------------------------------------------------
' 8. Hidden form inputs - emit inside the cart <form>.
'    instPrd.asp reads these and processes as a 3-item bundle.
'    (Shape matches the legacy inc_headerDAJS.asp formBundleOptions.)
' ------------------------------------------------------------
Sub mmEmitBundleHiddenInputs()
%>
  <input type="hidden" name="idproduct2" value="<%= mmBunSid %>">
  <input type="hidden" name="QtyM<%= mmBunSid %>" value="1">
  <input type="hidden" name="idproduct3" value="<%= mmBunMid %>">
  <input type="hidden" name="QtyM<%= mmBunMid %>" value="<%= mmBunMonCount %>">
  <input type="hidden" name="pCnt" value="3">
<%
End Sub
%>
