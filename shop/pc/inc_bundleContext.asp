<%
' ==============================================================
' inc_bundleContext.asp
' 2026 redesign - bundle end-page shared context.
'
' Included by viewPrd-<machine>-bundle.asp after common.asp.
' Reads the stand + monitor slugs from the URL (rewritten by IIS
' from /products/<machine>/<stand-slug>/<monitor-slug>/ into the
' querystring pair stand=&monitor=), looks up the products by
' pcUrlBundle, derives monitor count, maps to bundle discount, and
' exposes VBScript variables used throughout the bundle page.
' Also provides mmEmitBundleHiddenInputs() which writes the
' hidden <input> fields instPrd.asp consumes to process the
' add-to-cart as a bundle (pCnt=3 + idproduct2/3).
'
' Prerequisites - the including page must define BEFORE the
' include directive:
'   Const MM_PRODUCT_ID = <pc idProduct>
'   Const MM_VAT_RATE   = 1.2
' The PC identity comes from the page (and the URL path segment
' that the IIS rewrite matched), so no querystring carries it.
' ==============================================================

' ------------------------------------------------------------
' 1. Parse + validate slug querystrings (set by IIS rewrite)
'    Slug format guard: lowercase alphanumeric + hyphen only.
'    Anything else -> /bundles/ (treat as garbage).
' ------------------------------------------------------------
Dim mmBunStandSlug, mmBunMonSlug
mmBunStandSlug = LCase(Trim(Request.QueryString("stand") & ""))
mmBunMonSlug   = LCase(Trim(Request.QueryString("monitor") & ""))

If mmBunStandSlug = "" Then
  Response.Redirect "/bundles/"
End If
If Not mmBunIsValidSlug(mmBunStandSlug) Then
  Response.Redirect "/bundles/"
End If
If mmBunMonSlug = "" Then
  Response.Redirect "/bundles/"
End If
If Not mmBunIsValidSlug(mmBunMonSlug) Then
  Response.Redirect "/bundles/"
End If

' ------------------------------------------------------------
' 2. Stand + monitor DB lookup by slug (live prices + images).
'    pcprod_DisplayLayout filter rejects cross-type matches -
'    e.g. /products/trader-pc/trader-pc/anything/ wouldn't return
'    the Trader PC row as a stand.
' ------------------------------------------------------------
Dim mmBunSid, mmBunMid
Dim mmBunStandName, mmBunStandPriceInc, mmBunStandImg, mmBunStandImgLg, mmBunStandSku
Dim mmBunMonName,   mmBunMonPriceInc,   mmBunMonImg,   mmBunMonSku
mmBunSid = 0 : mmBunMid = 0
mmBunStandName = "" : mmBunStandPriceInc = 0 : mmBunStandImg = "" : mmBunStandSku = ""
mmBunMonName   = "" : mmBunMonPriceInc   = 0 : mmBunMonImg   = "" : mmBunMonSku   = ""

Dim mmBunSql, mmBunRs

mmBunSql = "SELECT idProduct, description, sku, price, smallImageUrl, imageUrl " & _
           "FROM products " & _
           "WHERE pcUrlBundle = '" & Replace(mmBunStandSlug, "'", "''") & "'" & _
           "  AND pcprod_DisplayLayout = 'stand'" & _
           "  AND active = -1 AND removed = 0"
Set mmBunRs = connTemp.Execute(mmBunSql)
If Not mmBunRs.EOF Then
  mmBunSid           = CLng(mmBunRs("idProduct"))
  mmBunStandName     = mmBunRs("description") & ""
  mmBunStandSku      = mmBunRs("sku") & ""
  mmBunStandPriceInc = CDbl(mmBunRs("price"))
  mmBunStandImg      = mmBunRs("smallImageUrl") & ""
  mmBunStandImgLg    = mmBunRs("imageUrl") & ""
End If
mmBunRs.Close : Set mmBunRs = Nothing

If mmBunStandName = "" Then
  Response.Redirect "/bundles/"
End If

mmBunSql = "SELECT idProduct, description, sku, price, smallImageUrl " & _
           "FROM products " & _
           "WHERE pcUrlBundle = '" & Replace(mmBunMonSlug, "'", "''") & "'" & _
           "  AND pcprod_DisplayLayout = 'monitor'" & _
           "  AND active = -1 AND removed = 0"
Set mmBunRs = connTemp.Execute(mmBunSql)
If Not mmBunRs.EOF Then
  mmBunMid         = CLng(mmBunRs("idProduct"))
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
' 4. Bundle discount & Cable cost by monitor count
'    (matches inc_headerDAJS.asp numBunDiscount)
' ------------------------------------------------------------
Dim mmBunDiscount
Select Case mmBunMonCount
  Case 2, 3 : mmBunDiscount = 25
  Case 4, 5 : mmBunDiscount = 50
  Case 6, 8 : mmBunDiscount = 100
  Case Else : mmBunDiscount = 0
End Select

Dim mmCableCost 
mmCableCost = mmBunMonCount * 15

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
mmBunMonDispName   = Replace(mmBunMonName, "Monitor", "") 

' ------------------------------------------------------------
' 7. Image URLs with fallback
' ------------------------------------------------------------
Dim mmBunStandImgSrc, mmBunMonImgSrc, mmBunStandImgLgSrc
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
If mmBunStandImgLg <> "" Then
  mmBunStandImgLgSrc = "/shop/pc/catalog/" & mmBunStandImgLg
Else
  mmBunStandImgLgSrc = "/shop/pc/catalog/no_image.gif"
End If

'DA Edit - Build up array picture based off product IDs
Dim mmBunArrImgSrc, mmBunArrMonCode, mmBunArrStdCode
Select Case mmBunSid
  Case 326 : mmBunArrStdCode = "s2v"
  Case 287 : mmBunArrStdCode = "s2h"
  Case 312 : mmBunArrStdCode = "s3h"
  Case 324 : mmBunArrStdCode = "s3p"
  Case 313 : mmBunArrStdCode = "s4s"
  Case 337 : mmBunArrStdCode = "s4sp"
  Case 325 : mmBunArrStdCode = "s4p"
  Case 327 : mmBunArrStdCode = "s4h"
  Case 318 : mmBunArrStdCode = "s5p"
  Case 338 : mmBunArrStdCode = "s6r"
  Case 314 : mmBunArrStdCode = "s6rp"
  Case 319 : mmBunArrStdCode = "s8r"
End Select

Select Case mmBunMid
  Case 304 : mmBunArrMonCode = "a22"
  Case 317 : mmBunArrMonCode = "a24"
  Case 328 : mmBunArrMonCode = "a27"
  Case 320 : mmBunArrMonCode = "i23"
  Case 344 : mmBunArrMonCode = "i27"
  Case 345 : mmBunArrMonCode = "i27"
End Select

mmBunArrImgSrc = "/images/bundles/" & mmBunArrStdCode & "-" & mmBunArrMonCode & "-blg.png"

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

' ------------------------------------------------------------
' 9. SEO helpers - <title> + meta description for bundle pages
'    pcv_PageName drives inc_headerV5.asp's <title> output.
'    pcv_DefaultDescription is picked up by GenerateMetaTags()
'    in include-metatags.asp because bundle pages don't pass
'    idproduct/idcategory in the querystring, so its product/
'    category branches all skip and execution lands on the
'    default-description fallback.
'
'    Dim'd at script scope here so mmSetBundleSeo's assignments
'    modify the page-level variables. Without these Dims VBScript
'    (no Option Explicit) would create sub-local variables that
'    vanish when the sub returns - and inc_headerV5.asp would see
'    pcv_PageName as empty, skipping the <title> output.
' ------------------------------------------------------------
Dim pcv_PageName, pcv_DefaultDescription
pcv_PageName = ""
pcv_DefaultDescription = ""

Function mmBunScreenWord(n)
  Select Case n
    Case 1 : mmBunScreenWord = "Single"
    Case 2 : mmBunScreenWord = "Dual"
    Case 3 : mmBunScreenWord = "Triple"
    Case 4 : mmBunScreenWord = "Quad"
    Case 5 : mmBunScreenWord = "Five"
    Case 6 : mmBunScreenWord = "Six"
    Case 8 : mmBunScreenWord = "Eight"
    Case Else : mmBunScreenWord = CStr(n)
  End Select
End Function

Sub mmSetBundleSeo(byVal machineName)
  Dim word : word = mmBunScreenWord(mmBunMonCount)
  pcv_PageName = machineName & " " & word & " Screen Bundle | Multiple Monitors"
  pcv_DefaultDescription = "Save £" & mmBunDiscount & _
    " on a complete trading setup: " & mmBunMonCount & " x " & _
    Trim(mmBunMonDispName) & ", " & Trim(mmBunStandDispName) & _
    " and " & machineName & ". Free UK delivery, built to order."
End Sub

' Slug guard - matches the URL segment shape we accept from IIS.
Function mmBunIsValidSlug(ByVal s)
  Dim i, ch, ok
  mmBunIsValidSlug = False
  If Len(s) = 0 Or Len(s) > 80 Then Exit Function
  For i = 1 To Len(s)
    ch = Mid(s, i, 1)
    ok = (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Or ch = "-"
    If Not ok Then Exit Function
  Next
  mmBunIsValidSlug = True
End Function
%>
