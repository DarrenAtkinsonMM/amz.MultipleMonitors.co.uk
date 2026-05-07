<%@ LANGUAGE="VBSCRIPT" %>
<% Response.Buffer = True %>
<!--#include file="../includes/common.asp"-->
<%
' ============================================================
' viewPrdRouter.asp
' Lightweight friendly-URL dispatcher. Resolves /products/<slug>/
' to a product row, then Server.Transfers to the right page.
' Server.Transfer gives each target page its own VBScript engine
' instance, so common.asp's `Dim` declarations don't collide.
'
' v2 target pages (monitor, stand) re-resolve their product from
' Request.QueryString("slug") — preserved across Server.Transfer —
' so no Session hand-off is needed and concurrent tabs cannot
' corrupt each other's idProduct.
'
' The legacy Case Else fallthrough into viewPrd.asp still relies
' on Session("idProductRedirect") (viewPrdCode.asp reads it before
' falling back to a `url=` querystring that this rewrite does not
' set), so the session write is kept on that branch only.
'
' mmSlugIsSafe is defined in shop/includes/stringfunctions.asp.
' ============================================================

Dim mmSlug, mmIdProduct, mmLayout, mmRouterRs, mmRouterSql

mmSlug = Trim(Request.QueryString("slug") & "")
If mmSlug = "" Or Not mmSlugIsSafe(mmSlug) Then
    Response.Redirect "/"
End If

mmRouterSql = "SELECT idProduct, pcprod_DisplayLayout FROM products " & _
              "WHERE pcUrl = '" & Replace(mmSlug, "'", "''") & "' " & _
              "  AND active = -1 AND removed = 0"
Set mmRouterRs = connTemp.Execute(mmRouterSql)
If mmRouterRs.EOF Then
    Set mmRouterRs = Nothing
    Response.Redirect "/shop/pc/msg.asp?message=88"
End If
mmIdProduct = CLng(mmRouterRs("idProduct"))
mmLayout    = LCase(mmRouterRs("pcprod_DisplayLayout") & "")
mmRouterRs.Close : Set mmRouterRs = Nothing

Select Case mmLayout
    Case "stand"
        Server.Transfer "viewprd-stand-v2.asp"
    Case "monitor"
        Server.Transfer "viewPrd-Monitor-v2.asp"

    Case Else
        ' Legacy viewPrd.asp / viewPrdCode.asp reads Session("idProductRedirect")
        ' first; its QueryString fallback looks for `url=`, not `slug=`.
        Session("idProductRedirect") = mmIdProduct
        Server.Transfer "viewPrd.asp"
End Select
%>
