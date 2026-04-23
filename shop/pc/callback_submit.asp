<%@ LANGUAGE="VBSCRIPT" %>
<%
' ============================================================
' callback_submit.asp
' AJAX endpoint for the sitewide callback modal
' (shop/pc/inc_callModal.asp). POST-only. Runs the anti-spam
' gate, validates the input, emails the sales team via
' sendMail(), returns a JSON response:
'   { "ok": true }                       -> success
'   { "ok": false, "error": "<code>" }   -> user-facing error
' Bot submissions are silent-accepted (fake 'ok:true') so the
' bot believes it succeeded and doesn't retry.
' ============================================================
Response.Buffer = True
%>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<%
Response.Clear()
Response.ContentType = "application/json"
Response.Charset = "UTF-8"
Response.CacheControl = "no-store"

' Only accept POST.
If UCase(Request.ServerVariables("REQUEST_METHOD") & "") <> "POST" Then
    Response.Status = "405 Method Not Allowed"
    Response.Write "{""ok"":false,""error"":""method""}"
    Response.End
End If

Dim pName, pPhone, pTime, pEmail, pTrap, pFillMs
pName   = Trim(Left(Request.Form("name") & "", 80))
pPhone  = Trim(Left(Request.Form("phone") & "", 40))
pTime   = Trim(Left(Request.Form("time") & "", 120))
pEmail  = Trim(Left(Request.Form("email") & "", 120))
pTrap   = Trim(Request.Form("website") & "")    ' honeypot
pFillMs = Trim(Request.Form("fillMs") & "")

' ---------------------------------------------------------------
' ANTI-SPAM GATE — silent reject (fake success) for every hit.
' ---------------------------------------------------------------

' 1. Honeypot: real users can't see the field; bots fill it.
If pTrap <> "" Then
    Response.Write "{""ok"":true}"
    Response.End
End If

' 2. Time check: submissions under 2500ms are bots, and so are
'    ones with no timing payload at all (no JS handler running).
If Not IsNumeric(pFillMs) Then
    Response.Write "{""ok"":true}"
    Response.End
End If
If CLng(pFillMs) < 2500 Then
    Response.Write "{""ok"":true}"
    Response.End
End If

' 3. URLs in the name field are a classic bot tell.
Dim lcName
lcName = LCase(pName)
If InStr(1, lcName, "http://", 1) > 0 _
   Or InStr(1, lcName, "https://", 1) > 0 _
   Or InStr(1, lcName, "www.", 1) > 0 Then
    Response.Write "{""ok"":true}"
    Response.End
End If

' ---------------------------------------------------------------
' VALIDATION — real errors, shown to the user.
' ---------------------------------------------------------------

If Len(pName) = 0 Or Len(pPhone) = 0 Or Len(pTime) = 0 Then
    Response.Write "{""ok"":false,""error"":""missing""}"
    Response.End
End If

' Phone sanity: at least 7 digits somewhere in the value.
Dim digits, i, ch
digits = 0
For i = 1 To Len(pPhone)
    ch = Mid(pPhone, i, 1)
    If ch >= "0" And ch <= "9" Then digits = digits + 1
Next
If digits < 7 Then
    Response.Write "{""ok"":false,""error"":""phone""}"
    Response.End
End If

' Email is optional, but if filled must look plausible.
If Len(pEmail) > 0 Then
    If InStr(pEmail, "@") = 0 Or InStr(pEmail, ".") = 0 Then
        Response.Write "{""ok"":false,""error"":""email""}"
        Response.End
    End If
End If

' ---------------------------------------------------------------
' BUILD + SEND EMAIL
' ---------------------------------------------------------------

Dim emailForBody, refPage, fromAddr, rcptAddr, mailSubject, mailBody

If pEmail <> "" Then
    emailForBody = pEmail
Else
    emailForBody = "(not provided)"
End If

refPage = Request.ServerVariables("HTTP_REFERER") & ""
If refPage = "" Then refPage = "(unknown)"

mailSubject = "Callback request from " & pName

' sendmail.asp's CDOSYS branch sets mail.HTMLBody, so use <br> for
' line breaks rather than vbCrLf (which would collapse in HTML).
mailBody = "<p>New callback request from the website.</p>"
mailBody = mailBody & "<p>"
mailBody = mailBody & "<b>Name:</b> "  & Server.HTMLEncode(pName)        & "<br>"
mailBody = mailBody & "<b>Phone:</b> " & Server.HTMLEncode(pPhone)       & "<br>"
mailBody = mailBody & "<b>Email:</b> " & Server.HTMLEncode(emailForBody) & "<br>"
mailBody = mailBody & "<b>When:</b> "  & Server.HTMLEncode(pTime)
mailBody = mailBody & "</p>"
mailBody = mailBody & "<hr>"
mailBody = mailBody & "<p style=""color:#666; font-size:12px;"">"
mailBody = mailBody & "Submitted: " & Now() & "<br>"
mailBody = mailBody & "From page: " & Server.HTMLEncode(refPage) & "<br>"
mailBody = mailBody & "IP: "        & Server.HTMLEncode(Request.ServerVariables("REMOTE_ADDR") & "")
mailBody = mailBody & "</p>"

' Reply-To is the user's email if they gave one, otherwise the
' store's own address so the header stays well-formed.
If pEmail <> "" Then
    fromAddr = pEmail
Else
    fromAddr = scEmail
End If

' Recipient: customer-service inbox from store settings, with a
' hard-coded fallback so a missing setting doesn't black-hole the
' request.
rcptAddr = scCustServEmail
If Len(Trim(rcptAddr & "")) = 0 Then rcptAddr = "sales@multiplemonitors.co.uk"

' sendmail.asp uses `on error resume next` throughout and captures
' whatever err happens to be in `pcv_errMsg` — which is often set by
' benign CDO property assignments even when the send succeeds. We
' can't distinguish a real send failure from a spurious warning
' without rewriting the shared helper, so we trust the call.
Call sendMail(pName, fromAddr, rcptAddr, mailSubject, mailBody)

Response.Write "{""ok"":true}"
%>
