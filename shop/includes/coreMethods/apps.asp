<%
Public Sub pcs_doEventHook(pEvent)
    Dim rs
    On Error Resume Next 

    query = "SELECT hook_Uri, hook_Method, hook_Lang FROM pcHooks WHERE hook_Event = '" & pEvent & "' "
    Set rs=connTemp.execute(query)
    If Not rs.Eof Then
        pcv_strUri = rs("hook_Uri")
        pcv_strMethod = rs("hook_Method")
        pcv_strLang = rs("hook_Lang")
    End If
    Set rs = Nothing

    If pcv_strLang = "ASP" Then
        Session("ehidOrder") = qry_ID
        Session("SFMethod") = pcv_strMethod  
        session("idProductRedirect") = pidProduct 
        session("pcGatewayDataIdOrder") = pcGatewayDataIdOrder  
        Session("pcIdCustomer") = pcIdCustomer        
        Session("pcBillingFirstName") = pcBillingFirstName
        Session("pcBillingLastName") = pcBillingLastName
        Session("pcBillingCompany") = pcBillingCompany
        Session("pcBillingAddress") = pcBillingAddress
        Session("pcBillingAddress2") = pcBillingAddress2
        Session("pcBillingCity") = pcBillingCity
        Session("pcBillingStateCode") = pcBillingStateCode
        Session("pcBillingProvince") = pcBillingProvince
        Session("pcBillingCity") = pcBillingCity
        Session("pcBillingPostalCode") = pcBillingPostalCode
        Session("pcBillingCountryCode") = pcBillingCountryCode
        Session("pcBillingPhone") = pcBillingPhone
        Session("pcShippingFirstName") = pcShippingFirstName
        Session("pcShippingLastName") = pcShippingLastName
        Session("pcShippingCompany") = pcShippingCompany
        Session("pcShippingAddress") = pcShippingAddress
        Session("pcShippingAddress2") = pcShippingAddress2
        Session("pcShippingCity") = pcShippingCity
        Session("pcShippingStateCode") = pcShippingStateCode
        Session("pcShippingProvince") = pcShippingProvince
        Session("pcShippingPostalCode") = pcShippingPostalCode
        Session("pcShippingCountryCode") = pcShippingCountryCode        
        Session("pcCustomerEmail") = pcCustomerEmail
        Session("pcBillingTotal") = pcBillingTotal
        Session("tempURL") = tempURL
        Server.Execute(pcv_strUri)    
    End If

    If pcv_strLang="PHP" Then
        response.Write(pcf_GetRequest(pcv_strUri, ""))
    End IF
    
    err.clear
    Session("ehidOrder") = ""
    session("pcGatewayDataIdOrder") = ""         
    Session("pcBillingFirstName") = ""
    Session("pcBillingLastName") = ""
    Session("pcBillingCompany") = ""
    Session("pcBillingAddress") = ""
    Session("pcBillingAddress2") = ""
    Session("pcBillingCity") = ""
    Session("pcBillingStateCode") = ""
    Session("pcBillingProvince") = ""
    Session("pcBillingCity") = ""
    Session("pcBillingPostalCode") = ""
    Session("pcBillingCountryCode") = ""
    Session("pcBillingPhone") = ""
    Session("pcShippingFirstName") = ""
    Session("pcShippingLastName") = ""
    Session("pcShippingCompany") = ""
    Session("pcShippingAddress") = ""
    Session("pcShippingAddress2") = ""
    Session("pcShippingCity") = ""
    Session("pcShippingStateCode") = ""
    Session("pcShippingProvince") = ""
    Session("pcShippingPostalCode") = ""
    Session("pcShippingCountryCode") = ""        
    Session("pcCustomerEmail") = ""
    Session("pcBillingTotal") = ""
    Session("tempURL") = ""
End Sub


Public Sub pcs_addWidget(code)
    Dim rs
    On Error Resume Next 

    query = "SELECT widget_Uri, widget_Method, widget_Lang FROM pcWidgets WHERE widget_Shortcode = '" & code & "' "
    Set rs=connTemp.execute(query)
    If Not rs.Eof Then
        pcv_strUri = rs("widget_Uri")
        pcv_strMethod = rs("widget_Method")
        pcv_strLang = rs("widget_Lang")
    End If
    Set rs = Nothing

    If pcv_strLang = "ASP" Then
        Session("SFMethod") = pcv_strMethod  
        session("idProductRedirect") = pidProduct  
        Server.Execute(pcv_strUri)    
    End If
    
    If pcv_strLang="PHP" Then
        response.Write(pcf_GetRequest(pcv_strUri, ""))
    End IF
    
    err.clear
End Sub


Public Function pcf_dynamicInclude(content)
    On Error Resume Next
    out=""   
    If Instr(content,"#include ")>0 Then
        response.Write "Error: include directive not permitted!"
        response.End
    End If     
    content=replace(content,"<"&"%=","<"&"%response.write ")   
    pos1=instr(content,"<%")
    pos2=instr(content,"%"& ">")
    If pos1>0 Then
      before= mid(content,1,pos1-1)
      before=replace(before,"""","""""")
      before=replace(before,vbcrlf,""""&vbcrlf&"response.write vbcrlf&""")
      before=vbcrlf & "response.write """ & before & """" &vbcrlf
      middle= mid(content,pos1+2,(pos2-pos1-2))
      after=mid(content,pos2+2,len(content))
      out=before & middle & pcf_dynamicInclude(after)
    Else
      content=replace(content,"""","""""")
      content=replace(content,vbcrlf,""""&vbcrlf&"response.write vbcrlf&""")
      out=vbcrlf & "response.write """ & content &""""
    End If
    pcf_dynamicInclude=out
End Function


Public Function pcf_getMappedFileAsString(byVal strFilename)
    On Error Resume Next
    Dim fso,td
    Set fso = Server.CreateObject("Scripting.FilesystemObject")
    Set ts = fso.OpenTextFile(Server.MapPath(strFilename), 1)
    pcf_getMappedFileAsString = ts.ReadAll
    ts.close  
    Set ts = nothing
    Set fso = Nothing
End Function
%>