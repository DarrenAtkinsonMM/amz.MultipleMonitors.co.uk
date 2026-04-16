<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<%
pcv_request=getUserInput(request("req"),0)
if pcv_request="" then
	response.redirect "default.asp"
end if
%>
<!--#include file="header_wrapper.asp"-->
<div id="pcMain">

  <div class="pcMainContent">
   	<h1><%=dictLanguage.Item(Session("language")&"_sds_confirmrequest_1")%></h1>
    
    <%IF request("action")="upd" THEN
      pcv_CustAllow=getUserInput(request("R1"),0)
      if (pcv_CustAllow="") OR (not IsNumeric(pcv_CustAllow)) then
        pcv_CustAllow="1"
      end if
      query="SELECT idorder FROM Orders WHERE pcOrd_CustRequestStr='" & pcv_request &"';"
      set rs=connTemp.execute(query)
      if err.number<>0 then
        call LogErrorToDatabase()
        set rs=nothing
        call closedb()
        response.redirect "techErr.asp?err="&pcStrCustRefID
      end if
      if rs.eof then%>
        <div class="pcErrorMessage">
          <%=dictLanguage.Item(Session("language")&"_sds_confirmrequest_5")%>
        </div>
      <%response.end
      else
        pcv_IDOrder=rs("idorder")
      end if
      set rs=nothing
      
      query="UPDATE Orders SET pcOrd_CustAllowSeparate=" & pcv_CustAllow & " WHERE pcOrd_CustRequestStr='" & pcv_request &"';"
      set rs=connTemp.execute(query)
  
      if err.number<>0 then
        call LogErrorToDatabase()
        set rs=nothing
        call closedb()
        response.redirect "techErr.asp?err="&pcStrCustRefID
      end if
      set rs=nothing		
      
      'Send Notification E-mail to Store Owner
      pcv_AdmSbj=replace(ship_dictLanguage.Item(Session("language")&"_notifyseparate_sbj_1"),"<ORDER_ID>",(scpre + int(pcv_IDOrder)))
      pcv_AdmMail=""
      if pcv_CustAllow="1" then
        pcv_AdmMail=replace(ship_dictLanguage.Item(Session("language")&"_notifyseparate_msg_1"),"<ORDER_ID>",(scpre + int(pcv_IDOrder))) & vbcrlf
      else
        pcv_AdmMail=replace(ship_dictLanguage.Item(Session("language")&"_notifyseparate_msg_2"),"<ORDER_ID>",(scpre + int(pcv_IDOrder))) & vbcrlf
      end if
          
      strPath=Request.ServerVariables("PATH_INFO")
      iCnt=0
      do while iCnt<2
        if mid(strPath,len(strPath),1)="/" then
          iCnt=iCnt+1
        end if
        if iCnt<2 then
          strPath=mid(strPath,1,len(strPath)-1)
        end if
      loop
    
      strPathInfo="http://" & Request.ServerVariables("HTTP_HOST") & strPath
          
      if Right(strPathInfo,1)="/" then
      else
        strPathInfo=strPathInfo & "/"
      end if
      
      strPathInfo=strPathInfo & scAdminFolderName & "/OrdDetails.asp?id=" & pcv_IDOrder
      pcv_AdmMail=pcv_AdmMail & strPathInfo
      call sendmail (scCompanyName, scEmail, scFrmEmail, pcv_AdmSbj, pcv_AdmMail)
      'End of Send Notification E-mail to Store Owner
      %>
        <div class="pcErrorMessage">
          <%=dictLanguage.Item(Session("language")&"_sds_confirmrequest_7")%>
          <br /><br />
          <a href="default.asp"><%=dictLanguage.Item(Session("language")&"_titles_5")%></a>
        </div>
    <%ELSE
      query="SELECT idorder,pcOrd_CustAllowSeparate FROM Orders WHERE pcOrd_CustRequestStr='" & pcv_request &"';"
      set rs=connTemp.execute(query)
  
      if err.number<>0 then
        call LogErrorToDatabase()
        set rs=nothing
        call closedb()
        response.redirect "techErr.asp?err="&pcStrCustRefID
      end if
    
      if rs.eof then%>
        <div class="pcErrorMessage">
          <%=dictLanguage.Item(Session("language")&"_sds_confirmrequest_5")%>
        </div>
      <%else
        pcv_IDOrder=rs("idorder")
        pcv_CustAllow=rs("pcOrd_CustAllowSeparate")
        if IsNull(pcv_CustAllow) or pcv_CustAllow="" then
          pcv_CustAllow=0
        end if
        if pcv_CustAllow>0 then%>
          <div class="pcErrorMessage">
            <%=dictLanguage.Item(Session("language")&"_sds_confirmrequest_6")%>
          </div>
        <%else%>
        <form method="post" action="sds_AllowSeparateShip.asp?action=upd" name="form1">
        	<div class="pcShowContent">
          	<div class="pcFormItem">
            	<%=dictLanguage.Item(Session("language")&"_sds_confirmrequest_2")%>&nbsp;<strong><%=(scpre + int(pcv_IDOrder))%></strong>
            </div>
            <br/>
            <div class="pcFormItem">
            	<%=dictLanguage.Item(Session("language")&"_sds_confirmrequest_2a")%>
            </div>
            <br/>
            <div class="pcFormItemFull">
            	<input type="radio" name="R1" value="1" id="R1" checked class="clearBorder">
							<label for="R1"><%=dictLanguage.Item(Session("language")&"_sds_confirmrequest_3")%></label>
            </div>
            
            <div class="pcFormItemFull">
            	<input type="radio" name="R1" value="2" id="R2" class="clearBorder">
							<%=dictLanguage.Item(Session("language")&"_sds_confirmrequest_4")%>
          		<input type="hidden" name="req" value="<%=pcv_request%>">
            </div>
            
           	<div class="pcSpacer"></div>
            
            <div class="pcFormButtons">
              <button class="pcButton pcButtonContinue" id="submit" name="Confirm">
                <img src="<%=pcf_getImagePath("",rslayout("submit"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_submit") %>" />
                <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_submit") %></span>
              </button>
            </div>
          </div>
        </form>
        <%end if
      end if
      set rs=nothing
    END IF%>
  </div>
</div>
<!--#include file="footer_wrapper.asp"-->
