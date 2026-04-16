<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="CustLIv.asp"-->
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/sendmail.asp"-->

<!--#include file="checkdate.asp" -->
<!--#include file="header_wrapper.asp"-->

<!--#include file="../htmleditor/editor.asp"-->

<div id="pcMain">
	<div class="pcMainContent">
		<h1><%= dictLanguage.Item(Session("language")&"_viewPostings_3")%></h1>
    
		<%
    'Allow upload: change to "0" to disallow
    AllowUpload="1"

    LngIdOrder=clng(getUserInput(request("IDOrder"),0))-clng(scpre)
    intIdFeedback=getUserInput(request("IDFeedback"),0)
    
    query="SELECT * FROM pcComments WHERE pcComm_IDFeedback=" & intIdFeedback & " and pcComm_IDParent=0 and pcComm_IDOrder=" & LngIdOrder & " and pcComm_IDUser=" & session("IDCustomer")
    set rs=connTemp.execute(query)
     
    if rs.eof then
      call closedb()
      response.redirect "userviewfeedback.asp?IDOrder=" & LngIdOrder & "&IDFeedback=" & intIdFeedback & "&r=1&msg="&dictLanguage.Item(Session("language")&"_editFeedback_b")
    end if
    
    'Update feedback
    if (request("action")="update") and (request("rewrite")="0") then
      LngIdOrder=getUserInput(request("IDOrder"),0)
      strFDesc=getUserInput(request("Description"),0)
      strFDetails=getUserInput(request("Details"),0)
      intFStatus=getUserInput(request("FStatus"),0)
      intFType=getUserInput(request("FType"),0)
      intPriority=getUserInput(request("Priority"),0)
      
      dtComDate=CheckDateSQL(now())
      
      query="UPDATE pcComments SET pcComm_EditedDate='" & dtComDate & "',pcComm_FType=" & intFType & ",pcComm_FStatus=" & intFStatus & ",pcComm_Priority=" & intPriority & ",pcComm_Description=N'" & strFDesc & "',pcComm_Details=N'" & strFDetails & "' WHERE pcComm_IDOrder=" & (LngIdOrder)-scpre & " and pcComm_IDFeedback=" & intIdFeedback
      set rs=connTemp.execute(query)
    
      if AllowUpload="1" then
        ACount=getUserInput(request("ACount"),0)
        if ACount<>"" then
          ACount1=clng(ACount)
          For k=1 to ACount1
            if request("AC" & k)="1" then
              query="UPDATE pcUploadFiles SET pcUpld_IDFeedback=" & intIdFeedback & " WHERE pcUpld_IDFile=" & getUserInput(request("AID" & k),0)
              set rs=connTemp.execute(query)
            else
              query="SELECT pcUpld_Filename FROM pcUploadFiles WHERE pcUpld_IDFile=" & getUserInput(request("AID" & k),0) & " and pcUpld_IDFeedback=" & intIdFeedback
              set rs=connTemp.execute(query)
              if not rs.eof then
                strFileName=rs("pcUpld_Filename")
                if strFileName<>"" then
                  QfilePath="Library/" & strFileName
                  findit = Server.MapPath(QfilePath)
                  Set fso = server.CreateObject("Scripting.FileSystemObject")
                  Set f = fso.GetFile(findit)
                  f.Delete
                  Set fso = nothing
                  Set f = nothing
                  Err.number=0
                  Err.Description=""
                end if
              end if
    
              query="DELETE FROM pcUploadFiles WHERE pcUpld_IDFeedback=" & intIdFeedback & " and pcUpld_IDFile=" & getUserInput(request("AID" & k),0)
              set rs=connTemp.execute(query)
            end if
          next
        end if
      end if
      %>
      <div class="pcErrorMessage">
        <%= dictLanguage.Item(Session("language")&"_editFeedback_a")%>
      </div>
      <%end if%>                

      <script type=text/javascript>
				function Form1_Validator(theForm)
				{
				<%if session("UserType")=3 then%>
						if (theForm.FType.value == "")
					{
								alert("<%= dictLanguage.Item(Session("language")&"_editFeedback_d")%>");
								theForm.FType.focus();
								return (false);
					}
							if (theForm.Priority.value == "")
					{
								alert("<%= dictLanguage.Item(Session("language")&"_editFeedback_e")%>");
								theForm.Priority.focus();
								return (false);
					}
								if (theForm.FStatus.value == "")
					{
								alert("<%= dictLanguage.Item(Session("language")&"_editFeedback_f")%>");
								theForm.FStatus.focus();
								return (false);
					}
				<%end if%>
				
							if (theForm.Description.value == "")
					{
								alert("<%= dictLanguage.Item(Session("language")&"_editFeedback_g")%>");
								theForm.Description.focus();
								return (false);
					}
					
							if (theForm.Details.value == "")
					{
								alert("<%= dictLanguage.Item(Session("language")&"_editFeedback_h")%>");
								theForm.Details.focus();
								return (false);
					}
					
				return (true);
				}
				
				function newWindow(file,window) {
					msgWindow=open(file,window,'resizable=no,width=400,height=500');
					if (msgWindow.opener == null) msgWindow.opener = self;
				}
      </script>
      
      <% query="SELECT pcComm_FType,pcComm_FStatus,pcComm_Priority,pcComm_Description,pcComm_Details FROM pcComments WHERE pcComm_IDFeedback=" & intIdFeedback & ";"
      set rs=connTemp.execute(query)
      intFType=rs("pcComm_FType")
      intFStatus=rs("pcComm_FStatus")
      intPriority=rs("pcComm_Priority")
      strDesc=rs("pcComm_Description")
      strDetails=rs("pcComm_Details")
      %>
      
      <form name="hForm" method="post" action="usereditFeedback.asp?action=update" onSubmit="return Form1_Validator(this)" class="pcForms">
        <input type="hidden" name=IDOrder value="<%=scpre+clng(LngIdOrder)%>">
        <input type="hidden" name=IDFeedback value="<%=intIdFeedback%>">
        <div class="pcSectionTitle"><%= dictLanguage.Item(Session("language")&"_editFeedback_c")%></div>
        <div class="pcShowContent">
        	<% 'Order Number %>
          <div class="pcFormItem">
          	<div class="pcFormLabelRight">
            	<%= dictLanguage.Item(Session("language")&"_viewFeedback_b")%>
            </div>
            <div class="pcFormField">
            	<b><%=scpre+clng(LngIdOrder)%></b>
            </div>
          </div>
          
        <%if (session("UserType")=3) then%>  
        	<% 'Message Type %>
          <div class="pcFormItem">
          	<div class="pcFormLabelRight">
              <%= dictLanguage.Item(Session("language")&"_viewFeedback_h")%>
            </div>
            <div class="pcFormField">
              <select name="FType">
                <option value=""></option>
                <% query="SELECT pcFType_IDType,pcFType_Name FROM pcFTypes"
                set rs=connTemp.execute(query)
                do while not rs.eof %>
                  <option value="<%=rs("pcFType_idtype")%>" <% if rs("pcFType_idtype")=intFType then%>selected<%end if%> ><%=rs("pcFType_name")%></option>
                  <%rs.MoveNext
                Loop%>
              </select>
            </div>
          </div>
          
        	<% 'Priority %>
          <div class="pcFormItem">
          	<div class="pcFormLabelRight">
              <%= dictLanguage.Item(Session("language")&"_viewFeedback_i")%>
            </div>
            <div class="pcFormField">
              <select name="Priority">
                <option value=""></option>
                <% query="SELECT pcPri_idPri,pcPri_name FROM pcPri_Priority"
                set rs=connTemp.execute(query)
                do while not rs.eof %>
                  <option value="<%=rs("pcPri_idPri")%>" <% if rs("pcPri_idPri")=intPriority then%>selected<%end if%>><%=rs("pcPri_name")%></option>
                  <%rs.MoveNext
                Loop%>
              </select>
            </div>
          </div>
          
        	<% 'Current Status %>
          <div class="pcFormItem">
          	<div class="pcFormLabelRight">
            	<%= dictLanguage.Item(Session("language")&"_viewFeedback_l")%>
            </div>
            <div class="pcFormField">
              <select name="FStatus">
                <option value=""></option>
                <% query="SELECT pcFStat_idStatus,cFStat_name FROM pcFStatus"
                set rs=connTemp.execute(query)
                do while not rs.eof
                  %>
                  <option value="<%=rs("pcFStat_idStatus")%>" <% if rs("pcFStat_idStatus")=intFStatus then%>selected<%end if%>><%=rs("pcFStat_name")%></option>
                  <%rs.MoveNext
                Loop%>
              </select> 
            </div>
          </div>
          
        <%else 'Not Admins%>
        	<% 'Message Type %>
          <div class="pcFormItem">
          	<div class="pcFormLabelRight">
              <%= dictLanguage.Item(Session("language")&"_viewFeedback_h")%>
            </div>
            <div class="pcFormField">
							<% query="SELECT pcFType_name FROM pcFTypes WHERE pcFType_IDType=" & intFType
              set rs=connTemp.execute(query)
              if not rs.eof then %>
                <%=rs("pcFType_name")%>
              <%else%>
                &nbsp;
              <% end if %>
            </div>
          </div>
          
        	<% 'Priority %>
          <div class="pcFormItem">
          	<div class="pcFormLabelRight">
            	<%= dictLanguage.Item(Session("language")&"_viewFeedback_i")%>
            </div>
            <div class="pcFormField">
							<% query="SELECT pcPri_name FROM pcPriority WHERE pcPri_IDPri=" & intPriority
              set rs=connTemp.execute(query)
              if not rs.eof then %>
                <%=rs("pcPri_name")%>
              <%end if%>
            </div>
          </div>
          
        	<% 'Current Status %>
          <div class="pcFormItem">
          	<div class="pcFormLabelRight">
              <%= dictLanguage.Item(Session("language")&"_viewFeedback_l")%>
            </div>
            <div class="pcFormField">
							<% query="SELECT pcFStat_name FROM pcFStatus WHERE pcFStat_IDStatus=" & intFStatus
              set rs=connTemp.execute(query)
              if not rs.eof then %>
                <%=rs("pcFStat_name")%>
              <%else%>
                &nbsp;
              <% end if %>
              <input type="hidden" name="FType" value="<%=intFType%>">
              <input type="hidden" name="FStatus" value="<%=intFStatus%>">
              <input type="hidden" name="Priority" value="<%=intPriority%>">
            </div>
          </div>
        <%end if 'Check Admins & Users%>
        
        <% 'Short Description %>
        <div class="pcFormItem">
          <div class="pcFormLabelRight">
            <%= dictLanguage.Item(Session("language")&"_viewFeedback_j")%>
          </div>
          <div class="pcFormField">
            <input name="Description" type="text" value="<%if request("Description")<>"" then%><%=request("Description")%><%else%><%=strDesc%><%end if%>" size="25" maxlength="100"> 
          </div>
        </div>
        
        <% 'Details %>
        <div class="pcFormItem">
          <div class="pcFormLabelRight">
            <%= dictLanguage.Item(Session("language")&"_viewFeedback_k")%>
          </div>
          <div class="pcFormField">
            <textarea class="htmleditor" name="Details" cols="56" rows="7" id="bugLongDsc"><%if request("Details")<>"" then%><%=request("Details")%><%else%><%=strDetails%><%end if%></textarea>
          </div>
        </div>
        
        <%if AllowUpload="1" then%>
         	<div class="pcFormItem">
          	<div class="pcFormLabelRight"><%= dictLanguage.Item(Session("language")&"_viewFeedback_s")%></div>
            <div class="pcFormField">
              <%query="SELECT pcUpld_IDFile,pcUpld_FileName FROM pcUploadFiles WHERE pcUpld_IDFeedback=" & intIdFeedback
              set rs=connTemp.execute(query)
              if rs.eof then%>
                <%= dictLanguage.Item(Session("language")&"_viewFeedback_6")%>
              <%else
                ACount=0
                do while not rs.eof
                  ACount=ACount+1 %>
                  <input type="hidden" name="AID<%=ACount%>" value="<%=rs("pcUpld_IDFile")%>">
                  <input type="checkbox" name="AC<%=ACount%>" value="1" checked class="clearBorder">
                  <%
                  strFileName= rs("pcUpld_FileName")
                  strFileName = mid(strFileName,instr(strFileName,"_")+1,len(strFileName))%>
                  <%=strFileName%><br>
                  <%rs.MoveNext
                loop%>
                <input type="hidden" name="ACount" value="<%=ACount%>">
              <%end if%>
              <script type=text/javascript>
                function newWindow1(file,window) {
									catWindow=open(file,window,'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=400,height=360');
									if (catWindow.opener == null) catWindow.opener = self;
                }
              </script>
              <br>
              <%= dictLanguage.Item(Session("language")&"_viewFeedback_7")%><a href="#" onClick="javascript:newWindow1('userfileuploada_popup.asp?IDFeedback=<%=intIdFeedback%>&amp;ReLink=<%=Server.URLencode("usereditfeedback.asp?IDOrder=" & scpre+clng(LngIdOrder) & "&IDFeedback=" & intIdFeedback)%>','window2')"><%= dictLanguage.Item(Session("language")&"_viewFeedback_8")%></a>
            </div>
          </div>
        <%end if%>
        
        <div class="pcSpacer"></div>
        
        <div class="pcFormItem">
        	<div class="pcFormLabelRight"></div>
          <div class="pcFormField">
            <button class="pcButtonUpdate" onClick="document.hForm.rewrite.value='0';" name="Submit" value="Update" >Update</button>
            <button class="pcButtonBack" onClick="location='userviewfeedback.asp?IDOrder=<%=scpre+clng(LngIdOrder)%>&amp;IDFeedback=<%=intIdFeedback%>';">Back</button>
            
            <%if session("IDOrder")>0 then%>
              <a class="pcButtonOrderMessages" href="userviewallposts.asp?IDOrder=<%=session("IDOrder")%>"><%=dictLanguage.Item(Session("language")&"_viewFeedback_10") %></a>
            <%end if%>
            
            <input type="hidden" name="uploaded" value="">
            <input type="hidden" name="rewrite" value="1">
          </div>
        </div>
    	</div>
  	</form>
	</div>
</div>
<!--#include file="footer_wrapper.asp" -->
