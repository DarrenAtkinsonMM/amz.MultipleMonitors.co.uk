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

    Dim lngIDOrder,lngIDFeedback,lngIDComment    
    
    lngIDOrder=Clng(getUserInput(request("IDOrder"),0))-clng(scpre)
    lngIDFeedback=getUserInput(request("IDFeedback"),0)
    lngIDComment=getUserInput(request("IDComment"),0)
    
     queryQ="select * from pcComments where pcComm_IDFeedback=" & lngIDComment & " and pcComm_IDParent=" & lngIDFeedback & " and pcComm_IDOrder=" & lngIDOrder & " and pcComm_IDUser=" & session("IDCustomer")
     set rstemp=connTemp.execute(queryQ)
    
     if rstemp.eof then
     call closedb()
     response.redirect "userviewfeedback.asp?IDOrder=" & clng(lngIDOrder) & "&IDFeedback=" & lngIDFeedback & "&r=1&msg=" & dictLanguage.Item(Session("language")&"_editFeedback_b")
     end if
    
    Dim strFDetails,dtComDate,ACount,ACount1
    
    'Create new feedback
    if (request("action")="update") and (request("rewrite")="0") then
      strFDetails=getUserInput(request("Details"),0)
      
      dtComDate=CheckDateSQL(now())
      
      queryQ="UPDATE pcComments SET pcComm_EditedDate='" & dtComDate & "', pcComm_Details=N'" & strFDetails & "' WHERE pcComm_IDOrder=" & lngIDOrder & " AND pcComm_IDFeedback=" & lngIDComment & " AND pcComm_IDParent=" & lngIDFeedback & ";"
    
      set rstemp=connTemp.execute(queryQ)
      
      queryQ="UPDATE pcComments SET pcComm_EditedDate='" & dtComDate & "' WHERE pcComm_IDOrder=" & lngIDOrder & " AND pcComm_IDFeedback=" & lngIDFeedback & ";"
      set rstemp=connTemp.execute(queryQ)
      
      if AllowUpload="1" then
      ACount=getUserInput(request("ACount"),0)
      if ACount<>"" then
      ACount1=clng(ACount)
      For k=1 to ACount1
      if request("AC" & k)="1" then
      queryQ="UPDATE pcUploadFiles SET pcUpld_IDFeedback=" & lngIDComment & " WHERE pcUpld_IDFile=" & getUserInput(request("AID" & k),0)
      set rstemp4=connTemp.execute(queryQ)
      else
      
      queryQ="SELECT * FROM pcUploadFiles WHERE pcUpld_IDFile=" & getUserInput(request("AID" & k),0) & " AND pcUpld_IDFeedback=" & lngIDComment
      set rstemp5=connTemp.execute(queryQ)
      if not rstemp5.eof then
       Filename=rstemp5("pcUpld_Filename")
       if Filename<>"" then
        QfilePath="Library/" & Filename
          findit = Server.MapPath(QfilePath)
        findit1 = findit
        Set fso = server.CreateObject("Scripting.FileSystemObject")
        Set f = fso.GetFile(findit)
        f.Delete
        Set fso = nothing
        Set f = nothing
        Err.number=0
        Err.Description=""
       end if
        end if
    
      queryQ="DELETE FROM pcUploadFiles WHERE pcUpld_IDFeedback=" & lngIDComment & " AND pcUpld_IDFile=" & getUserInput(request("AID" & k),0)
      set rstemp4=connTemp.execute(queryQ)
      
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
					// InnovaStudio HTML Editor Workaround for this keyword
					theForm = document.hForm;

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
				
				function newWindow1(file,window) {
				catWindow=open(file,window,'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=400,height=360');
				if (catWindow.opener == null) catWindow.opener = self;
				}
      </script>
      <%
        queryQ="SELECT * FROM pcComments WHERE pcComm_IDFeedback=" & lngIDComment & " AND pcComm_IDParent=" & lngIDFeedback & " AND pcComm_IDOrder=" & lngIDOrder
        set rstemp=connTemp.execute(queryQ)
        Details=rstemp("pcComm_Details")
       %>
      <div class="pcSectionTitle"><%= dictLanguage.Item(Session("language")&"_editFeedback_c")%></div>
      <form name="hForm" method="post" action="usereditComment.asp?action=update" onSubmit="return Form1_Validator(this)" class="pcForms">
				<input type="hidden" name="Priority" value="<%=Priority%>">
				<input type="hidden" name="FStatus" value="<%=FStatus%>">
				<input type="hidden" name="FType" value="<%=FType%>">
				<input type="hidden" name="IDOrder" value="<%=clng(lngIDOrder)+scpre%>">
				<input type="hidden" name="IDFeedback" value="<%=lngIDFeedback%>">
				<input type="hidden" name="IDComment" value="<%=lngIDComment%>">
				<div class="pcShowContent">
        	<% 'Order ID %>
        	<div class="pcFormItem">
          	<div class="pcFormLabelRight">
							<%= dictLanguage.Item(Session("language")&"_viewPostings_b")%>
            </div>
            <div class="pcFormField">
							<b><%=scpre+clng(lngIDOrder)%></b>
            </div>
          </div>
          
          <% 'Short Description %>
          <div class="pcFormItem">
          	<div class="pcFormLabelRight">
							<%= dictLanguage.Item(Session("language")&"_viewFeedback_j")%>
            </div>
            <div class="pcFormField">
							<%
								queryQ="SELECT * FROM pcComments WHERE pcComm_IDParent=0 AND pcComm_IDOrder=" & lngIDOrder & " AND pcComm_IDFeedback=" & lngIDFeedback
					 			set rstemp1=connTemp.execute(queryQ)
					 			response.write rstemp1("pcComm_Description")
							%>
            </div>
          </div>
          
          <% 'Details %>
          <div class="pcFormItem">
          	<div class="pcFormLabelRight">
							<%= dictLanguage.Item(Session("language")&"_viewFeedback_k")%>
            </div>
            <div class="pcFormField">
							<textarea class="htmleditor" name="Details" cols="56" rows="7" id="bugLongDsc"><%if request("Details")<>"" then%><%=request("Details")%><%else%><%=Details%><%end if%></textarea>
            </div>
          </div>
          
					<%if AllowUpload="1" then%>
          <div class="pcFormItem">
          	<div class="pcFormLabelRight">
							<%= dictLanguage.Item(Session("language")&"_viewFeedback_s")%>
            </div>
            <div class="pcFormField">
							<%queryQ="SELECT * FROM pcUploadFiles WHERE pcUpld_IDFeedback=" & lngIDComment
              set rstemp4=connTemp.execute(queryQ)
              if rstemp4.eof then%>
              <%= dictLanguage.Item(Session("language")&"_viewFeedback_6")%>
              <br>
              <%else
              ACount=0
              do while not rstemp4.eof
              ACount=ACount+1
              %>
              <input type="hidden" name="AID<%=ACount%>" value="<%=rstemp4("pcUpld_IDFile")%>">
              <input type="checkbox" name="AC<%=ACount%>" value="1" checked class="clearBorder">&nbsp;<%
              Filename= rstemp4("pcUpld_FileName")
              FileName = mid(FileName,instr(Filename,"_")+1,len(FileName))%>
              <%=FileName%>
              <br>
              <%rstemp4.MoveNext
              loop%>
              <input type="hidden" name="ACount" value="<%=ACount%>">
              <%end if%>
              <br>
              <%= dictLanguage.Item(Session("language")&"_viewFeedback_7")%> <a href="#" onClick="javascript:newWindow1('userfileuploada_popup.asp?IDFeedback=<%=lngIDComment%>&amp;ReLink=<%=Server.URLencode("usereditcomment.asp?IDComment=" & lngIDComment & "&IDOrder=" & scpre+clng(lngIDOrder) & "&IDFeedback=" & lngIDFeedback)%>','window2')"><%= dictLanguage.Item(Session("language")&"_viewFeedback_8")%></a>
              
            </div>
          </div>
					<%end if%>
        	<div class="pcSpacer"></div>
          <div class="pcFormItem">
          	<div class="pcFormLabelRight">
            </div>
            <div class="pcFormField">
            	<div class="pcFormButtons">
                <button class="pcButtonUpdate" onClick="document.hForm.rewrite.value='0';" name="Submit" value="Update" >Update</button>
                <button class="pcButtonBack" onClick="location='userviewfeedback.asp?IDOrder=<%=scpre+clng(lngIDOrder)%>&amp;IDFeedback=<%=lngIDFeedback%>';">Back</button>
                
                <%if session("IDOrder")>0 then%>
                  <a class="pcButton pcButtonOrderMessages" onclick="location='userviewallposts.asp?IDOrder=<%=session("IDOrder")%>'">dictLanguage.Item(Session("language")&"_viewFeedback_10")</button>
                <%end if%>
         
                <input type="hidden" name="uploaded" value="">
                <input type="hidden" name="rewrite" value="1">
              </div>
            </div>
          </div>
				</div>
		</form>
	</div>
</div>

<!--#include file="footer_wrapper.asp" -->
