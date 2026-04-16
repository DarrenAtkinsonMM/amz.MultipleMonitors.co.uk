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
		
		Dim r
		r=0		
		
		Dim pcv_idOrder
		
		pcv_idOrder=getUserInput(request("IDOrder"),0)
		if pcv_idOrder<>"" then
			session("IDOrder")=Clng(pcv_idOrder)
		else
			pcv_idOrder=session("IDOrder")
		end if
		
		pcv_idOrder=Clng(pcv_idOrder)-Clng(scpre)
		
		dim pcv_strTemp
		pcv_tempStr=" and idCustomer=" & session("idCustomer")
		query="SELECT idCustomer FROM Orders WHERE idOrder=" & pcv_idOrder & pcv_tempStr
		
		set rs=connTemp.execute(query)
		
		if rs.eof then
			response.redirect "userviewallposts.asp?IDOrder=" & pcv_idOrder
		end if
		
		'Create new feedback
		if (request("action")="add") and (request("rewrite")="0") then
			dim intPriority, strFPriority,dtComDate,strFDesc,FDetails,intFStatus,FType
			strFDesc=getUserInput(request("Description"),0)
			strFDetails=getUserInput(request("Details"),0)
			intFStatus=getUserInput(request("FStatus"),0)
			intFType=getUserInput(request("FType"),0)
			intPriority=getUserInput(request("Priority"),0)
			
			query="SELECT pcPri_name FROM pcPriority WHERE pcPri_IDPri=" & intPriority
			set rs=connTemp.execute(query)
		
			strFPriority=rs("pcPri_name")
			
			dtComDate=CheckDateSQL(now())
		
			query="INSERT Into pcComments (pcComm_IDOrder,pcComm_IDParent,pcComm_IDUser,pcComm_CreatedDate,pcComm_EditedDate,pcComm_FType,pcComm_FStatus,pcComm_Priority,pcComm_Description,pcComm_Details) VALUES (" & pcv_idOrder & ",0," & session("IDCustomer") & ",'" & dtComDate & "','" & dtComDate & "'," & intFType & "," & intFStatus & "," & intPriority & ",N'" & strFDesc & "',N'" & strFDetails & "')"
			set rs=connTemp.execute(query)
			
			query="SELECT pcComm_IDFeedback FROM pcComments WHERE pcComm_IDParent=0 and pcComm_IDOrder=" & pcv_idOrder & " and pcComm_IDUSer=" & session("IDCustomer") & " ORDER BY pcComm_IDFeedback DESC;"
			set rstemp=connTemp.execute(query)
			
			Dim strMsg, intLastFB
			
			if rstemp.eof then
				strMsg=dictLanguage.Item(Session("language")&"_addFB_s")
			else	
				intLastFB=rstemp("pcComm_IDFeedback")
				set rstemp=nothing
				IDFeedback=intLastFB
		
				'Generate View Feedback Link for Store Owner
				strPath=Request.ServerVariables("PATH_INFO")
				dim iCnt,strPathInfo '// v4.5 Removed "strPath" declaration as it is now in inc_headerv5.asp
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

				strURL=strPathInfo & scAdminFolderName & "/login_1.asp?redirectUrl=" & Server.URLEnCode(strPathInfo & scAdminFolderName & "/adminviewfeedback.asp?IDOrder=" & pcv_idOrder & "&IDFeedback=" & IDFeedback)
				
				if AllowUpload="1" then
					ACount=getUserInput(request("ACount"),0)
					if ACount<>"" then
						ACount1=clng(ACount)
						For k=1 to ACount1
							if request("AC" & k)="1" then
								query="UPDATE pcUploadFiles SET pcUpld_IDFeedback=" & intLastFB & " WHERE pcUpld_IDFile=" & getUserInput(request("AID" & k),0) & " and pcUpld_IDFeedback=0"
								set rs=connTemp.execute(query)
							end if
						next
						query="DELETE FROM pcUploadFiles WHERE pcUpld_IDFeedback=0"
						set rs=connTemp.execute(query)
					end if
				end if
			
				'Send mail to Store Owner
				Dim strMsgBody
				strMsgBody=""
				strMsgBody=dictLanguage.Item(Session("language")&"_addFB_email1") & clng(scpre)+clng(pcv_idOrder) & dictLanguage.Item(Session("language")&"_addFB_email2") & VBCrlf & VBCrlf 
				strMsgBody=strMsgBody & dictLanguage.Item(Session("language")&"_addFB_email3") & clng(scpre)+clng(pcv_idOrder) & vbcrlf
				strMsgBody=strMsgBody & dictLanguage.Item(Session("language")&"_addFB_email4") & strFDesc & vbcrlf
				
				query="SELECT Name,LastName,Email FROM Customers WHERE IDCustomer=" & session("IDCustomer")
				set rs=connTemp.execute(query)
				pc_name=rs("name")
				pc_lastname=rs("lastname")
		
				strMsgBody=strMsgBody & dictLanguage.Item(Session("language")&"_addFB_email5") & pc_name & " " & pc_lastname & vbcrlf
				strMsgBody=strMsgBody & dictLanguage.Item(Session("language")&"_addFB_email6") & strFPriority & vbcrlf
				strMsgBody=strMsgBody & dictLanguage.Item(Session("language")&"_addFB_email7") & vbcrlf
				strMsgBody=strMsgBody & dictLanguage.Item(Session("language")&"_addFB_email8") & strURL & VBCrlf & VBCrlf
				strMsgBody=strMsgBody & scCompanyName
				
				Dim strMsgBodyMain
				strMsgBodyMain=scCompanyName & ","&VBCrlf&VBCrlf&strMsgBody
				
				'// Prevent issues with Customer Service E-mail not being set (v4.5)
				Dim strCustServEmail
				strCustServEmail=scCustServEmail
				if trim(strCustServEmail)="" then strCustServEmail=scFrmEmail
				
				call sendmail(scCompanyName,scEmail,strCustServEmail,scCompanyName & dictLanguage.Item(Session("language")&"_addFB_email9") & clng(scpre)+clng(pcv_idOrder),strMsgBodyMain)
				call pcs_hookHelpDeskEmailSent(strCustServEmail)
		
				r=1
				strMsg=dictLanguage.Item(Session("language")&"_addFB_a")
			END IF%>
			 <div class="pcInfoMessage">
				<%=strMsg%>
			 </div>                
		<%end if
		
		if request("uploaded")<>"" then
			session("uploaded")="1"
		else
			session("uploaded")="0"
		end if
		
		
		'Delete Temponary uploaded files
		if request("k")="1" then
		else
			if session("uploaded")="1" then
				session("uploaded")="0"
			else
				query="SELECT pcUpld_Filename FROM pcUploadFiles WHERE pcUpld_IDFeedback=0"
				set rs=connTemp.execute(query)
				do while not rs.eof
					Dim strFilename
					strFilename=rs("pcUpld_Filename")
					on error resume next
					if strFilename<>"" then
						dim strQfileName, findit, f
						strQfileNameName="Library/" & strFilename
						findit = Server.MapPath(strQfileNameName)
						Set fso = server.CreateObject("Scripting.FileSystemObject")
						Set f = fso.GetFile(findit)
						f.Delete
						Set fso = nothing
						Set f = nothing
						Err.number=0
						Err.Description=""
					end if
					rs.MoveNext
				loop
				query="DELETE FROM pcUploadFiles WHERE pcUpld_IDFeedback=0"
				set rs=connTemp.execute(query)
				session("uploaded")="0"
			end if
		end if 
		%>
    
		<script type=text/javascript>
			function Form1_Validator(theForm)
			{
				// InnovaStudio HTML Editor Workaround for this keyword
				theForm = document.hForm;

					if (theForm.FType.value == "")
				{
							alert("Please select a feedback type.");
							theForm.FType.focus();
							return (false);
				}
				
						if (theForm.Priority.value == "")
				{
							alert("Please select a priority.");
							theForm.Priority.focus();
							return (false);
				}
				
				
						if (theForm.Description.value == "")
				{
							alert("Please enter a value for Short Description.");
							theForm.Description.focus();
							return (false);
				}
				
						if (theForm.Details.value == "")
				{
							alert("Please enter a value for Long Description.");
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
		
		<form name="hForm" method="post" action="useraddfeedback.asp?action=add" onSubmit="return Form1_Validator(this)" class="pcForms">
    
    <div class="pcShowContent">
    
      <div class="pcFormItem">
				<%= dictLanguage.Item(Session("language")&"_addFB_b")%>
			</div>
			<%if AllowUpload="1" then%>
      <div class="pcFormItem">
        <ol>
          <li><%= dictLanguage.Item(Session("language")&"_addFB_c")%></li>
          <li><%= dictLanguage.Item(Session("language")&"_addFB_d")%></li>
        </ol>
			</div>
			<%else%>
      <div class="pcFormItem">
				<%= dictLanguage.Item(Session("language")&"_addFB_e")%>
			</div>
			<%end if%>
      <div class="pcSpacer"></div>
			<%if AllowUpload="1" then%>
      <div class="pcFormItem">
      	<div class="pcFormLabel pcFormLabelRight">
					<%= dictLanguage.Item(Session("language")&"_addFB_f")%>
				</div>
				<div class="pcFormField">
					<%
						query="SELECT * FROM pcUploadFiles WHERE pcUpld_IDFeedback=0"
						set rs=connTemp.execute(query)
						if rs.eof then%>
						<%= dictLanguage.Item(Session("language")&"_addFB_g")%>
						<br>
						<% else
							ACount=0
							do while not rs.eof
								ACount=ACount+1 %>
								<input type="hidden" name="AID<%=ACount%>" value="<%=rs("pcUpld_IDFile")%>">
								<input type="checkbox" name="AC<%=ACount%>" value="1" checked class="clearBorder">
								<%
								strFilename= rs("pcUpld_FileName")
								strFilename = mid(strFilename,instr(strFilename,"_")+1,len(strFilename))%>
								<%=strFilename%>
								<br>
								<%
								rs.MoveNext
								loop
								%>
								<input type="hidden" name=ACount value="<%=ACount%>">
								<%end if%>
								<script type=text/javascript>
									function newWindow1(file,window) {
									catWindow=open(file,window,'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=400,height=380');
									if (catWindow.opener == null) catWindow.opener = self;
									}
								</script>
								<br>
								<%= dictLanguage.Item(Session("language")&"_addFB_h")%>
								<a href="#" onClick="javascript:newWindow1('userfileuploada_popup.asp?IDFeedback=0&ReLink=<%=Server.URLencode("useraddfeedback.asp?d=1")%>','window2')"><%= dictLanguage.Item(Session("language")&"_addFB_i")%></a>
             <%end if%>
        </div>
        <div class="pcSpacer"></div>
				<%
        'Default Status = Open
        pcv_tempStr=""
        query="SELECT * FROM pcFStatus"
        set rstemp=connTemp.execute(query)
        do while not rstemp.eof
          IDStatus=rstemp("pcFStat_idstatus")
          SName=ucase(rstemp("pcFStat_name"))
          if SName="OPEN" then
            pcv_tempStr="" & IDStatus
          end if
          rstemp.MoveNext
        Loop
        if pcv_tempStr="" then
          pcv_tempStr="1"
        end if%>
        <input type="hidden" name="FStatus" value="<%=pcv_tempStr%>">
      </div>
        
      <div class="pcFormItem">
      	<div class="pcFormLabel pcFormLabelRight">
					<%= dictLanguage.Item(Session("language")&"_addFB_j")%>
        </div>
        <div class="pcFormField">
          <b><%=scpre+int(pcv_idOrder)%></b>
          <input name="IDOrder" type="hidden" value="<%=scpre+clng(pcv_idOrder)%>">
        </div>
      </div>
      
      <div class="pcFormItem">
      	<div class="pcFormLabel pcFormLabelRight">
					<%= dictLanguage.Item(Session("language")&"_addFB_k")%>
        </div>
        <div class="pcFormField">
          <select name="FType">
            <option value=""></option>
            <% query="SELECT pcFType_idtype,pcFType_name FROM pcFTypes"
            set rstemp=connTemp.execute(query)
            do while not rstemp.eof
              pc_pcFType_idtype=rstemp("pcFType_idtype")
              pc_pcFType_name=rstemp("pcFType_name") %>
              <option value="<%=pc_pcFType_idtype%>" <%if request("FType")<>"" then%><%if clng(request("FType"))=clng(pc_pcFType_idtype) then%>selected<%end if%><%end if%>><%=pc_pcFType_name%></option>
              <%rstemp.MoveNext
            Loop%>
          </select>
        </div>
      </div>
      
      <div class="pcFormItem">
      	<div class="pcFormLabel pcFormLabelRight">
					<%= dictLanguage.Item(Session("language")&"_addFB_l")%>
        </div>
        <div class="pcFormField">
          <select name="Priority">
            <option value=""></option>
            <% query="SELECT * FROM pcPriority"
            set rstemp=connTemp.execute(query)
            dim pc_pcPri_idPri,pc_pcPri_name
            do while not rstemp.eof
              pc_pcPri_idPri=rstemp("pcPri_idPri")
              pc_pcPri_name=rstemp("pcPri_name") %>
             <option value="<%=pc_pcPri_idPri%>" <%if request("Priority")<>"" then%><%if clng(request("Priority"))=clng(pc_pcPri_idPri) then%>selected<%end if%><%end if%>><%=pc_pcPri_name%></option>
             <%rstemp.MoveNext
            Loop%>
          </select> 
        </div>
      </div>
      
      <div class="pcFormItem">
      	<div class="pcFormLabel pcFormLabelRight">
					<%= dictLanguage.Item(Session("language")&"_addFB_q")%>
        </div>
        <div class="pcFormField">
					<input name="Description" type="text" id="bugShortDsc" size="25" maxlength="100" value="<%=request("Description")%>"> 
        </div>
      </div>
      
      <div class="pcFormItem">
      	<div class="pcFormLabel pcFormLabelRight">
					<%= dictLanguage.Item(Session("language")&"_addFB_r")%>
        </div>
        <div class="pcFormField">
					<textarea class="htmleditor" name="details" cols="56" rows="7" id="bugLongDsc"><%=request("Details")%></textarea>
        </div>
      </div>
      
      <div class="pcSpacer"></div>
      
      <div class="pcFormItem">
      	<div class="pcFormLabel">&nbsp;</div>
        <div class="pcFormField">
          <div class="pcFormButtons">   
          	<button class="pcButtonAddFeedback" onClick="document.hForm.rewrite.value='0';">Add Feedback</button>
            <a class="pcButtonBack" onClick="JavaScript:history.back(); return false;">Back</a>
             
            <%if session("IDOrder")>0 then%>
            	<a class="pcButtonOtherMessages" href="userviewallposts.asp?IDOrder=<%=session("IDOrder")%>"><%=dictLanguage.Item(Session("language")&"_viewFeedback_10") %></a>
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
