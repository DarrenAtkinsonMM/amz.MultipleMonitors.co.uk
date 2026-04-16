<!DOCTYPE html>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="inc_srcImgsQuery.asp"-->
<% pageTitle = getUserInput(request("src_FormTitle2"),0)

totalrecords=0

Set rstemp=Server.CreateObject("ADODB.Recordset")

rstemp.CacheSize=iPageSize
rstemp.PageSize=iPageSize
rstemp.Open query, Conntemp, adOpenStatic, adLockReadOnly, adCmdText

if err.number <> 0 then
    call LogErrorToDatabase()
    set rstemp = Nothing
    call closeDb()
    response.redirect "techErr.asp?err="&pcStrCustRefID
end If
iPageCount=0
if not rsTemp.eof then 
	totalrecords=clng(rstemp.RecordCount)
	iPageCount=rstemp.PageCount
end if

'--- Get Search Form parameters ---

src_FormTitle1=getUserInput(request("src_FormTitle1"),0)
src_FormTitle2=getUserInput(request("src_FormTitle2"),0)
src_FormTips1=getUserInput(request("src_FormTips1"),0)
src_FormTips2=getUserInput(request("src_FormTips2"),0)
src_DisplayType=getUserInput(request("src_DisplayType"),0)
src_ShowLinks=getUserInput(request("src_ShowLinks"),0)
src_FromPage=getUserInput(request("src_FromPage"),0)
src_ToPage=getUserInput(request("src_ToPage"),0)
src_Button2=getUserInput(request("src_Button2"),0)
src_Button3=getUserInput(request("src_Button3"),0)
form_key1=getUserInput(request("key1"),0)
form_key2=getUserInput(request("key2"),0)
form_key3=getUserInput(request("key3"),0)
form_key4=getUserInput(request("key4"),0)
form_resultCnt=getUserInput(request("resultCnt"),0)
form_order=getUserInput(request("order"),0)
pshowimage=getUserInput(request("showimage"),0)
fid=request.QueryString("fid")
ffid=request.QueryString("ffid")
submit=request("Submit")
submit2=request("Submit2")


'--- End of Search Form parameters ---
Function URLDecode(tmpURL)
	Dim tmp1,tmpArr,i,icount	
	tmp1=tmpURL	
	if tmp1<>"" then
		tmp1=replace(tmp1,"+"," ")
		tmpArr=split(tmp1,"%")
		tmp1=tmpArr(0)
		icount=ubound(tmpArr)
		For i=1 to icount
			tmp1=tmp1 & Chr("&H" & Left(tmpArr(i),2)) & Right(tmpArr(i),len(tmpArr(i))-2)
		Next
	end if	
	URLDecode=tmp1
End Function

%>
<html lang="en">
<head>

<!--#include file="inc_headerV5.asp" -->
<style>
    body {
        background: #fff !important;
        margin: 0px
    }
</style>

<script type=text/javascript>

function openGalleryWindow(url) {
	if (document.all)
		var xMax = screen.width, yMax = screen.height;
	else
		if (document.layers)
			var xMax = window.outerWidth, yMax = window.outerHeight;
		else
			var xMax = 800, yMax=600;
	var xOffset = (xMax - 200)/2, yOffset = (yMax - 200)/2;
	var xOffset = 100, yOffset = 100;

	popupWin = window.open(url,'new_page','width=700,height=535,screenX='+xOffset+',screenY='+yOffset+',top='+yOffset+',left='+xOffset+',scrollbars=auto,toolbars=no,menubar=no,resizable=yes')
}

</script>
</head>
<body id="pcPopup">
<div id="pcMain">
	<div class="pcMainContent">

<% IF request("del")<>"" THEN  
        '=======================================
        ' START delete checked files
        '=======================================
	    IF session("admin")=-1 THEN '2 - Check for admin user
		    Count=request("dCount")
		    if Count>"0" then '3
			    pcv_TestP=0
			    pcv_ErrMsg=""

			    Set fso=server.CreateObject("Scripting.FileSystemObject")
			    PageName="catalog/testing.txt"
			    findit=Server.MapPath(PageName)
    			
			    Set f=fso.OpenTextFile(findit, 2, True)
			    f.Write "test done"
			    if Err.number>0 then
				    pcv_TestP=1
				    pcv_ErrMsg=dictLanguage.Item(Session("language")&"_alert_10")
				    Err.number=0
			    end if
			    Set f=nothing

			    IF pcv_TestP=0 THEN
				    Set f=fso.GetFile(findit)
				    Err.number=0
				    f.Delete
				    if Err.number>0 then
					    pcv_TestP=1
					    pcv_ErrMsg=dictLanguage.Item(Session("language")&"_alert_11")
					    Err.number=0
				    end if
			    END IF

			    Set f=nothing			
			    Set fso=nothing

			    IF pcv_TestP=0 THEN '4
			        pcv_ErrMsg =""
				    For i=1 to Count
					    if request("cimg"&i)<>"" then
						    Set fso = server.CreateObject("Scripting.FileSystemObject")
						    Set f = fso.GetFile(Server.MapPath(URLDecode(request("cimg"&i))))
						    Err.number=0
						    f.Attributes=vbArchive
						    f.Delete
					        if Err.number>0 then
					            pcv_ErrMsg=dictLanguage.Item(Session("language")&"_alert_11")
					            Err.number=0
					            exit for
				            end if
					        
					        fname=mid(request("cimg"&i),instr(1,request("cimg"&i),"/")+1)
							fname=URLDecode(fname)
                            query="DELETE from pcImageDirectory where pcImgDir_Name='"&fname&"'"
                            set rs=server.CreateObject("ADODB.RecordSet")
                            set rs=conntemp.execute(query)
                            set rs=nothing
					       
				            Set f=nothing
				            Set fso = nothing
					        pcv_ErrMsg = pcv_ErrMsg & URLDecode(fname) & " has been deleted.<br>"
					    end if
				    Next
				    Set f=nothing
				    Set fso = nothing
			    END IF '4 %>
			    	<div class="pcErrorMessage"><%=pcv_ErrMsg%></div>
		        <%
		    end if '3
	    END IF '2
        '=======================================
        ' END delete checked files
        '=======================================
   'end if
%>    

<%
		elseIF rstemp.eof THEN
			Dim intNoResults
			intNoResults=1
%>
	<div class="pcErrorMessage"><%= dictLanguage.Item(Session("language")&"_ShowSearch_5")%><br /><br /><a class="btn btn-default" href="imageDir.asp?ffid=smallImageUrl&fid=hForm">Back</a></div>
<% ELSE%>

	<%if src_FormTips2<>"" then%>
	<p><%=src_FormTips2%></p>
		
	<%end if%>

	<!--AJAX Functions-->
	<script type=text/javascript>
  	var iPageCount = <%=iPageCount%>
	</script>
	<!--End of AJAX Functions-->
    
  <span id="runmsg"></span>
		
	
	<% 
    response.write "<script type=text/javascript>"&vbCrlf&vbCrlf
    for i=1 to iPageSize
			response.write "function setForm"&i&"() {"&vbCrlf
			response.write "opener.document."&fid&"."&ffid&".value = document.inputForm"&i&".inputField"&i&".value;"&vbCrlf
			response.write "opener.document."&fid&"."&ffid&".focus();"&vbCrlf
			response.write "self.close();"&vbCrlf
			response.write "return false;"&vbCrlf
			response.write "}"&vbCrlf
    next
    response.write "</script>"&vbCrlf
    %>
		
    <form name="ajaxSearch" class="pcForms">
      <input type="hidden" name="src_DisplayType" value="<%=src_DisplayType%>">
      <input type="hidden" name="src_ShowLinks" value="<%=src_ShowLinks%>">
      <input type="hidden" name="src_FromPage" value="<%=src_FromPage%>">
      <input type="hidden" name="src_ToPage" value="<%=src_ToPage%>">
      <input type="hidden" name="src_Button2" value="<%=src_Button2%>">
      <input type="hidden" name="src_Button3" value="<%=src_Button3%>">
      <input type="hidden" name="key1" value="<%=form_key1%>">
      <input type="hidden" name="key2" value="<%=form_key2%>">
      <input type="hidden" name="key3" value="<%=form_key3%>">
      <input type="hidden" name="key4" value="<%=form_key4%>">
      <input type="hidden" name="resultCnt" value="<%=form_resultCnt%>">
      <input type="hidden" name="order" value="<%=form_order%>">
      <input type="hidden" name="fid" value="<%=fid%>">
      <input type="hidden" name="ffid" value="<%=ffid%>">
      <input type="hidden" name="showimage" value="<%=pshowimage%>">
      <input type="hidden" name="submit" value="<%=submit%>">
      <input type="hidden" name="submit2" value="<%=submit2%>">
      <input type="hidden" name="iPageCurrent" value="1">
      <input type="hidden" name="Imglist" value="">
    </form>

      <div id="resultarea"></div>
      <script type=text/javascript>
				$pc(function() {
					srcImgs();
				});
      </script>
      <%
      pcv_HaveResults=1
      END IF
      set rstemp=nothing%>
      
      <script type=text/javascript>
      var savelist="xml,";
      
      function getImglist()
      {
      var tmp2=savelist;
      var pos=0;
        pos=tmp2.indexOf("xml,");
        var out="xml,";
        var temp = "" + (tmp2.substring(0, pos) + tmp2.substring((pos + out.length), tmp2.length));
        return(temp);
      }
      </script>
      
      <% If iPageCount>1 and pcv_ErrMsg="" Then %>
        <div class="pcPageNav">
        Currently viewing page <span id="currentpage">1</span> of <%=iPageCount%><br>
          <%For I = 1 To iPageCount%>
          <a href="javascript:document.ajaxSearch.iPageCurrent.value='<%=I%>';srcImgs();"><%=I%></a> 
          <% Next %>
        </div>
      <% End If %>
      
      <% 
      set rstemp=nothing      
      %>

    <%if src_FromPage<>"" and intNoResults<>1 then%>
      <div style="text-align: center; padding: 10px;">
      	<form class="pcForms">
        	<input class="btn btn-default" type="button" value="<%=src_Button3%>" onClick="location.href='<%=src_FromPage%>'">
        </form>
  		</div>
    <% end if%>
      
	</div>
</div>
<!--#include file="inc_footer.asp" -->
</body>
</html>
<%
call closeDb()
%>
