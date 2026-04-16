<!DOCTYPE html>
<html>
<head>

<!--#include file="../includes/common.asp"-->
<!--#include file="inc_headerV5.asp"-->

<%
Response.Buffer = True
Server.ScriptTimeout = 9999

'-------------------------------
' declare local variables
'-------------------------------
Dim CurFile, PopFile,ShowSub, ShowPic, PictureNo, ref, pShowButton
Dim strPathInfo, strPhysicalPath, doIndex, imgName, lastIndex
Dim intTotPics, intPicsPerRow, intPicsPerPage, intTotPages, intPage, strPicArray()

fid=request.QueryString("fid")
ffid=request.QueryString("ffid")
imgIndex=request.QueryString("btnIndex")
ref=request.QueryString("ref")

'--- Get Search Form parameters ---
form_key1=getUserInput(request("key1"),0)
form_key2=getUserInput(request("key2"),0)
form_key3=getUserInput(request("key3"),0)
form_key4=getUserInput(request("key4"),0)
form_resultCnt=getUserInput(request("resultCnt"),0)
form_order=getUserInput(request("order"),0)
pshowimage=getUserInput(request("showimage"),0)
precords=form_resultCnt
pshowform=request.QueryString("ajaxSearch")


doIndex=request.QueryString("doIndex")
if doIndex="" then
    doIndex=0
else
    doIndex=-1
end if

if pshowimage="YES" then
	intPicsPerRow  = 4
	intPicsPerPage = precords
else
	intPicsPerRow  = 4
	intPicsPerPage = precords
end if

intPage = CInt(Request.QueryString("Page"))
If intPage = 0 Then
	intPage = 1
End If

CurFile = "ImageDir.asp"
PopFile = "showPicture.asp"


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Check if Index exists
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

query="SELECT pcImgDir_Name,pcImgDir_DateIndexed from pcImageDirectory"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
if err.number<>0 then
	'//Logs error to the database
	call LogErrorToDatabase()
	'//clear any objects
	set rs=nothing
	'//close any connections
	call closedb()
	'//redirect to error page
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
if rs.EOF then
    doIndex=-1
else
    lastIndex = rs("pcImgDir_DateIndexed")
end if
set rs=nothing
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Check if Index exists
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

%>

<script src="<%=pcf_getJSPath("../includes/javascripts","productcart.js")%>"></script>
<script src="<%=pcf_getJSPath("../includes/javascripts","productcartCP.js")%>"></script>
<script src="<%=pcf_getJSPath("../includes/javascripts","bootstrap.min.js")%>"></script>
<script src="<%=pcf_getJSPath("../includes/javascripts","bootstrap-datepicker.js")%>"></script>
<script src="<%=pcf_getJSPath("../includes/jquery/opentip","opentip-jquery.min.js")%>"></script>

<link href="<%=pcf_getCSSPath("../includes/jquery/opentip","opentip.css")%>" rel="stylesheet" type="text/css">
<link href="<%=pcf_getCSSPath("css","bootstrap.min.css")%>" rel="stylesheet" type="text/css">
<link href="<%=pcf_getCSSPath("css","datepicker3.css")%>" rel="stylesheet" type="text/css">

<style>
html {
    background-color: #FFF !important;
}
body {
    padding: 10px;
    background-color: #FFF !important;
}
table {
    border-collapse: separate !important;
    border-spacing: 2px !important;
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

function DoSubmit()
{
  document.frmIndex.btnIndex.style.display="none";
  document.frmIndex.btnWait.style.display="block";
	return(true);
    //document.frmIndex.submit();
}

function OnLoad()
{
    window.resizeTo(750, 800);
    
    //center window
    //var x = screen.width/2, y = screen.height/2;
    //window.moveTo(x-300,y-300);
}
</script>
</head>
<%
Set conlayout=Server.CreateObject("ADODB.Connection")
conlayout.Open scDSN
Set RSlayout = conlayout.Execute("Select * From layout Where layout.ID=2")
Set rsIconObj = conlayout.Execute("Select * From icons WHERE id=1")

'// Load validation resources
pcv_strRequiredIcon = rsIconObj("requiredicon")
pcv_strErrorIcon = rsIconObj("errorfieldicon")
%>
<body onLoad="Javascript:OnLoad();" id="pcPopup">

<div id="pcMain">
	<div class="pcShowContent">
		<h1 style="margin-top: 0px">Locate an Image</h1>
		<%

  	
    IF doIndex=-1 and imgIndex="" then
%>    
	    <form action="imageDir.asp" method="get" name="frmIndex">
	        <input type="hidden" name="fid" value="<%=fid%>">
	        <input type="hidden" name="ffid" value="<%=ffid%>">
	        <input type="hidden" name="ref" value="<%=ref%>">
        	
	        <div>
		        <p>
								Searching for an image requires that you first index your image directory.<br /><br />Click on the following button to create a searchable index of all the images contained in the &quot;catalog&quot; folder. Please note that this may take some time if you have a large amount of images in that folder.</p>

			            <input type="submit" name="btnIndex" class="btn btn-default" value="Index"  onclick="Javascript:DoSubmit();">
			            <input type="submit" name="btnWait" class="btn btn-default" value="Please wait...." disabled style="display:none" />
			  </div>
	    </form>
<%

   	elseIF imgIndex="Index" OR pshowform="submit"  THEN '1 - Check form submission
	
	    IF imgIndex="Index"  THEN '1 - Check form submission
            '=======================================
            ' START Index Files
            '=======================================
            strPhysicalPath = Server.MapPath(".\catalog")
            Set objFSO = CreateObject("Scripting.FileSystemObject")
            Set objFolder = objFSO.GetFolder(strPhysicalPath)
            Set objFolderContents = objFolder.Files

            Dim pTodayDate
            pTodayDate=Date()
            if SQL_Format="1" then
                pTodayDate=Day(pTodayDate)&"/"&Month(pTodayDate)&"/"&Year(pTodayDate)
            else
                pTodayDate=Month(pTodayDate)&"/"&Day(pTodayDate)&"/"&Year(pTodayDate)
            end if
            pTodayDate=pTodayDate&" "&Time()
            
            query="DELETE from pcImageDirectory"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
            set rs=nothing
            
            msg=""
            For Each objFileItem in objFolderContents
	            If Ucase(Right(objFileItem.Name,4))=".GIF" OR Ucase(Right(objFileItem.Name,4))=".JPG" OR Ucase(Right(objFileItem.Name,4))=".JPE" OR Ucase(Right(objFileItem.Name,5))=".JPEG" OR Ucase(Right(objFileItem.Name,4))=".PNG" THEN
	            	IF inStr(objFileItem.Name,"'")=0 AND inStr(objFileItem.Name,"%")=0 AND inStr(objFileItem.Name,",")=0 AND inStr(objFileItem.Name,"#")=0 AND inStr(objFileItem.Name,"+")=0 THEN
	            	    query="INSERT INTO pcImageDirectory (pcImgDir_Name,pcImgDir_Type,pcImgDir_Size,pcImgDir_DateUploaded,pcImgDir_DateIndexed) "
						if SQL_Format="1" then
							pDateCreated=Day(objFileItem.DateCreated)&"/"&Month(objFileItem.DateCreated)&"/"&Year(objFileItem.DateCreated)
						else
							pDateCreated=Month(objFileItem.DateCreated)&"/"&Day(objFileItem.DateCreated)&"/"&Year(objFileItem.DateCreated)
						end if
						pDateCreated=pDateCreated&" "&Time()	
               	        query=query&"VALUES('"&objFileItem.Name&"','"&objFileItem.Type&"',"&objFileItem.Size&",'"&pDateCreated&"','"&pTodayDate&"')"
                	    set rs=server.CreateObject("ADODB.RecordSet")
                	    set rs=conntemp.execute(query)
                	    if err.number<>0 then
                	        '//Logs error to the database
                	        call LogErrorToDatabase()
                	        '//clear any objects
                	        set rs=nothing
                	        '//close any connections
                	        call closedb()
                	        '//redirect to error page
                	        response.redirect "techErr.asp?err="&pcStrCustRefID
                	    end if
                	    set rs=nothing
					ELSE
						msg="One or more images in the 'pc/catalog' folder could not be indexed because the file names contained an apostrophe ('), percent sign (%), comma (,), number sign (#), or plus sign (+). Please rename these images and click on the 'Index Now' button again."
					END IF
	            End if
            Next

            set rs=nothing

            lastIndex = pTodayDate
            doIndex=0 

            '=======================================
            ' END Index Files
            '=======================================
        end if
        
    end if
    
    IF doIndex=0  THEN 
%>
        <form action="imageDir.asp" method="get" name="frmReIndex">
            <input type="hidden" name="fid" value="<%=fid%>">
            <input type="hidden" name="ffid" value="<%=ffid%>">
            <input type="hidden" name="ref" value="<%=ref%>">

            <div class="pcShowContent">
							<%if msg<>"" then%>
								<div class="pcSpacer"></div>
								<div class="pcErrorMessage"><%=msg%></div>
								<div class="pcSpacer"></div>
							<%end if%>
							<p>
								Where are the images that I just uploaded? If you don't see images that you recently uploaded, re-index the image directory using the &quot;Index now&quot; button below.
								<br /><br />
								Last Index: <%=lastIndex %>
								<br /><br />
								<input type="submit" name="doIndex" class="btn btn-default" value="Index now">
							</p>
							<div class="pcSpacer"></div>
            </div>
        </form>
		<%
			src_FormTitle1=""
			src_FormTitle2=""
			src_FormTips1="Use the following filters to look for images in your store."
			src_FormTips2=""
			src_IncNormal=1
			src_IncBTO=0
			src_IncItem=0
			src_DisplayType=0
			src_ShowLinks=0
			src_FromPage="imagedir.asp?ajaxSearch=submit&fid="&fid&"&ffid="&ffid&" "
			src_ToPage="imagedir.asp"
			src_Button1=" Search "
			src_Button2=" Continue "
			src_Button3=" Back "
			src_PageSize=""
			UseSpecial=1
			session("srcprd_from")=""
			session("srcprd_where")=""
		%>
        <!--#include file="inc_SrcImgs.asp"-->        
	<% 

	END IF %>
	</div>
</div>

</body>
</html>
<%
call closeDb()
%>