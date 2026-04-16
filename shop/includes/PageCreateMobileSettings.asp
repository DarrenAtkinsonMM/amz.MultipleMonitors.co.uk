<!--#include file="common.asp"-->
<!--#include file="adminv.asp"-->
<% 
'// Check permissions on include folder
Dim q, PageName, findit, Body, f, fso

'Mobile Commerce Add-on version
pcvMobileVersion = "5.2.00"

'// Request values
q=Chr(34)
PageName="pcMobileSettings.asp"
findit=Server.MapPath(PageName)

pcvMobileOn=session("adm_MobileOn")

If (len(pcvMobileOn)=0) Then
	response.End() '// Not a valid form.
End If

if pcvMobileOn="" then
	pcvMobileOn=0
end if

pcvMobilePay=session("adm_MobilePay")
if pcvMobilePay="" then
	pcvMobilePay=0
end if

pcvMobileLogo=session("adm_MobileLogo")
pcvMobileShowHomeNav=session("adm_MobileShowHomeNav")
if pcvMobileShowHomeNav="" then
	pcvMobileShowHomeNav=0
end if
pcvMobileShowHomeSP=session("adm_MobileShowHomeSP")
if pcvMobileShowHomeSP="" then
	pcvMobileShowHomeSP=0
end if
pcvMobileShowHomeNA=session("adm_MobileShowHomeNA")
if pcvMobileShowHomeNA="" then
	pcvMobileShowHomeNA=0
end if
pcvMobileShowHomeBS=session("adm_MobileShowHomeBS")
if pcvMobileShowHomeBS="" then
	pcvMobileShowHomeBS=0
end if
pcvMobileShowHomeFP=session("adm_MobileShowHomeFP")
if pcvMobileShowHomeFP="" then
	pcvMobileShowHomeFP=0
end if
pcvMobileShowNavTop=session("adm_MobileShowNavTop")
if pcvMobileShowNavTop="" then
	pcvMobileShowNavTop=0
end if
pcvMobileShowNavBot=session("adm_MobileShowNavBot")
if pcvMobileShowNavBot="" then
	pcvMobileShowNavBot=0
end if
pcvMobileIsApparelAddOn=session("adm_MobileIsApparelAddOn")
if pcvMobileIsApparelAddOn="" then
	pcvMobileIsApparelAddOn=0
end if
pcvMobilePayPalCardTypes=session("adm_MobilePayPalCardTypes")

query="SELECT * FROM pcMobileSettings;"
set rs=connTemp.execute(query)
if not rs.eof then
	query="UPDATE pcMobileSettings SET pcMS_TurnOn=" & pcvMobileOn & ",pcMS_Pay=" & pcvMobilePay & ",pcMS_Logo='" & pcvMobileLogo & "',pcMS_ShowHomeNav=" & pcvMobileShowHomeNav & ",pcMS_ShowHomeSP=" & pcvMobileShowHomeSP & ",pcMS_ShowHomeNA=" & pcvMobileShowHomeNA & ",pcMS_ShowHomeBS=" & pcvMobileShowHomeBS & ",pcMS_ShowHomeFP=" & pcvMobileShowHomeFP & ",pcMS_ShowNavTop=" & pcvMobileShowNavTop & ",pcMS_ShowNavBot=" & pcvMobileShowNavBot & ",pcMS_IsApparelAddOn=" & pcvMobileIsApparelAddOn & ",pcMS_PayPalCardTypes='" & pcvMobilePayPalCardTypes & "';"
	set rs=connTemp.execute(query)
else
	query="INSERT INTO pcMobileSettings (pcMS_TurnOn,pcMS_Pay,pcMS_Logo,pcMS_ShowHomeNav,pcMS_ShowHomeSP,pcMS_ShowHomeNA,pcMS_ShowHomeBS,pcMS_ShowHomeFP,pcMS_ShowNavTop,pcMS_ShowNavBot,pcMS_IsApparelAddOn,pcMS_PayPalCardTypes) VALUES (" & pcvMobileOn & "," & pcvMobilePay & ",'" & pcvMobileLogo & "'," & pcvMobileShowHomeNav & "," & pcvMobileShowHomeSP & "," & pcvMobileShowHomeNA & "," & pcvMobileShowHomeBS & "," & pcvMobileShowHomeFP & "," & pcvMobileShowNavTop & "," & pcvMobileShowNavBot & "," & pcvMobileIsApparelAddOn & ",'" & pcvMobilePayPalCardTypes & "');"
	set rs=connTemp.execute(query)
end if
set rs=nothing

Body=CHR(60)&CHR(37)&CHR(10)
Body=Body & "private const scMobileVersion = "&q&pcvMobileVersion&q&CHR(10)
Body=Body & "private const scMobileOn = "&q&pcvMobileOn&q&CHR(10)
Body=Body & "private const scMobileLogo = "&q&pcvMobileLogo&q&CHR(10)
Body=Body & "private const scMobileShowHomeNav = "&q&pcvMobileShowHomeNav&q&CHR(10)
Body=Body & "private const scMobileShowHomeSP = "&q&pcvMobileShowHomeSP&q&CHR(10)
Body=Body & "private const scMobileShowHomeNA = "&q&pcvMobileShowHomeNA&q&CHR(10)
Body=Body & "private const scMobileShowHomeBS = "&q&pcvMobileShowHomeBS&q&CHR(10)
Body=Body & "private const scMobileShowHomeFP = "&q&pcvMobileShowHomeFP&q&CHR(10)
Body=Body & "private const scMobileShowNavTop = "&q&pcvMobileShowNavTop&q&CHR(10)
Body=Body & "private const scMobileShowNavBot = "&q&pcvMobileShowNavBot&q&CHR(10)
Body=Body & "private const scMobileIsApparelAddOn = "&q&pcvMobileIsApparelAddOn&q&CHR(10)
Body=Body & "private const scMobilePay = "&q&pcvMobilePay&q&CHR(10)
Body=Body & "private const scMobilePayPalCardTypes = "&q&pcvMobilePayPalCardTypes&q&CHR(37)&CHR(62) 

'// Create the file using the FileSystemObject
'// On Error Resume Next
Set fso=server.CreateObject("Scripting.FileSystemObject")
Set f=fso.GetFile(findit)
Err.number=0
f.Delete
if Err.number>0 then
	response.redirect "../"&scAdminFolderName&"/techErr.asp?error="&Server.URLEncode("Permissions Not Set to Modify Constants")
end if
Set f=nothing

Set f=fso.OpenTextFile(findit, 2, True)
f.Write Body
f.Close
Set fso=nothing
Set f=nothing

if request.QueryString("refer")<>"" then
	response.redirect "../"&scAdminFolderName&"/"&request.QueryString("refer")
else
	response.redirect "../"&scAdminFolderName&"/MobileSettings.asp?msg=success"
end if
%>