<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<%PmAdmin=9%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<% pageTitle="Database Clean Up Tool" %>
<!--#include file="AdminHeader.asp"-->
<% 
on error resume next 

If request("purgeorder")<>"" then	

	'// 1.) Optimize Performance/ Purge Customer Sessions
	pcCustSession_Date=Date()
	pcCustSession_Date=dateadd("d",-2,pcCustSession_Date)
	if SQL_Format="1" then
		pcCustSession_Date=Day(pcCustSession_Date)&"/"&Month(pcCustSession_Date)&"/"&Year(pcCustSession_Date)
	else
		pcCustSession_Date=Month(pcCustSession_Date)&"/"&Day(pcCustSession_Date)&"/"&Year(pcCustSession_Date)
	end if

	strDtDelim="'"

	query="DELETE FROM pcCustomerSessions WHERE pcCustSession_Date<"&strDtDelim&pcCustSession_Date&strDtDelim&" ;"	
	set rstemp=server.CreateObject("ADODB.Recordset")
	set rstemp=conntemp.execute(query)

	'CONFIGURATOR ADDON-S
	if scBTO=1 then

	  '// 2.) Optimize Performance/ Purge Configuration Sessions
	  query="DELETE FROM configSessions WHERE (configSessions.idconfigSession NOT IN (SELECT ProductsOrdered.idconfigSession FROM ProductsOrdered WHERE ProductsOrdered.idconfigSession<>0)) AND (configSessions.idconfigSession NOT IN (SELECT SCArray16 FROM pcSavedCartArray));"	
	  set rstemp=server.CreateObject("ADODB.Recordset")
	  set rstemp=conntemp.execute(query)
	  
	  '// 3.) Optimize Performance/ Purge Configuration Wishlist Sessions
	  query="DELETE FROM configWishlistSessions WHERE (configWishlistSessions.idconfigWishlistSession NOT IN (SELECT idconfigWishlistSession FROM wishList));"	
	  set rstemp=server.CreateObject("ADODB.Recordset")
	  set rstemp=conntemp.execute(query)

	end if
	'CONFIGURATOR ADDON-E

	if err.number <> 0 then
		pcErrDescription = err.description
		set rstemp = nothing
		
		call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error in PurgeCustSessions.asp: "& pcErrDescription) 
	end If
	
	set rstemp=nothing	
	
	%>
      <div class="pcCPmessageSuccess">
            Clean Up Successful. 
            <br />
            The database has been purged of unused data. 
            <br />
            Return to the <a href="menu.asp">Start Page</a>.
      </div>
<% else %>
	<form action="PurgeCustSessions.asp" method="post" name="form" id="form" class="pcForms">
		<table class="pcCPcontent">
            <tr> 
				<td colspan="2">
                      <div class="pcCPmessage">Warning! Backup your database before using this tool.</div>
                      <p>                    
                      You may use this tool to completely remove unused data from your database. This tool will reduce the size of your database and may improve performance. This action is permanent and cannot be reversed. We strongly recommend that you backup your database before you run this clean up script. The following data will be removed when you run this tool:
                      </p>
                      <ul>
                          <li>Expired customer sessions (<em>pcCustomerSessions</em> table)</li>
						  <%
                          'CONFIGURATOR ADDON-S
                          if scBTO=1 then 
                              %>	
                              <li>Expired configurations (<em>configSessions</em> table) </li>
							  <li>Expired configurations of Wishlist (<em>configWishlistSessions</em> table) </li>
                              <%
                          end if
                          'CONFIGURATOR ADDON-E
                          %>                          
                      </ul>
					<p>To run the database clean up tool click the &quot;Clean Up Database&quot; button below.</p>
        		</td>
       		</tr>				
          	<tr> 
              	<td colspan="2" class="pcCPspacer"></td>
          	</tr>
          	<tr> 
              	<td colspan="2"><input name="purgeorder" type="submit" id="purgeorder" value="Clean Up Database" class="btn btn-primary"></td>
          	</tr>
		</table>
	</form>
<% 
end if
%>
<!--#include file="AdminFooter.asp"-->
