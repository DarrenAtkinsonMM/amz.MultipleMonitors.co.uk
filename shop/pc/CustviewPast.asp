<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="CustLIv.asp"-->
<%
iPageSize=25
iPageCurrent=getUserInput(request("iPageCurrent"),0)
if iPageCurrent="" then
	iPageCurrent=1
end if
if not IsNumeric(iPageCurrent) then
	response.redirect "CustPref.asp"
end if

query="SELECT idOrder,orderstatus,orderDate,total,ord_OrderName FROM orders WHERE idCustomer=" &Session("idcustomer") &" AND OrderStatus>1 ORDER BY idOrder DESC"
set rstemp=Server.CreateObject("ADODB.Recordset")
rstemp.CacheSize=iPageSize
rstemp.PageSize=iPageSize
rstemp.Open query, Conntemp, adOpenStatic, adLockReadOnly, adCmdText

if err.number <> 0 then
    call LogErrorToDatabase()
    set rstemp = Nothing
    call closeDb()
    response.redirect "techErr.asp?err="&pcStrCustRefID
end If

if rstemp.eof then
	set rstemp=nothing
	call closeDb()
 	response.redirect "msg.asp?message=34"     
end if

iPageCount=rstemp.PageCount

	If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=Cint(iPageCount)
	If Cint(iPageCurrent) < 1 Then iPageCurrent=Cint(1)
	rstemp.AbsolutePage=iPageCurrent
	pCnt=0         

%> 

<!--#include file="header_wrapper.asp"-->
<!-- Header: pagetitle -->
    <header id="pagetitle" class="pagetitle">
		<div class="pt-content">
			<div class="container">
				<div class="row">
					<div class="col-xs-12 pagetitle">
						<div class="wow fadeInDown" data-wow-offset="0" data-wow-delay="0.1s">
							<h3 class="color marginbot-0 h-semi" title="Contact Us">Customer Service Area</h3>
						</div>
					</div>				
				</div>		
			</div>		
		</div>	
    </header>
	<!-- /Header: pagetitle -->

	<section id="intWarranties" class="intWarranties paddingtop-30 paddingbot-70">	
           <div class="container">
				<div class="row">
                	<div class="col-sm-12 warrantyHeading wow fadeInUp" data-wow-offset="0" data-wow-delay="0.1s">
<div id="pcMain">
  <div class="pcMainContent">
    <h1>
			<%
      if session("pcStrCustName") <> "" then
        response.write(session("pcStrCustName") & " - " & dictLanguage.Item(Session("language")&"_CustviewPast_4"))
        else
        response.write(dictLanguage.Item(Session("language")&"_CustviewPast_4"))
      end if
      %>
    </h1>
    
    <div class="pcShowContent">
    
			<%
				col_OrderNumClass			= "col-xs-2 col-sm-2"
				col_OrderStatusClass	= "col-xs-3 col-sm-2"
				col_OrderNameClass		= "col-xs-2 col-sm-2 hidden-xs"
				col_OrderTotalClass		= "col-xs-2 col-sm-2 hidden-xs"
				col_OrderDateClass		= "col-xs-3 col-sm-2"
				col_OrderActionsClass	= "col-xs-4 col-sm-2"
			%>
    	<div id="pcTableCustViewPast" class="pcCartLayout container-fluid">
      	<div class="row pcTableHeader">
        	<div class="<%= col_OrderNumClass %>"><%= dictLanguage.Item(Session("language")&"_CustviewPast_5")%></div>
        	<div class="<%= col_OrderStatusClass %>"><%= dictLanguage.Item(Session("language")&"_CustviewOrd_11")%></div>
            <div class="<%= col_OrderNameClass %>">
						<%if scOrderName="1" then %>
							<%= dictLanguage.Item(Session("language")&"_CustviewPast_9")%>
						<% end if %>
            </div>
        	<div class="<%= col_OrderTotalClass %>"><%= dictLanguage.Item(Session("language")&"_CustviewPast_7")%></div>
        	<div class="<%= col_OrderDateClass %>"><%= dictLanguage.Item(Session("language")&"_CustviewPast_6")%></div>
            <div class="<%= col_OrderActionsClass %>">&nbsp;</div>
        </div>
				     
        <%
					do while not rstemp.eof and pCnt<iPageSize
          	pCnt=pCnt+1
            pIdOrder = rstemp("idOrder")
						porderstatus = rstemp("orderstatus")
						pOrderName = rstemp("ord_OrderName")
						pOrderTotal = rstemp("total")
						pOrderDate = rstemp("orderDate")
          
          	%>
          
            <div class="row">
              <div class="<%= col_OrderNumClass %>"><a href="CustviewPastD.asp?idOrder=<%= (scpre+int(pIdOrder))%>"><%= (scpre+int(pIdOrder))%></a></div>
              <div class="<%= col_OrderStatusClass %>"><!--#include file="inc_orderStatus.asp"--></div>            
              <div class="<%= col_OrderNameClass %>">
								<%if scOrderName="1" then %>
									<%=pOrderName%>
								<% end if %>
              </div>
              <div class="<%= col_OrderTotalClass %>"><%= scCurSign&money(pOrderTotal) %></div>
              <div class="<%= col_OrderDateClass %>"><%= showdateFrmt(pOrderDate) %></div>
      
              <div class="<%= col_OrderActionsClass %>">

                <div class="btn-group">
                  <button type="button" class="btn btn-default dropdown-toggle" data-toggle="dropdown">
                    <%= dictLanguage.Item(Session("language")&"_CustviewPast_10")%> <span class="caret"></span>
                  </button>
                  <ul class="dropdown-menu pull-right" role="menu">
                    <li>
											<a href="CustviewPastD.asp?idOrder=<%= (scpre+int(pIdOrder))%>">
												<span class="glyphicon glyphicon-file"></span>&nbsp;<%= dictLanguage.Item(Session("language")&"_CustviewPast_3")%>
											</a>
                    </li>
                    <li>
											<a href="RepeatOrder.asp?idOrder=<%=pIdOrder%>">
												<span class="glyphicon glyphicon-share-alt"></span>&nbsp;<%= dictLanguage.Item(Session("language")&"_CustviewPast_8")%>
											</a>
                    </li>
                    <% 'Hide/show link to Help Desk
                    If scShowHD <> 0 then %>
                      <li>
												<a href="userviewallposts.asp?idOrder=<%=clng(scpre)+clng(pIdOrder)%>">
													<span class="glyphicon glyphicon-question-sign"></span>&nbsp;<%= dictLanguage.Item(Session("language")&"_viewPostings_3")%>
												</a>
                      </li>                
                    <% end if %>
                  </ul>
                </div>

              </div>
            </div>
          
          <div class="row">
            <div class="col-xs-12"><hr></div>
          </div>
      
          	<%
						
						rstemp.movenext
						loop
						
						set rstemp = nothing
          %>
      </div>
      
			<%
      iRecSize=10

      '*******************************
      ' START Page Navigation
      '*******************************
			
      If iPageCount>1 then %>
        <div class="pcPageNav">
        <%=(dictLanguage.Item(Session("language")&"_advSrcb_4") & iPageCurrent & dictLanguage.Item(Session("language")&"_advSrcb_5") & iPageCount)%>
        &nbsp;-&nbsp;
          <% if iPageCount>iRecSize then %>
          <% if cint(iPageCurrent)>iRecSize then %>
            <a href="CustviewPast.asp?iPageCurrent=1"><%=dictLanguage.Item(Session("language")&"_PageNavigaion_1")%></a>&nbsp;
              <% end if %>
          <% if cint(iPageCurrent)>1 then
                  if cint(iPageCurrent)<iRecSize AND cint(iPageCurrent)<iRecSize then
                      iPagePrev=cint(iPageCurrent)-1
                  else
                      iPagePrev=iRecSize
                  end if %>
                  <a href="CustviewPast.asp?iPageCurrent=<%=cint(iPageCurrent)-iPagePrev%>"><%=dictLanguage.Item(Session("language")&"_PageNavigaion_2")%>&nbsp;<%=iPagePrev%>&nbsp;<%=dictLanguage.Item(Session("language")&"_PageNavigaion_3")%></a>
          <% end if
          if cint(iPageCurrent)+1>1 then
            intPageNumber=cint(iPageCurrent)
          else
            intPageNumber=1
          end if
        else
          intPageNumber=1
        end if

        if (cint(iPageCount)-cint(iPageCurrent))<iRecSize then
          iPageNext=cint(iPageCount)-cint(iPageCurrent)
        else
          iPageNext=iRecSize
        end if

        For pageNumber=intPageNumber To (cint(iPageCurrent) + (iPageNext))
          If Cint(pageNumber)=Cint(iPageCurrent) Then %>
            <strong><%=pageNumber%></strong> 
          <% Else %>
                <a href="CustviewPast.asp?iPageCurrent=<%=pageNumber%>"><%=pageNumber%></a>
          <% End If 
        Next

        if (cint(iPageNext)+cint(iPageCurrent))=iPageCount then
        else
          if iPageCount>(cint(iPageCurrent) + (iRecSize-1)) then %>
            <a href="CustviewPast.asp?iPageCurrent=<%=cint(intPageNumber)+iPageNext%>"><%=dictLanguage.Item(Session("language")&"_PageNavigaion_4")%>&nbsp;<%=iPageNext%>&nbsp;<%=dictLanguage.Item(Session("language")&"_PageNavigaion_3")%></a>
          <% end if

          if cint(iPageCount)>iRecSize AND (cint(iPageCurrent)<>cint(iPageCount)) then %>
              &nbsp;<a href="CustviewPast.asp?iPageCurrent=<%=cint(iPageCount)%>"><%=dictLanguage.Item(Session("language")&"_PageNavigaion_5")%></a>
            <% end if 
        end if 
			%>
      	</div>
      <%
      end if

      '*******************************
      ' END Page Navigation
      '*******************************
      %>
    	<div class="pcSpacer"></div>
			
      <a class="pcButton pcButtonBack" href="custPref.asp">
				<img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_back") %>">
				<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
      </a>
    	
    </div>

  </div>
</div>
				</div>
		    </div>
    </section>	
    <!-- /Section: Welcome -->

<!--#include file="footer_wrapper.asp"-->
