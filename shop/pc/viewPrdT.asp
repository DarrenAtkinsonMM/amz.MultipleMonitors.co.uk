<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2015. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<div id="pcViewProductT" class="pcViewProduct">

	<%if ppTop<>"" then%>
    <div class="pcPageTop">
      <%tmpListStr=split(ppTop,",")
      For El=0 to ubound(tmpListStr)
        if trim(tmpListStr(El))<>"" then
          DisplayZone(trim(tmpListStr(El)))
        end if
      Next%>
    </div>
		<div class="pcClear"></div>
  <%end if%>
  
  <%if ((ppTopLeft<>"") OR (ppTopRight<>"")) then%>
    <%if not IsNull(ppTopLeft) then%>
    <%end if%>
    <%if ppTopLeft<>"" then %>
      <div class="pcPageTopLeft">
      <%
      tmpListStr=split(ppTopLeft,",")
      For El=0 to ubound(tmpListStr)
        if trim(tmpListStr(El))<>"" then
          DisplayZone(trim(tmpListStr(El)))
        end if
      Next
      %>
				<div class="pcClear"></div>
      </div>
      <%
    end if%>      
    <%if Not IsNull(ppTopRight) then %>
      <div class="pcPageTopRight">
      <%
      tmpListStr=split(ppTopRight,",")
      For El=0 to ubound(tmpListStr)
        if trim(tmpListStr(El))<>"" then
          DisplayZone(trim(tmpListStr(El)))
        end if
      Next
      %>
				<div class="pcClear"></div>
      </div>
      <%
    end if%>
  <%end if%>
  
  <%if Not IsNull(ppMiddle) then%>
    <div class="pcClear"></div>
    <div class="pcPageMiddle">
      <%tmpListStr=split(ppMiddle,",")
      For El=0 to ubound(tmpListStr)
        if trim(tmpListStr(El))<>"" then
          DisplayZone(trim(tmpListStr(El)))
        end if
      Next%>
    </div>
  <%end if%>

  <%if ppTabs<>"" then%>
	<div class="pcClear"></div>
    <div class="pcPageTabs">
    <div id="prdtabs">
      <ul class="nav nav-tabs">
          <%
          tmpListStr=split(ppTabs,"||")
          pcv_strActiveTab = "in active"
          For El=0 to ubound(tmpListStr)
            if trim(tmpListStr(El))<>"" then
                tmpListStr1=split(tmpListStr(El),"``")%>
                <li class="<%=pcv_strActiveTab %>"><a data-toggle="tab" href="#tabs-<%=El%>"><%=tmpListStr1(0)%></a></li>
                <%
                pcv_strActiveTab = ""
            end if
          Next
          %>
      </ul>
      <div class="tab-content">
      <%
      pcv_strActiveTab = "in active"
      For El=0 to ubound(tmpListStr)
        if trim(tmpListStr(El))<>"" then
          tmpListStr1=split(tmpListStr(El),"``")          
          %>          
          <div class="tab-pane fade <%=pcv_strActiveTab%> " id="tabs-<%=El%>">
            <div class="tabWrapper">
              <%if tmpListStr1(1)<>"" then
                tmpListStr2=split(tmpListStr1(1),",")
                For Ele=0 to ubound(tmpListStr2)
                  if trim(tmpListStr2(Ele))<>"" then
                    if trim(tmpListStr2(Ele))="CUSTOMHTML" then%>
                      <%=tmpListStr1(2)%>
                    <%else
                      DisplayZone(trim(tmpListStr2(Ele)))
                    end if
                  end if
                Next
              end if
              
              pcv_strActiveTab = ""
              %>
            </div> 
			<div class="pcClear"></div>
          </div>          
        <%
        end if
      Next
      %>
      </div>
    </div>

    </div>
		<div class="pcClear"></div>
  <%end if%>
  
  <%if ppBottom<>"" then%>
    <div id="bottom" class="pcPageBottom">
      <%tmpListStr=split(ppBottom,",")
      For El=0 to ubound(tmpListStr)
        if trim(tmpListStr(El))<>"" then
          DisplayZone(trim(tmpListStr(El)))
        end if
      Next%>
    </div>
		<div class="pcClear"></div>
  <%end if%>
  
  <%pcs_BTOADDON%>
</div>