<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

  '// Locate preferred results count and load as default
    Dim pcIntPreferredCountSearch
    pcIntPreferredCountSearch =(scPrdRow*scPrdRowsPerPage)
%>
<form action="showsearchresults.asp" name="search" method="get">
	<input type="hidden" name="pageStyle" value="<%=bType%>">
	<input type="hidden" name="resultCnt" value="<%=pcIntPreferredCountSearch%>">
	<input type="Text" name="keyword" size="14" value="" id="smallsearchbox" >
    <a href="javascript:document.search.submit()" title="Search"><img src="<%=pcf_getImagePath("images","pc2009-search.png")%>" alt="Search" align="absbottom"></a>
    <div style="margin-top: 3px;">
		<a href="search.asp">More search options</a>
	</div>
</form>
