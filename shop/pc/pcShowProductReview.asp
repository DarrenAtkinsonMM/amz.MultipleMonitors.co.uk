<% 'PRV41 start %>
<%IF pRSActive AND pShowAvgRating THEN

    ' Assign pIDProduct to pcv_IDProduct - pcv_IDProduct is used in prv_incfunctions.asp
    pcv_IDProduct = pIDProduct

	IF pcv_RatingType="0" then

        query = "SELECT pcProd_AvgRating FROM Products WHERE idProduct=" & pIDProduct
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rs=nothing
			call closedb()
			response.redirect "techErr.asp?err=" & pcStrCustRefID
		end if
		pcv_tmpRating=Round(rs("pcProd_AvgRating"),1)
		intCount = pcf_TotalReviewCount(pIDProduct)
		set rs=Nothing		
		%>
		<% If pcv_tmpRating>"0" Then %>
		<div class="pcShowProductRating">
			<img class="pcReviewFeelIcon" src="<%=pcf_getImagePath("catalog",pcv_Img1)%>" alt="<%= pcv_MainRateTxt1 %>">
			<span itemprop="aggregateRating" itemscope itemtype="http://schema.org/AggregateRating"><meta itemprop="ratingValue" content="<%=pcv_tmpRating%>">%<meta itemprop="bestRating" content="100" /> <%=pcv_MainRateTxt1%> (<span itemprop="ratingCount"><%=intCount%></span>&nbsp;<%=dictLanguage.Item(Session("language")&"_prv_7")%>)</span>
		</div>
		<% End If %>
	<%
	ELSE
		if pcv_CalMain="1" Then
		
			query = "SELECT pcProd_AvgRating FROM Products WHERE idProduct=" & pIDProduct
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=connTemp.execute(query)

			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if

			pcv_tmpRating=Round(rs("pcProd_AvgRating"),1)

            query = "SELECT COUNT(*) as ct FROM pcReviews WHERE pcRev_IDProduct=" & pIDProduct & " AND pcRev_Active=1 AND pcRev_MainDRate>0"
            set rs=connTemp.execute(query)

            if err.number<>0 then
                call LogErrorToDatabase()
                set rs=nothing
                call closedb()
                response.redirect "techErr.asp?err="&pcStrCustRefID
            end if
			if not rs.eof then
	            intCount = clng(rs("ct"))
			end if
			set rs=nothing
			
			if pcv_tmpRating>"0" then%>
			<div class="pcShowProductRating">
				<%=dictLanguage.Item(Session("language")&"_prv_39")%>
				<% Call WriteStar(pcv_tmpRating,1) %>
				<span itemprop="aggregateRating" itemscope itemtype="http://schema.org/AggregateRating">
					<meta itemprop="worstRating" content = "1" />
					<meta itemprop="ratingValue" content="<%=pcv_tmpRating%>" />
					<meta itemprop="bestRating" content="<%=pcv_MaxRating%>" />
                    <meta itemprop="ratingCount" content="<%=intCount%>" />
				</span>
			</div>
			<%end if%>
		<% else
		    Call CreateList()
			pcv_SaveRating=CalRating()
			pcv_tmpRating=pcv_SaveRating
			if pcv_tmpRating>"0" then%>
			<div class="pcShowProductRating">
				<% Call WriteStar(pcv_tmpRating,1) %>
				<span itemprop="aggregateRating" itemscope itemtype="http://schema.org/AggregateRating">
					<meta itemprop="worstRating" content = "1" />
					<meta itemprop="ratingValue" content="<%=pcv_tmpRating%>" />
					<meta itemprop="bestRating" content="<%=pcv_MaxRating%>" />
                    <meta itemprop="ratingCount" content="<%=intCount%>" />
				</span>
			</div>
			<%end if%>
		<% end if
	END IF 'Main Rating

END IF%>
<% 'PRV41 end %>
