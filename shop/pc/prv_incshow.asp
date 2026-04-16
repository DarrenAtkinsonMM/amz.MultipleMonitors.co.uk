<% IF FCount>0 THEN
	query="SELECT pcRev_IDReview,pcRev_Date,pcRev_MainRate,pcRev_MainDRate FROM pcReviews where pcRev_IDProduct=" & pcv_IDProduct & " and pcRev_Active=1 ORDER BY pcRev_Date DESC"
	Set rs=Server.CreateObject("ADODB.Recordset")

	rs.CacheSize=pcv_CShow
	rs.PageSize=pcv_CShow
	rs.Open query, connTemp, 3, 1
	
	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	IF not rs.eof THEN
		RCount=0
		iPageCount=rs.PageCount
		If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=iPageCount
		If iPageCurrent < 1 Then iPageCurrent=1
		rs.AbsolutePage=iPageCurrent
		pcArrayR=rs.getRows(rs.PageSize)
		intReCount=Ubound(pcArrayR,2)
		set rs=nothing
		
		For v=0 to intReCount
			RCount=RCount+1

			Rev_ID=pcArrayR(0,v)
			Rev_Date=pcArrayR(1,v)
			Rev_MainRate=pcArrayR(2,v)
			Rev_MainDRate=pcArrayR(3,v)
			query="SELECT pcRD_Comment FROM pcReviewsData WHERE pcRD_IDReview=" & Rev_ID & " and pcRD_IDField=1"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=connTemp.execute(query)
			
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if

			Rev_CustName=rs("pcRD_Comment")

			query="SELECT pcRD_Comment FROM pcReviewsData WHERE pcRD_IDReview=" & Rev_ID & " and pcRD_IDField=2"
			set rs=server.CreateObject("ADODB.RecordSet")
			set rs=connTemp.execute(query)
			
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if

			Rev_Title=rs("pcRD_Comment")

			For m=0 to FCount-1
				if (Fi(m)<>"1") and (Fi(m)<>"2") then
					query="SELECT pcRD_Feel,pcRD_Rate,pcRD_Comment FROM pcReviewsData WHERE pcRD_IDReview=" & Rev_ID & " and pcRD_IDField=" & Fi(m)
					set rs=server.CreateObject("ADODB.RecordSet")
					set rs=connTemp.execute(query)
					
					if err.number<>0 then
						call LogErrorToDatabase()
						set rs=nothing
						call closedb()
						response.redirect "techErr.asp?err="&pcStrCustRefID
					end if
					
					IF not rs.eof then
	
					if FType(m)<"3" then
						FValue(m)=rs("pcRD_Comment")
					end if
					if FType(m)="3" then
						FValue(m)=rs("pcRD_Feel")
					end if
					if FType(m)="4" then
						FValue(m)=rs("pcRD_Rate")
					end if
					
					ELSE
					
					if FType(m)<"3" then
						FValue(m)=""
					end if
					if FType(m)="3" then
						FValue(m)=0
					end if
					if FType(m)="4" then
						FValue(m)=0
					end if
					
					END IF
				end if
			Next
			set rs=nothing
			'******* Display Review%>
			<div class="pcReview" itemprop="review" itemscope itemtype="http://schema.org/Review">
				<%if pcv_RatingType="0" then%>
                    <%if Rev_MainRate>"0" then%>
                        <img class="pcReviewFeelIcon" src="<%if Rev_MainRate="2" then%><%=pcf_getImagePath("catalog",pcv_Img1)%><%else%><%=pcf_getImagePath("catalog",pcv_Img2)%><%end if%>" alt="<%= Rev_Title %>">
                    <%end if%>
                <%end if%>
                <p itemprop="name"><b>"<%=Rev_Title%>"</b></p>
				<% If scDateFrmt="DD/MM/YY" then 
                    Rev_Date = day(Rev_Date) & "/" & month(Rev_Date) & "/" & year(Rev_Date) & " " & TimeValue(Rev_Date)
                Else
                    Rev_Date = month(Rev_Date) & "/" & day(Rev_Date) & "/" & year(Rev_Date) & " " & TimeValue(Rev_Date)
                End If
				meta_RevDate = year(Rev_Date) & "-" & month(Rev_Date) & "-" & day(Rev_Date)
				%>
                <p><span itemprop="author" itemscope itemtype="http://schema.org/Person"><span itemprop="name"><%=Rev_CustName%></span></span> on <meta itemprop="datePublished" content="<%=meta_RevDate%>"><%=Rev_Date%></p>
                <p>
				<%if pcv_RatingType="1" then
					if pcv_CalMain="1" then
						Rev_Rating=Rev_MainDRate
					else
						tmp1=0
						tmp2=0
						For m=0 to FCount-1
							if FType(m)="4" then
								if FValue(m)>"0" then
								tmp1=tmp1+1
								tmp2=tmp2+clng(FValue(m))
								end if
							end if
						Next
						if tmp2>"0" then
							Rev_Rating=tmp2/tmp1
						else
							Rev_Rating=0
						end if
					end if
					if ((pcv_CalMain<>"1") and (tmp1>1)) or (pcv_CalMain="1") then
						Call WriteStar(Rev_Rating,1)
					end if
	
				end if%>
                
                <% if Rev_Rating > 0 then %>
                	<span itemprop="reviewRating" itemscope itemtype="http://schema.org/Rating"><meta itemprop="ratingValue" content="<%=Rev_Rating%>" /></span>
                <% end if %>
                </p>

				<% For m=0 to FCount-1
					if (Fi(m)<>"1") and (Fi(m)<>"2") then
						IF FType(m)>"2" then %>
							<div class="pcReviewSubRating">
								<p><%=FName(m)%>:</p>
								<p>
								<% IF FType(m)="3" then
                                    if FValue(m)="2" then%>
                   		<img class="pcReviewFeelIcon" src="<%=pcf_getImagePath("catalog",pcv_Img1)%>" alt="<%= pcv_SubRateTxt1 %>"> <%=pcv_SubRateTxt1%>
									<%else
										if FValue(m)="1" then%>
											<img class="pcReviewFeelIcon" src="<%=pcf_getImagePath("catalog",pcv_Img2)%>" alt="<%= pcv_SubRateTxt2 %>"> <%=pcv_SubRateTxt2%>
										<%else%>
											<%=dictLanguage.Item(Session("language")&"_prv_15")%>
										<% end if
									end if
								ELSE
									Rev_Rating=FValue(m)
									Call WriteStar(Rev_Rating,0)
								END IF %>
								</p>
							</div>
						<% ELSE
							if trim(FValue(m))<>"" then %>
								<div class="pcReviewSpecialField">
									<%tmp_Review=FValue(m)
									if pcv_RevLenLimit>0 then
										tmp_Review1=ClearHTMLTags2(tmp_Review,0)
										if len(tmp_Review1)>pcv_RevLenLimit then
											tmp_Review=Left(tmp_Review1,pcv_RevLenLimit) & "...&nbsp;<a href='prv_allreviews.asp?IDProduct=" & pcv_IDProduct & "&IDCategory=" & pcv_IDCategory & "'>" & tmp_strMore & "</a>"
										end if
									end if%>
									<p><b><%=FName(m)%>:</b> <%=tmp_Review%></p>
								</div>
							<% end if
						END IF
					end if
				Next %>
            </div>
		<% Next
	ELSE
		set rs=nothing
	END IF
END IF
%>