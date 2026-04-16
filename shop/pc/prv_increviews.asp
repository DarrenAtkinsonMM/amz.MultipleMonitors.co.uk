<%
'Review Text Length Limitation
pcv_RevLenLimit=500
tmp_strMore=dictLanguage.Item(Session("language")&"_viewPrd_21")
tmp_strMore=replace(tmp_strMore,"...","")

IF pcv_Active="1" THEN
	query="SELECT pcRE_IDProduct FROM pcRevExc WHERE pcRE_IDProduct=" & pcv_IDProduct
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=connTemp.execute(query)

	if err.number<>0 then
		call LogErrorToDatabase()
		set rs=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if
	
	if rs.eof then
		Prv_Accept=1
	else
		Prv_Accept=0
	end if
	set rs=nothing
	
	IF Prv_Accept=1 THEN
		Call CreateList() %>

		<div class="pcShowProductReviews" id="productReviews">
			<div class="pcSectionTitle"><%=dictLanguage.Item(Session("language")&"_prv_1")%></div>
      
			<div class="pcSectionContents" id="pcReviews">
				
				<%
				intCount = pcf_TotalReviewCount(pIDProduct)

				if pcv_ShowRatSum="1" then
					pcv_SaveRating = CalRating()
					IF pcv_RatingType="0" then
						
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

						set rs=Nothing
						%>
						<div id="pcReviewAverage">
							<% if pcv_tmpRating>"0" then %>
								<%=dictLanguage.Item(Session("language")&"_prv_2")%>
								<img class="pcReviewFeelIcon" src="<%=pcf_getImagePath("catalog",pcv_Img1)%>" alt="<%=pcv_tmpRating%>% <%=pcv_MainRateTxt1%>"><%=pcv_tmpRating%>% <%=pcv_MainRateTxt1%> (<%=intCount%>&nbsp;<%=dictLanguage.Item(Session("language")&"_prv_7")%>)
							<% end if %>
						</div>
					<% ELSE
						if pcv_CalMain="1" then     ' Can be set independently of sub-ratings 
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

							set rs=nothing
							if pcv_tmpRating>"0" then%>
								<div id="pcReviewRating">
									<%=dictLanguage.Item(Session("language")&"_prv_39")%>
									<% Call WriteStar(pcv_tmpRating,1) %>
								</div>
							<%end if%>
						<% else    ' Will be calculated automatically by averaging sub-ratings
							pcv_tmpRating=pcv_SaveRating
							if pcv_tmpRating>"0" then %>
								<p><%=dictLanguage.Item(Session("language")&"_prv_2")%>
								<% Call WriteStar(pcv_tmpRating,1) %></p>
							<% end if %>
						<% end if
					END IF 'Main Rating %>

				
					<% '******** Display Sub-Rating
					if FCount>"0" then
						For m=0 to FCount-1
							if FType(m)>"2" then
						
								IF FType(m)="3" then
									if FRecord(m)>"0" then
										if FValue(m)="0" then
											pcv_Rpercent=0
										else
											pcv_Rpercent=Round((Fvalue(m)/FRecord(m))*100)
										end if
										if pcv_Rpercent<>"0" then %>
											<div class="pcReviewSubRating">
												<p><%=FName(m)%>:</p>
												<p>
													<img class="pcReviewFeelIcon" src="<%=pcf_getImagePath("catalog",pcv_Img1)%>" alt="<%=pcv_Rpercent%>%"> <%=pcv_Rpercent%>%&nbsp;
													<img class="pcReviewFeelIcon" src="<%=pcf_getImagePath("catalog",pcv_Img2)%>" alt="<%=100-pcv_Rpercent%>%"> <%=100-pcv_Rpercent%>%
												</p>
											</div>
										<% end if
									end if
								ELSE
									if FRecord(m)>"0" then
										if FValue(m)="0" then
											Rev_Rating=0
										else
											Rev_Rating=Fvalue(m)/FRecord(m)
										end if
									else
										Rev_Rating=0
									end if
									if Rev_Rating<>"0" then%>
										<div class="pcReviewSubRating">
											<p><%=FName(m)%>:</p>
											<% Call WriteStar(Rev_Rating,0) %>
										</div>
									<% end if %>
								<% END IF %>
							<% end if
						Next %>
					<% end if
				END IF 'Show Rating Sumary
				%>

				<div class="pcReviewActions">
					<%
						if pcv_RevCount="0" then 
							tmppcv_RevCount=1
						else
							tmppcv_RevCount=pcv_RevCount
						end if
						query="SELECT top " & tmppcv_RevCount & " pcRev_IDReview,pcRev_Date,pcRev_MainRate,pcRev_MainDRate FROM pcReviews where pcRev_IDProduct=" & pcv_IDProduct & " and pcRev_Active=1 ORDER BY pcRev_Date DESC"
						set rs=server.CreateObject("ADODB.RecordSet")
						set rs=connTemp.execute(query)
						if not rs.eof Then
							 If intCount>0 then%>
                                <a href="prv_allreviews.asp?IDProduct=<%=pcv_IDProduct%>&IDCategory=<%=pcv_IDCategory%>"><strong>&gt;&gt;&nbsp;<%=dictLanguage.Item(Session("language")&"_prv_3")%></strong> [<%=intCount%>]</a>
                                &nbsp;|&nbsp;
                            <% End If
						end if
						set rs=nothing%>
						<a href="javascript:openbrowser('<%= Server.HtmlEncode("prv_postreview.asp?IDPRoduct=" & pcv_IDProduct) %>');"><%=dictLanguage.Item(Session("language")&"_prv_4")%></a>
				</div>

				<% if pcv_RatingType="0" then %>
					<div class="pcReviewRate">
						<b><%=dictLanguage.Item(Session("language")&"_prv_5")%></b> 
					
						<input name="feel" id="feel" type="hidden" value="">
					
						<img src="<%=pcf_getImagePath("catalog",pcv_Img1)%>" class="pcReviewFeelIcon" alt="<%= pcv_MainRateTxt2 %>">&nbsp;<input name="feel1" value="2" type="radio" onclick="$pc('#feel').val('2');" class="clearBorder">&nbsp;<%=pcv_MainRateTxt2%>
						<img src="<%=pcf_getImagePath("catalog",pcv_Img2)%>" class="pcReviewFeelIcon" alt="<%= pcv_MainRateTxt3 %>">&nbsp;<input name="feel1" value="1" type="radio" onclick="$pc('#feel').val('1');" class="clearBorder">&nbsp;<%=pcv_MainRateTxt3%> 
						&nbsp;
						<input type="button" value="<%=dictLanguage.Item(Session("language")&"_prv_6")%>" onclick="javascript:openbrowser('<%= Server.HtmlEncode("prv_postreview.asp?IDPRoduct=" & pcv_IDProduct & "&feel=") %>' + $pc('#feel').val());" class="submit2">
					</div>
				<% else %>
					<div class="pcReviewRate">
        		        <b><%=dictLanguage.Item(Session("language")&"_prv_5")%></b> <input name="rate" id="rate" type="hidden" value="">
					
						<%if pcv_CalMain="1" then%>
							<%for k=1 to pcv_MaxRating%>
								<input name="rate1" value="<%=k%>" type="radio" onclick="$pc('#rate').val('<%=k%>');" class="clearBorder">&nbsp;<span class="pcSmallText"><%=k%></span>&nbsp;
							<%next%>
						<%end if%>
					
						<input type="button" value="<%=dictLanguage.Item(Session("language")&"_prv_6")%>" onclick="javascript:openbrowser('<%= Server.HtmlEncode("prv_postreview.asp?IDPRoduct=" & pcv_IDProduct & "&rate=") %>' + $pc('#rate').val());" class="submit2">
					 </div>
				<% end if %>
				<hr>
				<%IF pcv_RevCount>"0" then
                    pcv_CShow=pcv_RevCount
                    iPageCurrent=1%>
                    <!--#include file="prv_incshow.asp"-->
				<%END IF%>
                
			</div>

			<% if RCount>="5" then %>
				<div name="rating1" class="pcForms">
					<p>
					    <% query="SELECT pcRev_IDReview,pcRev_Date,pcRev_MainRate,pcRev_MainDRate FROM pcReviews where pcRev_IDProduct=" & pcv_IDProduct & " and pcRev_Active=1 ORDER BY pcRev_Date DESC"
					    set rs=connTemp.execute(query)
					    if not rs.eof then%>
						    <a href="prv_allreviews.asp?IDProduct=<%=pcv_IDProduct%>&IDCategory=<%=pcv_IDCategory%>"><strong>&gt;&gt;&nbsp;<%=dictLanguage.Item(Session("language")&"_prv_3")%></strong> [<%=intCount%>]</a>&nbsp;|&nbsp;
					    <% end if
					    set rs=nothing %>
					    <a href="javascript:openbrowser('<%= Server.HtmlEncode("prv_postreview.asp?IDPRoduct=" & pcv_IDProduct) %>');"><%=dictLanguage.Item(Session("language")&"_prv_4")%></a>
					</p>
				</div>
			
		<% end if %>
    </div>
	<% END IF
END IF %>