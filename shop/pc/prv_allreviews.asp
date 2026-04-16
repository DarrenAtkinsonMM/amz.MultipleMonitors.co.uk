<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%
dim pcv_IDProduct, pIDProduct, pcv_IDCategory

pcv_IDProduct=trim(request("IDProduct"))
	if not validNum(pcv_IDProduct) then
		response.redirect "msg.asp?message=85"
	end if
	pIDProduct=pcv_IDProduct
	
pcv_IDCategory=trim(request("IDCategory"))
	if not validNum(pcv_IDCategory) then
		pcv_IDCategory=0
	end if

queryQ="SELECT description FROM Products WHERE idProduct=" & pIDProduct & ";"
set rsQ=connTemp.execute(queryQ)

if not rsQ.eof then
	pDescription=rsQ("description")
end if
set rsQ=nothing
%>
<!--#include file="header_wrapper.asp"-->
<!--#include file="pcValidateHeader.asp"-->
<% 
'*******************************
' START: Check store on/off, start PC session, check affiliate ID
'*******************************
%>
<!--#include file="pcStartSession.asp"-->
<%
'*******************************
' END: Check store on/off, start PC session, check affiliate ID
'*******************************

' Check to see if the user is updating the product after adding it to the shopping cart
tIndex=0
tUpdPrd=request.QueryString("imode")
if tUpdPrd="updOrd" then
	tIndex=request.QueryString("index")
end if
%>
<!--#include file="prv_getsettings.asp"-->
<!--#include file="prv_incfunctions.asp"-->

<script type="text/javascript">
	var feelInput = null;
	var rateInput = null;
	var feelInput2 = null;
	var rateInput2 = null;

	$pc(document).ready(function() {
		feelInput = $pc("#rating1 [name='feel']");
		rateInput = $pc("#rating1 [name='rate']");
		feelInput2 = $pc("#rating2 [name='feel']");
		rateInput2 = $pc("#rating2 [name='rate']");
	});
</script>

<%
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
		set rs=nothing
		call closedb()
		response.redirect "viewPrd.asp?IDProduct=" & pcv_IDProduct & "&IDCategory=" & pcv_IDCategory
	end if
	
	IF Prv_Accept=1 THEN
		Call CreateList()
		query="SELECT description, active FROM products WHERE idproduct=" & pcv_IDProduct
		set rs=server.CreateObject("ADODB.RecordSet")
		set rs=connTemp.execute(query)
		
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if

		pcv_PrdName=rs("description") 
		pcIntProductStatus = rs("active")
			if pcIntProductStatus=0 or isNull(pcIntProductStatus) or pcIntProductStatus="" then
				set rs = nothing
				call closeDb()
				response.redirect "msg.asp?message=95"
			end if
			
		set rs=nothing
		%>
		<script type=text/javascript>
		function openbrowser(url) {
						self.name = "productPageWin";
						popUpWin = window.open(url,'rating','toolbar=0,location=0,directories=0,status=0,top=0,scrollbars=yes,resizable=1,width=705,height=535');
						if (navigator.appName == 'Netscape') {
										popUpWin.focus();
						}
		}
		</script>
		
		<div id="pcMain" class="pcPrvAllReviews">
			<div class="pcMainContent">
				<h1><%=dictLanguage.Item(Session("language")&"_prv_1")%></h1>

				<div class="pcFormItem">
					<div class="pcFormItemFull">
						<%=dictLanguage.Item(Session("language")&"_prv_10")%><span class="pcShowProductName"><a href="viewPrd.asp?IDProduct=<%=pcv_IDProduct%>&IDCategory=<%=pcv_IDCategory%>"><%=pcv_PrdName%></a></span>
					</div>
				</div>

				<div class="pcSpacer"></div>

				<div id="pcReviewDetails">
						
					<div id="pcReviews">

						<div class="pcReviewActions">
							<a href="javascript:openbrowser('prv_postreview.asp?IDPRoduct=<%=pcv_IDProduct%>');"><%=dictLanguage.Item(Session("language")&"_prv_4")%></a>&nbsp;|&nbsp;<a href="viewPrd.asp?IDProduct=<%=pcv_IDProduct%>&IDCategory=<%=pcv_IDCategory%>"><%=dictLanguage.Item(Session("language")&"_prv_30")%></a>
						</div>

						<%
						if pcv_ShowRatSum="1" then
							%>
							<div id="pcReviewRating">
							<%
							pcv_SaveRating = CalRating()
                            intCount = pcf_TotalReviewCount(pIDProduct)
							If pcv_RatingType="0" Then

                                    query = "SELECT pcProd_AvgRating FROM Products WHERE idProduct=" & pIDProduct
                                    set rs=server.CreateObject("ADODB.RecordSet")
                                    set rs=connTemp.execute(query)
                                    if err.number<>0 then
                                        call LogErrorToDatabase()
                                        set rs=nothing
                                        call closedb()
                                        response.redirect "techErr.asp?err="&pcStrCustRefID
                                    end if
                                    pcv_tmpRating = Round(rs("pcProd_AvgRating"), 1)
                                    set rs=Nothing
									%>
                                    
                                    <% If pcv_tmpRating>"0" Then %>
                                        <div class="pcShowProductRating">
                                            <img class="pcReviewFeelIcon" src="<%=pcf_getImagePath("catalog",pcv_Img1)%>" alt="<%= pcv_MainRateTxt1 %>">
                                            <span itemprop="aggregateRating" itemscope itemtype="http://schema.org/AggregateRating">
                                                <meta itemprop="ratingValue" content="<%=pcv_tmpRating%>">
                                                <meta itemprop="bestRating" content="100" />&nbsp;
                                                <%=pcv_MainRateTxt1%>&nbsp;(<span itemprop="ratingCount"><%=intCount%></span>&nbsp;<%=dictLanguage.Item(Session("language")&"_prv_7")%>)
                                            </span>
                                        </div>
                                    <% End If %>  

                            <% Else '// If pcv_RatingType="0" Then %>
								    <%
								    if pcv_CalMain="1" then														

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
										%>

                                        <% If pcv_tmpRating>"0" Then %>
                                            <div class="pcShowProductRating">
                                                <%=dictLanguage.Item(Session("language")&"_prv_2")%>
                                                <% Call WriteStar(pcv_tmpRating,1) %>
                                                <span itemprop="aggregateRating" itemscope itemtype="http://schema.org/AggregateRating">
                                                    <meta itemprop="worstRating" content = "1" />
                                                    <meta itemprop="ratingValue" content="<%=pcv_tmpRating%>" />
                                                    <meta itemprop="bestRating" content="<%=pcv_MaxRating%>" />
                                                    <meta itemprop="ratingCount" content="<%=intCount%>" />
                                                </span>
                                            </div>
                                        <% End If %>
                                        
									<%else
										pcv_tmpRating=pcv_SaveRating
										%>
                                        
                                        <% If pcv_tmpRating>"0" Then %>
                                            <div class="pcShowProductRating">
                                                <%=dictLanguage.Item(Session("language")&"_prv_2")%>
                                                <% Call WriteStar(pcv_tmpRating,1) %>
                                                <span itemprop="aggregateRating" itemscope itemtype="http://schema.org/AggregateRating">
                                                    <meta itemprop="worstRating" content = "1" />
                                                    <meta itemprop="ratingValue" content="<%=pcv_tmpRating%>" />
                                                    <meta itemprop="bestRating" content="<%=pcv_MaxRating%>" />
                                                    <meta itemprop="ratingCount" content="<%=intCount%>" />
                                                </span>
                                            </div>
                                        <% End If %>
                                        
									<% end if %>
                                    
                            <% End If '// If pcv_RatingType="0" Then %>

                            </div>

							<% '******** Display Sub-Rating
							if FCount>"0" then%>
									<%For m=0 to FCount-1
										if FType(m)>"2" then%>
										<div class="pcReviewSubRating">
											<p><%=FName(m)%>:</p>
											<p>
												<%IF FType(m)="3" then
													if FRecord(m)>"0" then
														if FValue(m)="0" then
															pcv_Rpercent=0
														else
															pcv_Rpercent=Round((FValue(m)/FRecord(m))*100)
														end if %>
														<img class="pcReviewFeelIcon" src="<%=pcf_getImagePath("catalog",pcv_Img1)%>" alt="<%=pcv_SubRateTxt1%>"> <%=pcv_Rpercent%>%
														<img class="pcReviewFeelIcon" src="<%=pcf_getImagePath("catalog",pcv_Img2)%>" alt="<%=pcv_SubRateTxt2%>"> <%=100-pcv_Rpercent%>%
													<%else%>
														<%=dictLanguage.Item(Session("language")&"_prv_15")%>
													<% end if
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
													Call WriteStar(Rev_Rating,0)
												END IF%>
											</p>
										</div>
									<%end if
								Next%>
							</div>
						<%end if
					END IF 'Show Rating Sumary%>

					<hr>
					
					<div class="pcReviewRate">
						<form name="rating1" id="rating1" class="pcForms">
							<%if pcv_RatingType="0" then%>
								<b><%=dictLanguage.Item(Session("language")&"_prv_5")%></b>
							
								<input name="feel" type="hidden" value="">
							
								<img class="pcReviewFeelIcon" src="<%=pcf_getImagePath("catalog",pcv_Img1)%>" alt="<%= pcv_MainRateTxt2 %>">&nbsp;<input name="feel1" value="2" type="radio" onClick="feelInput.val('2');" class="clearBorder">&nbsp;<%=pcv_MainRateTxt2%> 
								<img class="pcReviewFeelIcon" src="<%=pcf_getImagePath("catalog",pcv_Img2)%>" alt="<%= pcv_MainRateTxt3 %>">&nbsp;<input name="feel1" value="1" type="radio" onClick="feelInput.val('1');" class="clearBorder">&nbsp;<%=pcv_MainRateTxt3%> 
							
								&nbsp;

								<input type="button" value="<%=dictLanguage.Item(Session("language")&"_prv_6")%>" onClick="javascript: openbrowser('prv_postreview.asp?IDPRoduct=<%=pcv_IDProduct%>&feel=' + feelInput.val());" class="submit2">
							<%else%>
								<b><%=dictLanguage.Item(Session("language")&"_prv_5")%></b>
							
								<input name="rate" type="hidden" value="">
							
								<%if pcv_CalMain="1" then%><%for k=1 to pcv_MaxRating%><input name="rate1" value="<%=k%>" type="radio" onClick="rateInput.val('<%=k%>');" class="clearBorder">&nbsp;<span class="pcSmallText"><%=k%></span>&nbsp;<%next%><%end if%>
							
								<input type="button" value="<%=dictLanguage.Item(Session("language")&"_prv_6")%>" onClick="javascript: openbrowser('prv_postreview.asp?IDPRoduct=<%=pcv_IDProduct%>&rate=' + rateInput.val());" class="submit2">
						<%end if%>
						</form>
					</div>
						
					<hr>

					<%pcv_CShow=15
					iPageCurrent=request("page")
					if (iPageCurrent="") then
						iPageCurrent=1
					end if
					if not validNum(iPageCurrent) then
						iPageCurrent=1
					end if
					%>
					<!--#include file="prv_incshow.asp"-->
					<% If iPageCount>1 then
					iPageCurrent=clng(iPageCurrent) %>
					<!-- If Page count is more then 1 show page navigation -->
							<div class="pcReviewPagination"><b> 
								<% If iPageCurrent > 1 Then %>
									<a href="prv_allreviews.asp?IDProduct=<%=pcv_IDProduct%>&IDCategory=<%=pcv_IDCategory%>&page=<%=iPageCurrent -1 %>"><img src="<%=pcf_getImagePath("",rsIconObj("previousicon"))%>"></a> 
								<% End If
								For I=1 To iPageCount
									If I=iPageCurrent Then %>
										<%= I %> 
									<% Else %>
										<a class=privacy href="prv_allreviews.asp?IDProduct=<%=pcv_IDProduct%>&IDCategory=<%=pcv_IDCategory%>&page=<%=I%>"><%=I%></a> 
									<% End If 
								Next
								
								If iPageCurrent < iPageCount Then %>
									<a href="prv_allreviews.asp?IDProduct=<%=pcv_IDProduct%>&IDCategory=<%=pcv_IDCategory%>&page=<%=iPageCurrent + 1%>"><img src="<%=pcf_getImagePath("",rsIconObj("nexticon"))%>"></a> 
								<% End If %>
								</b>
							</div>
					<!-- end of page navigation -->
					<% end if%>
					
					<%if RCount>="5" then%>
						<div class="pcReviewActions">
							<a href="javascript:openbrowser('prv_postreview.asp?IDPRoduct=<%=pcv_IDProduct%>');"><%=dictLanguage.Item(Session("language")&"_prv_4")%></a>
						</div>

						<hr>

						<div class="pcReviewRate">
							<form name="rating2" id="rating2" class="pcForms">
								<%if pcv_RatingType="0" then%>
									<b><%=dictLanguage.Item(Session("language")&"_prv_5")%></b> 
					
									<input name="feel" type="hidden" value="">
					
									<img class="pcReviewFeelIcon" src="<%=pcf_getImagePath("catalog",pcv_Img1)%>" alt="<%= pcv_MainRateTxt2 %>">&nbsp;<input name="feel1" value="2" type="radio" onClick="feelInput2.val('2');" class="clearBorder">&nbsp;<%=pcv_MainRateTxt2%>  
									<img class="pcReviewFeelIcon" src="<%=pcf_getImagePath("catalog",pcv_Img2)%>" alt="<%= pcv_MainRateTxt3 %>">&nbsp;<input name="feel1" value="1" type="radio" onClick="feelInput2.val('1');" class="clearBorder">&nbsp;<%=pcv_MainRateTxt3%>
					
									<input type="button" value="<%=dictLanguage.Item(Session("language")&"_prv_6")%>" onClick="	javascript: openbrowser('prv_postreview.asp?IDPRoduct=<%=pcv_IDProduct%>&feel=' + feelInput2.val());" class="submit2">
								<%else%>
									<b><%=dictLanguage.Item(Session("language")&"_prv_5")%></b>
					
									<input name="rate" type="hidden" value="">
					
									<%if pcv_CalMain="1" then%>
										<%for k=1 to pcv_MaxRating%>
											<input name="rate1" value="<%=k%>" type="radio" onClick="rateInput2.val('<%=k%>');" class="clearBorder">&nbsp;<span class="pcSmallText"><%=k%></span>&nbsp;
										<%next%>
									<%end if%>

									<input type="button" value="<%=dictLanguage.Item(Session("language")&"_prv_6")%>" onClick="	javascript: openbrowser('prv_postreview.asp?IDPRoduct=<%=pcv_IDProduct%>&rate=' + rateInput2.val());" class="submit2">
								<%end if%>
							</form>
						</div>
					<%end if%>

					<div class="pcSpacer"></div>
					<div class="pcReviewActions">
						<a href="viewPrd.asp?IDProduct=<%=pcv_IDProduct%>&IDCategory=<%=pcv_IDCategory%>"><%=dictLanguage.Item(Session("language")&"_prv_30")%></a>
					</div>

			</div>
		</div>
	</div>
</div>
	<%
	END IF
END IF%>
<!--#include file="footer_wrapper.asp"-->
