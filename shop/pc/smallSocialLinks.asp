<%
call openDb()
%>
<%
	
	socialImageDir = session("pcsTheme") & "/images/social"
	
	Set connTemp2=Server.CreateObject("ADODB.Connection")
	connTemp2.Open scDSN
	
	'// Get Social Links
	query = "SELECT pcSocialLink_Name, pcSocialLink_Image, pcSocialLink_CustomImage, pcSocialLink_Url, pcSocialLink_Alt FROM pcSocialLinks ORDER BY pcSocialLink_Order, pcSocialLink_Name"
	set rs=connTemp2.execute(query)
	if not rs.eof then
		pcSocialLinksArr = rs.GetRows()
		pcSocialLinksCnt = UBound(pcSocialLinksArr, 2) + 1
	else
		pcSocialLinksCnt = 0
	end if
	set rs = nothing
	
	set connTemp2 = nothing
%>

<ul id="pcSocialLinks" itemscope itemtype="http://schema.org/Organization">  
  <li class="pcSocialLinkContactUs">
    <link itemprop="url" href="<%=homepageurl%>">
    <a href="contact.asp" title="Contact Us">
      <img src="<%=pcf_getImagePath(socialImageDir,"email.png")%>" alt="Contact Us" />
    </a>
  </li>
  
  <% 
    if pcSocialLinksCnt > 0 then
      for i = 0 to pcSocialLinksCnt - 1			
        slName = pcSocialLinksArr(0, i)
        slImage = pcSocialLinksArr(1, i)
        slCustomImage = pcSocialLinksArr(2, i)
        slUrl = pcSocialLinksArr(3, i)
        slAlt = pcSocialLinksArr(4, i)
				
				if len(slCustomImage) > 0 then
					slImage = slCustomImage
					slImageDir = "catalog"
				else
					slImageDir = socialImageDir
				end if
        
        if len(slUrl) > 0 then
        %>
          <li class="pcSocialLink<%= Replace(slName, " ", "") %>">
            <a itemprop="sameAs" href="<%= slUrl %>" title="<%= slAlt %>" target="_blank">
              <img src="<%=pcf_getImagePath(slImageDir,slImage)%>" alt="<%= slName %>"/>
            </a>
          </li>
        <%
        end if
      next	
    end if 
  %>
</ul>
<%
call closeDb()
%>