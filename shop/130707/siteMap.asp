<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Control Panel Links / Site Map" %>
<% section="" %>
<%PmAdmin=0%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="AdminHeader.asp"-->

<table class="pcCPcontent" >
	<tr>
		<td>
			<div class="pcCPsiteMap">
				<!--#include file="pcv4_navigation_links.asp"-->
			</div>
			<script type="text/javascript">
				$pc(document).ready(function () {
					$pc(".pcCPsiteMap > ul > li").each(function () {
						var titleItem = $pc(this).children(".pcCPlinkSectionTitle");
						if (titleItem.length > 0) {
							$pc(this).addClass("panel panel-default");
							titleItem.replaceWith("<div class='panel-heading'><h3 class='panel-title'>" + titleItem.html() + "</h3></div>");
						}
					});


					$pc(".pcCPsiteMap > ul > li > ul").each(function () {
						var cnt = 0;
						var otherLinksCnt = 0;

						$(this).children("li").each(function () {
							if (!$pc(this).hasClass("pcCPlinkSection")) {
								var parent = $pc(this).parent();
								if (parent.find(".pcCPlinkMain").length < 1) {
									parent.prepend("<li class='pcCPlinkMain'><a class='pcCPlinkSectionTitle'>Main Links</a><ul></ul></li>");
									cnt++;
								}
								parent.find(".pcCPlinkMain ul").append("<li>" + $(this).html() + "</li>");
								$(this).remove();
							} else {
								otherLinksCnt++;
								if (cnt == 3) {
									$("<li style='clear: both'></li>").insertAfter(this);
									cnt = 0;
								} else {
									cnt++;
								}
							}
						});

						if (otherLinksCnt < 1) {
							$(this).find(".pcCPlinkMain .pcCPlinkSectionTitle").hide();
						}
					});
				});
			</script>
		</td>
	</tr>
</table>

<!--#include file="AdminFooter.asp"-->
