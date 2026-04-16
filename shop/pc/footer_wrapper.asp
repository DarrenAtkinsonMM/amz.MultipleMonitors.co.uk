<% If session("Facebook")="1" Then %>
	
	<%
    server.Execute(pcv_theme & "/footer_facebook.asp")
    %>

<% ElseIf session("Mobile")="1" Then %>

    <%
    server.Execute(pcv_theme & "/footer_mobile.asp")
    %>

<% Else %>

    <%
    'server.Execute(pcv_theme & "/footer.asp")
    %>
    
<% End If %>


<!-- DA - EDIT -->
<footer>	
		<div class="container">
			<div class="row">
				<div class="col-sm-12 col-md-5 mobi-first-row">
					<div class="row">
						<div class="col-sm-6 fcol-custom">
							<div class="wow fadeInDown" data-wow-delay="0.1s">
								<div class="widget">
									<h5 class="text-white">Policies &amp; Legal</h5>
									<div class="footer-content">
										<ul class="footer-list">
											<li><a href="/pages/delivery/">Delivery Information</a></li>
											<li><a href="/pages/international/">International Orders</a></li>
											<li><a href="/pages/warranty/">Warranty Information</a></li>
											<li><a href="/pages/returns/">Returns Policy</a></li>
											<li><a href="/pages/privacy-policy/">Privacy Policy</a></li>
											<li><a href="/pages/terms/">Terms &amp; Conditions</a></li>
										</ul>
									</div>
								</div>
							</div>
						</div>
						<div class="col-sm-6 fcol-custom">
							<div class="wow fadeInDown" data-wow-delay="0.1s">
								<div class="widget">
									<h5 class="text-white">Recently Viewed</h5>
									<div class="footer-content">
										<ul class="footer-list">
											<!--#include file="smallRecentProducts.asp"-->
										</ul>
									</div>
								</div>
							</div>
						</div>
					</div>
				</div>
				<div class="col-sm-12 col-md-7">
					<div class="row">
						<div class="col-sm-7 fcol-custom">
							<div class="wow fadeInDown" data-wow-delay="0.1s">
								<div class="widget">
									<h5 class="text-white h-semi">Get In touch</h5>
									<div class="footer-content footer-git">
										<h4 class="h-bold font-light">
											<a href="tel:03302236655" class="hlink-contact hlink-phone text-white"><i class="fa fa-mobile-phone"></i> 0330 223 66 55</a>
										</h4>
										<h6 class="h-semi font-light footer-mail">
											<a href="mailto:sales@multiplemonitors.co.uk" class="hlink-contact hlink-mail text-white"><i class="fa  fa-envelope"></i> sales@multiplemonitors.co.uk</a>
										</h6>
									</div>
								</div>
							</div>
						</div>
						<div class="col-sm-5 fcol-custom">
							<div class="wow fadeInDown" data-wow-delay="0.1s">
								<div class="widget">
									<h5 class="text-white h-semi">FREE BUYERS GUIDE</h5>
									<div class="footer-content footer-subscribe">
										<p>Learn all about  multi-screen PC components with our FREE buyers guide.</p>
										<a href="https://www.getdrip.com/forms/206455195/submissions/new" data-drip-show-form="206455195" class="btn footer-submit bg-skin transition-nm medium manual-optin-trigger">Get It Now</a>
									</div>
								</div>
							</div>
						</div>
					</div>
				</div>
			</div>	
		</div>
		<div class="sub-footer">
		<div class="container">
			<div class="row">
				<div class="col-sm-5">
					<div class="wow fadeInLeft" data-wow-delay="0.1s">
					<div class="text-left">
					<p>&copy; Copyright <%= Year(Now()) %> - MultipleMonitors Ltd. All rights reserved.</p>
					</div>
					</div>
				</div>
				<div class="col-sm-7">
					<div class="wow fadeInRight" data-wow-delay="0.1s">
					<div class="text-right">
						<ul class="footer-base-menu">
							<li><a href="/">Home</a></li>
							<li><a href="/blog/">Blog</a></li>
							<li><a href="/shop/pc/checkout.asp?cmode=1">Support</a></li>
							<li><a href="/pages/about-us/">About Us</a></li>
							<li><a href="/pages/contact-us/">Contact Us</a></li>
							<li><a href="/pages/testimonials/">Testimonials</a></li>
							<li><a href="/pages/site-map/">Site Map</a></li>
						</ul>
						</ul>
					</div>
					</div>
				</div>
			</div>	
		</div>
		</div>
	</footer>

</div>
<!--#include file="inc_footer.asp" -->
<%
call closeDB()
%>
<a href="#" class="scrollup"><i class="fa fa-angle-up active"></i></a>
	<!-- Core JavaScript Files -->
<%
'if not pcStrPageName = "OnePageCheckout.asp" then]
if LCase(Request.ServerVariables("SCRIPT_NAME")) = "/shop/pc/checkout.asp" OR pcStrPageName = "onepagecheckout.asp" then
%>
    
    <script src="/js/jquery.easing.min.js"></script>
	<script src="/js/wow.min.js"></script>
	<script src="/js/jquery.scrollTo.js"></script>
	<script src="/js/ekko-lightbox.min.js"></script>
    <script src="/js/custom-opc.js"></script>
<%
else
%>
    
    <script src="/js/jquery.easing.min.js"></script>
	<script src="/js/wow.min.js"></script>
	<script src="/js/jquery.scrollTo.js"></script>
	<script src="/js/ekko-lightbox.min.js"></script>
    <script src="/js/custom.js"></script>
<%
end if
%>

<!-- Drip -->
<script type="text/javascript">
  var _dcq = _dcq || [];
  var _dcs = _dcs || {};
  _dcs.account = '1043541';

  (function() {
    var dc = document.createElement('script');
    dc.type = 'text/javascript'; dc.async = true;
    dc.src = '//tag.getdrip.com/1043541.js';
    var s = document.getElementsByTagName('script')[0];
    s.parentNode.insertBefore(dc, s);
  })();
</script>
<!-- end Drip -->
<script src="/js/cookie.js"></script>
</body>
</html>