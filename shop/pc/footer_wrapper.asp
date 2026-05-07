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
	<div class="mm-container">
		<div class="foot-grid">
			<div class="foot-brand">
				<div class="logo-w">
					<img src="/images/mm-logo-trans-w.png" alt="Multiple Monitors Ltd" style="height:40px; filter:brightness(0.8);">
				</div>
				<p>UK multi-screen specialists since 2008. Trading computers, monitor arrays, Synergy stands &amp; complete bundles &mdash; custom built, benchmark tested, supported for life.</p>
				<div class="contacts">
					<div><i class="fa fa-phone"></i><a href="tel:03302236655">0330 223 66 55</a></div>
					<div><i class="fa fa-envelope-o"></i><a href="mailto:sales@multiplemonitors.co.uk">sales@multiplemonitors.co.uk</a></div>
					<div><i class="fa fa-map-marker"></i>United Kingdom</div>
				</div>
			</div>

			<div class="foot">
				<h6>Shop</h6>
				<ul>
					<li><a href="/pages/trading-computers/">Trading Computers</a></li>
					<li><a href="/computers/">Computers</a></li>
					<li><a href="/display-systems/">Monitor Arrays</a></li>
					<li><a href="/stands/">Synergy Stands</a></li>
					<li><a href="/bundles/">Bundles</a></li>
				</ul>
			</div>

			<div class="foot">
				<h6>About</h6>
				<ul>
					<li><a href="/pages/about-us/">About us</a></li>
					<li><a href="/blog/">Blog</a></li>
					<li><a href="https://traderspec.com">TraderSpec.com</a></li>
					<li><a href="/pages/testimonials/">Testimonials</a></li>
					<li><a href="/pages/contact-us/">Contact</a></li>
				</ul>
			</div>

			<div class="foot">
				<h6>Policies</h6>
				<ul>
					<li><a href="/pages/delivery/">Delivery</a></li>
					<li><a href="/pages/international/">International</a></li>
					<li><a href="/pages/warranty/">Warranty</a></li>
					<li><a href="/pages/returns/">Returns</a></li>
					<li><a href="/pages/privacy-policy/">Privacy</a></li>
					<li><a href="/pages/terms/">Terms</a></li>
				</ul>
			</div>
		</div>

		<div class="foot-bottom">
			<div>&copy; Copyright <%= Year(Now()) %> Multiple Monitors Ltd. All rights reserved.</div>
			<div class="legal">
				UK Company No. 6863410 &middot; VAT Reg. 971 7562 87
			</div>
		</div>
	</div>
</footer>

</div>
<!--#include file="inc_footer.asp" -->
<%
call closeDB()
%>
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

<!-- 2026 redesign: reveal-on-scroll for .mm-site .reveal elements (no-op on legacy pages) -->
<script>
(function(){
	var els = document.querySelectorAll('.reveal');
	if (!els.length) return;
	if (!('IntersectionObserver' in window)) {
		els.forEach(function(e){ e.classList.add('is-in'); });
		return;
	}
	var io = new IntersectionObserver(function(entries){
		entries.forEach(function(entry){
			if (entry.isIntersecting) {
				entry.target.classList.add('is-in');
				io.unobserve(entry.target);
			}
		});
	}, { threshold: 0.12, rootMargin: '0px 0px -40px 0px' });
	els.forEach(function(e){ io.observe(e); });
})();
</script>

<!--#include file="inc_callModal.asp" -->
</body>
</html>