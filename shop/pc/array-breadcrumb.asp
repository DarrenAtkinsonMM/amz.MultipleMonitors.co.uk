<%
'DA Edit for Bundle Breadcrumb

'Check which page we are on
Select Case Request.ServerVariables("PATH_INFO")
	Case "/shop/pc/CUSTOMCAT-arrays1.asp"
		bunBCpage = "array1"
	Case "/shop/pc/CUSTOMCAT-arrays2.asp"
		bunBCpage = "array2"
	Case "/shop/pc/CUSTOMCAT-arrays3.asp"
		bunBCpage = "array3"
End Select

'fix missing querystring's - 2nd edit to check for SQL injections
if Request.querystring("mid") = "" Then
	bunBCmid = 0
else
	'Querystring has something check if a number, if not dump user to 404 page
	If IsNumeric(Request.querystring("mid")) Then
		bunBCmid = CInt(Request.querystring("mid"))
	Else
		response.redirect("/404.html")
	End If
end if

if Request.querystring("sid") = "" Then
	bunBCsid = 0
else
	'Querystring has something check if a number, if not dump user to 404 page
	If IsNumeric(Request.querystring("sid")) Then
		bunBCsid = CInt(Request.querystring("sid"))
	Else
		response.redirect("/404.html")
	End If
end if


'Work out correct stand
Select Case bunBCsid
	Case 326
		bunStandimgThm = "/images/bundles/bun-s2v-thm.jpg"
		bunStandimg = "/images/bundles/bun-s2v-med.png"
		bunStandimgLG = "/images/bundles/bun-s2v-lg.jpg"
		bunStandimgXlg = "/shop/pc/catalog/2v-front-angle-lg.jpg"
		bunStandimg2 = "<a href=""/images/bundles/syn-central-vesa-lg.jpg"" data-zoom=""/shop/pc/catalog/syn-central-vesa-lg.jpg""><img class=""arr-thumb"" src=""/images/bundles/syn-central-vesa-thm.jpg"" /></a>"
		bunStandimg3 = "<a href=""/images/bundles/syn-vesa-rotation-lg.jpg"" data-zoom=""/shop/pc/catalog/syn-vesa-rotation-lg.jpg""><img class=""arr-thumb"" src=""/images/bundles/syn-vesa-rotation-thm.jpg"" /></a>"
		bunStandTxt = "Dual Vertical"
		bunStandTxtShort = "Dual"
		bunStandOk = 1
		bunMonNum = "2"
		bunBCdiscount = 25
		bunArrStdImg = "s2v"
	Case 287
		bunStandimgThm = "/images/bundles/bun-s2h-thm.jpg"
		bunStandimg = "/images/bundles/bun-s2h-med.png"
		bunStandimgLG = "/images/bundles/bun-s2h-lg.jpg"
		bunStandimgXlg = "/shop/pc/catalog/2h-front-angle-lg.jpg"
		bunStandimg2 = "<a href=""/images/bundles/syn-arm-pivot-lg.jpg"" data-zoom=""/shop/pc/catalog/syn-arm-pivot-lg.jpg""><img class=""arr-thumb"" src=""/images/bundles/syn-arm-pivot-thm.jpg"" /></a>"
		bunStandimg3 = "<a href=""/images/bundles/syn-arm-vesa-height-lg.jpg"" data-zoom=""/shop/pc/catalog/syn-arm-vesa-height-lg.jpg""><img class=""arr-thumb"" src=""/images/bundles/syn-arm-vesa-height-thm.jpg"" /></a>"
		bunStandTxt = "Dual Horizontal"
		bunStandTxtShort = "Dual"
		bunStandOk = 1
		bunMonNum = "2"
		bunBCdiscount = 25
		bunArrStdImg = "s2h"
	Case 324
		bunStandimgThm = "/images/bundles/bun-s3p-thm.jpg"
		bunStandimg = "/images/bundles/bun-s3p-med.png"
		bunStandimgLG = "/images/bundles/bun-s3p-lg.jpg"
		bunStandimgXlg = "/shop/pc/catalog/3p-front-angle-lg.jpg"
		bunStandimg2 = "<a href=""/images/bundles/syn-arm-pivot-lg.jpg"" data-zoom=""/shop/pc/catalog/syn-arm-pivot-lg.jpg""><img class=""arr-thumb"" src=""/images/bundles/syn-arm-pivot-thm.jpg"" /></a>"
		bunStandimg3 = "<a href=""/images/bundles/syn-arm-vesa-height-lg.jpg"" data-zoom=""/shop/pc/catalog/syn-arm-vesa-height-lg.jpg""><img class=""arr-thumb"" src=""/images/bundles/syn-arm-vesa-height-thm.jpg"" /></a>"
		bunStandTxt = "Triple Pyramid"
		bunStandTxtShort = "Triple"
		bunStandOk = 1
		bunMonNum = "3"
		bunBCdiscount = 25
		bunArrStdImg = "s3p"
	Case 312
		bunStandimgThm = "/images/bundles/bun-s3h-thm.jpg"
		bunStandimg = "/images/bundles/bun-s3h-med.png"
		bunStandimgLG = "/images/bundles/bun-s3h-lg.jpg"
		bunStandimgXlg = "/shop/pc/catalog/3h-front-angle-lg.jpg"
		bunStandimg2 = "<a href=""/images/bundles/syn-arm-pivot-lg.jpg"" data-zoom=""/shop/pc/catalog/syn-arm-pivot-lg.jpg""><img class=""arr-thumb"" src=""/images/bundles/syn-arm-pivot-thm.jpg"" /></a>"
		bunStandimg3 = "<a href=""/images/bundles/syn-arm-vesa-height-lg.jpg"" data-zoom=""/shop/pc/catalog/syn-arm-vesa-height-lg.jpg""><img class=""arr-thumb"" src=""/images/bundles/syn-arm-vesa-height-thm.jpg"" /></a>"
		bunStandTxt = "Triple Horizontal"
		bunStandTxtShort = "Triple"
		bunStandOk = 1
		bunMonNum = "3"
		bunBCdiscount = 25
		bunArrStdImg = "s3h"
	Case 313
		bunStandimgThm = "/images/bundles/bun-s4s-thm.jpg"
		bunStandimg = "/images/bundles/bun-s4s-med.png"
		bunStandimgLG = "/images/bundles/bun-s4s-lg.jpg"
		bunStandimgXlg = "/shop/pc/catalog/4s-front-angle-lg.jpg"
		bunStandimg2 = "<a href=""/images/bundles/syn-arm-pivot-lg.jpg"" data-zoom=""/shop/pc/catalog/syn-arm-pivot-lg.jpg""><img class=""arr-thumb"" src=""/images/bundles/syn-arm-pivot-thm.jpg"" /></a>"
		bunStandimg3 = "<a href=""/images/bundles/syn-arm-vesa-height-lg.jpg"" data-zoom=""/shop/pc/catalog/syn-arm-vesa-height-lg.jpg""><img class=""arr-thumb"" src=""/images/bundles/syn-arm-vesa-height-thm.jpg"" /></a>"
		bunStandTxt = "Quad Square"
		bunStandTxtShort = "Quad"
		bunStandOk = 1
		bunMonNum = "4"
		bunBCdiscount = 50
		bunArrStdImg = "s4s"
	Case 337
		bunStandimgThm = "/images/bundles/bun-s4sp-thm.jpg"
		bunStandimg = "/images/bundles/bun-s4sp-med.png"
		bunStandimgLG = "/images/bundles/bun-s4sp-lg.jpg"
		bunStandimgXlg = "/shop/pc/catalog/4sp-front-angle-lg.jpg"
		bunStandimg2 = "<a href=""/images/bundles/syn-central-vesa-lg.jpg"" data-zoom=""/shop/pc/catalog/syn-central-vesa-lg.jpg""><img class=""arr-thumb"" src=""/images/bundles/syn-central-vesa-thm.jpg"" /></a>"
		bunStandimg3 = "<a href=""/images/bundles/syn-vesa-rotation-lg.jpg"" data-zoom=""/shop/pc/catalog/syn-vesa-rotation-lg.jpg""><img class=""arr-thumb"" src=""/images/bundles/syn-vesa-rotation-thm.jpg"" /></a>"
		bunStandTxt = "Quad Square"
		bunStandTxtShort = "Quad"
		bunStandOk = 1
		bunMonNum = "4"
		bunBCdiscount = 50
		bunArrStdImg = "s4sp"
	Case 327
		bunStandimgThm = "/images/bundles/bun-s4h-thm.jpg"
		bunStandimg = "/images/bundles/bun-s4h-med.png"
		bunStandimgLG = "/images/bundles/bun-s4h-lg.jpg"
		bunStandimgXlg = "/shop/pc/catalog/4h-front-angle-lg.jpg"
		bunStandTxt = "Quad Horizontal"
		bunStandTxtShort = "Quad"
		bunStandOk = 1
		bunMonNum = "4"
		bunBCdiscount = 50
		bunArrStdImg = "s4h"
	Case 325
		bunStandimgThm = "/images/bundles/bun-s4p-thm.jpg"
		bunStandimg = "/images/bundles/bun-s4p-med.png"
		bunStandimgLG = "/images/bundles/bun-s4p-lg.jpg"
		bunStandimgXlg = "/shop/pc/catalog/4p-front-angle-lg.jpg"
		bunStandimg2 = "<a href=""/images/bundles/syn-arm-pivot-lg.jpg"" data-zoom=""/shop/pc/catalog/syn-arm-pivot-lg.jpg""><img class=""arr-thumb"" src=""/images/bundles/syn-arm-pivot-thm.jpg"" /></a>"
		bunStandimg3 = "<a href=""/images/bundles/syn-arm-vesa-height-lg.jpg"" data-zoom=""/shop/pc/catalog/syn-arm-vesa-height-lg.jpg""><img class=""arr-thumb"" src=""/images/bundles/syn-arm-vesa-height-thm.jpg"" /></a>"
		bunStandTxt = "Quad Pyramid"
		bunStandTxtShort = "Quad"
		bunStandOk = 1
		bunMonNum = "4"
		bunBCdiscount = 50
		bunArrStdImg = "s4p"
	Case 318
		bunStandimgThm = "/images/bundles/bun-s5p-thm.jpg"
		bunStandimg = "/images/bundles/bun-s5p-med.png"
		bunStandimgLG = "/images/bundles/bun-s5p-lg.jpg"
		bunStandimgXlg = "/shop/pc/catalog/5p-front-angle-lg.jpg"
		bunStandimg2 = "<a href=""/images/bundles/syn-arm-pivot-lg.jpg"" data-zoom=""/shop/pc/catalog/syn-arm-pivot-lg.jpg""><img class=""arr-thumb"" src=""/images/bundles/syn-arm-pivot-thm.jpg"" /></a>"
		bunStandimg3 = "<a href=""/images/bundles/syn-arm-vesa-height-lg.jpg"" data-zoom=""/shop/pc/catalog/syn-arm-vesa-height-lg.jpg""><img class=""arr-thumb"" src=""/images/bundles/syn-arm-vesa-height-thm.jpg"" /></a>"
		bunStandTxt = "Five Pyramid"
		bunStandTxtShort = "Five"
		bunStandOk = 1
		bunMonNum = "5"
		bunBCdiscount = 50
		bunArrStdImg = "s5p"
	Case 338
		bunStandimgThm = "/images/bundles/bun-s6r-thm.jpg"
		bunStandimg = "/images/bundles/bun-s6r-med.png"
		bunStandimgLG = "/images/bundles/bun-s6r-lg.jpg"
		bunStandimgXlg = "/shop/pc/catalog/6r-front-angle-lg.jpg"
		bunStandimg2 = "<a href=""/images/bundles/syn-arm-pivot-lg.jpg"" data-zoom=""/shop/pc/catalog/syn-arm-pivot-lg.jpg""><img class=""arr-thumb"" src=""/images/bundles/syn-arm-pivot-thm.jpg"" /></a>"
		bunStandimg3 = "<a href=""/images/bundles/syn-arm-vesa-height-lg.jpg"" data-zoom=""/shop/pc/catalog/syn-arm-vesa-height-lg.jpg""><img class=""arr-thumb"" src=""/images/bundles/syn-arm-vesa-height-thm.jpg"" /></a>"
		bunStandTxt = "Six Way"
		bunStandTxtShort = "Six"
		bunStandOk = 1
		bunMonNum = "6"
		bunBCdiscount = 100
		bunArrStdImg = "s6r"
	Case 314
		bunStandimgThm = "/images/bundles/bun-s6rp-thm.jpg"
		bunStandimg = "/images/bundles/bun-s6rp-med.png"
		bunStandimgLG = "/images/bundles/bun-s6rp-lg.jpg"
		bunStandimgXlg = "/shop/pc/catalog/6rp-front-angle-lg.jpg"
		bunStandimg2 = "<a href=""/images/bundles/syn-central-vesa-lg.jpg"" data-zoom=""/shop/pc/catalog/syn-central-vesa-lg.jpg""><img class=""arr-thumb"" src=""/images/bundles/syn-central-vesa-thm.jpg"" /></a>"
		bunStandimg3 = "<a href=""/images/bundles/syn-vesa-rotation-lg.jpg"" data-zoom=""/shop/pc/catalog/syn-vesa-rotation-lg.jpg""><img class=""arr-thumb"" src=""/images/bundles/syn-vesa-rotation-thm.jpg"" /></a>"
		bunStandTxt = "Six Way"
		bunStandTxtShort = "Six"
		bunStandOk = 1
		bunMonNum = "6"
		bunBCdiscount = 100
		bunArrStdImg = "s6rp"
	Case 319
		bunStandimgThm = "/images/bundles/bun-s8r-thm.jpg"
		bunStandimg = "/images/bundles/bun-s8r-med.png"
		bunStandimgLG = "/images/bundles/bun-s8r-lg.jpg"
		bunStandimgXlg = "/shop/pc/catalog/8r-front-angle-lg.jpg"
		bunStandimg2 = "<a href=""/images/bundles/syn-arm-pivot-lg.jpg"" data-zoom=""/shop/pc/catalog/syn-arm-pivot-lg.jpg""><img class=""arr-thumb"" src=""/images/bundles/syn-arm-pivot-thm.jpg"" /></a>"
		bunStandimg3 = "<a href=""/images/bundles/syn-arm-vesa-height-lg.jpg"" data-zoom=""/shop/pc/catalog/syn-arm-vesa-height-lg.jpg""><img class=""arr-thumb"" src=""/images/bundles/syn-arm-vesa-height-thm.jpg"" /></a>"
		bunStandTxt = "Eight Way"
		bunStandTxtShort = "Eight"
		bunStandOk = 1
		bunMonNum = "8"
		bunBCdiscount = 100
		bunArrStdImg = "s8r"
	Case else
		bunStandimg = "/images/bundles/bun-question.jpg"
		bunStandOk = 0
		bunBCdiscount = 0
End Select

'Change monitor number numeric to text
Select Case bunMonNum
	Case 2
		bunMonNumTxt = "Two"
	Case 3
		bunMonNumTxt = "Three"
	Case 4
		bunMonNumTxt = "Four"
	Case 5
		bunMonNumTxt = "Five"
	Case 6
		bunMonNumTxt = "Six"
	Case 8
		bunMonNumTxt = "Eight"
End Select

'Work out correct monitor
Select Case bunBCmid
	Case 315
		bunMonitorimgThm = "/images/bundles/bun-acersq-thm.jpg"
		bunMonitorimg = "/images/bundles/bun-acersq-med.png"
		bunMonitorimgLG = "/images/bundles/bun-acersq-lg.jpg"
		bunMonitorimgXlg = "/shop/pc/catalog/acer17_detail.jpg"
		bunMonitorTxt = "Acer 17&quot;"
		bunMonitorTxt2 = "Monitors"
		bunMonitorBenefit = "Acer 17&quot; 1280 x 1024 monitors"
		bunArrMonImg = "a17"
		bunMonitorOk = 1
	Case 304
		bunMonitorimgThm = "/images/bundles/bun-acerwide-thm.jpg"
		bunMonitorimg = "/images/bundles/bun-acerwide-med.png"
		bunMonitorimgLG = "/images/bundles/bun-acerwide-lg.jpg"
		bunMonitorimgXlg = "/shop/pc/catalog/acer22_detail.jpg"
		bunMonitorTxt = "AOC 21.5&quot;"
		bunMonitorTxt2 = "Widescreens"
		bunMonitorBenefit = "AOC 21.5&quot; 1920 x 1080 (Full HD) widescreen monitors"
		bunMonitorOk = 1
		bunArrMonImg = "a22"
	Case 316
		bunMonitorimgThm = "/images/bundles/bun-acersq-thm.jpg"
		bunMonitorimg = "/images/bundles/bun-acersq-med.png"
		bunMonitorimgLG = "/images/bundles/bun-acersq-lg.jpg"
		bunMonitorimgXlg = "/shop/pc/catalog/acer17_detail.jpg"
		bunMonitorTxt = "Acer 19&quot;"
		bunMonitorTxt2 = "Monitors"
		bunMonitorBenefit = "Acer 19&quot; 1280 x 1024 monitors"
		bunArrMonImg = "a19"
		bunMonitorOk = 1
	Case 317
		bunMonitorimgThm = "/images/bundles/bun-acerwide-thm.jpg"
		bunMonitorimg = "/images/bundles/bun-acerwide-med.png"
		bunMonitorimgLG = "/images/bundles/bun-acerwide-lg.jpg"
		bunMonitorimgXlg = "/shop/pc/catalog/acer22_detail.jpg"
		bunMonitorTxt = "Acer 24&quot;"
		bunMonitorTxt2 = "Widescreens"
		bunMonitorBenefit = "Acer 24&quot; 1920 x 1080 (Full HD) widescreen monitors"
		bunMonitorOk = 1
		bunArrMonImg = "a24"
	Case 321
		bunMonitorimgThm = "/images/bundles/bun-iiyama-thm.jpg"
		bunMonitorimg = "/images/bundles/bun-iiyama-med.png"
		bunMonitorimgLG = "/images/bundles/bun-iiyama-lg.jpg"
		bunMonitorimgXlg = "/shop/pc/catalog/iiyama22ips-detail.jpg"
		bunMonitorTxt = "Iiyama 21.5&quot;"
		bunMonitorTxt2 = "Widescreens"
		bunMonitorBenefit = "Iiyama 21.5&quot; 1920 x 1080 (Full HD) IPS thin-bezel widescreens"
		bunArrMonImg = "i22"
		bunMonitorOk = 1
	Case 328
		bunMonitorimgThm = "/images/bundles/bun-acerwide-thm.jpg"
		bunMonitorimg = "/images/bundles/bun-acerwide-med.png"
		bunMonitorimgLG = "/images/bundles/bun-acerwide-lg.jpg"
		bunMonitorimgXlg = "/shop/pc/catalog/acer22_detail.jpg"
		bunMonitorTxt = "Acer 27&quot;"
		bunMonitorTxt2 = "Widescreens"
		bunMonitorBenefit = "Acer 27&quot; 1920 x 1080 (Full HD) widescreen monitors"
		bunArrMonImg = "a27"
		bunMonitorOk = 1
	Case 320
		bunMonitorimgThm = "/images/bundles/bun-iiyama-thm.jpg"
		bunMonitorimg = "/images/bundles/bun-iiyama-med.png"
		bunMonitorimgLG = "/images/bundles/bun-iiyama-lg.jpg"
		bunMonitorimgXlg = "/shop/pc/catalog/iiyama22ips-detail.jpg"
		bunMonitorTxt = "Iiyama 24&quot;"
		bunMonitorTxt2 = "Widescreens"
		bunMonitorBenefit = "Iiyama 24&quot; 1920 x 1080 (Full HD) IPS thin-bezel widescreens"
		bunMonitorOk = 1
		bunArrMonImg = "i23"
	Case 329
		bunMonitorimgThm = "/images/bundles/bun-iiyama-thm.jpg"
		bunMonitorimg = "/images/bundles/bun-iiyama-med.png"
		bunMonitorimgLG = "/images/bundles/bun-iiyama-lg.jpg"
		bunMonitorimgXlg = "/shop/pc/catalog/iiyama22ips-detail.jpg"
		bunMonitorTxt = "Iiyama 27&quot;"
		bunMonitorTxt2 = "Widescreens"
		bunMonitorBenefit = "Iiyama 27&quot; 1920 x 1080 (Full HD) IPS thin-bezel widescreens"
		bunMonitorOk = 1
		bunArrMonImg = "i27"
	Case 342
		bunMonitorimgThm = "/images/bundles/bun-viewsonic-thm.jpg"
		bunMonitorimg = "/images/bundles/bun-viewsonic-med.png"
		bunMonitorimgLG = "/images/bundles/bun-viewsonic-lg.jpg"
		bunMonitorimgXlg = "/shop/pc/catalog/vs-27-qhd_detail.jpg.jpg"
		bunMonitorTxt = "ViewSonic 27&quot;"
		bunMonitorTxt2 = "Widescreens"
		bunMonitorBenefit = "ViewSonic 27&quot; 2560 x 1440 (Quad HD) IPS thin-bezel widescreens"
		bunMonitorOk = 1
		bunArrMonImg = "i27"
	Case 344
		bunMonitorimgThm = "/images/bundles/bun-aoc-thm.jpg"
		bunMonitorimg = "/images/bundles/bun-aoc-med.png"
		bunMonitorimgLG = "/images/bundles/bun-aoc-lg.jpg"
		bunMonitorimgXlg = "/shop/pc/catalog/vs-27-qhd_detail.jpg.jpg"
		bunMonitorTxt = "AOC 27&quot;"
		bunMonitorTxt2 = "Widescreens"
		bunMonitorBenefit = "AOC 27&quot; 2560 x 1440 (Quad HD) IPS thin-bezel widescreens"
		bunMonitorOk = 1
		bunArrMonImg = "i27"
	Case 345
		bunMonitorimgThm = "/images/bundles/bun-iiyama-thm.jpg"
		bunMonitorimg = "/images/bundles/bun-iiyama-med.png"
		bunMonitorimgLG = "/images/bundles/bun-iiyama-lg.jpg"
		bunMonitorimgXlg = "/shop/pc/catalog/iiyama22ips-detail.jpg"
		bunMonitorTxt = "Iiyama 27&quot;"
		bunMonitorTxt2 = "Widescreens"
		bunMonitorBenefit = "Iiyama 27&quot; 2560 x 1440 (Quad HD) IPS thin-bezel widescreens"
		bunMonitorOk = 1
		bunArrMonImg = "i27"
	Case else
		bunMonitorimg = "/images/bundles/bun-question.jpg"
		bunMonitorOk = 0
End Select

'Pull db pricing data
if bunBCpage = "array3" then
'dim rstemp, conntemp, query, rstempd2, conntempd2, queryd2
call openDb()

'open product database and load recordset

dim queryBun, rsBun

queryBun="select idProduct, Price from products"
set rsBun=server.CreateObject("ADODB.RecordSet")
set rsBun=conntemp.execute(queryBun)
rsBun.MoveFirst

while not rsBun.eof

	if rsBun("idProduct") = bunBCsid then
		bunStandPrice = rsBun("Price")
	end if
	
	if rsBun("idProduct") = bunBCmid then
		bunMonitorPrice = rsBun("Price")
	end if

rsBun.MoveNext

wend

set rsBun=nothing

bunArrayPrice = bunStandPrice + (bunMonNum * bunMonitorPrice)

end if

if bunBCpage = "array1" then
	if bunBCmid = 0 then
	'We have no stand or monitor so show bundle start toolbar
%>
	<!-- Header: intro -->
    <header id="bundle-stands" class="bundle-wrap bg-lyt">
		<div class="intro-content paddingtop-20">
			<div class="container">
				<div class="row">
					<div class="col-md-12">
						<div class="wow fadeInDown text-center" data-wow-offset="0" data-wow-delay="0">
							<h1 class="h-bold bundle-heading text-uppercase color-focus"><i class="fa fa-cogs green-link"></i> <span>Create Your Own</span> Monitor Array</h1>
							<h5 class="text-uppercase color-med h-semi bundle-sub bundle-sub1">Select any Stand and Screen Combination to make your perfect Multi-Monitor Display Array</h5>
						</div>
						<div class="wow fadeInUp bd-wrap" data-wow-offset="0" data-wow-delay="0">
							<div class="row">
								<div class="col-md-3 bd-heading">
									<h2 class="color-med h-bold text-uppercase">Array <span class="color-focus">Bonus</span></h2>
								</div>
								<div class="col-md-9 bd-benifits">
									<div class="row">
										<div class="col-md-7 bd-benifit bd-freeCable">
											<i class="fa fa-code-fork color-focus"></i>
											<h4 class="f-light color-med h-semi">Free Long Length Cables Worth &pound;15 Per Screen</h4>
										</div>
										<div class="col-md-5 bd-benifit bd-freeDelivery">
											<i class="fa fa-truck color-focus"></i>
											<h4 class="f-light color-med h-semi">Reduced Delivery Fee</h4>
										</div>
									</div>
								</div>
							</div>
						</div>
						<div class="wow fadeInUp text-center" data-wow-delay="0">
							<h3 class="h-semi text-uppercase hd-message hm1 color-focus">To Get Started Simply Select A Stand For Your Monitor Array</h3><a name="arraystart"></a>
						</div>
					</div>				
				</div>		
			</div>
		</div>	
		<a href="#arraystart" id="wg-scrollDown">&#xf107;</a>
    </header>
	
	<!-- /Header: Bundle -->
<%
	else
	'Means we have a monitor but no stand
%>
	<!-- Header: intro -->
    <header id="bundle-stands" class="bundle-wrap bg-lyt">
		<div class="intro-content paddingtop-20"><a name="arraycustom"></a>
			<div class="container">
				<div class="row">
					<div class="col-md-12">
						<div class="wow fadeInDown text-center" data-wow-offset="0" data-wow-delay="0">
							<h1 class="h-bold bundle-heading text-uppercase color-focus"><i class="fa fa-cogs green-link"></i> <span>Create Your Own</span> Monitor Array</h1>
						</div>
						<div class="wow fadeInUp bd-productBox" data-wow-offset="0" data-wow-delay="0">
							<div class="bd-product bd-product-single">
								<div class="row">
									<div class="col-sm-3 bd-product-img">
										<img src="<%=bunMonitorimg%>">
									 </div>
									 <div class="col-sm-9 bd-product-text text-left">
										<h3 class="color-med h-bold">Monitors <span>Selected</span></h3>
										<h5 class="green-link h-bold"><%=bunMonitorTxt%> <span><%=bunMonitorTxt2%></span></h5>
										<a href="/display-systems/" class="btn btn-skin semi margintop-20">Restart Array Creation <i class="fa fa-angle-right"></i></a>
									 </div>
								</div>
							</div>
						</div>
						<div class="wow fadeInRight text-center" data-wow-delay="0">
							<h3 class="h-semi text-uppercase hd-message color-focus">Pick Your Preferred Stand, Select One Below</h3><a name="arraystart"></a>
						</div>
					</div>				
				</div>		
			</div>
		</div>	
		<a href="#arraystart" id="wg-scrollDown">&#xf107;</a>		
    </header>
	
	<!-- /Header: Bundle -->
<%
	end if
end if
if bunBCpage = "array2" then
%>
	<!-- Header: intro -->
    <header id="bundle-screens" class="bundle-wrap bg-lyt">
		<div class="intro-content">
			<div class="container">
				<div class="row">
					<div class="col-md-12">
						<div class="wow fadeInDown text-center" data-wow-offset="0" data-wow-delay="0">
							<h1 class="h-bold bundle-heading text-uppercase color-focus"><i class="fa fa-cogs green-link"></i> <span>Create Your Own</span> Monitor Array</h1>
						</div>
						<div class="wow fadeInUp bd-productBox" data-wow-offset="0" data-wow-delay="0">
							<div class="bd-product bd-product-single">
								<div class="row">
									<div class="col-sm-3 bd-product-img">
										<img src="<%=bunStandimg%>">
									 </div>
									 <div class="col-sm-9 bd-product-text text-left">
										<h3 class="color-med h-bold">Stand <span>Selected</span></h3>
										<h5 class="green-link h-bold"><%=bunStandTxt%> <span>Synergy Stand</span></h5>
										<a href="/display-systems/#arraystart" class="btn btn-skin semi margintop-20">Change Selection <i class="fa fa-angle-right"></i></a>
									 </div>
								</div>
							</div>
						</div>
						<div class="wow fadeInRight text-center" data-wow-delay="0">
							<h3 class="h-semi text-uppercase hd-message color-focus">Now Add Some Screens, Select Them Below</h3><a name="arraystart"></a>
						</div>
					</div>				
				</div>		
			</div>
		</div>	
		<a href="#arraystart" id="wg-scrollDown">&#xf107;</a>		
    </header>
	
	<!-- /Header: Bundle -->
<%
end if
if bunBCpage = "array3" then
%>
	<!-- Header: intro -->
    <header id="bundle-screens" class="bundle-wrap bg-lyt">
		<div class="intro-content">
			<div class="container">
				<div class="row">
					<div class="col-md-12">
						<div class="wow fadeInDown text-center" data-wow-offset="0" data-wow-delay="0">
							<h1 class="h-bold bundle-heading text-uppercase color-focus"><i class="fa fa-check-square-o green-link"></i> <%=bunStandTxtShort & " " & Replace(Replace(bunMonitorTxt,"Iiyama ",""),"Acer ","")%> Monitor Array</h1>
						</div>
						<div class="wow fadeInUp bd-ArrayBox" data-wow-offset="0" data-wow-delay="0">
							<div class="bd-product bd-product-array no-bg">
								<div class="row">
									<div class="col-sm-5 bd-product-img">
										<div class="productimage-box">
											<div id="product-zoom">
												<div class="pi-box">
													<div id="productbig-image" class="pi-boxfix">
														<a data-toggle="lightbox" href="/images/bundles/<%=bunArrStdImg%>-<%=bunArrMonImg%>-axl.jpg">
															<img id="img_01" src="/images/bundles/<%=bunArrStdImg%>-<%=bunArrMonImg%>-alg.jpg"> 
														</a>
													</div>
												</div>
												<div id="product-thumbs"> 
													<a class="act-thumb" href="/images/bundles/<%=bunArrStdImg%>-<%=bunArrMonImg%>-alg.jpg" data-zoom="/images/bundles/<%=bunArrStdImg%>-<%=bunArrMonImg%>-axl.jpg"><img class="arr-thumb" id="img_01" src="/images/bundles/<%=bunArrStdImg%>-<%=bunArrMonImg%>-atn.jpg" /></a>
													<a href="<%=bunStandimgLG%>" data-zoom="<%=bunStandimgXlg%>"><img class="arr-thumb" src="<%=bunStandimgThm%>" /></a>
                                                    <%=bunStandimg2%>
                                                    <%=bunStandimg3%>
													<a href="<%=bunMonitorimgLG%>" data-zoom="<%=bunMonitorimgXlg%>"><img class="arr-thumb" src="<%=bunMonitorimgThm%>" /></a>
                                                    <a href="/images/bundles/free-cables-lg.jpg" data-zoom="/images/bundles/free-cables-xlg.jpg"><img class="arr-thumb" src="/images/bundles/free-cables-thm.jpg" /></a>
												</div>
												<p class="text-center pz-info"><dfn>(Click to see larger image and other views)<dfn></p>
											</div>
										</div>
									 </div>
									 <div class="col-sm-7 bd-product-text text-left">
										<h5 class="green-link h-bold"><%=bunStandTxt%> <span>Synergy Stand</span></h5>
                                        <p class="benefit">A rock solid and ultra stable <%=bunStandTxt%> Synergy Stand which will hold your screens in perfect alignment.</p>
										<a href="/display-systems/?mid=<%=bunBCmid%>#arraystart" class="btn btn-skin semi margintop-10">Change Selection <i class="fa fa-angle-right"></i></a>
										<h5 class="green-link h-bold margintop-20"><%=bunMonNum%> X <%=bunMonitorTxt%> <span><%=bunMonitorTxt2%></span></h5>
                                        <p class="benefit"><%=bunMonNumTxt & " " & bunMonitorBenefit%> to layout your programs and data.</p>
										<a href="/display-systems-2/?sid=<%=bunBCsid%>" class="btn btn-skin semi margintop-10">Change Selection <i class="fa fa-angle-right"></i></a>
										<h5 class="green-link h-bold margintop-20"><%=bunMonNum%> X Extended Length <span>Video &amp; Power cables</span></h5>
                                        <p class="benefit"><span class="color-med"><strong>Free Gifts:</strong></span> 3 meter long high quality digital video cables and 3 meter power cables, a free set for each screen. <span class="color-focus"><strong>(Worth &pound;15 Per Set!)</strong></span></p>
                                        <h5 class="green-link h-bold margintop-20">Buy Your Monitor Array</h5>
                                        <p class="benefit">Order your new array and get the stand, screens, free cables and free (UK) delivery for this one inclusive price:</p>
										<div class="product-price">
											<label class="media-middle">Total Array Price:</label> <h3 class="price-info disp-inline h-semi color media-middle">&pound;<%=FormatNumber(bunArrayPrice/1.2)%> + VAT</h3><br /><span class="media-middle vat-info" style="margin-left:0px;">(&pound;<%=FormatNumber(bunArrayPrice)%> inc. VAT)</span>
										</div>
										<h6 class="h-semi green-link text-uppercase h-semi"> <strong>With Free Delivery <span class="color-focus">(Save &pound;10!)</span></strong></h6>
                                        <form method="post" action="/shop/pc/instPrd.asp" name="additem" >
                                        <input name="index" type="hidden" value="1">
        								<input type="hidden" name="idproduct1" value="<%=bunBCsid%>">
        								<input type="hidden" name="QtyM<%=bunBCsid%>" value="1">
        								<input type="hidden" name="idproduct2" value="<%=bunBCmid%>">
        								<input type="hidden" name="QtyM<%=bunBCmid%>" value="<%=bunMonNum%>">
       									<input type="hidden" name="pCnt" value="2">
										<input type="submit" value="ORDER YOUR NEW ARRAY >" class="btn btn-skin btn-wc semi order-btn margintop-30" />
                                        </form>
									 </div>
								</div>
							</div>
						</div>
					</div>				
				</div>		
			</div>
		</div>		
    </header>
	
	<!-- /Header: Bundle -->
<%
end if
%>