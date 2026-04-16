<%

dabfoffH2 = "Black Friday Deal:"
dabfoffH31 = "&nbsp;&nbsp;Free Memory (RAM) & SSD Upgrades On All PCs - <strong>Save Up To £160!</strong>"
dabfoffH32 = "Upgrades Automatically Applied // See Computer Pages For Full Details // Offer Ends Nov. 30th"
   
'Ultra  
if not InStr(pSku, "ULT1") = 0 Then
dabfoffH2 = "Black Friday Deal:"
 dabfoffH31 = "&nbsp;&nbsp;<span class=""cta-dbl"">Double RAM</span> & <span class=""cta-dbl"">Double SSD</span> Free Upgrades On This PC - <strong>Offer Worth £50!</strong>"
 dabfoffH32 = "<span class=""cta-dbl"">16GB RAM & 500GB SSD</span> Upgrades Automatically Applied // Our Best Ever Deal // Offer Ends Nov. 30th"
end if
 
 'Extreme  
if not InStr(pSku, "EXT1") = 0 Then
dabfoffH2 = "Black Friday Deal:"
 dabfoffH31 = "&nbsp;&nbsp;<span class=""cta-dbl"">Double RAM</span> & <span class=""cta-dbl"">Double SSD</span> Free Upgrades On This PC - <strong>Offer Worth £110!</strong>"
 dabfoffH32 = "<span class=""cta-dbl"">32GB DDR5 RAM & 1TB SSD</span> Upgrades Automatically Applied // Our Best Ever Deal // Offer Ends Nov. 30th"
end if
   
   'Trader  
if not InStr(pSku, "TRA1") = 0 Then
dabfoffH2 = "Black Friday Deal:"
 dabfoffH31 = "&nbsp;&nbsp;<span class=""cta-dbl"">Double RAM</span> & <span class=""cta-dbl"">Double SSD</span> Free Upgrades On This PC - <strong>Offer Worth £80!</strong>"
 dabfoffH32 = "<span class=""cta-dbl"">32GB RAM & 1TB SSD</span> Upgrades Automatically Applied // Our Best Ever Deal // Offer Ends Nov. 30th"
end if
	
	'Trader Pro  
if not InStr(pSku, "TRP1") = 0 Then
dabfoffH2 = "Black Friday Deal:"
 dabfoffH31 = "&nbsp;&nbsp;<span class=""cta-dbl"">Double RAM</span> & <span class=""cta-dbl"">Double SSD</span> Free Upgrades On This PC - <strong>Offer Worth £160!</strong>"
 dabfoffH32 = "<span class=""cta-dbl"">64GB DDR5 RAM & 2TB SSD</span> Upgrades Automatically Applied // Our Best Ever Deal // Offer Ends Nov. 30th"
end if
%>

<section id="callaction" class="ca-bf">	
           <div class="container">
				<div class="row">
					<div class="col-md-12">
						<div class="callaction">
							<div class="row">
								<div class="col-md-12">
									<div class="wow fadeInUp" data-wow-delay="0.1s">
									<div class="cta-text">
									<h2 class="h-bold font-light disp-inline"><%= dabfoffH2 %></h2>
									<h3 class="h-light font-light disp-inline"><%= dabfoffH31 %></h3>
                                    <h3 class="h-light font-light disp-inline"><%= dabfoffH32 %></h3>
									</div>
									</div>
								</div>
							</div>
						</div>
					</div>
				</div>
            </div>
	</section>