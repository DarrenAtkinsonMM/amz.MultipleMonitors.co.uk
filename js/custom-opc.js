/*global jQuery:false */
(function ($) {

	var wow = new WOW(
	  {
		boxClass:     'wow',      // animated element css class (default is wow)
		animateClass: 'animated', // animation css class (default is animated)
		offset:       0,          // distance to the element when triggering the animation (default is 0)
		mobile:       false       // trigger animations on mobile devices (true is default)
	  }
	);
	wow.init();
	
	
	$(".navbar-collapse").removeClass("in");
	
	//jQuery to collapse the navbar on scroll
	$(window).scroll(function() {
		if ($(".navbar").offset().top > 50) {
			$(".navbar-fixed-top").addClass("top-nav-collapse");
			$(".navbar-brand").addClass("reduce");
			$(".top-tagline-bar").addClass("collapse");

			$(".navbar-custom ul.nav ul.dropdown-menu").css("margin-top","11px");
		
		} else {
			$(".navbar-fixed-top").removeClass("top-nav-collapse");
			$(".navbar-brand").removeClass("reduce");
			$(".top-tagline-bar").removeClass("collapse");

			$(".navbar-custom ul.nav ul.dropdown-menu").css("margin-top","16px");
	
		}
	});
	jQuery('#hero-header').append('<a href="#wg-top" id="wg-toplink" class="scroll">&#xf107;</a>');
	$(".scroll").click(function(event) {
    event.preventDefault();
    $('html,body').animate( { scrollTop:$(this.hash).offset().top } , 1000);
    } );
	//scroll to top
	$(window).scroll(function(){
		if ($(this).scrollTop() > 100) {
			$('.scrollup').fadeIn();
			} else {
			$('.scrollup').fadeOut();
		}
	});
	$('.scrollup').click(function(){
		$("html, body").animate({ scrollTop: 0 }, 1000);
			return false;
	});
	


	//jQuery for page scrolling feature - requires jQuery Easing plugin
	$(function() {
		$('.navbar-nav li a').bind('click', function(event) {
			var $anchor = $(this);
			var nav = $($anchor.attr('href'));
			if (nav.length) {
			$('html, body').stop().animate({				
				scrollTop: $($anchor.attr('href')).offset().top				
			}, 1500, 'easeInOutExpo');
			
			event.preventDefault();
			}
		});
		$('.page-scroll a').bind('click', function(event) {
			var $anchor = $(this);
			$('html, body').stop().animate({
				scrollTop: $($anchor.attr('href')).offset().top
			}, 1500, 'easeInOutExpo');
			event.preventDefault();
		});
	});
	
	$(function(){
    $("#product-thumbs a").click(function(e){
        var href = $(this).attr("href");
        var zoom = $(this).attr("data-zoom");
        var headlin = $(this).attr("data-heading");
        $("#productbig-image a").attr("href", zoom);
        $("#productbig-image a").attr("data-title", headlin);
        $("#productbig-image img").attr("src", href);
		$(this).addClass('act-thumb').siblings().removeClass('act-thumb');
        e.preventDefault();
        return false;
    });
	});
	$(document).delegate('*[data-toggle="lightbox"]', 'click', function(event) {event.preventDefault();$(this).ekkoLightbox();});

	
})(jQuery);


//DA Addition to implement smooth inpage anchor scrolling
  $(function() {
    $('a[href*="#"]:not([href="#"]):not([data-toggle="tab"])').click(function() {
      if (location.pathname.replace(/^\//,'') == this.pathname.replace(/^\//,'') && location.hostname == this.hostname) {
        var target = $(this.hash);
        target = target.length ? target : $('[name=' + this.hash.slice(1) +']');
        if (target.length) {
          $('html, body').animate({
            scrollTop: target.offset().top
			}, 1500, 'easeInOutExpo');
          return false;
        }
      }
    });
  });	

	function faqjump(){
		$('#productSpecs li').removeClass('active');
		$('#productSpecs li.pscustom-wide a').attr("aria-expanded", "false");
		$('#productSpecs li.pscustom-wide2 a').attr("aria-expanded", "true");
		$('#productSpecs li.pscustom-wide2').addClass('active');
		$('#specsTabContent .tab-pane').removeClass('active in');
		$('#specsTabContent #faq.tab-pane').addClass('active in');
		$('html, body').stop().animate({
				scrollTop: $('#custom-order').offset().top
			}, 1500, 'easeInOutExpo');
        return false;
	};

	function specjump(){
		$('#productSpecs li').removeClass('active');
		$('#productSpecs li.pscustom-wide a').attr("aria-expanded", "true");
		$('#productSpecs li.pscustom-wide2 a').attr("aria-expanded", "false");
		$('#productSpecs li.pscustom-wide').addClass('active');
		$('#specsTabContent .tab-pane').removeClass('active in');
		$('#specsTabContent #fullSpecs.tab-pane').addClass('active in');
		$('html, body').stop().animate({
				scrollTop: $('#custom-order').offset().top
			}, 1500, 'easeInOutExpo');
        return false;
	};

	function dimensionjump(){
		$('#productSpecs li').removeClass('active');
		$('#productSpecs li.pscustom-wide a').attr("aria-expanded", "true");
		$('#productSpecs li.pscustom-wide2 a').attr("aria-expanded", "false");
		$('#productSpecs li.pscustom-wide').addClass('active');
		$('#specsTabContent .tab-pane').removeClass('active in');
		$('#specsTabContent #dimensions-tab.tab-pane').addClass('active in');
		$('html, body').stop().animate({
				scrollTop: $('#custom-order').offset().top
			}, 1500, 'easeInOutExpo');
        return false;
	};
	
	function videojump(){
		$('#productSpecs li').removeClass('active');
		$('#productSpecs li.pscustom-wide a').attr("aria-expanded", "false");
		$('#productSpecs li.pscustom-wide2 a').attr("aria-expanded", "true");
		$('#productSpecs li.pscustom-wide2').addClass('active');
		$('#specsTabContent .tab-pane').removeClass('active in');
		$('#specsTabContent #specVideos.tab-pane').addClass('active in');
		$('html, body').stop().animate({
				scrollTop: $('#custom-order').offset().top
			}, 1500, 'easeInOutExpo');
        return false;
	};

	function arraycustomjump(){
		$('html, body').stop().animate({
				scrollTop: $('#bundle-stands').offset().top -100
			}, 1500, 'easeInOutExpo');
        return false;
	};


	function tlpcustomjump(){
		$('html, body').stop().animate({
				scrollTop: $('#tlp-jump').offset().top -100
			}, 1500, 'easeInOutExpo');
        return false;
	};
