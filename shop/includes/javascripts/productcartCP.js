// Reuseable window method (Non Modal) for Control Panel.
function pcCPWindow(fileName, h, w)
{
    myFloater = window.open(fileName,'popup','toolbar=no,status=no,location=no,menubar=no,height=' + h + ',width=' + w + ',scrollbars=no')
}
// Method to change images in Control Panel
function chgWin(file, window) {
    msgWindow=open(file,window,'scrollbars=yes,resizable=yes,width=600,height=620');
    if (msgWindow.opener == null) msgWindow.opener = self;
}

// Methods for the genCatNavigation.asp page
function selectShowNavOptions(item) {
	if ($pc(item).val() == 0) {
		$pc("#OtherNavOptions").show();
		$pc("#JQNavOptions").hide();
	} else {
		$pc("#OtherNavOptions").hide();
		$pc("#JQNavOptions").show();
	}
}

// START: Initialize the Control Panel help tooltips
function initHelpTooltips() 
{
  var tooltipScope = ".pcCPcontent";
  var helpItem = ".pcCPhelp";

  // Load and display tooltips on hover
  $pc(tooltipScope).on("mouseenter", helpItem, function () {
    var item = $pc(this);
    var tooltip = item.data("opentips");

    if (tooltip == undefined) {
      var helpTip = new Opentip(item, { target: true });

      $pc.get(item.attr("href"), "ajax=1", function (data) {
        helpTip.setContent(data);

        if (item.is(":hover")) helpTip.show();
      });
    }
  });

  // Show popup window on click and hide all tooltips
  $pc(tooltipScope).on("click", helpItem, function (e) {
    win($pc(this).attr("href"));

    for (var i = 0; i < Opentip.tips.length; i++) {
      Opentip.tips[i].hide();
    }

    e.preventDefault();
  });
}
// END: Initialize the Control Panel help tooltips

// START: Sortable + Scrollable Drag Functions

var pcv_scrollableDragItem = null;
var pcv_scrollableDragMaxHeight = 0;
var pcv_scrollableDragViewportHeight = 0;
var pcv_scrollableDragMethod = "";

function pcf_initScrollableDrag()
{
  // Initialize the height of the page as well as the viewport height
  pcv_scrollableDragMaxHeight = $pc("body").height();
  pcv_scrollableDragViewportHeight = $pc(window).height();

  // Get scroll method by testing whether <html> or <body> gets the scroll property.
  // This is for compatibility with FF/IE (<html> method) versus Chrome/Safari (<body> method)
  var scroll = $pc(window).scrollTop();
  $pc("html,body").scrollTop(1);
  if ($pc("html").scrollTop() > 0) pcv_scrollableDragMethod = "html";
  if ($pc("body").scrollTop() > 0) pcv_scrollableDragMethod = "body";
  $pc("html,body").scrollTop(scroll);

  // Update document and viewport heights
  $pc(window).on("click", function () {
    pcv_scrollableDragMaxHeight = $pc("body").height();
    pcv_scrollableDragViewportHeight = $pc(window).height();
  });

  // Capture mouse move events
  $pc(document).on("mousemove", function (event) {
    if (pcv_scrollableDragItem == null) return;

    window.x = event.pageX;
    window.y = event.pageY;

    pcf_handleScrollableDrag();
  });
}

function pcf_handleScrollableDrag()
{
  if (pcv_scrollableDragItem != null) {
    var height = $pc(pcv_scrollableDragItem).height();
    var scroll = $pc(pcv_scrollableDragMethod).scrollTop();
    var offset = window.y - scroll;

    // Scroll Up
    if (offset - height <= 0) {
      var move_factor = -parseInt((offset - height) / 2);
      var move_pixels = scroll - move_factor;

      if (move_factor > 0 && move_pixels > 0) {
        $pc(pcv_scrollableDragMethod).stop().animate({
          scrollTop: move_pixels
        }, {
          duration: 5,
          complete: function () {
            if (window.y - move_factor > 0) {
              window.y = window.y - move_factor;

              // Propagate mousemove event
              e = $pc.Event("mousemove");
              e.pageX = window.x;
              e.pageY = window.y;
              $pc(document).trigger(e);
            }
          }
        });
      }
    }

    // Scroll Down
    if (offset + height >= pcv_scrollableDragViewportHeight) {
      var move_factor = -parseInt((pcv_scrollableDragViewportHeight - offset - height) / 2);
      var move_pixels = scroll + move_factor;

      if (move_factor > 0 && (scroll + pcv_scrollableDragViewportHeight) - height < pcv_scrollableDragMaxHeight) {
        $pc(pcv_scrollableDragMethod).stop().animate({
          scrollTop: move_pixels
        }, {
          duration: 5,
          complete: function () {
            window.y = window.y + move_factor;

            // Propagate mousemove event
            e = $pc.Event("mousemove");
            e.pageX = window.x;
            e.pageY = window.y;
            $pc(document).trigger(e);
          }
        });
      }
    }
  }
}

function pcf_initSortable()
{
  // Setup sortable list variables
  var sortableClass = ".pcCPsortable";
  var sortableHandleClass = ".pcCPsortableHandle";
  var sortableNumberClass = ".pcCPsortableIndex";
  var sortableInputClass = ".pcCPsortableOrder";
  var sortableExceptionClass = ".pcCPnotSortable";

  if ($pc(sortableClass).length > 0) {
    var sortableOptions = {
      distance: 5,
      onDrag: function (item, container, _super) {
        _super(item, container);

        pcv_scrollableDragItem = item;
      },
      onDrop: function (item, container, _super) {
        _super(item, container);

        pcv_scrollableDragItem = null;

        $pc(sortableClass).find("li").each(function (i) {
          var index = i + 1;

          $pc(this).find(sortableInputClass).val(index);
          $pc(this).find(sortableNumberClass).html(index);
        });
      }
    };
    
    // Only use the handle if it exists
    if ($pc(sortableHandleClass).length > 0) {
      sortableOptions.handle = sortableHandleClass;
    }

    if ($pc(sortableExceptionClass).length > 0) {
    	sortableOptions.except = sortableExceptionClass;
    }

    // Initialize sortable class
    $pc(sortableClass).sortable(sortableOptions);
  }
}
// END: Sortable + Scrollable Drag Functions

$pc(document).ready(function () {

  // SSL Checkbox
  $pc('#sslCheckbox').click(function(){
    if ( $pc(this).is(':checked') )
        $pc('#sslModal').modal();
  })

  // Intialize numbered sortable lists
  pcf_initSortable();

  // Initalize scrollable dragging
  pcf_initScrollableDrag();

  // Intialize tooltips
  initHelpTooltips();

  // START: AdminSettings.asp
  $pc(".pcCPsocialMoreLink").click(function (e) {
    $pc(this).toggleClass("expanded");
    $pc(this).parents('li').find(".pcCPsocialMore").slideToggle('fast');

    e.preventDefault();
  });

  var paymentOrigHeight;
  $pc(".pcCPpaymentCustomizeLink").click(function (e) {
    var litem = $pc(this).parents("li");
    var custom = litem.find(".pcCPpaymentCustom");

    if (custom.is(":visible")) {
      $pc(this).html("Customize &raquo;");
      custom.slideUp('fast');
    } else {
      $pc(this).html("Customize &laquo;");
      custom.slideDown('fast');
    }

    e.preventDefault();
  });
	// END: AdminSettings.asp

	// START: modifyProduct.asp
  $pc("#noShipping").click(function () {
  	if ($pc(this).is(":checked")) {
  		$pc("#noShippingSettings").show();
  	} else {
  		$pc("#noShippingSettings").hide();
  	}
	});

	// END: modifyProduct.asp


	// START: Auto Complete Off
    $pc('input[class="pcAutoCompleteOff"]').attr('autocomplete', 'new-password');
	// END: Auto Complete Off
   
 }); 


 
