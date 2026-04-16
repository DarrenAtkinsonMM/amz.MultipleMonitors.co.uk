<%
Dim editorClass, editorContainerClass

editorClass = "htmleditor"
editorContainerClass = "htmleditor_container"
%>

<script type="text/javascript" src="../htmleditor/scripts/innovaeditor.js"></script>

<script type=text/javascript>
	
	function htmleditorcontent(editorName) {
		return window[editorName].getHTMLBody();
	}

	function htmleditorjs(editorName, fieldName, container) {
		window[editorName] = new InnovaEditor(editorName);
		
    /* Get Textarea Size */
		window[editorName].width = '100%';
		window[editorName].height = 280;
      
    /* Settings */
		window[editorName].toolbarMode = 1; //activate tab toolbar
    window[editorName].enableCssButtons = true;
    window[editorName].enableTableAutoformat = true;
		window[editorName].disableFocusOnLoad = true;
		window[editorName].returnKeyMode = 2;
    
		/* Custom Buttons */
		window[editorName].arrCustomButtons = [["HTML5Video", "modalDialog('../htmleditor/scripts/common/webvideo.htm',700,640,'HTML5 Video');", "HTML5 Video", "btnVideo.gif"]];
				
    /* Add Groups */
    window[editorName].groups = [
      ["font", "", 
        ["FontName", "Paragraph", "ForeColor", "BackColor", "BRK", "Bold", "Italic", "Underline", "Strikethrough",]
      ],
      ["paragraph", "", 
        ["Bullets", "Numbering", "Indent", "Outdent", "BRK", "JustifyLeft", "JustifyCenter", "JustifyRight", "RemoveFormat" ]
      ],
      ["elements", "", 
        ["Table" ,"TableDialog", "HTML5Video", "BRK", "LinkDialog", "ImageDialog", "YoutubeDialog" ]
      ],
      ["tools", "",
        ["Undo", "Redo", "BRK", "FullScreen", "Preview" ]
      ],
			["advanced_fonts", "",
        ["TextDialog", "Superscript", "FontDialog", "Line", "CharsDialog" ]
      ],
			["advanced_elements", "",
        ["FlashDialog" ]
      ],
			["advanced_tools", "",
        ["SourceDialog", "SearchDialog" ]
      ]
		];
		 
		window[editorName].tabs = [
			["tabCommon", "Common", ["font","paragraph","elements", "tools"]],
			["tabAdvanced", "Advanced", ["advanced_fonts", "advanced_elements", "advanced_tools"]]
		];
  
    /* Files */
    window[editorName].css = "../htmleditor/styles/simple.css";
    window[editorName].fileBrowser = "../../addons/assetmanager/asset.asp";

    /*Render the editor*/
		if (container.length > 0) {
    	window[editorName].REPLACE(fieldName, container);
			
			// Bind extra stuff to it
			
		} else {
			window[editorName].REPLACE(fieldName);
		}
	}

	$pc(document).ready(function () {
		$pc(".<%= editorClass %>").each(function(idx) {
			var editorId = $pc(this).attr("id");
			var tabIndex = $pc(this).attr("tabindex");
			
			// Help out and set the ID to the name
			if (editorId.length < 1) {
				editorId = $pc(this).attr("name");
				$pc(this).attr("id") = editorId;
			}
			
			if (editorId.length > 0) {
				
				// Add tab index
				tabIndexStr = "";
				if (tabIndex !== undefined) {
					tabIndexStr = "tabindex='" + tabIndex + "'";	
				}
				
				editorContainerId = editorId + "Div";
				
				// Add container div
				$pc(this).after("<div id='" + editorContainerId + "' class='<%= editorContainerClass %>' " + tabIndexStr + "></div>");
				
				htmleditorjs("editor" + idx, editorId, editorContainerId);
			}
		});
		
		// ProductCart Add-On: Take care of tabindex functionality for the HTML Editor	
		$pc(".<%= editorContainerClass %>").each(function() {
			if ($pc(this).contents().length > 1) {
				var editorName = $pc(this).find("table").attr("id").replace("idArea", "");
				var tabIndex = parseInt($pc(this).attr("tabindex"));
				
				$pc(this).focus(function() {
					window[editorName].focus();
				});
				
				if (!isNaN(tabIndex)) {
					window[editorName].onKeyPress = function(e) {
						if (e.keyCode == 9) {							
							var nextTabIndex = tabIndex + 1;
							var nextItem = $pc("[tabindex='" + nextTabIndex + "']:visible");
							
							if (nextItem.length > 0) {
								nextItem.focus();
								
								if (e.preventDefault) { 
									e.preventDefault();
								} else {
									e.returnValue = false;
								}
							}
						}
					}
				}
			} else {
				$pc(this).hide();
			}
		});
	});
</script>

<%

Function htmleditor(editorName, fieldName)
%>       
	<script type=text/javascript>
		htmleditorjs("<%= editorName %>", "<%= fieldName %>", "")
  </script>
<%
End Function
%>