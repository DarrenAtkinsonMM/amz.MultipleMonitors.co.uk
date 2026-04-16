<!-- #include file="inc_productcart.asp" -->

<% 
    assetType = "Files"
    If Request("img") = "yes" Then
        assetType = "Images"
    End If
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>

    <link href="../../scripts/style/editor.css" rel="stylesheet" type="text/css" />
    <style type="text/css">
			#inpFolder {
				border: 1px inset #ddd;
				font-size: 12px;
				-moz-border-radius: 3px;
				-webkit-border-radius: 3px;
				padding-left: 7px;
			}
    </style>
    <script src="jquery/jquery-1.7.min.js" type="text/javascript"></script>
    <script src="jqueryFileTree/jqueryFileTree.js" type="text/javascript"></script>
    <link href="jqueryFileTree/jqueryFileTree.css" rel="stylesheet" type="text/css" />
    <script src="jqueryFileTree/jquery.easing.js" type="text/javascript"></script>

    <!--use protocol relative url-->
    <link href="//fonts.googleapis.com/css?family=Arvo" rel="stylesheet" type="text/css" />

    <link href="uploadifive/uploadifive.css" rel="stylesheet" type="text/css" />
    <script src="uploadifive/jquery.uploadifive.min.js" type="text/javascript"></script>

    <link href="uploadify/uploadify.css" rel="stylesheet" type="text/css" />
    <script src="uploadify/jquery.uploadify.min.js" type="text/javascript"></script>

    <script type="text/javascript">

        var catalogURL = "<%= catalogURL %>";
		var base = "<%= imageBaseFolder %>"; /*Relative to Root*/

        var readonly = false;
        var fullpath = <%= imageFullPath %>;
        var flashUpload = true;
			
        function autoResize(id) {
        	var newheight;
        	var newwidth;

        	if (document.getElementById) {
        		newheight = document.getElementById(id).contentWindow.document.form1.offsetHeight;
        		newwidth = document.getElementById(id).contentWindow.document.form1.offsetWidth;
        	}

        	document.getElementById(id).style.height = "";

        	document.getElementById(id).height = (newheight) + "px";
        	document.getElementById(id).width = (newwidth) + "px";
        }

        $.urlParam = function (name) {
            var results = new RegExp('[\\?&]' + name + '=([^&#]*)').exec(window.location.href);
            if (!results) {
                return 0;
            }
            return results[1] || 0;
        }
        var qs = "";
        if ($.urlParam('img') == "yes") qs = "?img=yes";


        $(document).ready(function () {

            $("#active_folder").val(base); /*default value*/

            renderTree(false, $("#search").val());

            if (readonly) {
                $("#lnkNewFolder").css("display", "none");
                $("#lnkUpload").css("display", "none");
            }

            $("#search").keydown(function(e) {
            	if (e.keyCode == 13) {
            		var searchTerms = $(this).val();

            		renderTree(false, searchTerms);
            	}
            });
        });

        function renderTree(bPreview, searchTerms, addedFileName) {
        	var loadingMsg = "Loading...";
        	if (searchTerms && searchTerms.length > 0) {
        		loadingMsg = "Searching...";
        	}

            $('#container_id').fileTree({
                root: base + '/',
                script: 'jqueryFileTree/jqueryFileTree.asp'+qs,
                expandSpeed: 750,
                collapseSpeed: 750,
                expandEasing: 'easeOutBounce',
                collapseEasing: 'easeOutBounce',
                loadMessage: loadingMsg,
                multiFolder: true,
								searchTerms: searchTerms
            }, function (file) {
              var fileurl = '';
              var ext = file.split('.').pop().toLowerCase();
              var filename = file.substr(file.lastIndexOf("/") + 1);
              if ($.inArray(ext, ['gif', 'png', 'jpg', 'jpeg']) != -1) {
                  $("#preview_id").html("<table><tr><td><a id='idFile' href='" + file + "' target='_blank'><img id='imgFile' src='" + file + "' style='width:70px;padding:4px;border:#cccccc 1px solid;background:#ffffff;margin-bottom:3px' /></a></td><td style='padding-left:20px;padding-right:20px;width:100%;text-align:left;word-break:break-all;'>" + filename.replace("%23", "#") + "<br /><a id='lnkDelFile' style='font-weight:normal;font-size:10px;color:#c90000;word-spacing:2px;white-space:nowrap;' href='javascript:deleteFile()'>DELETE FILE</a></td></tr></table>");
                  if (fullpath) { fileurl = file }
                  else { fileurl = file };
                  try {
                      parent.fileclick(fileurl, catalogURL);
                  }
                  catch (e) { }
              }
              else {
                  if (ext.indexOf("/") == -1) {
                  		$("#preview_id").html("<table><tr><td><a id='idFile' target='_blank' href='" + file + "' style='color:#000000;background:#ffffff;margin-right:5px;'>" + filename.replace("%23", "#") + "</a></td><td>&nbsp;&nbsp;<a id='lnkDelFile' style='font-weight:normal;font-size:10px;color:#c90000;word-spacing:2px;white-space:nowrap;' href='javascript:deleteFile()'>DELETE FILE</a></td></tr></table>");
                      if (fullpath) { fileurl = window.location.protocol + "//" + window.location.host + file }
                      else { fileurl = file };
                      try {
                          parent.fileclick(fileurl, catalogURL);
                      }
                      catch (e) { }
                  }
              }

              preview();

              if (file.substr(file.length - 1) == "/") {
                  //folder is selected
                  $("#preview_id").html("");
              }

              var active_folder = file.substr(0, file.lastIndexOf('/')); /* ex. /images/sample */
              $("#active_folder").val(active_folder);
              $("#folder_id").html(active_folder.replace(base, '') + "/   &nbsp;&nbsp;&nbsp; <a id='lnkDelFolder' href='javascript:deleteFolder()' style='display:none;font-weight:normal;font-size:10px;color:#c90000;word-spacing:2px'>DELETE&nbsp;FOLDER</a>");

              if ($("#active_folder").val() == base) {
                  $("#lnkDelFolder").css("display", "none");
              }
              else {
                  $("#lnkDelFolder").css("display", "inline");
              }
              $("#lnkNewFolder").css("display", "inline");

              if (readonly) {
                  $("#lnkDelFile").css("display", "none");
                  $("#lnkDelFolder").css("display", "none");
                  $("#lnkNewFolder").css("display", "none");
                  $("#lnkUpload").css("display", "none");
              }

            });

            jQuery("#divNewFolder").hide();
            jQuery("#divUpload").hide();
            if (!bPreview) jQuery("#divPreview").hide();
        }

        function deleteFile() {
            if (confirm("Are you sure you want to delete this file?")) {
                $.post('server/delfile.asp', { file: $("#idFile").attr("href") },
                function (data) {
                    refresh();
                    jQuery("#divPreview").hide();
                });
            }
        }

        function deleteFolder() {
            if (confirm("Are you sure you want to delete this folder?")) {
                $.post('server/delfolder.asp', { folder: $("#active_folder").val() },
                function (data) {
                    var active_folder = $("#active_folder").val(); //data.substr(0, data.lastIndexOf('/'));
					active_folder = active_folder.substr(0, active_folder.lastIndexOf('/'));
                    $("#active_folder").val(active_folder);
                    $("#folder_id").html(active_folder.replace(base, '') + "/   &nbsp;&nbsp;&nbsp; <a id='lnkDelFolder' href='javascript:deleteFolder()' style='display:none;font-weight:normal;font-size:10px;color:#c90000;word-spacing:2px'>DELETE&nbsp;FOLDER</a>");

                    refresh();
                });
            }
        }

        function panelFolder() {
            jQuery("#divUpload").hide();
            jQuery("#divPreview").hide();
            $("#divNewFolder").slideToggle(750, 'easeOutBounce');
        }

        function createFolder() {
            $.post('server/newfolder.asp', { folder: $("#active_folder").val() + "/" + $("#inpFolder").val() },
                function (data) {
                    refresh();
                    $("#inpFolder").val("");
                });
        }


        function refresh(addedFile) {
            if (base == $("#active_folder").val()) {
            	renderTree(true, $("#search").val(), addedFile); /*Refresh Root*/
            } else {
            	var rel = $("#active_folder").val() + '/';
            	$('a[rel="' + rel + '"]').trigger("click");
            }

            //$("#preview_id").html('');
        }

        function upload() {
            jQuery("#divNewFolder").hide();
            jQuery("#divPreview").hide();
            if (!$("#divUpload").is(":visible")) {
            	panelUpload();
            }
            $("#divUpload").slideToggle(750, 'easeOutBounce');
        }

        function uploadComplete(file, data) {
        	var errorStr = "";

        	if (data == "FILENAME") {
        		errorStr = "The name of this file is invalid. Please check the filename and try again.";
        	} else if (data == "FILETYPE") {
        		errorStr = "You are not allowed to upload files of type <strong>*." + file.name.split(".").pop() + "</strong> to this store. Please use a different file format and try again.";
        	} else if (data == "NOTALLOWED") {
        		errorStr = "You do you not have the neccessary permissions to upload files.";
        	} else {
        		fileName	= file.name;
        		fileExt		= fileName.split('.').pop();
        		filePath	= $("#active_folder").val() + "/" + fileName;

						// Remove it from the current list if it's a re-upload
        		$(".jqueryFileTree").find("li.file").each(function() {
        			var itemName = $.trim($(this).find("a").html());
							
        			if (itemName == fileName) {
        				$(this).remove();
        			}
        		});

						// Add and highlight temp item
        		$("<li class='file ext_" + fileExt + " added'><a href='#' rel='" + filePath + "'>" + file.name + "</a></li>").insertBefore($(".jqueryFileTree").find("li.file,li.message").first());

						// Remove any messages
        		$(".jqueryFileTree").find("li.message").remove();

        		$(".added a").trigger('click');
        	}

        	return errorStr;
        }

        function panelUpload() {
        	if(flashUpload) {
        		$("#divUpload").html("<h3 style='margin-top:0px'>Upload Files</h3><div id='queue'></div><input id='File1' type='file' />");
        		<%
							timestamp = now()
						%>
        		$("#File1").uploadifive({
        			'auto': true,
        			'multi': true,
        			'formData': {
        				'timestamp':	'<%= timestamp %>',
        			},
							'queueID': 'queue',
							'uploadScript': 'server/upload.asp?folder=' + $("#active_folder").val(),
							'onProgress': function(file, e) {
								$(".uploadifive-queue-item.error").remove();
							},
							'onUploadComplete': function(file, data) {
								var errorMsg = uploadComplete(file, data);

								if (errorMsg.length > 0) {
									file.queueItem.addClass("error");
									file.queueItem.find(".fileinfo").html(" - " + errorMsg);
								}
							},
							'onFallback': function() {
								$("#File1").uploadify({
									'swf': 'uploadify/uploadify.swf',
									'uploader': 'server/upload.asp?folder=' + $("#active_folder").val() ,
									'formData': { 
										'timestamp':	'<%= timestamp %>',
									},
									'queueID': 'queue',
									'multi': true,
									'auto': true,
									'removeCompleted': false,
									'onUploadProgress': function(file, bytesUploaded, bytesTotal, totalBytesUploaded, totalBytesTotal) {
										$(".uploadify-queue-item.error").remove();
									},
									'onUploadSuccess': function (file, data, response) {
										var errorMsg = uploadComplete(file, data);

										if (errorMsg.length > 0) {
											$('#' + file.id).addClass("error");
											$('#' + file.id).find(".uploadify-progress").fadeOut();
											$('#' + file.id).find(".data").html(" - " + errorMsg);
										}
									}
								});
							}
        		});
        	}  else {

        		$("#frmUpload").attr("src", "basic_upload.asp?folder=" + $("#active_folder").val());

            }

        }

        function preview() {
            if ($("#divPreview").css('display') == 'block') return;
            jQuery("#divNewFolder").hide();
            jQuery("#divUpload").slideUp();
            $("#divPreview").slideToggle(750, 'easeOutBounce');
        }
    </script>
</head>
<body style="margin:0px;background:#ffffff;font-family:Arvo;font-size:12px">
    <form id="form1">
    <input id="active_folder" name="folder" type="hidden" />

    <div id="topPanel" style="display:block;position:fixed;top:0px;left:0px;width:100%;padding:15px;background:#fcfcfc;border-bottom:#f7f7f7 1px solid;border-right:#f7f7f7 1px solid">

        <div style="float: left; margin-top:5px;margin-bottom:5px;">
            <a id="lnkNewFolder" href="javascript:panelFolder()" style="margin-right:10px;font-size:10px;color:#000;">NEW FOLDER</a>
            <a id="lnkUpload" href="javascript:upload();" style="margin-right:10px;font-size:10px;color:#000;">UPLOAD</a>
        </div>

				<div style="float: right; margin-right: 20px">
					<input type="text" style="padding: 2px" id="search" placeholder="Search <%= assetType %>" size="30" />
				</div>

        <div id="divPreview" style="clear: both; margin-top:15px;padding:15px;padding-bottom:14px;border:#f3f3f3 1px solid;background:#fefefe;">
            <div style="font-weight:bold;font-size:12pt;margin-bottom:5px;">Folder: <span id="folder_id">/</span> &nbsp; <a href="javascript:refresh()" style="margin-right:10px;font-size:10px;color:#000;font-weight:normal;">REFRESH</a></div>
            <div id="preview_id"></div>
						<div style="text-align: right;margin-right:10px;margin-bottom:-10px">
							<a href="#" onclick="$('#divPreview').slideToggle(750, 'easeOutBounce'); return false;">close</a>
						</div>
        </div>
        <div id="divUpload" style="clear: both; margin-top:15px;padding:15px;border:#f3f3f3 1px solid;background:#fefefe;">
        	<iframe id="frmUpload" border="0" style="border:none;height:35px;width:98%" src="about:blank" onload="autoResize('frmUpload');"></iframe>
        </div>

        <div id="divNewFolder" style="clear: both; margin-top:15px;padding:15px;height:65px;border:#f3f3f3 1px solid;background:#fefefe;">
            <h3 style="margin-top:0px">New Folder</h3>
            <input type="text" id="inpFolder" style="width:120px;height:26px;float:left" value="">
            <input type="button" id="btnAddFolder" value=" create " onclick="createFolder()" class="inpBtn" style="width:70px;height:30px;margin-right:0px" onmouseover="this.className='inpBtnOver';" onmouseout="this.className='inpBtnOut'">
        </div>

    </div>

    <div id="container_id"  style="padding:15px;margin-top:52px;">
    </div>


    </form>
</body>
</html>
