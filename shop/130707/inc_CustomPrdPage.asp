<%'//Get Default Layout
dpTop=""
dpTopLeft=""
dpTopRight=""
dpBottom=""
dpTabs=""
query="SELECT pcDPL_Top,pcDPL_TopLeft,pcDPL_TopRight,pcDPL_Middle,pcDPL_Bottom,pcDPL_Tabs FROM pcDefaultPrdLayout;"
set rsQ=connTemp.execute(query)
if not rsQ.eof then
	dpTop=rsQ("pcDPL_Top")
	dpTopLeft=rsQ("pcDPL_TopLeft")
	dpTopRight=rsQ("pcDPL_TopRight")
	dpMiddle=rsQ("pcDPL_Middle")
	dpBottom=rsQ("pcDPL_Bottom")
	dpTabs=rsQ("pcDPL_Tabs")
end if
set rsQ=nothing%>
<link href="../pc/css/bootstrap-editable.css" rel="stylesheet"/>
<script type="text/javascript" src="../includes/javascripts/bootstrap-editable.js"></script>
<%
ReDim elementList(0)
pcv_intCounter = 0

query = "SELECT widget_Shortcode, widget_Desc, widget_Uri, widget_Method, widget_Lang FROM pcWidgets"
set rs = Server.CreateObject("ADODB.Recordset")  
rs.Open query, connTemp, adOpenStatic, adLockReadOnly, adCmdText
If Not rs.Eof Then
    pcv_intTotalCount = rs.RecordCount
    ReDim elementList(pcv_intTotalCount - 1)    
    Do While Not rs.Eof         
        pcv_strShortcode = rs("widget_Shortcode")
        pcv_strDesc = rs("widget_Desc")
        pcv_strUri = rs("widget_Uri")
        pcv_strMethod = rs("widget_Method")
        pcv_strLang = rs("widget_Lang")
        elementList(pcv_intCounter) = Array(pcv_strShortcode, pcv_strDesc)
        pcv_intCounter = pcv_intCounter + 1
        rs.movenext
    Loop
End If
Set rs = Nothing

columnCount = CInt(UBound(elementList) / 3)

tmpEleList = ""
tmpEleListDD = ""
tmpEleList = tmpEleList & "<ul class='OptBtnCol'>"
count = 0
If pcv_intCounter > 0 Then
	For Each element In elementList
		
		'// Button list
		tmpEleList = tmpEleList & "<li><button class='OptBtn btn btn-default' value='" & element(0) & "'>" & element(1) & "</button></li>"

		If count = columnCount Then
			tmpEleList = tmpEleList & "</ul>"
			tmpEleList = tmpEleList & "<ul class='OptBtnCol'>"
			count = 0
		Else
			count = count + 1
		End If
		
		'// Drop down
		If Element(0) <> "CUSTOMHTML" Then
			tmpEleListDD = tmpEleListDD & "<option value='" & Element(0) & "'>" & Element(1) & "</option>"
		End If
	Next
End If
tmpEleList = tmpEleList & "</ul>"
tmpEleList = tmpEleList & "<div style='clear: both'></div>"	


noElementsHTML = "<li class='notSortable' style='width: 100%'><div class='pcCPmessageInfo'>No Elements. Use the 'Add new element' section or drag-and-drop from another area to add elements.</div></li>"
noTabsHTML = "<div class='pcCPmessageInfo'>No Tabs. Click the add button above to get started.</div>"

%>
<tr>
	<td colspan="2" class="pcCPspacer">
		<input type="hidden" name="ppTop" id="ppTop" value="" />
		<input type="hidden" name="ppTopLeft" id="ppTopLeft" value="" />
		<input type="hidden" name="ppTopRight" id="ppTopRight" value="" />
		<input type="hidden" name="ppMiddle" id="ppMiddle" value="" />
		<input type="hidden" name="ppTabs" id="ppTabs" value="" />
		<input type="hidden" name="ppBottom" id="ppBottom" value="" />
        <input type="hidden" name="saveDefault" id="saveDefault" value="" />
	</td>
</tr>
<tr>
	<td colspan="2">
		<table id="TabWorking" class="pcCPcontent">
			<tbody class="CustomPrdLayout" <%= pcv_strCustPrdDisplayStyle %>>
				<tr>
					<th colspan="2">
						Customize Product Details Page
					</th>
				</tr>
				<tr>
					<td colspan="2">
						<div class="bs-callout bs-callout-info">							
							<strong>Please Note: </strong>If a feature is disabled in the storefront settings or the product doesn't contain data for any of the elements below, it may not appear on the <a href="../pc/viewPrd.asp?idproduct=<%= pIdProduct %>&adminPreview=1" target="_blank">storefront</a>.
							<ul>
								<li>To <strong>reset</strong> to the default tabbed display layout, <a href="javascript:DefaultPageSettings();">click here</a></li>
								<%if pidProduct>"0" then%>
								<li>To <strong>copy</strong> the layout from another product, <a href="CopyLayoutFromPrd.asp?idproduct=<%=pidProduct%>">click here</a></li>
								<%end if%>
								<%if ppTop&ppTopLeft&ppTopRight&ppMiddle&ppBottom&ppTabs<>"" then%>
								<li>To <strong>apply</strong> this layout to other products, <a href="ApplyLayoutToPrds.asp?idproduct=<%=pidProduct%>">click here</a></li>
								<%end if%>
								<li>To <strong>clear</strong> this layout and start from scratch, <a href="javascript:ClearCustomLayout();">click here</a></li>
                                <li>To <strong>save</strong> this layout as the store default, <a href="javascript:SaveAsDefaultLayout();">click here</a></li>
							</ul>
						</div>
					</td>
				</tr>
				<tr>
					<td colspan="2" class="pcCPspacer"></td>
				</tr>
				<tr>
					<td colspan="2">
						<div class="panel panel-default">
							<div class="panel-heading">Add new element</div>
							<div class="panel-body">
								<div style="clear:both; padding-top: 10px; padding-bottom: 10px;">
									Position:&nbsp; 
									<select name="PrdEPos" id="PrdEPos">
										<option value="0">Top Area</option>
										<option value="1">Top-Left Area</option>
										<option value="2">Top-Right Area</option>
										<option value="3">Middle Area</option>
										<option value="4">Bottom Area</option>
									</select>&nbsp;
									Element:&nbsp;
									<select name="PrdEle" id="PrdEle">
										<%=tmpEleListDD%>
									</select>&nbsp;
									<input type="button" class="btn btn-default"  name="addele" value="Add new" onClick="addElementToList()">
								</div>
							</div>
						</div>
					</td>
				</tr>
				<tr>
					<th colspan="2">Top of Page's Elements</th>
				</tr>
				<tr>

					<% '// Top Top Area %>
					<td colspan="2" align="center">
						<div align="left" class="pcCPsectionTitle">Top</div>
						<div id="TopArea0" class="AreaList" style="clear:both">
							<ul id='TT' class='connectedSortable sortable'>
								<%= noElementsHTML %>
							</ul>
						</div>
					</td>

				</tr>
				<tr>

					<% '// Top-Left Area %>
					<td style="width: 50%; vertical-align: top;">
						<div class="pcCPsectionTitle">Top-Left</div>
						<div id="TopArea1" class="AreaList" style="clear:both">
							<ul id='TL' class='connectedSortable sortable'>
								<%= noElementsHTML %>
							</ul>
						</div>
					</td>

					<% '// Top-Right Area %>
					<td style="width: 50%; vertical-align: top;">
						<div class="pcCPsectionTitle">Top-Right</div>
						<div id="TopArea2" class="AreaList" style="clear:both">
							<ul id='TR' class='connectedSortable sortable'>
								<%= noElementsHTML %>
							</ul>
						</div>
					</td>

				</tr>
	
				<tr valign="top">
					<% '// Middle Area %>
					<td colspan="2" style="text-align: center; vertical-align: top;">
						<div align="left" class="pcCPsectionTitle">Middle</div>
						<div id="TopArea3" class="AreaList" style="clear:both">
							<ul id='ML' class='connectedSortable sortable'>
								<%= noElementsHTML %>
							</ul>
						</div>
					</td>
				</tr>

				<tr>
					<td colspan="2" class="pcCPspacer"></td>
				</tr>
			</tbody>

			<% '// Product Tabs Area %>
			<tbody class="CustomTabsLayout" <%= pcv_strCustTabDisplayStyle %>>
				<tr>
					<th colspan="2">Product Tabs</th>
				</tr>
				<tr>
					<td colspan="2">
						<div id="TabsArea" style="clear:both"></div>
					</td>
				</tr>
				<tr>
					<td colspan="2" class="pcCPspacer"></td>
				</tr>
			</tbody>
			
			<% '// Bottom Area %>
			<tbody class="CustomPrdLayout" <%= pcv_strCustPrdDisplayStyle %>>
				<tr>
					<th colspan="2">Bottom of Page's Elements</th>
				</tr>
				<tr>
					<td colspan="2">
						<div id="BottomArea" class="AreaList" style="clear:both">
							<ul id='BL' class='connectedSortable sortable'>
								<%= noElementsHTML %>
							</ul>
						</div>
					</td>
				</tr>
				<tr>
					<td colspan="2">
						<div class="bs-callout bs-callout-warning">
							<h4>Tips and Tricks</h4>
							<ul>
								<li>All changes are applied only when you save the product.</li>
								<li>You can drag and drop elements to order them or between lists to change their position.</li>
								<li>You can rename product tabs by clicking the tab name when underlined (and when you see the edit text cursor).</li>
								<li>Zone Rules:
									<ul>
										<li>"Add To Cart Zone" must be placed below "Options" and "Custom Input Fields"</li>
										<li>"Wishlist Zone" must be placed below "Options"</li>
									</ul>
								</li>
							</ul>
						</div>
					</td>
				</tr>
			</tbody>
		</table>
	</td>
</tr>
<tr>
	<td colspan="2">
		
		<script type=text/javascript>

			var TopTop = "<%=ppTop%>";
			var TopLeft = "<%=ppTopLeft%>";
			var TopRight = "<%=ppTopRight%>";
			var Middle = "<%=ppMiddle%>";
			var BottomL = "<%=ppBottom%>";
			var PrdTabs=new Array();
			var PrdTabsSorted = new Array();

			<%tmpTabCount=0
			if ppTabs<>"" then
			tmpStr1=split(ppTabs,"||")
			for i=lbound(tmpStr1) to ubound(tmpStr1)
				if tmpStr1(i)<>"" then
					tmpStr2=split(tmpStr1(i),"``")
					tmpHTMLContent=tmpStr2(2)
					if tmpHTMLContent<>"" then
						tmpHTMLContent=replace(tmpHTMLContent,"""","\""")
						tmpHTMLContent=replace(tmpHTMLContent,vbCrLf,"")
						tmpHTMLContent=replace(tmpHTMLContent,CHR(60) & "script" & CHR(62),"\<\script\>")
						tmpHTMLContent=replace(tmpHTMLContent,CHR(60) & "/script" & CHR(62),"\<\/\script\>")
					end if
					%>
					PrdTabs[<%=tmpTabCount%>]=new Array();
					PrdTabs[<%=tmpTabCount%>][0]="Tab<%=tmpTabCount+1%>";
					PrdTabs[<%=tmpTabCount%>][1]="<%=tmpStr2(0)%>";
					PrdTabs[<%=tmpTabCount%>][2]="<%=tmpStr2(1)%>";
					PrdTabs[<%=tmpTabCount%>][3]="<%=tmpHTMLContent%>";
					<%
					tmpTabCount=tmpTabCount+1
				end if
			next
		end if%>
			
		var PTCount=<%=tmpTabCount%>;
		var OTCount=<%=tmpTabCount%>;

		function LoadSettingsForLayout(layout, addTabs) {
			if (layout == '') {
				layout = '<%= scViewPrdStyle %>';
			}

			switch (layout.toLowerCase()) {
			case 't':
				TopTop = "PrdName,CatTree,";
				TopLeft = "PrdSKU,PrdRate,PrdW,PrdBrand,PrdStock,PrdDesc,PrdConfig,PrdSearch,PrdRP,PrdPrice,PrdSB,PrdPromo,PrdNoShip,PrdOSM,PrdBOM,PrdOpt,PrdATC,PrdWL,";
				TopRight = "PrdImg,PrdQDisc,";
				Middle = "PrdBtns,";
				BottomL = "";
				PrdTabs=new Array();
				PTCount=0;
				OTCount=0;
				break;
			case 'c':
				TopTop = "PrdName,CatTree,";
				TopLeft = "PrdSKU,PrdRate,PrdW,PrdBrand,PrdStock,PrdDesc,PrdConfig,PrdSearch,PrdRP,PrdPrice,PrdSB,PrdPromo,PrdNoShip,PrdOSM,PrdBOM,PrdOpt,PrdATC,PrdWL,";
				TopRight = "PrdImg,PrdQDisc,";
				Middle = "PrdBtns,PrdCS,PrdLDesc,PrdRev,";
				BottomL = "";
				PrdTabs=new Array();
				PTCount=0;
				OTCount=0;
				break;
			case 'l':
				TopTop = "PrdName,CatTree,";
				TopLeft = "PrdImg,PrdQDisc,";
				TopRight = "PrdSKU,PrdRate,PrdW,PrdBrand,PrdStock,PrdDesc,PrdConfig,PrdSearch,PrdRP,PrdPrice,PrdSB,PrdPromo,PrdNoShip,PrdOSM,PrdBOM,PrdOpt,PrdATC,PrdWL,";
				Middle = "PrdBtns,PrdCS,PrdLDesc,PrdRev,";
				BottomL = "";
				PrdTabs=new Array();
				PTCount=0;
				OTCount=0;
				break;
			case 'o':
				TopTop = "PrdName,CatTree,";
				TopLeft = "";
				TopRight = "";
				Middle = "PrdImg,PrdQDisc,PrdSKU,PrdRate,PrdW,PrdBrand,PrdStock,PrdDesc,PrdConfig,PrdSearch,PrdRP,PrdPrice,PrdSB,PrdPromo,PrdNoShip,PrdOSM,PrdBOM,PrdOpt,PrdATC,PrdWL,PrdBtns,";
				BottomL = "PrdCS,PrdLDesc,PrdRev,";
				PrdTabs=new Array();
				PTCount=0;
				OTCount=0;
				break;
			}

			if (addTabs) {
				if (layout.toLowerCase() == 'o') {
					BottomL = "";
				} else {
					Middle = "PrdBtns,";
				}

				PrdTabs=new Array();
				PTCount=4;
				OTCount=4;
				PrdTabs[0]=new Array();
				PrdTabs[0][0]="Tab1";
				PrdTabs[0][1]="Overview";
				PrdTabs[0][2]="PrdLDesc,";
				PrdTabs[0][3]="";
				PrdTabs[1]=new Array();
				PrdTabs[1][0]="Tab2";
				PrdTabs[1][1]="Input Fields";
				PrdTabs[1][2]="PrdInput,";
				PrdTabs[1][3]="";
				PrdTabs[2]=new Array();
				PrdTabs[2][0]="Tab3";
				PrdTabs[2][1]="Reviews";
				PrdTabs[2][2]="PrdRev,";
				PrdTabs[2][3]="";
				PrdTabs[3]=new Array();
				PrdTabs[3][0]="Tab4";
				PrdTabs[3][1]="Related Products";
				PrdTabs[3][2]="PrdCS,";
				PrdTabs[3][3]="";
			}

			GenWorkingArea();
		}
	
		function LoadDefaultPageSettings() {
			<%if trim(dpTop&dpTopLeft&dpTopRight&dpMiddle&dpBottom&dbTabs)<>"" then%>
				TopTop = "<%=dpTop%>";
				TopLeft = "<%=dpTopLeft%>";
				TopRight = "<%=dpTopRight%>";
				TopRight = "<%=dpMiddle%>";
				BottomL = "<%=dpBottom%>";
				PrdTabs=new Array();
				<%tmpTabCount=0
				if dpTabs<>"" then
					tmpStr1=split(dpTabs,"||")
					for i=lbound(tmpStr1) to ubound(tmpStr1)
						if tmpStr1(i)<>"" then
							tmpStr2=split(tmpStr1(i),"``")
							%>
							PrdTabs[<%=tmpTabCount%>]=new Array();
							PrdTabs[<%=tmpTabCount%>][0]="Tab<%=tmpTabCount+1%>";
							PrdTabs[<%=tmpTabCount%>][1]="<%=tmpStr2(0)%>";
							PrdTabs[<%=tmpTabCount%>][2]="<%=tmpStr2(1)%>";
							PrdTabs[<%=tmpTabCount%>][3]="<%=replace(tmpStr2(2),"""","\""")%>";
							<%
							tmpTabCount=tmpTabCount+1
						end if
					next
				end if%>
				PTCount=<%=tmpTabCount%>;
				OTCount=<%=tmpTabCount%>;

				GenWorkingArea();
			<%else%>
				LoadSettingsForLayout('t');
			<%end if%>
			}

			function DefaultPageSettings() 
			{
				if (!confirm("Are you sure you want to reset your current layout and load the defaults? This action cannot be undone.")) {
					return;
				}

				LoadDefaultPageSettings();
			}
	
			function ClearCustomLayout() 
			{
				if (!confirm("Are you sure you want to completely clear your current layout and remove all elements? This action cannot be undone.")) {
					return;
				}

				TopTop = "";
				TopLeft = "";
				TopRight = "";
				Middle = "";
				BottomL = "";
				PrdTabs=new Array();
				PTCount=0;
				OTCount=0;
		
				document.hForm.ppTop.value=TopTop;
				document.hForm.ppTopLeft.value=TopLeft;
				document.hForm.ppTopRight.value=TopRight;
				document.hForm.ppMiddle.value=Middle;
				document.hForm.ppBottom.value=BottomL;
				document.hForm.ppTabs.value="";

				document.getElementById("TabWorking").style.display='none';

				GenWorkingArea();
			}
			
			function SaveAsDefaultLayout()
			{
				if (!confirm("Are you sure you want to save your current layout as the default product layout?")) {
					return;
				}
				SavePPToFields();
				document.hForm.saveDefault.value = "1";
				document.hForm.submit();
			}
		
			function findItemName(tmpN)
			{
				var x="";
				switch (tmpN)
				{
					<%
					If pcv_intCounter > 0 Then
						For Each Element In elementList
							Response.Write "case '" & Element(0) & "': x='" & Element(1) & "'; break;" & vbCrLf
						Next
					End If
					%>
				}
				return(x);
			}

			function CreateItemsString(List)
			{
				var tmpListStr="";
				var hasItems = false;

				if (List!="")
				{
					var tmp1=List.split(",");
					for (i=0;i<tmp1.length;i++)
					{
						if (tmp1[i]!="")
						{
							if (tmp1[i] == 'PrdFreeShip') {
								tmp1[i] = 'PrdNoShip';
							}
							tmpListStr=tmpListStr + "<li id='" + tmp1[i] + "'>" + findItemName(tmp1[i]) + "<a href='#' title='Remove Item' onclick='javascript:RemoveItem($(this).parent()); return false;'><span class='glyphicon glyphicon-remove'></span></a></li>";
							hasItems = true;
						}
					}
				}
				if (!hasItems) {
					tmpListStr=tmpListStr + "<%= noElementsHTML %>";
				}

				return tmpListStr;
			}

			function CreateItemLists()
			{
				$pc("#TT").html(CreateItemsString(TopTop));
				$pc("#TL").html(CreateItemsString(TopLeft));
				$pc("#TR").html(CreateItemsString(TopRight));
				$pc("#ML").html(CreateItemsString(Middle));
				$pc("#BL").html(CreateItemsString(BottomL));
			}

			function RemoveItem(item)
			{
				if (!confirm("Are you sure you want to remove this element?")) {
					return;
				}

				// Get list the item is in
				var list = null;
				var elementID = item.attr("id");
				switch (item.parent().attr("id")) 
				{
					case "TT":
						TopTop=TopTop.replace(elementID+",","");
						break;
					case "TL":
						TopLeft=TopLeft.replace(elementID+",","");
						break;
					case "TR":
						TopRight=TopRight.replace(elementID+",","");
						break;
					case "ML":
						Middle=Middle.replace(elementID+",","");
						break;
					case "BL":
						BottomL=BottomL.replace(elementID+",","");
						break;
				}

				SavePPToFields();
				CreateItemLists();
			}

			function handleConnectedSortable()
			{
				TopTop = getSortableResults($pc("#TT"));
				TopLeft = getSortableResults($pc("#TL"));
				TopRight = getSortableResults($pc("#TR"));
				Middle = getSortableResults($pc("#ML"));
				BottomL = getSortableResults($pc("#BL"));
			}

			function getSortableResults(list) 
			{
				ListSerial = list.sortable( "serialize" );
				var ListCount=ListSerial.get().length;
				if (ListCount>0 && !(ListSerial.get()[0] instanceof Object))
				{
					results=ListSerial.get().join()+",";

					list.find(".notSortable").remove();
				}
				else
				{
					results="";
					if (list.html().length < 1) {
						list.html("<%= noElementsHTML %>");
					}
				}

				return results;
			}
	
			var PTabs;
			var tabSortableEnabled = true;

			function TabShowEditable(tab)
			{
				tab.editable("show");
		
				// Also disable sortable
				$pc("#TabsArea .nav-tabs").sortable('disable');
		
				tabSortableEnabled = false;
			}
	
			function CreatePrdTabs()
			{
				document.getElementById("TabsArea").innerHTML="";

				var tmpTabTitle="<ul class='nav nav-tabs'>";
				var tmpTabContent="";
				var activeTab = "active";
			 
				for(var i = 0; i < PTCount; i++)
				{
					tmpTabTitle=tmpTabTitle + "<li class=\"" + activeTab + "\" id='" + PrdTabs[i][0] + "'>";
					tmpTabTitle=tmpTabTitle + "<a href='#" + PrdTabs[i][0] + "Tab' data-toggle='tab'><span class='TabName' data-title='Enter Tab Name' type='text' >" + PrdTabs[i][1] + "</span></a>";
					tmpTabTitle=tmpTabTitle + "<span onclick='javascript:RemovePrdTab(\"" + PrdTabs[i][0] + "\");' title='Remove this Tab' class='TabRemove glyphicon glyphicon-remove-circle'></span>";
					tmpTabTitle=tmpTabTitle + "</li>";

					tmpTabContent=tmpTabContent + "<div id='" + PrdTabs[i][0] + "Tab' class=\"tab-pane " + activeTab + "\">";
					tmpTabContent=tmpTabContent + "<div class='pcCPsectionTitle'>Tab's Elements</div>";
					tmpTabContent=tmpTabContent + "<div id='" + PrdTabs[i][0] + "List' style='clear:both'></div>";
					tmpTabContent=tmpTabContent + "<div style='clear:both'></div>";
					tmpTabContent=tmpTabContent + "</div>";
					activeTab = '';
				}
				
				tabElementToolbox = "";
				tabElementToolbox=tabElementToolbox + "<div id='TabElementToolbox'>";
				tabElementToolbox=tabElementToolbox + "<div class='pcCPsectionTitle'>Add new element</div>";
				tabElementToolbox=tabElementToolbox + "<%=tmpEleList%>";
				tabElementToolbox=tabElementToolbox + "</div>";

				tmpTabTitle=tmpTabTitle+"<li class='TabAdd'><a href='#' onclick='addNewTab(); return false;' title='Add New Tab'><span class='glyphicon glyphicon-plus'></span></a></li>";
				tmpTabTitle=tmpTabTitle+"</ul>"
				document.getElementById("TabsArea").innerHTML="<div id='TabbedPanels2' class=\"tabbable\">" + tmpTabTitle + "<div class=\"tab-content\">" + tmpTabContent + "</div></div>" + tabElementToolbox;

				for(var i = 0; i < PTCount; i++)
				{
					CreateTabList(i,PrdTabs[i][0],PrdTabs[i][0]+"List",PrdTabs[i][2]);
				}

				$pc(".TabName").editable({
					toggle: 'manual'
				}).on("hidden", function(e, reason) {
					$pc("#TabsArea .nav-tabs").sortable('enable');

					tabSortableEnabled = true;

					if (reason == 'save') {
						var tabName = $pc(this).parent().parent().attr('id');

						PrdTabs[findTabIndex(tabName)][1] = $pc(this).html();
					}
				});

				$pc("#TabsArea").on("click", ".nav-tabs li.active .TabName", function(e) {
					TabShowEditable($pc(this));
				}).on('mouseover', '.nav-tabs li a', function(e) {
					if (tabSortableEnabled && !$pc(this).parent().hasClass("TabAdd")) {
						$pc(this).css("cursor", "move");
					} else {
						$pc(this).css("cursor", "pointer");
					}
				});

				PTabs = $pc( "#TabbedPanels2" ).tab('show');
				$pc("#TabsArea .nav-tabs").sortable({
					distance: 10,
					vertical: false,
					exclude: '.TabAdd span, .TabAdd a',
					serialize: function (parent, children, isContainer) {
						return isContainer ? children.join() : parent.attr("id");
					},
					onDrop: function(item, container, _super) {
						_super(item, container);
        
						var tmpTabArr = $pc( "#TabsArea .nav-tabs" ).sortable( "serialize" ).get();
						if (!(tmpTabArr[0] instanceof Object)) {
							tmpTabSort=tmpTabArr.join()+",";
							var tmp1=tmpTabSort.split(",");
							var tmpTabs=new Array();
							var tmpCount=-1;
							for (i=0;i<tmp1.length;i++)
							{
								for (j=0;j<PTCount;j++)
								{
									if ((tmp1[i]!="") && (tmp1[i]==PrdTabs[j][0]))
									{
										tmpCount=tmpCount+1;
										tmpTabs[tmpCount]=new Array();
										tmpTabs[tmpCount][0]=PrdTabs[j][0];
										tmpTabs[tmpCount][1]=PrdTabs[j][1];
										tmpTabs[tmpCount][2]=PrdTabs[j][2];
										tmpTabs[tmpCount][3]=PrdTabs[j][3];
						
									}
								}
							}
							PrdTabs=tmpTabs;
						}
					}
				});

				if (PTCount < 1) {
					$pc("#TabElementToolbox").hide();
					$pc("#TabsArea").append("<%= noTabsHTML %>");
				} else {
					$pc("#TabElementToolbox").show();
				}
				
				$pc("#TabsArea .nav-tabs a").bind('click', function() {
					var tabName = $pc(this).parent().attr("id");
					var button = $pc("button[value='CUSTOMHTML']");

					if ($pc("#" + tabName + "Tab").find("#CUSTOMHTML").length < 1) {
						button.attr("disabled", false);
					} else {
						button.attr("disabled", true);
					}
				});
				PrdTabsSorted=PrdTabs;
			}
	
			function RemovePrdTab(tabName)
			{
				if (!confirm("Are you sure you want to remove this tab and all its elements? WARNING: Any Custom HTML Elements contained in this tab will be deleted, and the contents cannot be recovered later! ")) {
					return;
				}

				var curTab = $pc("#TabsArea li.active").attr("id");

				for(var i = 0; i < PTCount; i++)
				{
					if (PrdTabs[i][0]==tabName)
					{
						for(var j = i+1; j < PTCount; j++)
						{
							PrdTabs[j-1][0]=PrdTabs[j][0];
							PrdTabs[j-1][1]=PrdTabs[j][1];
							PrdTabs[j-1][2]=PrdTabs[j][2];
							PrdTabs[j-1][3]=PrdTabs[j][3];
						}
						PTCount=PTCount-1;
						break;
					}
				}
				SavePPToFields();
				CreatePrdTabs();

				$pc("#" + curTab).find("a").tab("show");
			}
	
			function findTabIndex(tabName)
			{
				for(var i = 0; i < PTCount; i++)
				{
					if (PrdTabs[i][0]==tabName)
					{
						return(i);
						break;
					}
				}
			}
	
			function showCusHTMLElement(item) 
			{
				var cusHTMLDiv = $pc(item).parent().find(".CusHTMLDiv");

				if ($pc(item).html() == "Edit") {
					cusHTMLDiv.slideDown();
					$pc(item).html("Hide");
				} else {
					cusHTMLDiv.slideUp();
					$pc(item).html("Edit");
				}
			}

			function CreateTabList(tmpindex,tabName,listName,listItems)
			{
				var tmpHasCusHTML=false;
				var tmpListStr="";

				tmpListStr=tmpListStr + "<ul id='" + listName + "Show' class='TabList connectedSortable sortable'>";

				if (listItems!="")
				{
					var tmp1=listItems.split(",");
					for (i=0;i<tmp1.length;i++)
					{
						if (tmp1[i]!="")
						{
							if (tmp1[i] == 'PrdFreeShip') {
								tmp1[i] = 'PrdNoShip';
							}
							tmpListStr=tmpListStr + "<li id='" + tmp1[i] + "'>";
							tmpListStr=tmpListStr + findItemName(tmp1[i]);
							tmpListStr=tmpListStr + "<a href='#' title='Remove Item' onclick='javascript:RemoveFromTab(" + tmpindex + ", \"" + tmp1[i] + "\"); return false;'><span class='glyphicon glyphicon-remove'></span></a>";
							if (tmp1[i] == "CUSTOMHTML") {
								tmpListStr=tmpListStr + "<a class='CusHTMLEdit' href='#' onclick='showCusHTMLElement(this); return false;'>Edit</a>";

								tmpListStr=tmpListStr + "<div class='CusHTMLDiv'>";
								tmpListStr=tmpListStr + "<div id='CusHTMLContainer" + tmpindex + "'></div>";
								tmpListStr=tmpListStr + "<textarea cols='80' rows='7' id='CusHTML" + tmpindex + "' name='CusHTML" + tmpindex + "'>" + PrdTabs[tmpindex][3] + "</textarea>";
								tmpListStr=tmpListStr + "</div>";

								tmpHasCusHTML = true;
							}
							tmpListStr=tmpListStr + "</li>";
						}
					}
				}
				tmpListStr=tmpListStr + "</ul>";
				document.getElementById(listName).innerHTML=tmpListStr;

				$pc("#" + tabName + "ListShow li").each(function() {
					var elementID = $pc(this).attr("id");
					var button = $pc("#TabElementToolbox").find("button[value='" + elementID + "']");
					if (button.val() == "CUSTOMHTML") {
						var tabName = $pc("#TabList li.active").attr("id");
						if ($pc("#" + tabName + "Tab").find("#CUSTOMHTML").length < 1) {
							button.attr("disabled", false);
						} else {
							button.attr("disabled", true);
						}
					} else {
						button.attr('disabled', true);
					}
				});
		
				var tmpList = $pc("#" + listName + "Show");
				
				// Re-initialize sortable
				initConnectedSortable();

				if (tmpHasCusHTML) {
					htmleditorjs("CusHTMLEditor" + tmpindex, "CusHTML" + tmpindex, "CusHTMLContainer" + tmpindex);
				}

				if (listItems=="")
				{
					tmpList.append("<%= noElementsHTML %>");
				}
			}

			function RemoveFromTab(tmpindex,tmpItem)
			{
				if (tmpItem != "CUSTOMHTML") {
					if (!confirm("Are you sure you want to remove this element?")) {
						return;
					}
				} else {
					if (!confirm("Are you sure you want to remove this Custom HTML Element? WARNING: If you remove this element all the contents will be completely deleted and cannot be recovered later!")) {
						return;
					}
				}

				var removeItem = $pc("#" + tmpItem).closest('li');
				var btnItem = $pc("button[value='" + removeItem.attr('id') + "']");

				var tmplistItems=PrdTabs[tmpindex][2];
				tmplistItems=tmplistItems.replace(removeItem.attr('id')+",","");
				PrdTabs[tmpindex][2]=tmplistItems;
				removeItem.remove();
				
				btnItem.attr('disabled', false);

				SavePPToFields();
				CreateTabList(tmpindex,PrdTabs[tmpindex][0],PrdTabs[tmpindex][0]+"List",PrdTabs[tmpindex][2]);
			}
	
			function RemoveSelectedFromTab(tmpindex,tmpCheck)
			{
				var tmp1=0;
				var tmplistItems=PrdTabs[tmpindex][2];
				var inputs = document.getElementsByTagName("input");
				for(var i = 0; i < inputs.length; i++) {
					if ((inputs[i].type == "checkbox") && (inputs[i].name.indexOf(tmpCheck)==0) && (inputs[i].checked)) {
						tmplistItems=tmplistItems.replace(inputs[i].value+",","");
						if (inputs[i].value=="CUSTOMHTML") document.getElementById(PrdTabs[tmpindex][0] + "CB").checked=false;
						PrdTabs[tmpindex][2]=tmplistItems;
						tmp1=1;
					}  
				}
				if (tmp1==1)
				{
					SavePPToFields();
					CreateTabList(tmpindex,PrdTabs[tmpindex][0],PrdTabs[tmpindex][0]+"List",PrdTabs[tmpindex][2]);
				}
			}
	
			function ShowRemoveBtnTab(tmpBtnArea,tmpCheck)
			{
				var inputs = document.getElementsByTagName("input");
				for(var i = 0; i < inputs.length; i++) {
					if ((inputs[i].type == "checkbox") && (inputs[i].name.indexOf(tmpCheck)==0) && (inputs[i].checked)) {
						document.getElementById(tmpBtnArea).style.display="";
						return(true);
					}  
				}
				document.getElementById(tmpBtnArea).style.display='none';
			}
	
			function addElementToList()
			{
				tmp1=document.getElementById("PrdEPos").value;
				switch (tmp1)
				{
					case "0":
						TopTop=TopTop+document.getElementById("PrdEle").value+",";
						break;
					case "1":
						TopLeft=TopLeft+document.getElementById("PrdEle").value+",";
						break;
					case "2":
						TopRight=TopRight+document.getElementById("PrdEle").value+",";
						break;
					case "3":
						Middle=Middle+document.getElementById("PrdEle").value+",";
						break;
					case "4":
						BottomL =BottomL+document.getElementById("PrdEle").value+",";
						break;
				}
				
				SavePPToFields();
				CreateItemLists();
			}

			function addElementToTabList(tmpindex, value) 
			{
				PrdTabs[tmpindex][2]=PrdTabs[tmpindex][2]+value+",";
				SavePPToFields();
				CreateTabList(tmpindex,PrdTabs[tmpindex][0],PrdTabs[tmpindex][0]+"List",PrdTabs[tmpindex][2]);

				// Automatically show HTML editor when adding
				if (value == "CUSTOMHTML") {
					$pc("#" + PrdTabs[tmpindex][0] + "List").find(".CusHTMLEdit").click();
				}
			}

			$pc(document).ready(function() {
				var customPrdLayout = $pc(".CustomPrdLayout,.CustomTabsLayout");

				$pc("#displayLayout").change(function () {
					if ($pc(this).val() == "t") {
						$pc("#customizeButtons").hide();
						customPrdLayout.show();

						if (TopTop == "" && TopLeft == "" && TopRight == "" && Middle == "" && BottomL == "") {
							LoadDefaultPageSettings();
						}
					} else {
						$pc("#customizeButtons").show();
						customPrdLayout.hide();
					}
				});

				$pc("#customizeLayout").click(function(e) {
					var displayLayout = $pc("#displayLayout");
					var selectedLayout = displayLayout.val();
					var layoutName = "'" + displayLayout.find(":selected").text() + "'";

					if (selectedLayout == '') {
						layoutName = "default";
					}

					if (confirm("Are you sure you want to customize the " + layoutName + " layout for this product? WARNING: If you have setup a previous custom layout for this product it will be overwritten! Would you like to continue?")) {
						customPrdLayout.show();
						displayLayout.val('t');
						$pc("#customizeButtons").hide();

						LoadSettingsForLayout(selectedLayout, false);
					}

					e.preventDefault();
				});

				$pc("#addTabsToLayout").click(function(e) {
					var displayLayout = $pc("#displayLayout");
					var selectedLayout = displayLayout.val();
					var layoutName = "'" + displayLayout.find(":selected").text() + "'";
					
					if (selectedLayout == '') {
						layoutName = "default";
					}

					if (confirm("Are you sure you want customize and add tabs to the " + layoutName + " layout for this product? WARNING: If you have setup a previous custom layout for this product it will be overwritten! Would you like to continue?")) {
						customPrdLayout.show();
						displayLayout.val('t');
						$pc("#customizeButtons").hide();

						LoadSettingsForLayout(selectedLayout, true);
					}

					e.preventDefault();
				});


				$pc(document).on("click", ".OptBtn", function(e) {
					var tabName = $pc("#TabsArea .nav-tabs li.active").attr("id");
					var tabIndex = findTabIndex(tabName);
					
					// Disable this button
					$pc(this).attr('disabled', true);

					// Add to list!
					addElementToTabList(tabIndex, $pc(this).val());

					e.preventDefault();
				});

				initConnectedSortable();
			});

			function initConnectedSortable()
			{
				$pc( ".connectedSortable" ).sortable({
					distance: 10,
					group: "connectedSortable",
					exclude: ".notSortable div",
					serialize: function (parent, children, isContainer) 
					{
						return isContainer ? children.join() : parent.attr("id");
					},
					onDragStart: function($item, $container, _super, event)
					{
						_super($item, $container);

						var parent = $item.parent();
						if (parent.hasClass("TabList")) {
							var index = findTabIndex(parent.attr("id").replace("ListShow", ""));
							UpdateCustomHtmlContent(index);
						}
						
						if ($item.attr("id") == "CUSTOMHTML") {
							$("#TT,#TL,#TR,#ML,#BL").sortable("disable");
						}
					},
					onDrop: function($item, $container, _super)
					{
						_super($item, $container);
						
						var parent = $item.parent();
						if (parent.hasClass("TabList")) {
							var index = findTabIndex(parent.attr("id").replace("ListShow", ""));
							PrdTabs[index][2] = getSortableResults(parent);
							if ($item.attr("id") == "CUSTOMHTML") {
								htmleditorjs("CusHTMLEditor" + index, "CusHTML" + index, "CusHTMLContainer" + index);
							}
						}

						if ($item.attr("id") == "CUSTOMHTML") {
							$("#TT,#TL,#TR,#ML,#BL").sortable("enable");
						}

						handleConnectedSortable();
					}
				});
			}

			function addNewTab()
			{
				SavePPToFields();
				newTabName="Tab " + (PTCount + 1);
				PTCount=PTCount+1;
				OTCount=OTCount+1
				PrdTabs[PTCount-1]=new Array();
				PrdTabs[PTCount-1][0]="Tab"+OTCount;
				PrdTabs[PTCount-1][1]=newTabName;
				PrdTabs[PTCount-1][2]="";
				PrdTabs[PTCount-1][3]="";
				CreatePrdTabs();
				SavePPToFields();

				$pc("#Tab" + OTCount).find("a").tab("show");
				TabShowEditable($pc("#Tab" + OTCount).find(".TabName"));
			}
	
			function GenWorkingArea()
			{
				document.getElementById("TabWorking").style.display='';
				CreateItemLists();
				CreatePrdTabs();
			}

			function UpdateCustomHtmlContent(index)
			{
				var editorContent = "";

				if ($("#idAreaCusHTMLEditor" + index).length > 0) 
				{
					editorContent = htmleditorcontent("CusHTMLEditor" + index);
					$pc("#CusHTML" + index).val(editorContent);
				}

				if ($pc("#CusHTML" + index).length > 0) 
				{
					PrdTabs[findTabIndex(PrdTabsSorted[index][0])][3]= $pc("#CusHTML" + index).val();
				}
			}
	
			function SavePPToFields()
			{	
				document.hForm.ppTop.value=TopTop;
				document.hForm.ppTopLeft.value=TopLeft;
				document.hForm.ppTopRight.value=TopRight;
				document.hForm.ppMiddle.value=Middle;
				document.hForm.ppBottom.value=BottomL;
				var tmp1="";
				for(var i = 0; i < PTCount; i++)
				{
					UpdateCustomHtmlContent(i);

					// Automatically pull the content from the HTML editor if we're using it
					//if (PrdTabs[i][2].indexOf("CUSTOMHTML") > 0 && $pc("#CusHTMLEditor" + i).length > 0) 
					//{
					//	$pc("#CusHTML" + i).val($pc("#CusHTMLEditor" + i).getHTMLBody());
					//}

					// Add content from custom HTML input if it exists
					//if ($pc("#CusHTML" + i).length > 0) 
					//{
					//	PrdTabs[findTabIndex(PrdTabs[i][0])][3]= $pc("#CusHTML" + i).val();
					//}

			
					tmp1=tmp1 + PrdTabs[i][1] + "``" + PrdTabs[i][2] + "``" + PrdTabs[i][3] + "||";
				}
				document.hForm.ppTabs.value=tmp1;
			}
	
			GenWorkingArea();
		
		</script>
	</td>
</tr>
