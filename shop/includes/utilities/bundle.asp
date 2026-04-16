<!--#include file="../common.asp"--> 
<%
Dim strArray
Dim strItem
Dim pcv_strJS
pcv_strJS = ""


'/////////////////////////////////////////////////////////////
'// START: CSS FOLDER
'/////////////////////////////////////////////////////////////
If Request("action")="CSS" OR Request("action")="ALL" Then

    pcv_boolMinifyCSS = True
    
    Set pcd_CSSBundles = Server.CreateObject("Scripting.Dictionary")
    
    '// Storefront
    pcd_CSSBundles.Add "animate.css", "../../pc/css/"
    pcd_CSSBundles.Add "bootstrap.css", "../../pc/css/"
    pcd_CSSBundles.Add "configuratorPricingBox.css", "../../pc/css/"
    pcd_CSSBundles.Add "datepicker3.css", "../../pc/css/"
    pcd_CSSBundles.Add "onepagecheckout.css", "../../pc/css/"
    pcd_CSSBundles.Add "pcBTO.css", "../../pc/css/"
    pcd_CSSBundles.Add "pcCM.css", "../../pc/css/"
    pcd_CSSBundles.Add "pcSearchFields.css", "../../pc/css/"
    pcd_CSSBundles.Add "pcStorefront.css", "../../pc/css/"
    pcd_CSSBundles.Add "quickview.css", "../../pc/css/"
    pcd_CSSBundles.Add "screen.css", "../../pc/css/"
    pcd_CSSBundles.Add "mojozoom.css", "../mojozoom/"
	pcd_CSSBundles.Add "awesome.css", "../../htmleditor/scripts/style/"
    
    '// Control Panel
    'pcd_CSSBundles.Add "pcPrint.css", "../../pc/css/"
    
    '// Other
    'pcd_CSSBundles.Add "pcSyndication.css", "../../pc/css/"
    pcd_CSSBundles.Add "search.css", "../../pc/css/"
    pcd_CSSBundles.Add "slidebars.css", "../../pc/css/"
    
    pcv_strAllFiles = pcd_CSSBundles.Keys   
    pcv_strAllPaths = pcd_CSSBundles.Items 
    
    For i = 0 To pcd_CSSBundles.Count - 1 
    
        pcv_strDataMinified = ""
        pcv_strData = ""
        pcv_strPath = pcv_strAllPaths(i)  
        pcv_strFilename = pcv_strAllFiles(i)  
        pcv_strFilenameMin = replace(pcv_strFilename,".css",".min.css")   
        pcv_strData = pcf_OpenUTF8(pcv_strPath & pcv_strFilename, pcv_strPath & pcv_strFilename)
        
        If pcv_boolMinifyCSS Then        
            pcv_strDataMinified = CompileCSS(pcv_strData)
            If len(pcv_strDataMinified)>0 Then
                call pcs_SaveUTF8(pcv_strPath & pcv_strFilenameMin, pcv_strPath & pcv_strFilenameMin, pcv_strDataMinified)
                pcv_strCSS = pcv_strCSS & pcv_strDataMinified
            Else
                pcv_strCSS = pcv_strCSS & pcv_strData
            End If            
        Else
            pcv_strCSS = pcv_strCSS & pcv_strData
        End If
    
    Next
    
    If len(pcv_strCSS)>0 Then
        call pcs_SaveUTF8("../../pc/css/combined.min.css", "../../pc/css/combined.min.css", pcv_strCSS)
    End If

End If
'/////////////////////////////////////////////////////////////
'// END: CSS FOLDER
'/////////////////////////////////////////////////////////////





'/////////////////////////////////////////////////////////////
'// START: JS FOLDER
'/////////////////////////////////////////////////////////////
If Request("action")="JS" OR Request("action")="ALL" Then

    pcv_boolMinifyJS = True
    
    Set pcd_JSBundles = Server.CreateObject("Scripting.Dictionary")

    '// Storefront - previously in header
    pcd_JSBundles.Add "jquery.validate.js", "../jquery/"
    pcd_JSBundles.Add "jquery.form.js", "../jquery/"
    pcd_JSBundles.Add "jquery.touchSwipe.js", "../jquery/"
    
    pcd_JSBundles.Add "ddsmoothmenu.js", "../jquery/smoothmenu/"
    pcd_JSBundles.Add "jquery.nivo.slider.pack.js", "../jquery/nivo-slider/"
    
    pcd_JSBundles.Add "bootstrap.js", "../javascripts/"
    pcd_JSBundles.Add "bootstrap-tabcollapse.js", "../javascripts/"
    pcd_JSBundles.Add "bootstrap-datepicker.js", "../javascripts/"
    
    '// Storefront - previously in footer
    pcd_JSBundles.Add "highslide.html.unpacked.js", "../javascripts/"
    pcd_JSBundles.Add "jquery.blockUI.js", "../javascripts/"
    pcd_JSBundles.Add "json3.js", "../javascripts/"
    pcd_JSBundles.Add "accounting.js", "../javascripts/"
    pcd_JSBundles.Add "productcart.js", "../javascripts/"
    
    pcd_JSBundles.Add "service.js", "../../pc/service/app/"
    pcd_JSBundles.Add "quickcart.js", "../../pc/service/app/"
    pcd_JSBundles.Add "viewcart.js", "../../pc/service/app/"
    pcd_JSBundles.Add "search.js", "../../pc/service/app/"
    pcd_JSBundles.Add "onepagecheckout.js", "../../pc/service/app/"
    pcd_JSBundles.Add "order.js", "../../pc/service/app/"

    '// Storefront - Other
    pcd_JSBundles.Add "mojozoom.js", "../mojozoom/"
    pcd_JSBundles.Add "opc_validation.js", "../javascripts/"

    pcv_strAllFiles = pcd_JSBundles.Keys   
    pcv_strAllPaths = pcd_JSBundles.Items 
    
    For i = 0 To pcd_JSBundles.Count - 1 
    
        pcv_strDataMinified = ""
        pcv_strData = ""
        pcv_strPath = pcv_strAllPaths(i)  
        pcv_strFilename = pcv_strAllFiles(i)  
        pcv_strFilenameMin = replace(pcv_strFilename,".js",".min.js")   
        pcv_strData = pcf_OpenUTF8(pcv_strPath & pcv_strFilename, pcv_strPath & pcv_strFilename)

        If pcv_boolMinifyJS Then        
            pcv_strDataMinified = CompileJS(pcv_strData)
            If len(pcv_strDataMinified)>0 Then
                call pcs_SaveUTF8(pcv_strPath & pcv_strFilenameMin, pcv_strPath & pcv_strFilenameMin, pcv_strDataMinified)
                pcv_strJS = pcv_strJS & pcv_strDataMinified
            Else
                pcv_strJS = pcv_strJS & pcv_strData
            End If
        Else
            pcv_strJS = pcv_strJS & pcv_strData
        End If
    
    Next
    
    If len(pcv_strJS)>0 Then
        pcv_strJS = replace(pcv_strJS, vbCrLf,"")
        pcv_strJS = replace(pcv_strJS, vbCr,"")
        pcv_strJS = replace(pcv_strJS, vbLf,"")
        call pcs_SaveUTF8("../../pc/js/combined.min.js", "../../pc/js/combined.min.js", pcv_strJS)
    End If
    
End If
'/////////////////////////////////////////////////////////////
'// END: JS FOLDER
'/////////////////////////////////////////////////////////////





'/////////////////////////////////////////////////////////////
'// START: COMBINE ALL
'/////////////////////////////////////////////////////////////


response.End()

'/////////////////////////////////////////////////////////////
'// END: COMBINE ALL
'/////////////////////////////////////////////////////////////





'/////////////////////////////////////////////////////////////
'// START: COMPILE
'/////////////////////////////////////////////////////////////

Public Function CompileCSS(str)

    str = pcf_PostForm("input=" & Server.URLEncode(str), "https://cssminifier.com/raw", "")

    CompileCSS = str
    
End Function

Public Function CompileJS(str)

    str = pcf_PostForm("js_code=" & Server.URLEncode(str) & "&output_format=xml&output_info=compiled_code&compilation_level=SIMPLE_OPTIMIZATIONS", "https://closure-compiler.appspot.com/compile", "")

    CompileJS = pcf_GetNode(str, "compiledCode", "//compilationResult")
    
End Function

Function pcf_GetNode(responseXML, nodeName, nodeParent)
    Set myXmlDoc = Server.CreateObject("Msxml2.DOMDocument"&scXML)				 
    myXmlDoc.loadXml(responseXML)
    Set Nodes = myXmlDoc.selectnodes(nodeParent)	
    For Each Node In Nodes	
        pcf_GetNode = pcf_CheckNode(Node,nodeName,"")				
    Next
    Set Node = Nothing
    Set Nodes = Nothing
    Set myXmlDoc = Nothing
End Function

Function pcf_CheckNode(Node,tagName,default)		
    Dim tmpNode
    Set tmpNode=Node.selectSingleNode(tagName)
    If tmpNode is Nothing Then
        pcf_CheckNode=default
    Else
        pcf_CheckNode=Node.selectSingleNode(tagName).text
    End if
End Function

'/////////////////////////////////////////////////////////////
'// END: COMPILE
'/////////////////////////////////////////////////////////////
%>