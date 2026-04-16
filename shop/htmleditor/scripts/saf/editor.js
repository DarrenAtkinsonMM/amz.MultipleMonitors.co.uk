/***********************************************************
InnovaStudio Live Editor 1.7.3 - Saf
© 2014, InnovaStudio (www.innovastudio.com). All rights reserved.
************************************************************/
var editor = new Array();
var oUtil = new InnovaEditorUtil();

function InnovaEditorUtil() {
    this.langDir = "en-US";
    try {
        if (LanguageDirectory) this.langDir = LanguageDirectory
    } catch (e) {}
    var oScripts = document.getElementsByTagName("script");
    for (var i = 0; i < oScripts.length; i++) {
        var sSrc = oScripts[i].src.toLowerCase();
        if (sSrc.indexOf("saf/editor.js") != -1) {
            this.scriptPath = oScripts[i].src.replace(/editor.js/, "")
        } else if (sSrc.indexOf("scripts/innovaeditor.js") != -1) {
            if (!this.scriptPath) this.scriptPath = oScripts[i].src.replace(/innovaeditor.js/, "") + "saf/"
        }
    }
    this.scriptPathLang = this.scriptPath.replace(/\/saf/, "") + "language/" + this.langDir + "/";
    if (this.langDir == "en-US") document.write("<scr" + "ipt src='" + this.scriptPathLang + "editor_lang.js'></scr" + "ipt>");
    this.oName;
    this.oEditor;
    this.obj;
    this.oSel;
    this.sType;
    this.bInside = bInside;
    this.useSelection = true;
    this.arrEditor = [];
    this.onSelectionChanged = function () {
        return true
    };
    this.activeElement;
    this.activeModalWin;
    this.setEdit = setEdit;
    this.bOnLoadReplaced = false;
    this.Table;
    this.protocol = window.location.protocol;
    this.spcCharCode = [
        [169, "&copy;"],
        [163, "&pound;"],
        [174, "&reg;"],
        [233, "&eacute;"],
        [201, "&Eacute;"],
        [8364, "&euro;"],
        [8220, "\""]
    ];
    this.spcChar = [];
    this.loadSpecialCharCode = function (spcCharCodes) {
        if (spcCharCodes != null) {
            this.spcCharCode = spcCharCodes;
            this.spcChar = []
        }
        for (var i = 0; i < this.spcCharCode.length; i++) {
            this.spcChar[i] = [new RegExp(String.fromCharCode(this.spcCharCode[i][0]), "g"), this.spcCharCode[i][1]]
        }
    };
    for (var i = 0; i < this.spcCharCode.length; i++) {
        this.spcChar[i] = [new RegExp(String.fromCharCode(this.spcCharCode[i][0]), "g"), this.spcCharCode[i][1]]
    }
    this.replaceSpecialChar = function (sHTML) {
        for (var i = 0; i < this.spcChar.length; i++) {
            sHTML = sHTML.replace(this.spcChar[i][0], this.spcChar[i][1])
        }
        sHTML = sHTML.replace(/class="Apple-style-span"/gi, "");
        return sHTML
    };
    this.initializeEditor = function (tselector, opt) {
        var allText = [],
            txt, edtCnt;
        if (typeof (tselector) == "object" && tselector.tagName && tselector.tagName == "TEXTAREA") {
            allText[0] = tselector
        } else if (tselector.substr(0, 1) == "#") {
            txt = document.getElementById(tselector.substr(1));
            if (!txt) return;
            allText[0] = txt
        } else {
            var all = document.getElementsByTagName("TEXTAREA");
            for (var i = 0; i < all.length; i++) {
                if (all[i].className == tselector) {
                    allText[allText.length] = all[i]
                }
            }
        }
        for (var i = 0; i < allText.length; i++) {
            txt = allText[i];
            if (txt.id || txt.id == "") txt.id = "editorarea" + i;
            edtCnt = document.createElement("DIV");
            edtCnt.id = "innovaeditor" + i;
            txt.parentNode.insertBefore(edtCnt, txt);
            window["oEdit" + i] = new InnovaEditor("oEdit" + i);
            var objStyle;
            if (window.getComputedStyle) {
                objStyle = window.getComputedStyle(txt, null)
            } else if (txt.currentStyle) {
                objStyle = txt.currentStyle
            } else {
                objStyle = {
                    width: window["oEdit" + i].width,
                    height: window["oEdit" + i].height
                }
            }
            window["oEdit" + i].width = objStyle.width;
            window["oEdit" + i].height = objStyle.height;
            if (opt) {
                for (var it in opt) {
                    window["oEdit" + i][it] = opt[it]
                }
            }
            window["oEdit" + i].REPLACE(txt.id, "innovaeditor" + i)
        }
    }
};

function bInside(oElement) {
    while (oElement != null) {
        if (oElement.designMode && oElement.designMode == "on") return true;
        oElement = oElement.parentNode
    }
    return false
};

function checkFocus() {
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    var oSel = oEditor.getSelection();
    var parent = getSelectedElement(oSel);
    if (parent != null) {
        if (!bInside(parent)) return false
    } else {
        if (!bInside(parent)) return false
    }
    return true
};

function iwe_focus(type) {
    var oEditor = document.getElementById("idContent" + this.oName);
    if (oEditor) {
        oEditor = oEditor.contentWindow;
        oEditor.focus()
    }
    if (type && oEditor) {
        var edtWin;
        var range = oEditor.document.createRange();
        if (oEditor.document.body.childNodes.length > 0) {
            var nd = null;
            if (type == "start") {
                nd = oEditor.document.body.childNodes[0];
                range.selectNode(nd);
                range.collapse(true)
            } else {
                nd = oEditor.document.body.childNodes[oEditor.document.body.childNodes.length - 1];
                range.selectNode(nd);
                if (nd.nodeType == 1 && nd.tagName == "BR") {
                    range.collapse(true)
                } else {
                    range.collapse(false)
                }
            }
        } else {
            range.selectNodeContents(oEditor.document.body);
            range.collapse(false)
        }
        var sel = oEditor.getSelection();
        sel.removeAllRanges();
        sel.addRange(range);
        var evt = document.createEvent('TextEvent');
        evt.initTextEvent('textInput', true, true, window, ' ');
        oEditor.document.dispatchEvent(evt);
        range = sel.getRangeAt(0);
        range.setStart(range.startContainer, range.startOffset - 1);
        range.deleteContents();
        sel.removeAllRanges();
        sel.addRange(range)
    }
};

function setFocus(type) {
    this.focus(type)
};

function setEdit(oName) {
    if ((oName != null) && (oName != "")) {
        try {
            var wnd = document.getElementById("idContent" + oName);
            wnd.contentDocument.body.contentEditable = true;
            wnd.focus()
        } catch (e) {}
    } else {
        for (var i = 0; i < this.arrEditor.length; i++) {
            try {
                var wnd = document.getElementById("idContent" + this.arrEditor[i]).contentWindow;
                wnd.document.body.contentEditable = true;
                var r = wnd.getSelection().getRangeAt(0);
                r.selectNode(wnd.document.getElementsByTagName("Body")[0]);
                r.collapse(true);
                wnd.focus()
            } catch (e) {}
        }
    }
};
var iconHeight;

function InnovaEditor(oName) {
    this.oName = oName;
    this.init = initISEditor;
    this.RENDER = RENDER;
    this.onRender = function () {};
    this.loadHTML = loadHTML;
    this.loadHTMLFull = loadHTMLFull;
    this.getHTMLBody = getHTMLBody;
    this.getHTML = getHTML;
    this.getXHTMLBody = getXHTMLBody;
    this.getXHTML = getXHTML;
    this.getTextBody = getTextBody;
    this.putHTML = putHTML;
    this.css = "";
    this.onKeyPress = function (evt) {
        return true
    };
    this.styleSelectionHoverBg = "#cccccc";
    this.styleSelectionHoverFg = "white";
    this.styleSelectorPrefix = "";
    this.cleanEmptySpan = cleanEmptySpan;
    this.cleanFonts = cleanFonts;
    this.cleanTags = cleanTags;
    this.replaceTags = replaceTags;
    this.cleanDeprecated = cleanDeprecated;
    this.doClean = doClean;
    this.applySpanStyle = applySpanStyle;
    this.bInside = bInside;
    this.checkFocus = checkFocus;
    this.focus = iwe_focus;
    this.setFocus = setFocus;
    this.disableFocusOnLoad = false;
    this.doCmd = doCmd;
    this.applyParagraph = applyParagraph;
    this.applyFontName = applyFontName;
    this.applyFontSize = applyFontSize;
    this.applyBullets = applyBullets;
    this.applyOutdent = applyOutdent;
    this.applyNumbering = applyNumbering;
    this.applyJustifyLeft = applyJustifyLeft;
    this.applyJustifyCenter = applyJustifyCenter;
    this.applyJustifyRight = applyJustifyRight;
    this.applyJustifyFull = applyJustifyFull;
    this.applyBlockDirLTR = applyBlockDirLTR;
    this.applyBlockDirRTL = applyBlockDirRTL;
    this.applySpan = applySpan;
    this.makeAbsolute = makeAbsolute;
    this.insertHTML = insertHTML;
    this.clearAll = clearAll;
    this.insertCustomTag = insertCustomTag;
    this.selectParagraph = selectParagraph;
    this.applyFormattingStyle = applyFormattingStyle;
    this.hide = hide;
    this.width = "700px";
    this.height = "350px";
    this.publishingPath = "";
    var oScripts = document.getElementsByTagName("script");
    for (var i = 0; i < oScripts.length; i++) {
        var sSrc = oScripts[i].src.toLowerCase();
        if (sSrc.indexOf("saf/editor.js") != -1) {
            this.scriptPath = oScripts[i].src.replace(/editor.js/, "");
            break
        } else if (sSrc.indexOf("innovaeditor.js") != -1) {
            this.scriptPath = oScripts[i].src.replace(/innovaeditor.js/, "saf/");
            break
        }
    }
    this.dialogPath = this.scriptPath.substring(0, this.scriptPath.indexOf("saf/")) + "common/";
    this.GetEmoticons = GetEmoticons;
    this.insertEmoticon = insertEmoticon;
    this.applyQuote = applyQuote;
    this.iconPath = "icons/";
    this.iconWidth = 29;
    this.iconHeight = 25;
    this.dropTopAdjustment_moz = 0;
    this.dropLeftAdjustment_moz = 0;
    this.applyColor = applyColor;
    this.oColor1 = new ColorPicker("oColor1", this.oName);
    this.oColor2 = new ColorPicker("oColor2", this.oName);
    this.expandSelection = expandSelection;
    this.useLastForeColor = false;
    this.useLastBackColor = false;
    this.stateForeColor = "";
    this.stateBackColor = "";
    this.fullScreen = fullScreen;
    this.stateFullScreen = false;
    this.getElm = iwe_getElm;
    this.features = [];
    this.btnParagraph = false;
    this.btnFontName = false;
    this.btnFontSize = false;
    this.btnCut = false;
    this.btnCopy = false;
    this.btnPaste = false;
    this.btnPasteText = false;
    this.btnUndo = false;
    this.btnRedo = false;
    this.btnBold = false;
    this.btnItalic = false;
    this.btnUnderline = false;
    this.btnStrikethrough = false;
    this.btnSuperscript = false;
    this.btnSubscript = false;
    this.btnJustifyLeft = false;
    this.btnJustifyCenter = false;
    this.btnJustifyRight = false;
    this.btnJustifyFull = false;
    this.btnNumbering = false;
    this.btnBullets = false;
    this.btnIndent = false;
    this.btnOutdent = false;
    this.btnLTR = false;
    this.btnRTL = false;
    this.btnForeColor = false;
    this.btnBackColor = false;
    this.btnTable = false;
    this.btnLine = false;
    this.tabs = [
        ["tabHome", "Home", ["group1", "group2", "group4"]],
        ["tabStyle", "Insert", ["group3"]]
    ];
    this.groups = [
        ["group1", "", ["Bold", "Italic", "Underline", "FontDialog", "ForeColor", "TextDialog", "RemoveFormat"]],
        ["group2", "", ["Bullets", "Numbering", "JustifyLeft", "JustifyCenter", "JustifyRight"]],
        ["group3", "", ["LinkDialog", "ImageDialog", "Table", "TableDialog"]],
        ["group4", "", ["Undo", "Redo", "FullScreen", "SourceDialog"]]
    ];
    this.toolbarMode = 2;
    this.showResizeBar = true;
    this.pasteTextOnCtrlV = false;
    this.dialogSize = {
        "Preview": {
            w: 900,
            h: 600
        },
        "TableDialog": {
            w: 785,
            h: 500
        },
        "ImageDialog": {
            w: 755,
            h: 545
        },
        "TextDialog": {
            w: 375,
            h: 475
        },
        "YoutubeDialog": {
            w: 421,
            h: 545
        },
        "LinkDialog": {
            w: 605,
            h: 475
        },
        "SourceDialog": {
            w: 700,
            h: 450
        },
        "CompleteTextDialog": {
            w: 815,
            h: 470
        },
        "FontDialog": {
            w: 500,
            h: 470
        },
        "FlashDialog": {
            w: 390,
            h: 195
        },
        "BookmarkDialog": {
            w: 360,
            h: 240
        },
        "CharsDialog": {
            w: 700,
            h: 122
        },
        "SearchDialog": {
            w: 370,
            h: 140
        }
    };
    this.setDialogSize = function (name, dim) {
        this.dialogSize[name] = dim
    };
    this.fileBrowser = "";
    this.enableFlickr = false;
    this.enableImageStyles = true;
    this.enableYTVideoStyles = true;
    this.enableCssButtons = false;
    this.enableLightbox = false;
    this.enableTableAutoformat = false;
    this.flickrUser = "ysw.insite";
    this.cmdAssetManager = "";
    this.cmdFileManager = "";
    this.cmdImageManager = "";
    this.cmdMediaManager = "";
    this.cmdFlashManager = "";
    this.btnContentBlock = false;
    this.cmdContentBlock = ";";
    this.btnInternalLink = false;
    this.cmdInternalLink = ";";
    this.insertLink = insertLink;
    this.btnCustomObject = false;
    this.cmdCustomObject = ";";
    this.btnInternalImage = false;
    this.cmdInternalImage = ";";
    this.arrStyle = [];
    this.isCssLoaded = false;
    this.openStyleSelect = openStyleSelect;
    this.arrParagraph = [
        [getTxt("Heading 1"), "H1"],
        [getTxt("Heading 2"), "H2"],
        [getTxt("Heading 3"), "H3"],
        [getTxt("Heading 4"), "H4"],
        [getTxt("Heading 5"), "H5"],
        [getTxt("Heading 6"), "H6"],
        [getTxt("Preformatted"), "PRE"],
        [getTxt("Normal (P)"), "P"],
        [getTxt("Normal (DIV)"), "DIV"]
    ];
    this.arrFontName = ["Impact, Charcoal, sans-serif", "Palatino Linotype, Book Antiqua, Palatino, serif", "Tahoma, Geneva, sans-serif", "Century Gothic, sans-serif", "Lucida Sans Unicode, Lucida Grande, sans-serif", "Times New Roman, Times, serif", "Arial Narrow, sans-serif", "Verdana, Geneva, sans-serif", "Copperplate Gothic Light, sans-serif", "Lucida Console, Monaco, monospace", "Gill Sans MT, sans-serif", "Trebuchet MS, Helvetica, sans-serif", "Courier New, Courier, monospace", "Arial, Helvetica, sans-serif", "Georgia, Serif", "Garamond, Serif"];
    this.arrFontSize = [
        [getTxt("Size 1"), "8pt"],
        [getTxt("Size 2"), "10pt"],
        [getTxt("Size 3"), "12pt"],
        [getTxt("Size 4"), "14pt"],
        [getTxt("Size 5"), "18pt"],
        [getTxt("Size 6"), "24pt"],
        [getTxt("Size 7"), "36pt"]
    ];
    this.arrCustomTag = [];
    this.docType = "<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Transitional//EN\" \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd\">";
    this.html = "<html>";
    this.headContent = "";
    this.preloadHTML = "";
    this.originalContent = "";
    this.isContentChanged = isContentChanged;
    this.onSave = function () {
        document.getElementById("iwe_btnSubmit" + this.oName).click()
    };
    this.useBR = false;
    this.useDIV = true;
    this.returnKeyMode = -1;
    this.onFullScreen = function () {
        return true
    };
    this.onNormalScreen = function () {
        return true
    };
    this.initialRefresh = false;
    this.doUndo = doUndo;
    this.doRedo = doRedo;
    this.saveForUndo = saveForUndo;
    this.doUndoRedo = doUndoRedo;
    this.arrUndoList = [];
    this.arrRedoList = [];
    this.useTagSelector = true;
    this.TagSelectorPosition = "bottom";
    this.moveTagSelector = moveTagSelector;
    this.selectElement = selectElement;
    this.removeTag = removeTag;
    this.doClick_TabCreate = doClick_TabCreate;
    this.doRefresh_TabCreate = doRefresh_TabCreate;
    this.arrCustomButtons = [
        ["CustomName1", "alert(0)", "caption here", "btnSave.gif"],
        ["CustomName2", "alert(0)", "caption here", "btnSave.gif"]
    ];
    this.customDialogShow = customDialogShow;
    this.customDialog = [];
    this.onSelectionChanged = function () {
        return true
    };
    this.spellCheckMode = "ieSpell";
    this.encodeIO = false;
    this.changeHeight = changeHeight;
    this.fixWord = fixWord;
    this.REPLACE = REPLACE;
    this.mode = "XHTMLBody";
    this.idTextArea;
    var me = this;
    this.tbar = new ISToolbarManager(this.oName);
    this.tbar.iconPath = this.scriptPath.substring(0, this.scriptPath.indexOf("saf/")) + this.iconPath;
    editor[editor.length] = this;
    return this
};

function changeActiveEditor(oName) {
    var edtObj = eval(oName);
    var edtFrm = document.getElementById("idContent" + oName);
    oUtil.activeElement = null;
    oUtil.oName = oName;
    oUtil.oEditor = edtFrm.contentWindow;
    oUtil.obj = edtObj;
    edtObj.hide()
}

function saveForUndo() {
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    var obj = this;
    if (obj.arrUndoList[0])
        if (oEditor.document.body.innerHTML == obj.arrUndoList[0][0]) return;
    for (var i = 20; i > 1; i--) obj.arrUndoList[i - 1] = obj.arrUndoList[i - 2];
    obj.focus();
    var oSel = oEditor.getSelection();
    var range = oSel.getRangeAt(0);
    obj.arrUndoList[0] = [oEditor.document.body.innerHTML, range.cloneRange()];
    this.arrRedoList = [];
    if (this.btnUndo) this.tbar.btns["btnUndo" + this.oName].setState(1);
    if (this.btnRedo) this.tbar.btns["btnRedo" + this.oName].setState(5)
};

function doUndo() {
    this.doUndoRedo(this.arrUndoList, this.arrRedoList)
};

function doRedo() {
    this.doUndoRedo(this.arrRedoList, this.arrUndoList)
};

function doUndoRedo(listA, listB) {
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    var obj = this;
    if (!listA[0]) return;
    for (var i = 20; i > 1; i--) listB[i - 1] = listB[i - 2];
    var oSel = oEditor.getSelection();
    var range = oSel.getRangeAt(0);
    listB[0] = [oEditor.document.body.innerHTML, range.cloneRange()];
    sHTML = listA[0][0];
    oEditor.document.body.innerHTML = sHTML;
    oSel = oEditor.getSelection();
    oSel.removeAllRanges();
    oSel.addRange(listA[0][1]);
    for (var i = 0; i < 19; i++) listA[i] = listA[i + 1];
    listA[19] = null;
    realTime(this)
};
var bOnSubmitOriginalSaved = false;

function REPLACE(idTextArea, dvId) {
    this.idTextArea = idTextArea;
    var oTextArea = document.getElementById(idTextArea);
    oTextArea.style.display = "none";
    var oForm = oTextArea.form;
    if (oForm) {
        if (!bOnSubmitOriginalSaved) {
            if (oForm.onsubmit) onsubmit_original = oForm.onsubmit;
            bOnSubmitOriginalSaved = true
        }
        oForm.onsubmit = new Function("return onsubmit_new()")
    }
    var sContent = document.getElementById(idTextArea).value + " ";
    sContent = sContent.replace(/&/g, "&amp;");
    sContent = sContent.replace(/</g, "&lt;");
    sContent = sContent.replace(/>/g, "&gt;");
    this.RENDER(sContent, dvId)
};

function isContentChanged() {
    var sContent = "";
    if (this.mode == "HTMLBody") {
        sContent = this.getHTMLBody()
    } else if (this.mode == "HTML") {
        sContent = this.getHTML()
    } else if (this.mode == "XHTMLBody") {
        sContent = this.getXHTMLBody()
    } else if (this.mode == "XHTML") {
        sContent = this.getXHTML()
    }
    if (sContent != this.originalContent) {
        return true
    }
    return false
};

function onsubmit_new() {
    var sContent;
    for (var i = 0; i < oUtil.arrEditor.length; i++) {
        if (!document.getElementById("idContent" + oUtil.arrEditor[i])) continue;
        var oEdit = eval(oUtil.arrEditor[i]);
        if (oEdit.mode == "HTMLBody") sContent = oEdit.getHTMLBody();
        if (oEdit.mode == "HTML") sContent = oEdit.getHTML();
        if (oEdit.mode == "XHTMLBody") sContent = oEdit.getXHTMLBody();
        if (oEdit.mode == "XHTML") sContent = oEdit.getXHTML();
        document.getElementById(oEdit.idTextArea).value = sContent
    }
    if (onsubmit_original) return onsubmit_original()
};

function onsubmit_original() {};

function RENDER(sPreloadHTML, dvId) {
    iconHeight = this.iconHeight;
    if (sPreloadHTML.substring(0, 4) == "<!--" && sPreloadHTML.substring(sPreloadHTML.length - 3) == "-->") sPreloadHTML = sPreloadHTML.substring(4, sPreloadHTML.length - 3);
    if (sPreloadHTML.substring(0, 4) == "<!--" && sPreloadHTML.substring(sPreloadHTML.length - 6) == "--&gt;") sPreloadHTML = sPreloadHTML.substring(4, sPreloadHTML.length - 6);
    sPreloadHTML = sPreloadHTML.replace(/&lt;/g, "<");
    sPreloadHTML = sPreloadHTML.replace(/&gt;/g, ">");
    sPreloadHTML = sPreloadHTML.replace(/&amp;/g, "&");
    this.preloadHTML = sPreloadHTML;
    var sHTMLDropMenus = "";
    var sHTMLIcons = "";
    var sTmp = "";
    this.oColor1.onShow = new Function(this.oName + ".hide()");
    this.oColor1.onPickColor = new Function(this.oName + ".applyColor('ForeColor', " + this.oName + ".oColor1.color)");
    this.oColor1.onRemoveColor = new Function(this.oName + ".applyColor('ForeColor','')");
    this.oColor2.onShow = new Function(this.oName + ".hide()");
    this.oColor2.onPickColor = new Function(this.oName + ".applyColor('hilitecolor', " + this.oName + ".oColor2.color)");
    this.oColor2.onRemoveColor = new Function(this.oName + ".applyColor('HiliteColor','')");
    var me = this;
    var tmp = null,
        tmpTb, grpMap = new Object();
    for (var i = 0; i < this.groups.length; i++) {
        tmp = this.groups[i];
        tmpTb = this.tbar.createToolbar(this.oName + "tbar" + tmp[0]);
        tmpTb.onClick = function (id) {
            tbAction(tmpTb, id, me, me.oName)
        };
        tmpTb.style.toolbar = "main_istoolbar";
        tmpTb.iconPath = this.scriptPath.substring(0, this.scriptPath.indexOf("saf/")) + this.iconPath;
        tmpTb.btnWidth = this.iconWidth;
        tmpTb.btnHeight = this.iconHeight;
        for (var j = 0; j < tmp[2].length; j++) {
            eval(this.oName + ".btn" + tmp[2][j] + "=true")
        }
        buildToolbar(tmpTb, this, tmp[2]);
        grpMap[tmp[0]] = tmp[1]
    }
    if (this.toolbarMode == 1) {
        var eTab = this.tbar.createTbTab("tabCtl" + this.oName),
            tmpGrp;
        for (var i = 0; i < this.tabs.length; i++) {
            tmp = this.tabs[i];
            tmpGrp = this.tbar.createTbGroup(this.oName + "grp" + tmp[0]);
            for (var j = 0; j < tmp[2].length; j++) {
                tmpGrp.addGroup(this.oName + tmp[2][j], grpMap[tmp[2][j]], this.oName + "tbar" + tmp[2][j])
            }
            eTab.addTab(this.oName + tmp[0], tmp[1], tmpGrp)
        }
    } else if (this.groups.length > 0) {
        var tmpGrp;
        tmpGrp = this.tbar.createTbGroup(this.oName + "grp");
        for (var i = 0; i < this.groups.length; i++) {
            tmp = this.groups[i];
            tmpGrp.addGroup(this.oName + tmp[0], grpMap[tmp[0]], this.oName + "tbar" + tmp[0])
        }
        if (this.toolbarMode == 3) {
            tmpGrp.groupFlow = true
        }
        if (this.toolbarMode == 4) {
            tmpGrp.groupFlow = true;
            tmpGrp.draggable = true
        }
    }
    var sHTML = "";
    var icPath = this.scriptPath.substring(0, this.scriptPath.indexOf("saf/")) + this.iconPath;
    if (!document.getElementById("id_refresh_z_index")) sHTML += "<div id=id_refresh_z_index style='margin:0px'></div>";
    sHTML += "<table id=idArea" + this.oName + " name=idArea" + this.oName + " border='0px' " + "cellpadding=0 cellspacing=0 width='" + this.width + "' height='" + this.height + "' style='width:" + this.width + ";height:" + this.height + ";border:none;border-bottom:solid 1px #cfcfcf'>" + "<tr><td colspan=2 style=\"padding:0px;border:#cfcfcf 0px solid;background:url('" + icPath + "bg.gif')\">" + "<table cellpadding=0 cellspacing=0 border=0 width='100%'  style='border:none;margin:0px'><tr><td dir=ltr style='padding:0px;'>" + this.tbar.render() + "</td></tr></table>" + "</td></tr>" + "<tr id=idTagSelTopRow" + this.oName + "><td colspan=2 id=idTagSelTop" + this.oName + " height=0px style='color:#000000;padding:0px;'></td></tr>";
    sHTML += "<tr style='width:100%;height:100%'><td colspan=2 valign=top height=100% style='padding:0px;background:white;padding-right:0px;'>";
    sHTML += "<table id='cntContainer" + this.oName + "' cellpadding=0 cellspacing=0 width='100%' height='100%' style='margin-top:0px;border:none;'><tr style='width:100%;height:100%'><td width='100%' height='100%' style='padding:0px;border:solid 1px #cfcfcf;border-bottom:none'>";
    sHTML += "<iframe frameborder='no' style='width:100%;height:100%;' " + " name=idContent" + this.oName + " id=idContent" + this.oName + " src='" + this.scriptPath + "blank.gif'></iframe>";
    sHTML += "<iframe style='display:none;width:1px;height:1px;overflow:auto;border:0px' id=\"myStyle" + this.oName + "\" name=\"myStyle" + this.oName + "\" src='" + this.scriptPath + "blank.gif'></iframe>";
    sHTML += "</td><td id=idStyles" + this.oName + " style='padding:0px;background:#f4f4f4'></td></tr></table>";
    sHTML += "</td></tr>";
    sHTML += "<tr id=idTagSelBottomRow" + this.oName + "><td colspan=2 id=idTagSelBottom" + this.oName + " style='color:#000000;padding:0px;'></td></tr>";
    if (this.showResizeBar) {
        sHTML += "<tr id=idResizeBar" + this.oName + "><td colspan=2 style='padding:0px;'><div  style='cursor:n-resize;' class='resize_bar' onmousedown=\"onEditorStartResize(event, this, '" + this.oName + "')\" ></div></td></tr>"
    }
    sHTML += "</table>";
    sHTML += sHTMLDropMenus;
    sHTML += "<input type=submit name=iwe_btnSubmit" + this.oName + " id=iwe_btnSubmit" + this.oName + " value=SUBMIT style='display:none' >";
    if (dvId) {
        var edtStr = [];
        edtStr[0] = sHTML;
        document.getElementById(dvId).innerHTML = edtStr.join("")
    } else {
        document.write(sHTML)
    }
    this.init()
};

function onEditorStartResize(ev, elm, oName) {
    document.onmousemove = onEditorResize;
    document.onmouseup = onEditorStopResize;
    document.onselectstart = function () {
        return false
    };
    document.ondragstart = function () {
        return false
    };
    document.body.style.cursor = "n-resize";
    oUtil.currentResized = eval(oName);
    oUtil.resizeInit = {
        x: ev.screenX,
        y: ev.screenY
    };
    if (!oUtil.isWindow) oUtil.isWindow = new ISWindow(oName);
    oUtil.isWindow.showOverlay()
};

function onEditorStopResize(event) {
    oUtil.resizeOffset = {
        dx: event.screenX - oUtil.resizeInit.x,
        dy: event.screenY - oUtil.resizeInit.y
    };
    oUtil.currentResized.changeHeight(oUtil.resizeOffset.dy);
    oUtil.isWindow.hideOverlay();
    document.onmousemove = null;
    document.onmouseup = null;
    document.body.style.cursor = "default"
};

function onEditorResize(event) {
    oUtil.resizeOffset = {
        dx: event.screenX - oUtil.resizeInit.x,
        dy: event.screenY - oUtil.resizeInit.y
    };
    oUtil.currentResized.changeHeight(oUtil.resizeOffset.dy);
    oUtil.resizeInit = {
        x: event.screenX,
        y: event.screenY
    }
};

function initISEditor() {
    if (this.useTagSelector) {
        if (this.TagSelectorPosition == "bottom") this.TagSelectorPosition = "top";
        else this.TagSelectorPosition = "bottom";
        this.moveTagSelector()
    }
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    oUtil.oName = this.oName;
    oUtil.oEditor = oEditor;
    oUtil.obj = this;
    oUtil.arrEditor.push(this.oName);
    var arrA = String(this.preloadHTML).match(/<HTML[^>]*>/ig);
    if (arrA) {
        this.loadHTMLFull(this.preloadHTML)
    } else {
        this.loadHTML(this.preloadHTML)
    } if (!oUtil.bOnLoadReplaced) {
        if (window.onload) onload_original = window.onload;
        window.onload = new Function("onload_new()");
        oUtil.bOnLoadReplaced = true
    }
    try {
        oEditor.document.body.contentEditable = true
    } catch (e) {}
    try {
        var cnt = oEditor.document.body.innerHTML;
        cnt = cnt.replace(/\s+/gi, "");
        if (cnt == "") {
            oEditor.document.body.innerHTML = "<br class=\"innova\" />"
        }
        var range = oEditor.document.createRange();
        range.selectNode(oEditor.document.body.childNodes[0]);
        range.collapse(true);
        var sel = oEditor.getSelection();
        sel.removeAllRanges();
        sel.addRange(range)
    } catch (e) {}
    switch (this.mode) {
    case "HTMLBody":
        this.originalContent = this.getHTMLBody();
        break;
    case "HTML":
        this.originalContent = this.getHTML();
        break;
    case "XHTMLBody":
        this.originalContent = this.getXHTMLBody();
        break;
    case "XHTML":
        this.originalContent = this.getXHTML();
        break
    };
    if (this.returnKeyMode == -1) {
        if (this.useDIV) this.returnKeyMode = 1;
        else if (this.useBR) this.returnKeyMode = 2;
        else if (!this.useDIV && !this.useBR) this.returnKeyMode = 3
    }
    if (this.disableFocusOnLoad == false) {
        this.focus()
    }
};

function buildToolbar(tb, oEdt, btnMap) {
    var oName = oEdt.oName;
    for (var i = 0; i < btnMap.length; i++) {
        sButtonName = btnMap[i];
        switch (sButtonName) {
        case "|":
            tb.addSeparator();
            break;
        case "BRK":
            tb.addBreak();
            break;
        case "FullScreen":
            tb.addButton("btnFullScreen" + oName, "btnFullScreen.gif", getTxt("Full Screen"));
            break;
        case "Print":
            tb.addButton("btnPrint" + oName, "btnPrint.gif", getTxt("Print"));
            break;
        case "Search":
            tb.addButton("btnSearch" + oName, "btnSearch.gif", getTxt("Search"));
            break;
        case "SpellCheck":
            if (oEdt.spellCheckMode != "ieSpell") tb.addButton("btnSpellCheck" + oName, "btnSpellCheck.gif", getTxt("Check Spelling"));
            break;
        case "Styles":
            tb.addButton("btnStyles" + oName, "btnStyleSelect.gif", getTxt("Style Selection"));
            break;
        case "Paragraph":
            tb.addDropdownButton("btnParagraph" + oName, "ddParagraph" + oName, "btnParagraph.gif", getTxt("Paragraph"), 37);
            var ddPar = new ISDropdown("ddParagraph" + oName);
            ddPar.onClick = function (id) {
                ddAction(tb, id, oEdt, oEdt.oName)
            };
            for (var j = 0; j < oEdt.arrParagraph.length; j++) {
                ddPar.addItem("btnParagraph_" + j + oName, "<" + oEdt.arrParagraph[j][1] + " style=\"\margin-bottom:4px\"  unselectable=on> " + oEdt.arrParagraph[j][0] + "</" + oEdt.arrParagraph[j][1] + ">")
            }
            break;
        case "FontName":
            tb.addDropdownButton("btnFontName" + oName, "ddFontName" + oName, "btnFontName.gif", getTxt("Font Name"), 37);
            var ddFont = new ISDropdown("ddFontName" + oName);
            ddFont.onClick = function (id) {
                ddAction(tb, id, oEdt, oEdt.oName)
            };
            for (var j = 0; j < oEdt.arrFontName.length; j++) {
                ddFont.addItem("btnFontName_" + j + oName, "<span style='font-family:" + oEdt.arrFontName[j] + ";font-size:12px' unselectable=on>" + oEdt.arrFontName[j] + "</span>")
            }
            break;
        case "FontSize":
            tb.addDropdownButton("btnFontSize" + oName, "ddFontSize" + oName, "btnFontSize.gif", getTxt("Font Size"), 37);
            var ddFs = new ISDropdown("ddFontSize" + oName);
            ddFs.onClick = function (id) {
                ddAction(tb, id, oEdt, oEdt.oName)
            };
            for (var j = 0; j < oEdt.arrFontSize.length; j++) {
                ddFs.addItem("btnFontSize_" + j + oName, "<div unselectable=on style=\"font-size:" + oEdt.arrFontSize[j][1] + "\">" + oEdt.arrFontSize[j][0] + "</div>")
            }
            break;
        case "Undo":
            tb.addButton("btnUndo" + oName, "btnUndo.gif", getTxt("Undo"));
            break;
        case "Redo":
            tb.addButton("btnRedo" + oName, "btnRedo.gif", getTxt("Redo"));
            break;
        case "Paste":
            tb.addDropdownButton("btnPaste" + oName, "ddPaste" + oName, "btnPaste.gif", getTxt("Paste"));
            var pvDD = new ISDropdown("ddPaste" + oName);
            pvDD.iconPath = tb.iconPath;
            pvDD.addItem("btnPasteClip" + oName, getTxt("Paste"), "btnPasteClip.gif");
            pvDD.addItem("btnPasteText" + oName, getTxt("Paste Text"), "btnPasteText.gif");
            pvDD.onClick = function (id) {
                ddAction(tb, id, oEdt, oEdt.oName)
            };
            break;
        case "Emoticons":
            tb.addDropdownButton("btnEmoticons" + oName, "ddEmoticons" + oName, "btnEmoticons.gif", getTxt("Emoticons"));
            var ddTable = new ISDropdown("ddEmoticons" + oName);
            ddTable.add(new ISCustomDDItem("btnInsertEmoticons", oEdt.GetEmoticons()));
            break;
        case "Quote":
            tb.addToggleButton("btnQuote" + oName, "", false, "btnQuote.gif", getTxt("Quote"));
            break;
        case "Bold":
            tb.addToggleButton("btnBold" + oName, "", false, "btnBold.gif", getTxt("Bold"));
            break;
        case "Italic":
            tb.addToggleButton("btnItalic" + oName, "", false, "btnItalic.gif", getTxt("Italic"));
            break;
        case "Underline":
            tb.addToggleButton("btnUnderline" + oName, "", false, "btnUnderline.gif", getTxt("Underline"));
            break;
        case "Strikethrough":
            tb.addToggleButton("btnStrikethrough" + oName, "", false, "btnStrikethrough.gif", getTxt("Strikethrough"));
            break;
        case "Superscript":
            tb.addToggleButton("btnSuperscript" + oName, "", false, "btnSuperscript.gif", getTxt("Superscript"));
            break;
        case "Subscript":
            tb.addToggleButton("btnSubscript" + oName, "", false, "btnSubscript.gif", getTxt("Subscript"));
            break;
        case "JustifyLeft":
            tb.addToggleButton("btnJustifyLeft" + oName, "align", false, "btnLeft.gif", getTxt("Justify Left"));
            break;
        case "JustifyCenter":
            tb.addToggleButton("btnJustifyCenter" + oName, "align", false, "btnCenter.gif", getTxt("Justify Center"));
            break;
        case "JustifyRight":
            tb.addToggleButton("btnJustifyRight" + oName, "align", false, "btnRight.gif", getTxt("Justify Right"));
            break;
        case "JustifyFull":
            tb.addToggleButton("btnJustifyFull" + oName, "align", false, "btnFull.gif", getTxt("Justify Full"));
            break;
        case "Numbering":
            tb.addToggleButton("btnNumbering" + oName, "bullet", false, "btnNumber.gif", getTxt("Numbering"));
            break;
        case "Bullets":
            tb.addToggleButton("btnBullets" + oName, "bullet", false, "btnList.gif", getTxt("Bullets"));
            break;
        case "Indent":
            tb.addButton("btnIndent" + oName, "btnIndent.gif", getTxt("Indent"));
            break;
        case "Outdent":
            tb.addButton("btnOutdent" + oName, "btnOutdent.gif", getTxt("Outdent"));
            break;
        case "LTR":
            tb.addToggleButton("btnLTR" + oName, "dir", false, "btnLTR.gif", getTxt("Left To Right"));
            break;
        case "RTL":
            tb.addToggleButton("btnRTL" + oName, "dir", false, "btnRTL.gif", getTxt("Right To Left"));
            break;
        case "ForeColor":
            tb.addDropdownButton("btnForeColor" + oName, "ddForeColor" + oName, "btnForeColor.gif", getTxt("Foreground Color"));
            var ddTable = new ISDropdown("ddForeColor" + oName);
            ddTable.add(new ISCustomDDItem("btnInsertForeColor", oEdt.oColor1.generateHTML()));
            break;
        case "BackColor":
            tb.addDropdownButton("btnBackColor" + oName, "ddBackColor" + oName, "btnBackColor.gif", getTxt("Background Color"));
            var ddTable = new ISDropdown("ddBackColor" + oName);
            ddTable.add(new ISCustomDDItem("btnInsertBackColor", oEdt.oColor2.generateHTML()));
            break;
        case "FontDialog":
            tb.addButton("btnFontDialog" + oName, "btnFont.gif", getTxt("Fonts"));
            break;
        case "TextDialog":
            tb.addButton("btnTextDialog" + oName, "btnText.gif", getTxt("Text"));
            break;
        case "CompleteTextDialog":
            tb.addButton("btnCompleteTextDialog" + oName, "btnText.gif", getTxt("Text"));
            break;
        case "LinkDialog":
            tb.addButton("btnLinkDialog" + oName, "btnHyperlink.gif", getTxt("Link"));
            break;
        case "ImageDialog":
            tb.addButton("btnImageDialog" + oName, "btnImage.gif", getTxt("Image"));
            break;
        case "YoutubeDialog":
            //tb.addButton("btnYoutubeDialog" + oName, "btnYoutubeVideo.gif", getTxt("YoutubeVideo"));
            break;
        case "TableDialog":
            tb.addButton("btnTableDialog" + oName, "btnTableEdit.gif", getTxt("Table"));
            break;
        case "FlashDialog":
            tb.addButton("btnFlashDialog" + oName, "btnFlash.gif", getTxt("Flash"));
            break;
        case "CharsDialog":
            tb.addButton("btnCharsDialog" + oName, "btnSymbol.gif", getTxt("Special Characters"));
            break;
        case "SearchDialog":
            tb.addButton("btnSearchDialog" + oName, "btnSearch.gif", getTxt("Search & Replace"));
            break;
        case "SourceDialog":
            tb.addButton("btnSourceDialog" + oName, "btnSource.gif", getTxt("HTML Editor"));
            break;
        case "BookmarkDialog":
            tb.addButton("btnBookmarkDialog" + oName, "btnBookmark.gif", getTxt("Bookmark"));
            break;
        case "Preview":
            tb.addButton("btnPreview" + oName, "btnPreview.gif", getTxt("Preview"));
            break;
        case "CustomTag":
            tb.addDropdownButton("btnCustomTag" + oName, "ddCustomTag" + oName, "btnCustomTag.gif", getTxt("Tags"), 37);
            var ddCustomTag = new ISDropdown("ddCustomTag" + oName);
            ddCustomTag.onClick = function (id) {
                ddAction(tb, id, oEdt, oEdt.oName)
            };
            for (var j = 0; j < oEdt.arrCustomTag.length; j++) {
                ddCustomTag.addItem("btnCustomTag_" + j + oName, oEdt.arrCustomTag[j][0])
            }
            break;
        case "ContentBlock":
            tb.addButton("btnContentBlock" + oName, "btnContentBlock.gif", getTxt("Content Block"));
            break;
        case "InternalLink":
            tb.addButton("btnInternalLink" + oName, "btnInternalLink.gif", getTxt("Internal Link"));
            break;
        case "InternalImage":
            tb.addButton("btnInternalImage" + oName, "btnInternalImage.gif", getTxt("Internal Image"));
            break;
        case "CustomObject":
            tb.addButton("btnCustomObject" + oName, "btnCustomObject.gif", getTxt("Object"));
            break;
        case "Table":
            var sdd = [],
                sZ = 0;
            sdd[sZ++] = "<table width='175' id='dropTableCreate" + oName + "' style='cursor:default;background:#f3f3f8;' cellpadding=0 cellspacing=1 unselectable=on>";
            for (var m = 0; m < 8; m++) {
                sdd[sZ++] = "<tr>";
                for (var n = 0; n < 8; n++) {
                    sdd[sZ++] = "<td onclick='" + oName + ".doClick_TabCreate(this)' onmouseover='doOver_TabCreate(this);event.cancelBubble=true;' style='background:#ffffff;font-size:1px;border:#d3d3d3 1px solid;width:20px;height:20px;' unselectable=on>&nbsp;</td>"
                }
                sdd[sZ++] = "</tr>"
            }
            sdd[sZ++] = "<tr><td colspan=8 onclick=\"" + oName + ".hide();modelessDialogShow('" + oEdt.dialogPath + "webtable.htm',785, 500);\" onmouseover=\"doOut_TabCreate(document.getElementById('dropTableCreate" + oName + "'));this.innerHTML='" + getTxt("Advanced Table Insert") + "';this.style.border='#777777 1px solid';this.style.backgroundColor='#444444';this.style.color='#ffffff'\" onmouseout=\"this.style.border='#f3f3f8 1px solid';this.style.backgroundColor='#f3f3f8';this.style.color='#000000'\" align=center style='font-family:verdana;font-size:10px;color:#000000;border:#f3f3f8 1px solid;padding:1px 1px 1px 1px' unselectable=on>" + getTxt("Advanced Table Insert") + "</td></tr>";
            sdd[sZ++] = "</table>";
            tb.addDropdownButton("btnTable" + oName, "ddTable" + oName, "btnTable.gif", getTxt("Insert Table"));
            var ddTable = new ISDropdown("ddTable" + oName);
            ddTable.add(new ISCustomDDItem("btnInsertTable", sdd.join("")));
            break;
        case "Absolute":
            tb.addButton("btnAbsolute" + oName, "btnAbsolute.gif", getTxt("Absolute"));
            break;
        case "Line":
            tb.addButton("btnLine" + oName, "btnLine.gif", getTxt("Line"));
            break;
        case "RemoveFormat":
            tb.addButton("btnRemoveFormat" + oName, "btnRemoveFormat.gif", getTxt("Remove Formatting"));
            break;
        case "ClearAll":
            tb.addButton("btnClearAll" + oName, "btnDelete.gif", getTxt("Clear All"));
            break;
        default:
            for (var j = 0; j < oEdt.arrCustomButtons.length; j++) {
                if (sButtonName == oEdt.arrCustomButtons[j][0]) {
                    sCbName = oEdt.arrCustomButtons[j][0];
                    sCbCaption = oEdt.arrCustomButtons[j][2];
                    sCbImage = oEdt.arrCustomButtons[j][3];
                    if (oEdt.arrCustomButtons[j].length < 5) {
                        if (oEdt.arrCustomButtons[j][4]) tb.addButton(sCbName + oName, sCbImage, sCbCaption, oEdt.arrCustomButtons[j][4]);
                        else tb.addButton(sCbName + oName, sCbImage, sCbCaption)
                    } else {
                        if (oEdt.arrCustomButtons[j][4] != 0) tb.addDropdownButton(sCbName + oName, "dd" + sCbName + oName, sCbImage, sCbCaption, oEdt.arrCustomButtons[j][4]);
                        else tb.addDropdownButton(sCbName + oName, "dd" + sCbName + oName, sCbImage, sCbCaption);
                        var ddTable = new ISDropdown("dd" + sCbName + oName);
                        var arrItems = oEdt.arrCustomButtons[j][5];
                        for (var k = 0; k < arrItems.length; k++) {
                            tmp = arrItems[k];
                            ddTable.addItem('item_' + sCbName + k + oName, '<div style="margin:4px 0;padding:2px 0;font-size:13px;"  unselectable=on onclick="' + tmp[1] + '" > ' + tmp[2] + '</div>')
                        }
                    }
                }
            }
            break
        }
    }
};

function iwe_getElm(s) {
    return document.getElementById(s + this.oName)
};

function onload_new() {
    onload_original()
};

function onload_original() {};
var arrColorPickerObjects = [];

function ColorPicker(sName, sParent) {
    this.oParent = sParent;
    if (sParent) {
        this.oName = sParent + "." + sName;
        this.oRenderName = sName + sParent
    } else {
        this.oName = sName;
        this.oRenderName = sName
    }
    arrColorPickerObjects.push(this.oName);
    this.onShow = function () {
        return true
    };

    this.onHide = function () {
        return true
    };
    this.onPickColor = function () {
        return true
    };
    this.onRemoveColor = function () {
        return true
    };
    this.hide = hideColorPicker;
    this.hideAll = hideColorPickerAll;
    this.color;
    this.isActive = false;
    this.generateHTML = generateHTML
};

function generateHTML() {
    var arrColors = [
        ["#ef001b", "#cc0017", "#a60012", "#83000e", "#ef0078", "#ce0067", "#ad0057", "#8b0045", "#e301ed", "#c501ce", "#a401ab", "#88018e", "#6716ef", "#5913ce", "#4b10af", "#3e0d90"],
        ["#f67684", "#e36875", "#ca5965", "#b34e59", "#f563ac", "#de599b", "#cc5490", "#b24d7f", "#ee68f4", "#db5fe1", "#c759cc", "#b255b6", "#a779f5", "#976cdf", "#8d68cc", "#7f5eb7"],
        ["#fcc0c6", "#eea8af", "#dd959c", "#ce8c93", "#fec7e2", "#f4b8d6", "#e5a6c6", "#d495b4", "#fabffd", "#eeaff1", "#e19fe4", "#cf90d2", "#e0c3fd", "#d1b1f1", "#c1a0e2", "#b192d1"],
        ["#fef5f6", "#fdeced", "#f7dee0", "#eacedc", "#fef3f8", "#fbe8f1", "#efd0e0", "#e6c7d6", "#fef2fe", "#fae6fb", "#f1d3f2", "#e3c1e4", "#f5edfe", "#f0e5fb", "#e1d3ef", "#d9cbe7"],
        ["#028b6c", "#02775d", "#02644e", "#015441", "#1882ed", "#1574d4", "#115eab", "#0e4f90", "#0040eb", "#0039d0", "#0030b1", "#002892", "#50509e", "#46468b", "#3a3a73", "#303060"],
        ["#69baa7", "#61a898", "#57998a", "#508b7d", "#7bb8f5", "#6ea7e0", "#6195c9", "#5684b2", "#6d92f5", "#5f82e0", "#5675c9", "#4d68b2", "#9b9bc9", "#8b8bb6", "#7e7ea5", "#747496"],
        ["#d0eae4", "#b3d7cf", "#9bc4ba", "#8fb4ac", "#c3dffc", "#aacdf0", "#9bbde0", "#97b4d1", "#bdcdfb", "#a8bbef", "#96aae1", "#8a9bcb", "#d8d8eb", "#c7c7dc", "#b5b5cc", "#a5a5bc"],
        ["#f0f8f6", "#deedea", "#d7e6e2", "#ceddda", "#f1f7fe", "#e5f0fb", "#d8e5f2", "#cfdbe7", "#eff3fe", "#e5eafa", "#dde3f4", "#d2d8ea", "#f4f4f9", "#e5e5ef", "#dbdbe5", "#d6d6df"],
        ["#00a000", "#008d00", "#007700", "#006000", "#86d800", "#73ba00", "#629e00", "#528400", "#eded00", "#cece00", "#afaf00", "#909000", "#e3ab00", "#c79600", "#aa8000", "#856400"],
        ["#68c868", "#5cb65c", "#56a456", "#4b924b", "#b7e768", "#a8d45f", "#97c056", "#86aa4d", "#f1f164", "#e1e15d", "#caca58", "#b2b24d", "#eecc65", "#dabc5e", "#c7ac59", "#b09850"],
        ["#c6ecc6", "#addead", "#96cd96", "#87b987", "#e1f6c0", "#d0eba6", "#c1d99a", "#b1c88c", "#fbfbad", "#f1f194", "#e2e28e", "#cece8c", "#faeaba", "#f2dfa7", "#e6d090", "#cbbb8b"],
        ["#eef9ee", "#dff1df", "#d5e8d5", "#c6dbc6", "#f1fbe2", "#e9f5d5", "#dfebcd", "#d4e1c0", "#fefef0", "#fafae3", "#f0f0cb", "#e4e4c5", "#fdf8ea", "#f9f2de", "#eee4c7", "#dfd7bf"],
        ["#818181", "#676767", "#494949", "#000000", "#783c00", "#673300", "#562b00", "#472300", "#eb4600", "#cd3d00", "#ad3300", "#8f2a00", "#ed7700", "#d26900", "#af5800", "#904800"],
        ["#c9c9c9", "#a9a9a9", "#919191", "#787878", "#af8b68", "#a28264", "#917458", "#856d55", "#f19068", "#dd8561", "#c97654", "#b47053", "#f5ac63", "#e1a05f", "#ca9259", "#b78451"],
        ["#efefef", "#dcdcdc", "#c1c1c1", "#9d9d9d", "#dbcab9", "#ccb8a5", "#bda792", "#a3917f", "#fbcebc", "#f1bba5", "#e1aa93", "#ce9f8b", "#fcd7b3", "#f3caa2", "#e7b98c", "#c8a078"],
        ["#ffffff", "#f7f7f7", "#ededed", "#dddddd", "#f4efeb", "#efe8e1", "#e6ded6", "#dbd3cc", "#fef5f2", "#fae8e1", "#f0dbd3", "#e1cbc2", "#fef7f0", "#faecde", "#f1e2d3", "#e3d3c3"]
    ];
    var sHTMLColor = "<table id=dropColor" + this.oRenderName + " style=\"background-color:#fcfcfc;\" unselectable=on cellpadding=0 cellspacing=0 width=140px><tr><td unselectable=on style='padding:2px;'>";
    sHTMLColor += "<table align=center cellpadding=0 cellspacing=0 border=0 unselectable=on>";
    for (var i = 0; i < arrColors.length; i++) {
        sHTMLColor += "<tr>";
        for (var j = 0; j < arrColors[i].length; j++) sHTMLColor += "<td onclick=\"" + this.oName + ".color='" + arrColors[i][j] + "';" + this.oName + ".onPickColor();" + this.oName + ".hideAll()\" style=\"\" unselectable=on>" + "<table style='-moz-border-radius:0px;-webkit-border-radius: 0px;margin:0px;padding:0px;width:12px;height:12px;background:" + arrColors[i][j] + ";' cellpadding=0 cellspacing=0 unselectable=on>" + "<tr><td unselectable=on style='cursor:pointer;margin:0px;padding:0px;line-height:normal;font-size:10px;'>&nbsp;</td></tr>" + "</table></td>";
        sHTMLColor += "</tr>"
    }
    sHTMLColor += "<tr>";
    sHTMLColor += "<td colspan=16 unselectable=on style='padding:0px;'>" + "<table style='padding:0px;margin-left:1px;width:100%;height:14px;background:#f4f4f4;' cellpadding=0 cellspacing=0 unselectable=on>" + "<tr><td onclick=\"" + this.oName + ".onRemoveColor();" + this.oName + ".hideAll()\" onmouseover=\"this.style.border='#777777 1px solid'\" onmouseout=\"this.style.border='white 1px solid'\" style=\"cursor:default;padding:1px;border:white 1px solid;font-family:verdana;font-size:10px;color:#000000;line-height:9px;\" align=center valign=top unselectable=on>x</td></tr>" + "</table></td>";
    sHTMLColor += "</tr>";
    sHTMLColor += "</table>";
    sHTMLColor += "</td></tr></table>";
    return sHTMLColor
};

function hideColorPicker() {
    this.onHide();
    return;
    var box = document.getElementById("dropColor" + this.oRenderName);
    box.style.display = "none";
    this.isActive = false
};

function hideColorPickerAll() {
    return;
    for (var i = 0; i < arrColorPickerObjects.length; i++) {
        var box = document.getElementById("dropColor" + eval(arrColorPickerObjects[i]).oRenderName);
        box.style.display = "none";
        eval(arrColorPickerObjects[i]).isActive = false
    }
};

function loadHTML(sHTML) {
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    oEditor.document.open("text/html", "replace");
    var defaultEditCss = "<link id='tmp_xxinnova' href='" + this.scriptPath + "default_edit.css' type='text/css' rel='stylesheet' />";
    var hdr = "";
    if (this.publishingPath != "") {
        var arrA = String(this.preloadHTML).match(/<base[^>]*>/ig);
        if (!arrA) {
            hdr = this.docType + "<HTML><HEAD><BASE HREF=\"" + this.publishingPath + "\"/>" + defaultEditCss + this.headContent + "</HEAD><BODY>" + "</BODY></HTML>"
        }
    } else {
        hdr = this.docType + "<HTML><HEAD>" + defaultEditCss + this.headContent + "</HEAD><BODY>" + "</BODY></HTML>"
    }

		// START ProductCart Mod: Fix YouTube and Vimeo URLs
    if (sHTML !== undefined && document.location.protocol == "https:") {
    	sHTML = sHTML.replace("http://www.youtube.com/", "https://www.youtube.com/");
    	sHTML = sHTML.replace("http://player.vimeo.com/", "https://player.vimeo.com/");
    }
		// END ProductCart Mod: Fix HTML content when using HTTPS.

    oEditor.document.write(hdr);
    oEditor.document.close();
    oEditor.document.body.innerHTML = sHTML;
    oEditor.document.body.contentEditable = true;
    var me = this;
    oEditor.document.addEventListener("keyup", new Function("editorDoc_onkeyup(" + this.oName + ")"), true);
    oEditor.document.addEventListener("mouseup", function (e) {
        editorDoc_onmouseup(e, me.oName)
    }, true);
    oEditor.document.addEventListener("keydown", new Function("var e=arguments[0];doKeyPress(eval(e), " + this.oName + ")"), false);
    if (typeof this.css == 'object') {
        for (var j = 0; j < this.css.length; j++) {
            var objL = oEditor.document.createElement("LINK");
            objL.href = this.css[j];
            objL.rel = "StyleSheet";
            oEditor.document.documentElement.childNodes[0].appendChild(objL)
        }
    } else {
        if (this.css != "") {
            var objL = oEditor.document.createElement("LINK");
            objL.href = this.css;
            objL.rel = "StyleSheet";
            oEditor.document.documentElement.childNodes[0].appendChild(objL)
        }
    } if (this.arrStyle.length > 0) {
        var oElement = oEditor.document.createElement("STYLE");
        oEditor.document.documentElement.childNodes[0].appendChild(oElement);
        var sCssText = "";
        for (var i = 0; i < this.arrStyle.length; i++) {
            selector = this.arrStyle[i][0];
            style = this.arrStyle[i][3];
            sCssText += selector + " { " + style + " } "
        }
        oElement.appendChild(oEditor.document.createTextNode(sCssText))
    }
    this.cleanDeprecated()
};

function loadHTMLFull(sHTML, firstLoad) {
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    var arrA = String(sHTML).match(/<!DOCTYPE[^>]*>/ig);
    if (arrA)
        for (var i = 0; i < arrA.length; i++) {
            this.docType = arrA[i]
        } else this.docType = "";
    var arrB = String(sHTML).match(/<HTML[^>]*>/ig);
    if (arrB)
        for (var i = 0; i < arrB.length; i++) {
            s = arrB[i];
            s = s.replace(/\"[^\"]*\"/ig, function (x) {
                x = x.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/'/g, "&apos;").replace(/\s+/ig, "#_#");
                return x
            });
            s = s.replace(/<([^ >]*)/ig, function (x) {
                return x.toLowerCase()
            });
            s = s.replace(/ ([^=]+)=([^" >]+)/ig, " $1=\"$2\"");
            s = s.replace(/ ([^=]+)=/ig, function (x) {
                return x.toLowerCase()
            });
            s = s.replace(/#_#/ig, " ");
            this.html = s
        } else this.html = "<html>";
    if (this.publishingPath != "") {
        var arrA = sHTML.match(/<base[^>]*>/ig);
        if (!arrA) {
            sHTML = "<BASE HREF=\"" + this.publishingPath + "\"/>" + sHTML
        }
    }
    this.css = "";
    var arrExtCss = String(sHTML).match(/<link[^>]*>/ig);
    if (arrExtCss)
        for (var i = 0; i < arrExtCss.length; i++) {
            var s = arrExtCss[i];
            var arrTmp = s.split("href=\"");
            if (arrTmp.length > 1) {
                s = arrTmp[1];
                arrTmp = s.split("\"");
                if (arrTmp.length > 1) this.css = arrTmp[0]
            }
        }
    var p1 = 0,
        p2 = sHTML.length;
    var tmp = sHTML.match(/<body[^>]*>/i);
    if (tmp) {
        p1 = sHTML.indexOf(tmp[0]) + tmp[0].length
    }
    tmp = sHTML.match(/<\/body>/i);
    if (tmp) {
        p2 = sHTML.indexOf(tmp[0])
    }
    var head = sHTML.substring(0, p1) + "</html>";
    var cnt = sHTML.substring(p1, p2);
    var foot = sHTML.substr(p2);
    var defaultEditCss = "<link id='tmp_xxinnova' href='" + this.scriptPath + "default_edit.css' type='text/css' rel='stylesheet' />";
    if (head.match(/<head>/gi)) {
        head = head.replace(/<head>/gi, "<head>" + defaultEditCss)
    } else {
        head = head.replace(/<body/gi, "<head>" + defaultEditCss + "</head><body")
    }
    oEditor.document.open("text/html", "replace");
    oEditor.document.write(head + foot);
    oEditor.document.close();
    oEditor.document.body.innerHTML = cnt;
    oEditor.document.body.contentEditable = true;
    var me = this;
    oEditor.document.addEventListener("keyup", new Function("editorDoc_onkeyup(" + me.oName + ")"), true);
    oEditor.document.addEventListener("mouseup", function (e) {
        editorDoc_onmouseup(e, me.oName)
    }, true);
    oEditor.document.addEventListener("keydown", new Function("var e=arguments[0];doKeyPress(eval(e), " + this.oName + ")"), false);
    this.cleanDeprecated()
};

function putHTML(sHTML) {
    this.loadHTMLFull(sHTML, true)
};

function encodeHTMLCode(sHTML) {
    return sHTML.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
};

function cleanHTML(sHTML) {
    var h = sHTML.replace(/<br class="innova" \/>/gi, "").replace(/<br class="innova">/gi, "");
    var sMatch = h.match(/<link[^>]*>/ig);
    if (sMatch) {
        for (var i = 0; i < sMatch.length; i++) {
            if (sMatch[i].indexOf("tmp_xxinnova") != -1) {
                h = h.replace(new RegExp(sMatch[i], "gi"), "")
            }
        }
    }
    return h
};

function getTextBody() {
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    return oEditor.document.body.textContent
};

function getHTML(bEdit) {
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    this.cleanDeprecated();
    sHTML = getOuterHTML(oEditor.document.documentElement);
    sHTML = String(sHTML).replace(/\<PARAM NAME=\"Play\" VALUE=\"0\">/ig, "<PARAM NAME=\"Play\" VALUE=\"-1\">");
    sHTML = this.docType + sHTML;
    sHTML = oUtil.replaceSpecialChar(sHTML);
    if (this.encodeIO) sHTML = encodeHTMLCode(sHTML);
    sHTML = sHTML.replace(/\u00A0+/gi, " ");
    sHTML = cleanHTML(sHTML);
    if (!bEdit) {
        sHTML = sHTML.replace(/ contenteditable=\"true\"/ig, "");
        sHTML = sHTML.replace(/ contenteditable=\"false\"/ig, "")
    }
    return sHTML
};

function getHTMLBody(bEdit) {
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    this.cleanDeprecated();
    sHTML = oEditor.document.body.innerHTML;
    sHTML = String(sHTML).replace(/\<PARAM NAME=\"Play\" VALUE=\"0\">/ig, "<PARAM NAME=\"Play\" VALUE=\"-1\">");
    sHTML = oUtil.replaceSpecialChar(sHTML);
    if (this.encodeIO) sHTML = encodeHTMLCode(sHTML);
    sHTML = sHTML.replace(/\u00A0+/gi, " ");
    sHTML = cleanHTML(sHTML);
    if (!bEdit) {
        sHTML = sHTML.replace(/ contenteditable=\"true\"/ig, "");
        sHTML = sHTML.replace(/ contenteditable=\"false\"/ig, "")
    }
    return sHTML
};
var sBaseHREF = "";

function getXHTML(bEdit) {
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    this.cleanDeprecated();
    sHTML = getOuterHTML(oEditor.document.documentElement);
    var arrTmp = sHTML.match(/<BASE([^>]*)>/ig);
    if (arrTmp != null) sBaseHREF = arrTmp[0];
    var arrBase = oEditor.document.getElementsByTagName("BASE");
    if (arrBase.length != null) {
        for (var i = 0; i < arrBase.length; i++) {
            arrBase[i].parentNode.removeChild(arrBase[i])
        }
    }
    sBaseHREF = sBaseHREF.replace(/<([^ >]*)/ig, function (x) {
        return x.toLowerCase()
    });
    sBaseHREF = sBaseHREF.replace(/ [^=]+="[^"]+"/ig, function (x) {
        x = x.replace(/\s+/ig, "#_#");
        x = x.replace(/^#_#/, " ");
        return x
    });
    sBaseHREF = sBaseHREF.replace(/ ([^=]+)=([^" >]+)/ig, " $1=\"$2\"");
    sBaseHREF = sBaseHREF.replace(/ ([^=]+)=/ig, function (x) {
        return x.toLowerCase()
    });
    sBaseHREF = sBaseHREF.replace(/#_#/ig, " ");
    sBaseHREF = sBaseHREF.replace(/>$/ig, " \/>").replace(/\/ \/>$/ig, "\/>");
    sHTML = recur(oEditor.document.documentElement, "");
    sHTML = this.docType + this.html + sHTML + "\n</html>";
    sHTML = sHTML.replace(/<head>/i, "<head>" + sBaseHREF);
    sHTML = sHTML.replace(/\u00A0+/gi, " ");
    sHTML = oUtil.replaceSpecialChar(sHTML);
    if (this.encodeIO) sHTML = encodeHTMLCode(sHTML);
    sHTML = cleanHTML(sHTML);
    if (!bEdit) {
        sHTML = sHTML.replace(/ contenteditable=\"true\"/ig, "");
        sHTML = sHTML.replace(/ contenteditable=\"false\"/ig, "")
    }
    return sHTML
};

function getXHTMLBody(bEdit) {
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    this.cleanDeprecated();
    sHTML = getOuterHTML(oEditor.document.documentElement);
    var arrTmp = sHTML.match(/<BASE([^>]*)>/ig);
    if (arrTmp != null) sBaseHREF = arrTmp[0];
    sHTML = recur(oEditor.document.body, "");
    sHTML = oUtil.replaceSpecialChar(sHTML);
    if (this.encodeIO) sHTML = encodeHTMLCode(sHTML);
    sHTML = sHTML.replace(/\u00A0+/gi, " ");
    sHTML = cleanHTML(sHTML);
    if (!bEdit) {
        sHTML = sHTML.replace(/ contenteditable=\"true\"/ig, "");
        sHTML = sHTML.replace(/ contenteditable=\"false\"/ig, "")
    }
    return sHTML
};

function ApplyCSS(oName) {
    var sTmp = "";
    var myStyle = document.getElementById("myStyle" + oName).contentWindow;
    for (var i = 0; i < myStyle.document.styleSheets[0].cssRules.length; i++) {
        var sSelector = myStyle.document.styleSheets[0].cssRules[i].selectorText;
        if (sSelector != undefined) {
            sCssText = myStyle.document.styleSheets[0].cssRules[i].style.cssText.replace(/"/g, "&quot;");
            var itemCount = sSelector.split(".").length;
            if (itemCount > 1) {
                sCaption = sSelector.split(".")[1];
                sTmp += ",[\"" + sSelector + "\",true,\"" + sCaption + "\",\"" + sCssText + "\"]"
            } else sTmp += ",[\"" + sSelector + "\",false,\"\",\"" + sCssText + "\"]"
        }
    }
    var arrStyle = eval("[" + sTmp.substr(1) + "]");
    if (arrStyle.length > 0) {
        var oEditor = document.getElementById("idContent" + oName).contentWindow;
        var oElement = oEditor.document.createElement("STYLE");
        oEditor.document.documentElement.childNodes[0].appendChild(oElement);
        var sCssText = "";
        for (var i = 0; i < arrStyle.length; i++) {
            selector = arrStyle[i][0];
            style = arrStyle[i][3].replace(/&quot;/g, "\"");
            sCssText += selector + " { " + style + " } "
        }
        oElement.appendChild(oEditor.document.createTextNode(sCssText))
    }
};

function ApplyExternalStyle(oName) {
    var edtObj = eval(oName);
    var oEditor = document.getElementById("idContent" + oName).contentWindow;
    var edtHost = oEditor.location.href.match(/:\/\/(.[^/]+)/)[1];
    var sTmp = "";
    var prefixes = edtObj.styleSelectorPrefix.split(","),
        found = false;
    for (var j = 0; j < oEditor.document.styleSheets.length; j++) {
        var myStyle = oEditor.document.styleSheets[j];
        if (myStyle.href && myStyle.href.match(/:\/\/(.[^/]+)/)[1] != edtHost) {
            continue
        }
        for (var i = 0; i < myStyle.cssRules.length; i++) {
            sSelector = myStyle.cssRules[i].selectorText;
            if (sSelector != undefined) {
                if (sSelector.match(/table\./gi)) {
                    continue
                }
                found = false;
                if (prefixes.length > 0) {
                    for (var ix = 0; ix < prefixes.length; ix++) {
                        if (prefixes[ix] != "" && sSelector.indexOf("." + prefixes[ix]) >= 0) {
                            found = true;
                            break
                        }
                    }
                }
                if (!found && edtObj.styleSelectorPrefix != "") continue;
                sSelector = sSelector.replace(/"/g, "\\\"");
                sCssText = myStyle.cssRules[i].style.cssText.replace(/"/g, "&quot;");
                var itemCount = sSelector.split(".").length;
                if (itemCount > 1) {
                    sCaption = sSelector.split(".")[1];
                    sTmp += ",[\"" + sSelector + "\",true,\"" + sCaption + "\",\"" + sCssText + "\"]"
                } else sTmp += ",[\"" + sSelector + "\",false,\"\",\"" + sCssText + "\"]"
            }
        }
    }
    var arrStyle = eval("[" + sTmp.substr(1) + "]");
    for (var i = 0; i < arrStyle.length; i++) {
        for (var j = 0; j < edtObj.arrStyle.length; j++) {
            if (arrStyle[i][0].toLowerCase() == edtObj.arrStyle[j][0].toLowerCase()) {
                arrStyle[i][1] = edtObj.arrStyle[j][1]
            }
        }
    }
    edtObj.arrStyle = arrStyle
};

function doApplyStyle(oName, sClassName) {
    var oEditor = document.getElementById("idContent" + oName).contentWindow;
    var oSel = oEditor.getSelection();
    eval(oName).saveForUndo();
    var element;
    if (oUtil.activeElement) {
        element = oUtil.activeElement;
        element.className = sClassName
    } else {
        element = getSelectedElement(oSel);
        if (isTextSelected(oSel)) {
            if (oSel != "") {
                eval(oName).applySpanStyle([], sClassName)
            } else {
                if (element.tagName == "BODY") return;
                element.className = sClassName
            }
        } else {
            if (element.tagName == "BODY") return;
            element.className = sClassName
        }
    }
    realTime(eval(oName))
};

function openStyleSelect() {
    if (!this.isCssLoaded) ApplyExternalStyle(this.oName);
    this.isCssLoaded = true;
    var idStyles = document.getElementById("idStyles" + this.oName);
    if (idStyles.innerHTML != "") {
        if (idStyles.style.display == "") idStyles.style.display = "none";
        else idStyles.style.display = "";
        return
    }
    idStyles.style.display = "";
    var h = document.getElementById("idContent" + this.oName).offsetHeight - 27;
    var arrStyle = this.arrStyle;
    var sHTML = "";
    sHTML += "<div unselectable=on style='margin:5px;margin-top:0px;margin-right:1px' align='left'>";
    sHTML += "<table style='margin:1px;margin-top:2px;margin-bottom:3px;width:14px;height:14px;background:#f4f4f4;' cellpadding=0 cellspacing=0 unselectable=on>" + "<tr><td onclick=\"" + this.oName + ".openStyleSelect();\" onmouseover=\"this.style.border='#708090 1px solid';this.style.color='white';this.style.backgroundColor='9FA7BB'\" onmouseout=\"this.style.border='white 1px solid';this.style.color='black';this.style.backgroundColor=''\" style=\"cursor:default;padding:1px;border:white 1px solid;font-family:verdana;font-size:10px;color:#000000;line-height:9px;\" align=center valign=top unselectable=on>x</td></tr>" + "</table>";
    var sBody = "";
    for (var i = 0; i < arrStyle.length; i++) {
        sSelector = arrStyle[i][0];
        if (sSelector == "body") sBody = arrStyle[i][3]
    }
    sHTML += "<div unselectable=on style='overflow:auto;width:200px;height:" + h + "px;'>";
    sHTML += "<table name='tblStyles" + this.oName + "' id='tblStyles" + this.oName + "' cellpadding=0 cellspacing=0 style='background:#fcfcfc;" + sBody + ";height:100%;width:100%;margin:0;'>";
    for (var i = 0; i < arrStyle.length; i++) {
        sSelector = arrStyle[i][0];
        isOnSelection = arrStyle[i][1];
        sCssText = arrStyle[i][3];
        sCaption = arrStyle[i][2];
        if (isOnSelection) {
            if (sSelector.split(".").length > 1) {
                var tmpSelector = sSelector;
                if (sSelector.indexOf(":") > 0) tmpSelector = sSelector.substring(0, sSelector.indexOf(":"));
                if (tmpSelector.indexOf("awesome") == -1) {
                    sHTML += "<tr style=\"cursor:default\" onmouseover=\"if(this.style.marginRight!='1px'){this.style.background='" + this.styleSelectionHoverBg + "';this.style.color='" + this.styleSelectionHoverFg + "'}\" onmouseout=\"if(this.style.marginRight!='1px'){this.style.background='';this.style.color=''}\">";
                    sHTML += "<td unselectable=on onclick=\"doApplyStyle('" + this.oName + "','" + tmpSelector.split(".")[1] + "')\" style='padding:2px'>";
                    if (sSelector.split(".")[0] == "") sHTML += "<span unselectable=on style=\"" + sCssText + ";margin:0;\">" + sCaption + "</span>";
                    else sHTML += "<span unselectable=on style=\"" + sCssText + ";margin:0;\">" + sSelector + "</span>";
                    sHTML += "</td></tr>"
                }
            }
        }
    }
    sHTML += "<tr><td height=100%>&nbsp;</td></tr></table></div>";
    sHTML += "</div>";
    document.getElementById("idStyles" + this.oName).innerHTML = sHTML
};

function cleanFonts() {
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    var allFonts = oEditor.document.body.getElementsByTagName("FONT");
    if (allFonts.length == 0) return false;
    var f;
    var range;
    while (allFonts.length > 0) {
        f = allFonts[0];
        if (f.hasChildNodes && f.childNodes.length == 1 && f.childNodes[0].nodeType == 1 && f.childNodes[0].nodeName == "SPAN") {
            var theSpan = f.childNodes[0];
            copyAttribute(theSpan, f);
            range = oEditor.document.createRange();
            range.selectNode(f);
            range.insertNode(theSpan);
            range.selectNode(f);
            range.deleteContents()
        } else if (f.parentNode.nodeName == "SPAN" && f.parentNode.childNodes.length == 1) {
            var theSpan = f.parentNode;
            copyAttribute(theSpan, f);
            theSpan.innerHTML = f.innerHTML
        } else {
            var newSpan = oEditor.document.createElement("SPAN");
            copyAttribute(newSpan, f);
            newSpan.innerHTML = f.innerHTML;
            f.parentNode.replaceChild(newSpan, f)
        }
    }
    return true
};

function cleanTags(elements, sVal) {
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    if (elements.length == 0) return false;
    var f;
    var range;
    while (elements.length > 0) {
        f = elements[0];
        if (f.hasChildNodes && f.childNodes.length == 1 && f.childNodes[0].nodeType == 1 && f.childNodes[0].nodeName == "SPAN") {
            var theSpan = f.childNodes[0];
            if (sVal == "bold") theSpan.style.fontWeight = "bold";
            if (sVal == "italic") theSpan.style.fontStyle = "italic";
            if (sVal == "line-through") theSpan.style.textDecoration = "line-through";
            if (sVal == "underline") theSpan.style.textDecoration = "underline";
            range = oEditor.document.createRange();
            range.selectNode(f);
            range.insertNode(theSpan);
            range.selectNode(f);
            range.deleteContents()
        } else if (f.parentNode.nodeName == "SPAN" && f.parentNode.childNodes.length == 1) {
            var theSpan = f.parentNode;
            if (sVal == "bold") theSpan.style.fontWeight = "bold";
            if (sVal == "italic") theSpan.style.fontStyle = "italic";
            if (sVal == "line-through") theSpan.style.textDecoration = "line-through";
            if (sVal == "underline") theSpan.style.textDecoration = "underline";
            theSpan.innerHTML = f.innerHTML
        } else {
            var newSpan = oEditor.document.createElement("SPAN");
            if (sVal == "bold") newSpan.style.fontWeight = "bold";
            if (sVal == "italic") newSpan.style.fontStyle = "italic";
            if (sVal == "line-through") newSpan.style.textDecoration = "line-through";
            if (sVal == "underline") newSpan.style.textDecoration = "underline";
            newSpan.innerHTML = f.innerHTML;
            f.parentNode.replaceChild(newSpan, f)
        }
    }
    return true
};

function replaceTags(sFrom, sTo) {
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    var elements = oEditor.document.body.getElementsByTagName(sFrom);
    while (elements.length > 0) {
        f = elements[0];
        var newSpan = oEditor.document.createElement(sTo);
        newSpan.innerHTML = f.innerHTML;
        f.parentNode.replaceChild(newSpan, f)
    }
};

function cleanDeprecated() {
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    var elements;
    elements = oEditor.document.body.getElementsByTagName("B");
    this.cleanTags(elements, "bold");
    elements = oEditor.document.body.getElementsByTagName("STRIKE");
    this.cleanTags(elements, "line-through");
    elements = oEditor.document.body.getElementsByTagName("S");
    this.cleanTags(elements, "line-through");
    elements = oEditor.document.body.getElementsByTagName("U");
    this.cleanTags(elements, "underline");
    this.replaceTags("DIR", "DIV");
    this.replaceTags("MENU", "DIV");
    this.replaceTags("CENTER", "DIV");
    this.replaceTags("XMP", "PRE");
    this.replaceTags("BASEFONT", "SPAN");
    elements = oEditor.document.body.getElementsByTagName("APPLET");
    while (elements.length > 0) {
        var f = elements[0];
        theParent = f.parentNode;
        theParent.removeChild(f)
    }
    this.cleanFonts();
    this.cleanEmptySpan();
    return true
};

function applySpanStyle(arrStyles, sClassName, blockTag) {
    var useBlock = "SPAN";
    if (blockTag != null && blockTag != "") useBlock = blockTag;
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    var oSel = oEditor.getSelection();
    var range;
    var oElement;
    if (!isTextSelected(oSel)) {
        range = oSel.getRangeAt(0);
        oElement = getSelectedElement(oSel);
        if (oElement.nodeName == useBlock) return oElement;
        return
    }
    this.hide();
    this.doCmd("fontname", "verdana");
    oElement = getSelectedElement(oSel);
    oElement.style.fontFamily = "";
    replaceWithSpan(oEditor, arrStyles, sClassName, useBlock);
    realTime(this)
};

function doClean() {
    this.saveForUndo();
    if (oUtil.activeElement) var element = oUtil.activeElement;
    else {
        var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
        var oSel = oEditor.getSelection();
        element = getSelectedElement(oSel);
        if (isTextSelected(oSel)) {
            if (oSel != "") {
                var range = oSel.getRangeAt(0);
                if (element.innerHTML != range.toString()) {
                    this.doCmd('RemoveFormat');
                    realTime(this);
                    return
                }
            } else {
                if (element.tagName == "BODY") return
            }
        } else {
            if (element.tagName == "BODY") return
        }
    }
    element.removeAttribute("className");
    element.removeAttribute("style");
    if (element.tagName == "H1" || element.tagName == "H2" || element.tagName == "H3" || element.tagName == "H4" || element.tagName == "H5" || element.tagName == "H6" || element.tagName == "PRE" || element.tagName == "P" || element.tagName == "DIV") {
        this.doCmd('FormatBlock', '<P>')
    }
    this.doCmd('RemoveFormat');
    realTime(this)
};

function cleanEmptySpan() {
    var bReturn = false;
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    var reg = /<\s*SPAN\s*>/gi;
    while (true) {
        var allSpans = oEditor.document.getElementsByTagName("SPAN");
        if (allSpans.length == 0) break;
        var emptySpans = [];
        for (var i = 0; i < allSpans.length; i++) {
            if (allSpans[i].className == "Apple-style-span") allSpans[i].removeAttribute("class", 0);
            if (allSpans[i].style.cssText == "") allSpans[i].removeAttribute("style", 0);
            if (getOuterHTML(allSpans[i]).search(reg) == 0) emptySpans[emptySpans.length] = allSpans[i]
        }
        if (emptySpans.length == 0) break;
        var theSpan, theParent;
        for (var i = 0; i < emptySpans.length; i++) {
            theSpan = emptySpans[i];
            theParent = theSpan.parentNode;
            if (!theParent) continue;
            if (theSpan.hasChildNodes()) {
                var range = oEditor.document.createRange();
                range.selectNodeContents(theSpan);
                var docFrag = range.extractContents();
                theParent.replaceChild(docFrag, theSpan)
            } else {
                theParent.removeChild(theSpan)
            }
            bReturn = true
        }
    }
    return bReturn
};

function copyStyleClass(newSpan, arrStyles, sClassName) {
    if (arrStyles) {
        for (var i = 0; i < arrStyles.length; i++) {
            newSpan.style[arrStyles[i][0]] = arrStyles[i][1]
        }
        if (newSpan.style.cssText == "") newSpan.removeAttribute("style", 0)
    }
    if (sClassName) newSpan.className = sClassName
};

function copyAttribute(newSpan, f) {
    if ((f.face != null) && (f.face != "")) newSpan.style.fontFamily = f.face;
    if ((f.size != null) && (f.size != "")) {
        var nSize = "";
        if (f.size == 1) nSize = "8pt";
        else if (f.size == 2) nSize = "10pt";
        else if (f.size == 3) nSize = "12pt";
        else if (f.size == 4) nSize = "14pt";
        else if (f.size == 5) nSize = "18pt";
        else if (f.size == 6) nSize = "24pt";
        else if (f.size >= 7) nSize = "36pt";
        else if (f.size <= -2 || f.size == "0") nSize = "8pt";
        else if (f.size == "-1") nSize = "10pt";
        else if (f.size == 0) nSize = "12pt";
        else if (f.size == "+1") nSize = "14pt";
        else if (f.size == "+2") nSize = "18pt";
        else if (f.size == "+3") nSize = "24pt";
        else if (f.size == "+4" || f.size == "+5" || f.size == "+6") nSize = "36pt";
        else nSize = ""; if (nSize != "") newSpan.style.fontSize = nSize
    }
    if ((f.style.backgroundColor != null) && (f.style.backgroundColor != "")) newSpan.style.backgroundColor = f.style.backgroundColor;
    if ((f.color != null) && (f.color != "")) newSpan.style.color = f.color;
    if ((f.className != null) && (f.className != "")) newSpan.className = f.className
};

function replaceWithSpan(oEditor, arrStyles, sClassName, blockTag) {
    var useBlock = "SPAN";
    if (blockTag != null && blockTag != "") useBlock = blockTag;
    var oSel = oEditor.getSelection();
    var elm = getSelectedElement(oSel);
    if (arrStyles || sClassName) copyStyleClass(elm, arrStyles, sClassName)
};

function editorDoc_onkeyup(oEdt) {
    if (oEdt.tmKeyup) {
        clearTimeout(oEdt.tmKeyup);
        oEdt.tmKeyup = null
    }
    if (!oEdt.tmKeyup) oEdt.tmKeyup = setTimeout(function () {
        realTime(oEdt)
    }, 1000)
};

function editorDoc_onmouseup(e, oName) {
    var edtObj = eval(oName);
    var edtFrm = document.getElementById("idContent" + oName);
    oUtil.activeElement = null;
    oUtil.oName = oName;
    oUtil.oEditor = edtFrm.contentWindow;
    oUtil.obj = edtObj;
    edtObj.hide();
    if (e.target && e.target.tagName && e.target.tagName == 'IMG') {
        var oSel = edtFrm.contentWindow.getSelection();
        var range = oSel.getRangeAt(0);
        oSel.removeAllRanges();
        range.selectNodeContents(e.target);
        range.setEndAfter(e.target);
        oSel.addRange(range)
    }
    realTime(edtObj)
};

function setActiveEditor(edtObj) {
    oUtil.oName = edtObj.oName;
    oUtil.oEditor = document.getElementById("idContent" + edtObj.oName).contentWindow;
    oUtil.obj = edtObj
};
var arrTmp = [];

function GetElement(oElement, sMatchTag) {
    while (oElement != null && oElement.tagName != sMatchTag) {
        if (oElement.tagName == "BODY") return null;
        oElement = oElement.parentNode
    }
    return oElement
};
var arrTmp2 = [];

function realTime(edtObj) {
    try {
        var oName = edtObj.oName;
        var oEditor = document.getElementById("idContent" + oName).contentWindow;
        var oSel = oEditor.getSelection();
        var element = getSelectedElement(oSel);
        var obj = edtObj;
        var tbar = obj.tbar;
        var btn = null;
        var oTable = GetElement(element, "TABLE");
        if (oTable) obj.Table = oTable;
        var doc = oEditor.document;
        if (obj.btnParagraph) {
            btn = tbar.btns["btnParagraph" + oName];
            btn.setState(doc.queryCommandEnabled("FormatBlock") ? 1 : 5)
        }
        if (obj.btnFontName) {
            btn = tbar.btns["btnFontName" + oName];
            btn.setState(doc.queryCommandEnabled("FontName") ? 1 : 5)
        }
        if (obj.btnFontSize) {
            btn = tbar.btns["btnFontSize" + oName];
            btn.setState(doc.queryCommandEnabled("FontSize") ? 1 : 5)
        }
        if (obj.btnUndo) {
            btn = tbar.btns["btnUndo" + oName];
            btn.setState(!obj.arrUndoList[0] ? 5 : 1)
        }
        if (obj.btnRedo) {
            btn = tbar.btns["btnRedo" + oName];
            btn.setState(!obj.arrRedoList[0] ? 5 : 1)
        }
        if (obj.btnBold) {
            btn = tbar.btns["btnBold" + oName];
            btn.setState(doc.queryCommandEnabled("Bold") ? (doc.queryCommandState("Bold") ? 4 : 1) : 5)
        }
        if (obj.btnItalic) {
            btn = tbar.btns["btnItalic" + oName];
            btn.setState(doc.queryCommandEnabled("Italic") ? (doc.queryCommandState("Italic") ? 4 : 1) : 5)
        }
        if (obj.btnUnderline) {
            btn = tbar.btns["btnUnderline" + oName];
            btn.setState(doc.queryCommandEnabled("Underline") ? (doc.queryCommandState("Underline") ? 4 : 1) : 5)
        }
        if (obj.btnStrikethrough) {
            btn = tbar.btns["btnStrikethrough" + oName];
            btn.setState(doc.queryCommandEnabled("Strikethrough") ? (doc.queryCommandState("Strikethrough") ? 4 : 1) : 5)
        }
        if (obj.btnSuperscript) {
            btn = tbar.btns["btnSuperscript" + oName];
            btn.setState(doc.queryCommandEnabled("Superscript") ? (doc.queryCommandState("Superscript") ? 4 : 1) : 5)
        }
        if (obj.btnSubscript) {
            btn = tbar.btns["btnSubscript" + oName];
            btn.setState(doc.queryCommandEnabled("Subscript") ? (doc.queryCommandState("Subscript") ? 4 : 1) : 5)
        }
        if (obj.btnNumbering) {
            btn = tbar.btns["btnNumbering" + oName];
            btn.setState(doc.queryCommandEnabled("InsertOrderedList") ? (doc.queryCommandState("InsertOrderedList") ? 4 : 1) : 5)
        }
        if (obj.btnBullets) {
            btn = tbar.btns["btnBullets" + oName];
            btn.setState(doc.queryCommandEnabled("InsertUnorderedList") ? (doc.queryCommandState("InsertUnorderedList") ? 4 : 1) : 5)
        }
        if (obj.btnJustifyLeft) {
            btn = tbar.btns["btnJustifyLeft" + oName];
            btn.setState(doc.queryCommandEnabled("JustifyLeft") ? (doc.queryCommandState("JustifyLeft") ? 4 : 1) : 5)
        }
        if (obj.btnJustifyCenter) {
            btn = tbar.btns["btnJustifyCenter" + oName];
            btn.setState(doc.queryCommandEnabled("JustifyCenter") ? (doc.queryCommandState("JustifyCenter") ? 4 : 1) : 5)
        }
        if (obj.btnJustifyRight) {
            btn = tbar.btns["btnJustifyRight" + oName];
            btn.setState(doc.queryCommandEnabled("JustifyRight") ? (doc.queryCommandState("JustifyRight") ? 4 : 1) : 5)
        }
        if (obj.btnJustifyFull) {
            btn = tbar.btns["btnJustifyFull" + oName];
            btn.setState(doc.queryCommandEnabled("JustifyFull") ? (doc.queryCommandState("JustifyFull") ? 4 : 1) : 5)
        }
        if (obj.btnIndent) {
            btn = tbar.btns["btnIndent" + oName];
            btn.setState(doc.queryCommandEnabled("Indent") ? 1 : 5)
        }
        if (obj.btnOutdent) {
            btn = tbar.btns["btnOutdent" + oName];
            btn.setState(doc.queryCommandEnabled("Outdent") ? 1 : 5)
        }
        if (obj.btnLTR) {
            btn = tbar.btns["btnLTR" + oName];
            btn.setState(element.dir ? (element.dir.toLowerCase() == "ltr" ? 4 : 1) : 1)
        }
        if (obj.btnRTL) {
            btn = tbar.btns["btnRTL" + oName];
            btn.setState(element.dir ? (element.dir.toLowerCase() == "rtl" ? 4 : 1) : 1)
        }
        var v = (element ? 1 : 5);
        if (obj.btnForeColor) tbar.btns["btnForeColor" + oName].setState(v);
        if (obj.btnBackColor) tbar.btns["btnBackColor" + oName].setState(v);
        if (obj.btnLine) tbar.btns["btnLine" + oName].setState(v);
        try {
            oUtil.onSelectionChanged()
        } catch (e) {}
        try {
            obj.onSelectionChanged()
        } catch (e) {}
        var idStyles = document.getElementById("idStyles" + oName);
        if (idStyles.innerHTML != "") {
            var oElement;
            if (oUtil.activeElement) oElement = oUtil.activeElement;
            else oElement = getSelectedElement(oSel);
            var sCurrClass = oElement.className;
            var oRows = document.getElementById("tblStyles" + oName).rows;
            for (var i = 0; i < oRows.length - 1; i++) {
                sClass = oRows[i].childNodes[0].childNodes[0].innerHTML;
                if (sClass.split(".").length > 1 && sClass != "") sClass = sClass.split(".")[1];
                if (sCurrClass == sClass) {
                    oRows[i].style.marginRight = "1px";
                    oRows[i].style.backgroundColor = obj.styleSelectionHoverBg;
                    oRows[i].style.color = obj.styleSelectionHoverFg
                } else {
                    oRows[i].style.marginRight = "";
                    oRows[i].style.backgroundColor = "";
                    oRows[i].style.color = ""
                }
            }
        }
        if (obj.useTagSelector) {
            oElement = element;
            var sHTML = "";
            var i = 0;
            arrTmp2 = [];
            while (oElement != null && oElement.tagName != "BODY" && oElement.nodeType == 1) {
                arrTmp2[i] = oElement;
                sHTML = "&nbsp; &lt;<span id=tag" + oName + i + " unselectable=on style='text-decoration:underline;cursor:pointer' onclick=\"" + oName + ".selectElement(" + i + ")\">" + oElement.tagName + "</span>&gt;" + sHTML;
                oElement = oElement.parentNode;
                i++
            }
            sHTML = "&nbsp;&lt;BODY&gt;" + sHTML;
            document.getElementById("idElNavigate" + oName).innerHTML = sHTML;
            document.getElementById("idElCommand" + oName).style.display = "none";
            for (i = 0; i < arrTmp2.length; i++) {
                document.getElementById("tag" + oName + i).addEventListener("click", new Function(oName + ".selectElement(" + i + ")"), true)
            }
        }
        btn = tbar.btns["btnQuote" + oName];
        if (btn) {
            var oQuote = GetElement(element, "BLOCKQUOTE");
            btn.setState(oQuote ? 4 : 1)
        }
    } catch (e) {}
};

function realtimeFontSelect(oName) {
    var oEditor = document.getElementById("idContent" + oName).contentWindow;
    var sFontName = oEditor.document.queryCommandValue("FontName");
    var edt = eval(oName);
    var found = false;
    for (var i = 0; i < edt.arrFontName.length; i++) {
        if (sFontName == edt.arrFontName[i]) {
            found = true;
            break
        }
    }
    if (found) {
        isDDs["ddFontName" + oName].selectItem("btnFontName_" + i + oName, true)
    } else {
        isDDs["ddFontName" + oName].clearSelection()
    }
};

function realtimeSizeSelect(oName) {
    var oEditor = document.getElementById("idContent" + oName).contentWindow;
    var sFontSize = oEditor.document.queryCommandValue("FontSize");
    var edt = eval(oName);
    var found = false;
    for (var i = 0; i < edt.arrFontSize.length; i++) {
        if (sFontSize == edt.arrFontSize[i][1]) {
            found = true;
            break
        }
    }
    if (found) {
        isDDs["ddFontSize" + oName].selectItem("btnFontSize_" + i + oName, true)
    } else {
        isDDs["ddFontSize" + oName].clearSelection()
    }
};

function moveTagSelector() {
    var edtArea = document.getElementById("idArea" + this.oName);
    var nWidth = edtArea.offsetWidth - 90;
    var icPath = this.scriptPath.substring(0, this.scriptPath.indexOf("saf/")) + this.iconPath;
    var sTagSelTop = "<table unselectable=on ondblclick='" + this.oName + ".moveTagSelector()' width='100%' cellpadding=0 cellspacing=0 style='border:none;border:#cfcfcf 1px solid;border-bottom:none;margin:0px'><tr style='background-color:#f4f4f4;font-family:arial;font-size:10px;color:black;'>" + "<td id=idElNavigate" + this.oName + " style='color:inherit;line-height:normal;font-size:inherit;padding:1px;width:" + nWidth + "px' valign=top>&nbsp;</td>" + "<td align=right valign='center' nowrap style='color:inherit;padding:0;margin;0;font-size:inherit;line-height:normal'>" + "<span id=idElCommand" + this.oName + " unselectable=on style='vertical-align:middle;display:none;text-decoration:underline;cursor:pointer;padding-right:5px;' onclick='" + this.oName + ".removeTag()'>" + getTxt("Remove Tag") + "</span>" + "</td></tr></table>";
    var sTagSelBottom = "<table unselectable=on ondblclick='" + this.oName + ".moveTagSelector()' width='100%' cellpadding=0 cellspacing=0 style='border:none;border-left:#cfcfcf 1px solid;border-right:#cfcfcf 1px solid;margin:0px'><tr style='background-color:#f4f4f4;font-family:arial;font-size:10px;color:black;'>" + "<td id=idElNavigate" + this.oName + " style='color:inherit;line-height:normal;font-size:inherit;padding:1px;width:" + nWidth + "px' valign=top>&nbsp;</td>" + "<td align=right valign='center' nowrap style='color:inherit;padding:0;margin;0;font-size:inherit;line-height:normal'>" + "<span id=idElCommand" + this.oName + " unselectable=on style='vertical-align:middle;display:none;text-decoration:underline;cursor:pointer;padding-right:5px;' onclick='" + this.oName + ".removeTag()'>" + getTxt("Remove Tag") + "</span>" + "</td></tr></table>";
    var selTop = document.getElementById("idTagSelTop" + this.oName);
    var selTopRow = document.getElementById("idTagSelTopRow" + this.oName);
    var selBottom = document.getElementById("idTagSelBottom" + this.oName);
    var selBottomRow = document.getElementById("idTagSelBottomRow" + this.oName);
    if (this.TagSelectorPosition == "top") {
        selTop.innerHTML = "";
        selBottom.innerHTML = sTagSelBottom;
        selTopRow.style.display = "none";
        selBottomRow.style.display = "";
        this.TagSelectorPosition = "bottom"
    } else {
        selTop.innerHTML = sTagSelTop;
        selBottom.innerHTML = "";
        selTopRow.style.display = "";
        selBottomRow.style.display = "none";
        this.TagSelectorPosition = "top"
    }
};

function selectElement(i) {
    if (!this.useTagSelector) return;
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    var range = oEditor.document.createRange();
    range.selectNode(arrTmp2[i]);
    var oSel = oEditor.getSelection();
    oSel.removeAllRanges();
    oSel.addRange(range);
    oActiveElement = arrTmp2[i];
    if (oActiveElement.tagName != "TD" && oActiveElement.tagName != "TR" && oActiveElement.tagName != "TBODY" && oActiveElement.tagName != "LI") document.getElementById("idElCommand" + this.oName).style.display = "";
    for (var j = 0; j < arrTmp2.length; j++) document.getElementById("tag" + this.oName + j).style.background = "";
    document.getElementById("tag" + this.oName + i).style.background = "DarkGray";
    if (oActiveElement) oUtil.activeElement = oActiveElement;
    this.focus()
};

function removeTag() {
    this.saveForUndo();
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    var oSel = oEditor.getSelection();
    var element = getSelectedElement(oSel);
    var nearElement = element.nextSibling == null ? (element.previousSibling == null ? element.parentNode : element.previousSibling) : element.nextSibling;
    switch (element.nodeName) {
    case "TABLE":
        ;
    case "IMG":
        ;
    case "INPUT":
        ;
    case "FORM":
        ;
    case "SELECT":
        element.parentNode.removeChild(element);
        break;
    default:
        oSel = oEditor.getSelection();
        var range = oSel.getRangeAt(0);
        var docFrag = range.createContextualFragment(element.innerHTML);
        range.selectNode(element);
        range.deleteContents();
        range.insertNode(docFrag);
        try {
            oEditor.document.designMode = "on"
        } catch (e) {}
        break
    }
    oSel = oEditor.getSelection();
    oSel.removeAllRanges();
    var range = oEditor.document.createRange();
    range.setStart(nearElement, 0);
    range.setEnd(nearElement, 0);
    oSel.addRange(range);
    this.focus();
    realTime(this)
};

function doCmd(sCmd, sOption) {
    if (sCmd == "Bold" || sCmd == "Italic" || sCmd == "Underline" || sCmd == "Strikethrough" || sCmd == "Superscript" || sCmd == "Subscript" || sCmd == "Indent" || sCmd == "Outdent" || sCmd == "InsertHorizontalRule") this.saveForUndo();
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    oEditor.document.execCommand("styleWithCSS", false, true);
    oEditor.document.execCommand(sCmd, false, sOption);
    realTime(this)
};

function applyColor(sType, sColor) {
    this.hide();
    this.focus();
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    this.saveForUndo();
    if (sColor != null && sColor != "") {
        oEditor.document.execCommand("styleWithCSS", false, true);
        oEditor.document.execCommand(sType, false, sColor);
        var sel = oEditor.getSelection();
        var range = sel.getRangeAt(0);
        if (range.startContainer.nodeType == Node.TEXT_NODE) {
            var el = range.startContainer.nextSibling;
            if (el) {
                range = oEditor.document.createRange();
                range.selectNode(el);
                sel = oEditor.getSelection();
                sel.removeAllRanges();
                sel.addRange(range)
            }
        }
    } else {
        var el = getSelectedElement(oEditor.getSelection());
        if (sType == "ForeColor") {
            el.style.color = ""
        } else if (sType == "HiliteColor") {
            el.style.backgroundColor = ""
        }
    }
    realTime(this);
    if (sColor == "") return
};

function applyParagraph(val) {
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    var oSel = oEditor.getSelection();
    this.hide();
    this.saveForUndo();
    this.doCmd("FormatBlock", val)
};

function applyFontName(val) {
    this.hide();
    this.saveForUndo();
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    var oSel = oEditor.getSelection();
    var oElement = getSelectedElement(oSel);
    this.doCmd("fontname", val)
};

function applyFontSize(val) {
    this.hide();
    this.saveForUndo();
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    var oSel = oEditor.getSelection();
    this.doCmd("fontsize", 1);
    var range = oSel.getRangeAt(0);
    var commonParent = range.commonAncestorContainer;
    if (commonParent.nodeType != 1) {
        commonParent = commonParent.parentNode
    }
    if (commonParent != null) {
        commonParent.style.fontSize = val
    }
    var spans = commonParent.getElementsByTagName("SPAN");
    for (var i = 0; i < spans.length; i++) {
        if (oSel.containsNode(spans[i], true)) {
            spans[i].style.fontSize = val
        }
    }
    realTime(this)
};

function mapFontSize(v) {
    var s = "8pt";
    if (v == 1) s = "8pt";
    else if (v == 2) s = "10pt";
    else if (v == 3) s = "12pt";
    else if (v == 4) s = "14pt";
    else if (v == 5) s = "18pt";
    else if (v == 6) s = "24pt";
    else if (v >= 7) s = "36pt";
    else if (v <= -2 || v == "0") s = "8pt";
    else if (v == "-1") s = "10pt";
    else if (v == 0) s = "12pt";
    else if (v == "+1") s = "14pt";
    else if (v == "+2") s = "18pt";
    else if (v == "+3") s = "24pt";
    else if (v == "+4" || v == "+5" || v == "+6") s = "36pt";
    return s
};

function applyFormattingStyle(cmd, opt) {
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    var oSel = oEditor.getSelection();
    var oElement = getSelectedElement(oSel);
    var doc = oEditor.document;
    var range = oSel.getRangeAt(0);
    if (oElement != null && oElement.tagName == "SPAN" && oElement.textContent == range.toString()) {
        if (cmd == "Bold" || cmd == "Italic" || cmd == "Underline" || cmd == "Strikethrough" || cmd == "Superscript" || cmd == "Subscript") {
            this.saveForUndo();
            switch (cmd) {
            case "Bold":
                oElement.style.fontWeight = (!doc.queryCommandState("Bold") ? "bold" : "");
                break;
            case "Italic":
                oElement.style.fontStyle = (!doc.queryCommandState("Italic") ? "italic" : "");
                break;
            case "Underline":
                oElement.style.textDecoration = (!doc.queryCommandState("Underline") ? "underline" : "");
                break;
            case "Strikethrough":
                oElement.style.textDecoration = (!doc.queryCommandState("Strikethrough") ? "line-through" : "");
                break;
            case "Superscript":
                oElement.style.verticalAlign = (!doc.queryCommandState("Superscript") ? "super" : "");
                break;
            case "Subscript":
                oElement.style.verticalAlign = (!doc.queryCommandState("Subscript") ? "sub" : "");
                break
            }
            realTime(this)
        }
    } else {
        this.doCmd(cmd, opt)
    }
};

function applyBullets() {
    this.saveForUndo();
    this.doCmd("InsertUnOrderedList");
    this.tbar.btns["btnNumbering" + this.oName].setState(1);
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    var oSel = oEditor.getSelection();
    var oElement = getSelectedElement(oSel);
    var li = GetElement(oElement, "LI");
    while (oElement != null && oElement.tagName != "OL" && oElement.tagName != "UL") {
        if (oElement.tagName == "BODY") return;
        oElement = oElement.parentNode
    }
    oElement.removeAttribute("type", 0);
    oElement.style.listStyleImage = "";
    var container = oElement.parentNode;
    if (container.firstChild == container.lastChild && container.tagName != "BODY") {
        container.parentNode.insertBefore(oElement, container);
        container.parentNode.removeChild(container)
    }
    if (li) {
        if (li.firstChild && (li.firstChild == li.lastChild) && li.firstChild.tagName == 'SPAN') {
            var old = li.removeChild(li.firstChild);
            li.innerHTML = old.innerHTML
        }
        if (li.textContent == "") li.innerHTML = "<br />";
        oSel = oEditor.getSelection();
        var range = oEditor.document.createRange();
        range.selectNodeContents(li);
        range.collapse(false);
        oSel.removeAllRanges();
        oSel.addRange(range)
    }
};

function applyNumbering() {
    this.saveForUndo();
    this.doCmd("InsertOrderedList");
    this.tbar.btns["btnBullets" + this.oName].setState(1);
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    var oSel = oEditor.getSelection();
    var oElement = getSelectedElement(oSel);
    var li = GetElement(oElement, "LI");
    while (oElement != null && oElement.tagName != "OL" && oElement.tagName != "UL") {
        if (oElement.tagName == "BODY") return;
        oElement = oElement.parentNode
    }
    oElement.removeAttribute("type", 0);
    oElement.style.listStyleImage = "";
    var container = oElement.parentNode;
    if (container.firstChild == container.lastChild && container.tagName != "BODY") {
        container.parentNode.insertBefore(oElement, container);
        container.parentNode.removeChild(container)
    }
    if (li) {
        if (li.firstChild && (li.firstChild == li.lastChild) && li.firstChild.tagName == 'SPAN') {
            var old = li.removeChild(li.firstChild);
            li.innerHTML = old.innerHTML
        }
        if (li.textContent == "") li.innerHTML = "";
        oSel = oEditor.getSelection();
        var range = oEditor.document.createRange();
        range.selectNodeContents(li);
        range.collapse(false);
        oSel.removeAllRanges();
        oSel.addRange(range)
    }
};

function applyOutdent() {
    this.doCmd("Outdent");
    cleanWebkitUselessSpan(this)
};

function applyJustifyLeft() {
    this.saveForUndo();
    this.doCmd("JustifyLeft")
};

function applyJustifyCenter() {
    this.saveForUndo();
    this.doCmd("JustifyCenter")
};

function applyJustifyRight() {
    this.saveForUndo();
    this.doCmd("JustifyRight")
};

function applyJustifyFull() {
    this.saveForUndo();
    this.doCmd("JustifyFull")
};

function applyBlockDirLTR() {
    this.saveForUndo();
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    var oSel = oEditor.getSelection();
    var oEl = getSelectedElement(oSel);
    if (oEl.dir) oEl.removeAttribute("dir");
    else oEl.dir = "ltr";
    this.focus()
};

function applyBlockDirRTL() {
    this.saveForUndo();
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    var oSel = oEditor.getSelection();
    var oEl = getSelectedElement(oSel);
    if (oEl.dir) oEl.removeAttribute("dir");
    else oEl.dir = "rtl";
    this.focus()
};

function insertCustomTag(index) {
    this.insertHTML(this.arrCustomTag[index][1]);
    this.hide();
    this.focus()
};

function expandSelection() {
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    var oSel = oEditor.getSelection();
    var range = oSel.getRangeAt(0);
    if (range.startContainer.nodeType == Node.TEXT_NODE) {
        if (range.toString() == "") {
            var sPos = range.startOffset;
            var ePos = range.endOffset;
            var startCont = range.startContainer;
            var str = startCont.nodeValue;
            sPos = str.substring(0, range.startOffset).search(/(\W+\w*)$/i);
            sPos = sPos == -1 ? 0 : sPos;
            var tPos = str.substring(sPos, range.startOffset).search(/\w+/i);
            sPos += (tPos == -1 ? str.substring(sPos, range.startOffset).length : tPos);
            ePos = str.substr(range.endOffset).search(/\W+/i);
            ePos = ePos == -1 ? str.length : ePos + range.endOffset;
            range = oEditor.document.createRange();
            range.setStart(startCont, sPos);
            range.setEnd(startCont, ePos);
            oSel = oEditor.getSelection();
            oSel.removeAllRanges();
            oSel.addRange(range)
        }
    }
    return
};

function selectParagraph() {
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    var oSel = oEditor.getSelection();
    var selParent = getSelectedElement(oSel);
    if (selParent) {
        if (oSel.getRangeAt(0).toString() == "") {
            var oElement = selParent;
            while (oElement != null && oElement.tagName != "H1" && oElement.tagName != "H2" && oElement.tagName != "H3" && oElement.tagName != "H4" && oElement.tagName != "H5" && oElement.tagName != "H6" && oElement.tagName != "PRE" && oElement.tagName != "P" && oElement.tagName != "DIV") {
                if (oElement.tagName == "BODY") return;
                oElement = oElement.parentNode
            }
            var range = oEditor.document.createRange();
            range.selectNode(oElement);
            oSel = oEditor.getSelection();
            oSel.removeAllRanges();
            oSel.addRange(range)
        }
    }
};

function insertHTML(sHTML) {
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    var oSel = oEditor.getSelection();
    var range = oSel.getRangeAt(0);
    this.saveForUndo();
    range.extractContents();
    range.collapse(true);
    var docFrag = range.createContextualFragment(sHTML);
    var lastNode = docFrag.lastChild;
    range.insertNode(docFrag);
    try {
        oEditor.document.designMode = "on"
    } catch (e) {}
    range.setStartAfter(lastNode);
    range.setEndAfter(lastNode);
    if (sHTML == "<br />" && range.endContainer.nodeType == 1) {
        if (range.endOffset == range.endContainer.childNodes.length - 1) {
            range.insertNode(range.createContextualFragment("<br />"));
            range.setStartAfter(lastNode);
            range.setEndAfter(lastNode)
        }
    }
    var comCon = range.commonAncestorContainer;
    if (comCon && comCon.parentNode) {
        try {
            comCon.parentNode.normalize()
        } catch (e) {}
    }
    oSel.removeAllRanges();
    oSel.addRange(range)
};

function insertLink(url, title, target, rel) {
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    var oSel = oEditor.getSelection();
    var range = oSel.getRangeAt(0);
    this.saveForUndo();
    var oImg = getSelectedElement(oSel);
    if (oImg) {
        if (oImg.tagName == "IMG") oImg.style.border = "none"
    }
    var emptySel = false;
    if (range.toString() == "") {
        if (range.startContainer.nodeType == Node.ELEMENT_NODE) {
            if (range.startContainer.childNodes[range.startOffset].nodeType != Node.TEXT_NODE) {
                if (range.startContainer.childNodes[range.startOffset].nodeName == "BR") emptySel = true;
                else emptySel = false
            } else {
                emptySel = true
            }
        } else {
            emptySel = true
        } if (emptySel) {
            var cap = (title != "" && title != null ? title : url);
            var node = oEditor.document.createTextNode(cap);
            range.insertNode(node);
            try {
                oEditor.document.designMode = "on"
            } catch (e) {}
            range = oEditor.document.createRange();
            range.setStart(node, 0);
            range.setEnd(node, cap.length);
            oSel = oEditor.getSelection();
            oSel.removeAllRanges();
            oSel.addRange(range)
        }
    }
    if (oImg && oImg.tagName == "IMG") {
        var newA = oEditor.document.createElement("A");
        newA.href = url;
        range.selectNode(oImg);
        range.surroundContents(newA);
        oImg.style.border = "none";
        range.selectNodeContents(oImg);
        range.setEndAfter(oImg);
        oSel.removeAllRanges();
        oSel.addRange(range)
    } else {
        var isSelInMidText = (range.startContainer.nodeType == Node.TEXT_NODE) && (range.startOffset > 0);
        oEditor.document.execCommand("styleWithCSS", false, true);
        oEditor.document.execCommand("CreateLink", false, url)
    }
    oSel = oEditor.getSelection();
    oEl = GetElement(getSelectedElement(oSel), "A");
    if (oEl) {
        if (target != "" && target != undefined) oEl.target = target;
        if (rel != "" && rel != undefined) oEl.setAttribute("rel", rel)
    }
};

function clearAll() {
    if (confirm(getTxt("Are you sure you wish to delete all contents?")) == true) {
        this.saveForUndo();
        var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
        oEditor.document.body.innerHTML = "<BR>";
        oEditor.focus()
    }
};

function applySpan(blockTag) {
    var useBlock = "SPAN";
    if (blockTag != null && blockTag != "") useBlock = blockTag;
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    var oSel = oEditor.getSelection();
    var range;
    var oElement;
    var sHTML;
    if (!isTextSelected(oSel)) {
        range = oSel.getRangeAt(0);
        oElement = getSelectedElement(oSel);
        if (oElement.nodeName == useBlock) return oElement;
        return
    }
    range = oSel.getRangeAt(0);
    sHTML = range.toString();
    var docFrag = range.extractContents();
    var idSpan = oEditor.document.createElement(useBlock);
    idSpan.appendChild(docFrag);
    range.insertNode(idSpan);
    try {
        oEditor.document.designMode = "on"
    } catch (e) {}
    range = oEditor.document.createRange();
    range.selectNode(idSpan);
    oSel = oEditor.getSelection();
    oSel.removeAllRanges();
    oSel.addRange(range);
    return idSpan
};

function makeAbsolute() {
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    var oSel = oEditor.getSelection();
    var oEl = getSelectedElement(oSel);
    this.saveForUndo();
    if (oEl != null && oEl.nodeName != "BODY") {
        if (oEl.style.position == "fixed") oEl.style.position = "";
        else oEl.style.position = "fixed"
    }
};

function doOver_TabCreate(el) {
    var oTD = el;
    var oTable = oTD.parentNode.parentNode.parentNode;
    var nRow = oTD.parentNode.rowIndex;
    var nCol = oTD.cellIndex;
    var rows = oTable.rows;
    var custCell = rows[rows.length - 1].childNodes[0];
    custCell.innerHTML = "<div>" + (nRow * 1 + 1) + " x " + (nCol * 1 + 1) + " " + getTxt("Table Dimension Text") + "</div>";
    custCell.style.backgroundColor = "";
    custCell.style.color = "#000000";
    custCell.style.border = "0px";
    for (var i = 0; i < rows.length - 1; i++) {
        var oRow = rows[i];
        for (var j = 0; j < oRow.childNodes.length; j++) {
            var oCol = oRow.childNodes[j];
            if (i <= nRow && j <= nCol) oCol.style.backgroundColor = "#316ac5";
            else oCol.style.backgroundColor = "#ffffff"
        }
    }
};

function doOut_TabCreate(el) {
    var oTable = el;
    if (oTable.nodeName != "TABLE") return;
    var rows = oTable.rows;
    for (var i = 0; i < rows.length - 1; i++) {
        var oRow = rows[i];
        for (var j = 0; j < oRow.childNodes.length; j++) {
            var oCol = oRow.childNodes[j];
            oCol.style.backgroundColor = "#ffffff"
        }
    }
};

function doRefresh_TabCreate() {
    if (!this.btnTable) return;
    var oTable = document.getElementById("dropTableCreate" + this.oName);
    var rows = oTable.rows;
    for (var i = 0; i < rows.length - 1; i++) {
        var oRow = rows[i];
        for (var j = 0; j < oRow.childNodes.length; j++) {
            var oCol = oRow.childNodes[j];
            oCol.style.backgroundColor = "#ffffff"
        }
    }
};

function doClick_TabCreate(el) {
    this.hide();
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    var oTD = el;
    var nRow = oTD.parentNode.rowIndex + 1;
    var nCol = oTD.cellIndex + 1;
    this.saveForUndo();
    var sHTML = "<table style='border-collapse:collapse;width:100%;' selThis='selThis'>";
    for (var i = 1; i <= nRow; i++) {
        sHTML += "<tr>";
        for (var j = 1; j <= nCol; j++) {
            sHTML += "<td><br/></td>"
        }
        sHTML += "</tr>"
    }
    sHTML += "</table>";
    var oSel = oEditor.getSelection();
    var range = oSel.getRangeAt(0);
    range.collapse(true);
    var docFrag = range.createContextualFragment(sHTML);
    range.insertNode(docFrag);
    try {
        oEditor.document.designMode = "on"
    } catch (e) {}
    var allTabs = oEditor.document.getElementsByTagName("TABLE");
    for (var i = 0; i < allTabs.length; i++) {
        if (allTabs[i].getAttribute("selThis") == "selThis") {
            range = oEditor.document.createRange();
            range.selectNode(allTabs[i]);
            oSel = oEditor.getSelection();
            oSel.removeAllRanges();
            oSel.addRange(range);
            allTabs[i].removeAttribute("selThis");
            break
        }
    }
    realTime(this)
};

function doKeyPress(evt, obj) {
    if (!obj.arrUndoList[0]) {
        obj.saveForUndo()
    }
    if (evt.ctrlKey || evt.metaKey) {
        switch (evt.keyCode) {
        case 89:
            obj.doRedo();
            break;
        case 90:
            obj.doUndo();
            break;
        case 86:
            if (obj.pasteTextOnCtrlV) {
                evt.preventDefault();
                modelessDialogShow(obj.dialogPath + "webpastetext.htm", 400, 280)
            } else {
                var mainEditor = document.getElementById("idContent" + obj.oName).contentWindow;
                var oSel = mainEditor.getSelection();
                var range = oSel.getRangeAt(0);
                mainEditor.document.body.contentEditable = false;
                var scrTop = (mainEditor.document.body.scrollTop);
                var pasteArea = mainEditor.document.getElementById("pasteArea");
                if (!pasteArea) {
                    pasteArea = document.createElement("div");
                    pasteArea.id = "pasteArea";
                    pasteArea.style.position = "fixed";
                    pasteArea.style.top = scrTop + "px";
                    pasteArea.style.width = "1px";
                    pasteArea.style.height = "1px";
                    pasteArea.style.border = "none";
                    pasteArea.style.overflow = "hidden";
                    pasteArea.contentEditable = true;
                    mainEditor.document.body.appendChild(pasteArea)
                }
                pasteArea.focus();
                window.setTimeout(function () {
                    var str = obj.fixWord();
                    var mainEditor = document.getElementById("idContent" + obj.oName).contentWindow;
                    mainEditor.document.body.contentEditable = true;
                    var oSel = mainEditor.getSelection();
                    oSel.removeAllRanges();
                    oSel.addRange(range);
                    obj.insertHTML(str)
                }, 100)
            }
            break
        }
    }
    if (evt.keyCode == 37 || evt.keyCode == 38 || evt.keyCode == 39 || evt.keyCode == 40) {
        obj.saveForUndo()
    }
    if (evt.keyCode == 13) {
        obj.saveForUndo();
        if (evt.shiftKey) {
            obj.insertHTML("<br />");
            evt.preventDefault()
        } else {
            if (obj.returnKeyMode == 3) {
                $applyReturnEvent(evt, obj, "p")
            } else if (obj.returnKeyMode == 2) {
                $applyReturnEvent(evt, obj, "br")
            } else {
                $applyReturnEvent(evt, obj, "div")
            }
        }
    }
    obj.onKeyPress(evt)
};

function $getBlockElement(el) {
    if (el == null) return null;
    var block = "|H1|H2|H3|H4|H5|H6|PRE|DIV|P|TD|";
    while (el.tagName != "BODY") {
        if (block.indexOf("|" + el.tagName + "|") != -1) {
            return el
        }
        el = el.parentNode
    };
    return null
};

function $applyReturnEvent(evt, obj, tag) {
    var wnd = document.getElementById("idContent" + obj.oName).contentWindow;
    var sel = wnd.getSelection();
    var selElm = getSelectedElement(sel);
    if (selElm) {
        var li = GetElement(selElm, "LI");
        if (li) {
            if (li.textContent == "") {
                var range = sel.getRangeAt(0);
                var list = li.parentNode;
                range.setStartAfter(list);
                range.setEndAfter(list);
                li.parentNode.removeChild(li);
                if (tag == "br") {
                    docFrag = range.createContextualFragment("<" + tag + " />")
                } else {
                    docFrag = range.createContextualFragment("<" + tag + "></" + tag + ">")
                }
                var cnt = docFrag.firstChild;
                if (cnt.innerHTML == "") cnt.innerHTML = "&nbsp;";
                range.insertNode(docFrag);
                range.selectNodeContents(cnt);
                cnt.normalize();
                cnt.parentNode.normalize();
                range.collapse(true);
                sel.removeAllRanges();
                sel.addRange(range);
                evt.preventDefault();
                return
            } else {
                var range = sel.getRangeAt(0);
                range.setEndAfter(li);
                var docFrag = range.extractContents();
                var newLi = docFrag.firstChild;
                if (newLi.textContent == "") newLi.innerHTML = "<br />";
                range.insertNode(docFrag);
                range.selectNodeContents(newLi);
                range.collapse(false);
                sel.removeAllRanges();
                sel.addRange(range);
                evt.preventDefault();
                return
            }
        }
    }
    if (tag == "br") {
        obj.insertHTML("<br />");
        evt.preventDefault();
        return
    }
    var el = $getBlockElement(getSelectedElement(sel));
    if (el && (el.tagName == "H1" || el.tagName == "H2" || el.tagName == "H3" || el.tagName == "H4" || el.tagName == "H5" || el.tagName == "H6" || el.tagName == "P" || el.tagName == "DIV" || el.tagName == "PRE")) {} else {
        obj.doCmd('FormatBlock', tag);
        sel = wnd.getSelection();
        el = $getBlockElement(getSelectedElement(sel))
    } if (el) {
        var r = sel.getRangeAt(0);
        r.setEndAfter(el);
        var docFrag = r.extractContents();
        if (el.innerHTML == "") {
            el.innerHTML = "&nbsp;";
            el.parentNode.normalize();
            el.normalize()
        }
        r.setStartAfter(el);
        r.setEndAfter(el);
        r.collapse(true);
        var isH = (",H1,H2,H3,H4,H5,H6,".indexOf("," + el.tagName + ",") >= 0);
        if (docFrag && docFrag.firstChild && (isH || docFrag.firstChild.nodeType == 3) && docFrag.firstChild.textContent == "") {
            docFrag = r.createContextualFragment("<" + tag + "></" + tag + ">")
        }
        var cnt = docFrag.firstChild;
        if (cnt.textContent == "") cnt.innerHTML = "<br />";
        r.insertNode(docFrag);
        r.selectNodeContents(cnt);
        cnt.normalize();
        cnt.parentNode.normalize();
        r.collapse(true);
        sel.removeAllRanges();
        sel.addRange(r);
        ensureVisible(cnt, document.getElementById("idContent" + obj.oName));
        evt.preventDefault()
    }
};

function fullScreen() {
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    var edtArea = document.getElementById("idArea" + this.oName);
    this.hide();
    if (this.stateFullScreen) {
        this.onNormalScreen();
        this.stateFullScreen = false;
        document.body.style.overflow = "";
        document.getElementById("id_refresh_z_index").style.margin = "0px";
        edtArea.style.position = "";
        edtArea.style.top = "0px";
        edtArea.style.left = "0px";
        if (this.width.toString().indexOf("%") != -1) edtArea.style.width = this.width;
        else edtArea.style.width = parseInt(this.width) + "px"; if (this.height.toString().indexOf("%") != -1) edtArea.style.height = this.height;
        else edtArea.style.height = parseInt(this.height) + "px";
        for (var i = 0; i < oUtil.arrEditor.length; i++) {
            if (oUtil.arrEditor[i] != this.oName) {
                document.getElementById("idArea" + oUtil.arrEditor[i]).style.display = "";
                try {
                    document.getElementById("idContent" + oUtil.arrEditor[i]).contentWindow.document.designMode = "on"
                } catch (e) {}
            }
        }
    } else {
        this.onFullScreen();
        this.stateFullScreen = true;
        window.scroll(0, 0);
        document.body.style.overflow = "hidden";
        document.getElementById("id_refresh_z_index").style.margin = "70px";
        edtArea.style.position = "fixed";
        edtArea.style.top = "0px";
        edtArea.style.left = "0px";
        edtArea.style.width = window.innerWidth + "px";
        edtArea.style.height = window.innerHeight + "px";
        edtArea.style.zIndex = 2000;
        for (var i = 0; i < oUtil.arrEditor.length; i++) {
            if (oUtil.arrEditor[i] != this.oName) document.getElementById("idArea" + oUtil.arrEditor[i]).style.display = "none"
        }
        oEditor.focus()
    }
    try {
        oEditor.document.designMode = "on"
    } catch (e) {}
    var idStyles = document.getElementById("idStyles" + this.oName);
    idStyles.innerHTML = ""
};

function modelessDialogShow(url, width, height, p, opt) {
    modalDialog(url, width, height)
};

function modalDialogShow(url, width, height, p, opt) {
    modalDialog(url, width, height)
};

function windowOpen(url, wd, hg, ov, p, opt) {
    var id = "ID" + (new Date()).getTime();
    var f = new ISWindow(id);
    f.iconPath = oUtil.scriptPath.substring(0, oUtil.scriptPath.indexOf("saf/")) + "icons/";
    f.show({
        width: wd + "px",
        height: hg + "px",
        overlay: ov,
        center: true,
        url: url,
        openerWin: p,
        options: opt
    })
};

function hide() {
    hideAllDD();
    this.oColor1.hide();
    this.oColor2.hide();
    if (this.btnTable) this.doRefresh_TabCreate()
};

function lineBreak1(tag) {
    arrReturn = ["\n", "", ""];
    if (tag == "A" || tag == "B" || tag == "CITE" || tag == "CODE" || tag == "EM" || tag == "FONT" || tag == "I" || tag == "SMALL" || tag == "STRIKE" || tag == "BIG" || tag == "STRONG" || tag == "SUB" || tag == "SUP" || tag == "U" || tag == "SAMP" || tag == "S" || tag == "VAR" || tag == "BASEFONT" || tag == "KBD" || tag == "TT" || tag == "SPAN" || tag == "IMG") arrReturn = ["", "", ""];
    if (tag == "TEXTAREA" || tag == "TABLE" || tag == "THEAD" || tag == "TBODY" || tag == "TR" || tag == "OL" || tag == "UL" || tag == "DIR" || tag == "MENU" || tag == "FORM" || tag == "SELECT" || tag == "MAP" || tag == "DL" || tag == "HEAD" || tag == "BODY" || tag == "HTML") arrReturn = ["\n", "", "\n"];
    if (tag == "STYLE" || tag == "SCRIPT") arrReturn = ["\n", "", ""];
    if (tag == "BR" || tag == "HR") arrReturn = ["", "\n", ""];
    return arrReturn
};

function fixAttr(s) {
    s = String(s).replace(/&/g, "&amp;");
    s = String(s).replace(/</g, "&lt;");
    s = String(s).replace(/"/g, "&quot;");
    return s
};

function fixVal(s) {
    s = String(s).replace(/&/g, "&amp;");
    s = String(s).replace(/</g, "&lt;");
    var x = escape(s);
    x = unescape(x.replace(/\%A0/gi, "-*REPL*-"));
    s = x.replace(/-\*REPL\*-/gi, "&nbsp;");
    return s
};

function recur(oEl, sTab) {
    var sHTML = "";
    for (var i = 0; i < oEl.childNodes.length; i++) {
        var oNode = oEl.childNodes[i];
        if (oNode.nodeType == 1) {
            var sTagName = oNode.nodeName;
            var bDoNotProcess = false;
            if (sTagName.substring(0, 1) == "/") {
                bDoNotProcess = true
            } else {
                var sT = sTab;
                sHTML += lineBreak1(sTagName)[0];
                if (lineBreak1(sTagName)[0] != "") sHTML += sT
            } if (bDoNotProcess) {} else if (sTagName == "OBJECT" || sTagName == "EMBED") {
                s = getOuterHTML(oNode);
                s = s.replace(/\"[^\"]*\"/ig, function (x) {
                    x = x.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/'/g, "&apos;").replace(/\s+/ig, "#_#").replace(/&amp;amp;/gi, "&amp;");
                    return x
                });
                s = s.replace(/<([^ >]*)/ig, function (x) {
                    return x.toLowerCase()
                });
                s = s.replace(/ ([^=]+)=([^"' >]+)/ig, " $1=\"$2\"");
                s = s.replace(/ ([^=]+)=/ig, function (x) {
                    return x.toLowerCase()
                });
                s = s.replace(/#_#/ig, " ");
                s = s.replace(/<param([^>]*)>/ig, "\n<param$1 />").replace(/\/ \/>$/ig, " \/>");
                if (sTagName == "EMBED")
                    if (oNode.innerHTML == "") s = s.replace(/>$/ig, " \/>").replace(/\/ \/>$/ig, "\/>");
                s = s.replace(/<param name=\"Play\" value=\"0\" \/>/, "<param name=\"Play\" value=\"-1\" \/>");
                sHTML += s
            } else if (sTagName == "TITLE") {
                sHTML += "<title>" + oNode.innerHTML + "</title>"
            } else {
                if (sTagName == "AREA") {
                    var sCoords = oNode.coords;
                    var sShape = oNode.shape
                }
                var oNode2 = oNode.cloneNode(false);
                s = getOuterHTML(oNode2).replace(/<\/[^>]*>/, "");
                if (sTagName == "STYLE") {
                    var arrTmp = s.match(/<[^>]*>/ig);
                    s = arrTmp[0]
                }
                s = s.replace(/\"[^\"]*\"/ig, function (x) {
                    x = x.replace(/&/g, "&amp;").replace(/&amp;amp;/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\s+/ig, "#_#");
                    return x
                });
                s = s.replace(/<([^ >]*)/ig, function (x) {
                    return x.toLowerCase()
                });
                s = s.replace(/ ([^=]+)=([^" >]+)/ig, " $1=\"$2\"");
                s = s.replace(/ ([^=]+)=/ig, function (x) {
                    return x.toLowerCase()
                });
                s = s.replace(/#_#/ig, " ");
                s = s.replace(/(<hr[^>]*)(noshade="")/ig, "$1noshade=\"noshade\"");
                s = s.replace(/(<input[^>]*)(checked="")/ig, "$1checked=\"checked\"");
                s = s.replace(/(<select[^>]*)(multiple="")/ig, "$1multiple=\"multiple\"");
                s = s.replace(/(<option[^>]*)(selected="")/ig, "$1selected=\"true\"");
                s = s.replace(/(<input[^>]*)(readonly="")/ig, "$1readonly=\"readonly\"");
                s = s.replace(/(<input[^>]*)(disabled="")/ig, "$1disabled=\"disabled\"");
                s = s.replace(/(<td[^>]*)(nowrap="" )/ig, "$1nowrap=\"nowrap\" ");
                s = s.replace(/(<td[^>]*)(nowrap=""\>)/ig, "$1nowrap=\"nowrap\"\>");
                if (sTagName == "AREA") {
                    s = s.replace(/ coords=\"0,0,0,0\"/ig, " coords=\"" + sCoords + "\"");
                    s = s.replace(/ shape=\"RECT\"/ig, " shape=\"" + sShape + "\"")
                }
                var bClosingTag = true;
                if (sTagName == "IMG" || sTagName == "BR" || sTagName == "AREA" || sTagName == "HR" || sTagName == "INPUT" || sTagName == "BASE" || sTagName == "LINK") {
                    s = s.replace(/>$/ig, " \/>").replace(/\/ \/>$/ig, "\/>");
                    bClosingTag = false
                }
                sHTML += s;
                if (sTagName != "TEXTAREA") sHTML += lineBreak1(sTagName)[1];
                if (sTagName != "TEXTAREA")
                    if (lineBreak1(sTagName)[1] != "") sHTML += sT;
                if (bClosingTag) {
                    s = getOuterHTML(oNode);
                    if (sTagName == "SCRIPT") {
                        s = s.replace(/<script([^>]*)>[\n+\s+\t+]*/ig, "<script$1>");
                        s = s.replace(/[\n+\s+\t+]*<\/script>/ig, "<\/script>");
                        s = s.replace(/<script([^>]*)>\/\/<!\[CDATA\[/ig, "");
                        s = s.replace(/\/\/\]\]><\/script>/ig, "");
                        s = s.replace(/<script([^>]*)>/ig, "");
                        s = s.replace(/<\/script>/ig, "");
                        s = s.replace(/^\s+/, '').replace(/\s+$/, '');
                        sHTML += "\n" + sT + "//<![CDATA[\n" + sT + s + "\n" + sT + "//]]>\n" + sT
                    }
                    if (sTagName == "STYLE") {
                        s = s.replace(/<style([^>]*)>[\n+\s+\t+]*/ig, "<style$1>");
                        s = s.replace(/[\n+\s+\t+]*<\/style>/ig, "<\/style>");
                        s = s.replace(/<style([^>]*)><!--/ig, "");
                        s = s.replace(/--><\/style>/ig, "");
                        s = s.replace(/<style([^>]*)>/ig, "");
                        s = s.replace(/<\/style>/ig, "");
                        s = s.replace(/^\s+/, "").replace(/\s+$/, "");
                        sHTML += "\n" + sT + "<!--\n" + sT + s + "\n" + sT + "-->\n" + sT
                    }
                    if (sTagName == "DIV" || sTagName == "P") {
                        if (oNode.innerHTML == "" || oNode.innerHTML == "&nbsp;") {
                            sHTML += "&nbsp;"
                        } else sHTML += recur(oNode, sT + "\t")
                    } else if (sTagName == "STYLE" || sTagName == "SCRIPT") {} else {
                        sHTML += recur(oNode, sT + "\t")
                    } if (sTagName != "TEXTAREA") sHTML += lineBreak1(sTagName)[2];
                    if (sTagName != "TEXTAREA")
                        if (lineBreak1(sTagName)[2] != "") sHTML += sT;
                    sHTML += "</" + sTagName.toLowerCase() + ">"
                }
            }
        } else if (oNode.nodeType == 3) {
            sHTML += fixVal(oNode.nodeValue).replace(/^[\t\r\n\v\f]*/, "").replace(/[\t\r\n\v\f]*$/, "")
        } else if (oNode.nodeType == 8) {
            if (getOuterHTML(oNode).substring(0, 2) == "<" + "%") {
                sTmp = (getOuterHTML(oNode).substring(2));
                sTmp = sTmp.substring(0, sTmp.length - 2);
                sTmp = sTmp.replace(/^\s+/, "").replace(/\s+$/, "");
                var sT = sTab;
                sHTML += "\n" + sT + "<%\n" + sT + sTmp + "\n" + sT + "%>\n" + sT
            } else {
                var sT = sTab;
                sTmp = oNode.nodeValue;
                sTmp = sTmp.replace(/^\s+/, "").replace(/\s+$/, "");
                sHTML += "\n" + sT + "<!--\n" + sT + sTmp + "\n" + sT + "-->\n" + sT
            }
        } else {}
    }
    return sHTML
};

function getSelectedElement(sel) {
    var range = sel.getRangeAt(0);
    var node = range.startContainer;
    if (node.nodeType == Node.ELEMENT_NODE) {
        if (range.startOffset >= node.childNodes.length) return node;
        node = node.childNodes[range.startOffset]
    }
    if (node.nodeType == Node.TEXT_NODE) {
        if (node.nodeValue.length == range.startOffset) {
            var el = node.nextSibling;
            if (el && el.nodeType == Node.ELEMENT_NODE) {
                if (range.endContainer.nodeType == Node.TEXT_NODE && range.endContainer.nodeValue.length == range.endOffset) {
                    if (el == range.endContainer.parentNode) {
                        return el
                    }
                }
            }
        }
        while (node != null && node.nodeType != Node.ELEMENT_NODE) {
            node = node.parentNode
        }
    }
    return node
};

function isTextSelected(sel) {
    var range = sel.getRangeAt(0);
    if (range != null && range.startContainer != null) {
        return (range.startContainer.nodeType == Node.TEXT_NODE)
    }
    return false
};

function getOuterHTML(node) {
    var sHTML = "";
    switch (node.nodeType) {
    case Node.ELEMENT_NODE:
        sHTML = "<" + node.nodeName;
        var tagVal = "";
        for (var atr = 0; atr < node.attributes.length; atr++) {
            if (node.attributes[atr].nodeName.substr(0, 4) == "_moz") continue;
            if (node.attributes[atr].nodeValue.substr(0, 4) == "_moz") continue;
            if (node.nodeName == 'TEXTAREA' && node.attributes[atr].nodeName.toLowerCase() == 'value') {
                tagVal = node.attributes[atr].nodeValue
            } else {
                sHTML += ' ' + node.attributes[atr].nodeName + '="' + node.attributes[atr].nodeValue + '"'
            }
        }
        sHTML += '>';
        sHTML += (node.nodeName != 'TEXTAREA' ? node.innerHTML : tagVal);
        sHTML += "</" + node.nodeName + ">";
        break;
    case Node.COMMENT_NODE:
        sHTML = "<!" + "--" + node.nodeValue + "--" + ">";
        break;
    case Node.TEXT_NODE:
        sHTML = node.nodeValue;
        break
    }
    return sHTML
};

function tbAction(tb, id, edt, sfx) {
    var e = edt,
        oN = sfx,
        btn = id.substring(0, id.lastIndexOf(oN));
    switch (btn) {
    case "btnFullScreen":
        e.fullScreen();
        break;
    case "btnPrint":
        e.focus();
        document.getElementById("idContent" + edt.oName).contentWindow.print();
        break;
    case "btnSpellCheck":
        e.hide();
        if (e.spellCheckMode == "ieSpell");
        else if (e.spellCheckMode == "NetSpell") checkSpellingById("idContent" + edt.oName);
        break;
    case "btnCut":
        e.doCmd("Cut");
        break;
    case "btnCopy":
        e.doCmd("Copy");
        break;
    case "btnUndo":
        e.doUndo();
        break;
    case "btnRedo":
        e.doRedo();
        break;
    case "btnBold":
        e.applyFormattingStyle("Bold");
        break;
    case "btnItalic":
        e.applyFormattingStyle("Italic");
        break;
    case "btnUnderline":
        e.applyFormattingStyle("Underline");
        break;
    case "btnStrikethrough":
        e.applyFormattingStyle("Strikethrough");
        break;
    case "btnSuperscript":
        e.applyFormattingStyle("Superscript");
        break;
    case "btnSubscript":
        e.applyFormattingStyle("Subscript");
        break;
    case "btnJustifyLeft":
        e.applyJustifyLeft();
        break;
    case "btnJustifyCenter":
        e.applyJustifyCenter();
        break;
    case "btnJustifyRight":
        e.applyJustifyRight();
        break;
    case "btnJustifyFull":
        e.applyJustifyFull();
        break;
    case "btnNumbering":
        e.applyNumbering();
        break;
    case "btnBullets":
        e.applyBullets();
        break;
    case "btnIndent":
        e.doCmd("Indent");
        break;
    case "btnOutdent":
        e.applyOutdent();
        break;
    case "btnLTR":
        e.applyBlockDirLTR();
        break;
    case "btnRTL":
        e.applyBlockDirRTL();
        break;
    case "btnFontDialog":
        e.hide();
        modelessDialogShow(e.dialogPath + "webfonts.htm", e.dialogSize["FontDialog"].w, e.dialogSize["FontDialog"].h);
        break;
    case "btnTextDialog":
        e.hide();
        modelessDialogShow(e.dialogPath + "webtext.htm", e.dialogSize["TextDialog"].w, e.dialogSize["TextDialog"].h);
        break;
    case "btnCompleteTextDialog":
        e.hide();
        modelessDialogShow(e.dialogPath + "webtextcomplete.htm", e.dialogSize["CompleteTextDialog"].w, e.dialogSize["CompleteTextDialog"].h);
        break;
    case "btnLinkDialog":
        e.hide();
        if (e.fileBrowser == "" && e.enableCssButtons == false) {
            modelessDialogShow(e.dialogPath + "weblink.htm", 300, 400)
        } else {
            modelessDialogShow(e.dialogPath + "weblink.htm", e.dialogSize["LinkDialog"].w, e.dialogSize["LinkDialog"].h)
        }
        break;
    case "btnImageDialog":
        e.hide();
        modelessDialogShow(e.dialogPath + "webimage.htm", e.dialogSize["ImageDialog"].w, e.dialogSize["ImageDialog"].h);
        break;
    case "btnYoutubeDialog":
        e.hide();
        modelessDialogShow(e.dialogPath + "webyoutube.htm", e.dialogSize["YoutubeDialog"].w, e.dialogSize["YoutubeDialog"].h);
        break;
    case "btnTableDialog":
        e.hide();
        modelessDialogShow(e.dialogPath + "webtable.htm", e.dialogSize["TableDialog"].w, e.dialogSize["TableDialog"].h);
        break;
    case "btnFlashDialog":
        e.hide();
        modelessDialogShow(e.dialogPath + "webflash.htm", e.dialogSize["FlashDialog"].w, e.dialogSize["FlashDialog"].h);
        break;
    case "btnCharsDialog":
        e.hide();
        modelessDialogShow(e.dialogPath + "webchars.htm", e.dialogSize["CharsDialog"].w, e.dialogSize["CharsDialog"].h);
        break;
    case "btnSearchDialog":
        e.hide();
        modelessDialogShow(e.dialogPath + "websearch.htm", e.dialogSize["SearchDialog"].w, e.dialogSize["SearchDialog"].h);
        break;
    case "btnSourceDialog":
        e.hide();
        changeActiveEditor(e.oName);
        modelessDialogShow(e.dialogPath + "websource.htm", e.dialogSize["SourceDialog"].w, e.dialogSize["SourceDialog"].h);
        break;
    case "btnBookmarkDialog":
        e.hide();
        modelessDialogShow(e.dialogPath + "webbookmark.htm", e.dialogSize["BookmarkDialog"].w, e.dialogSize["BookmarkDialog"].h);
        break;
    case "btnPreview":
        e.hide();
        modelessDialogShow(e.dialogPath + "webpreview.htm", e.dialogSize["Preview"].w, e.dialogSize["Preview"].h);
        break;
    case "btnContentBlock":
        e.hide();
        eval(e.cmdContentBlock);
        break;
    case "btnInternalLink":
        e.hide();
        eval(e.cmdInternalLink);
        break;
    case "btnInternalImage":
        e.hide();
        eval(e.cmdInternalImage);
        break;
    case "btnCustomObject":
        e.hide();
        eval(e.cmdCustomObject);
        break;
    case "btnGuidelines":
        e.runtimeBorder(true);
        break;
    case "btnAbsolute":
        e.makeAbsolute();
        break;
    case "btnLine":
        e.doCmd("InsertHorizontalRule");
        break;
    case "btnRemoveFormat":
        e.doClean();
        break;
    case "btnClearAll":
        e.clearAll();
        break;
    case "btnStyles":
        e.hide();
        e.openStyleSelect();
        break;
    case "btnParagraph":
        e.hide();
        e.selectParagraph();
        break;
    case "btnFontName":
        e.hide();
        e.expandSelection();
        realtimeFontSelect(e.oName);
        break;
    case "btnFontSize":
        e.hide();
        e.expandSelection();
        realtimeSizeSelect(e.oName);
        break;
    case "btnCustomTag":
        e.hide();
        break;
    case "btnQuote":
        e.applyQuote();
        break;
    default:
        for (var i = 0; i < e.arrCustomButtons.length; i++) {
            if (e.arrCustomButtons[i][0] == btn) {
                eval(e.arrCustomButtons[i][1]);
                break
            }
        }
    }
};

function ddAction(tb, id, edt, sfx) {
    var oN = sfx;
    var e = edt;
    var btn = id.substring(0, id.lastIndexOf(oN));
    switch (btn) {
    case "btnPasteClip":
        alert(getTxt("PasteWarning"));
        break;
    case "btnPasteWord":
        if (!e.customDialogShow("PasteWord")) modelessDialogShow(e.scriptPath + "paste_word.htm", 400, 280);
        break;
    case "btnPasteText":
        modelessDialogShow(e.dialogPath + "webpastetext.htm", 400, 280);
        break
    }
    var idx = 0;
    if (btn.indexOf("btnParagraphFormatting") != -1) {} else if (btn.indexOf("btnParagraph") != -1) {
        idx = btn.substr(btn.indexOf("_") + 1);
        e.applyParagraph("<" + e.arrParagraph[parseInt(idx)][1] + ">")
    } else if (btn.indexOf("btnFontName") != -1) {
        idx = btn.substr(btn.indexOf("_") + 1);
        e.applyFontName(e.arrFontName[parseInt(idx)])
    } else if (btn.indexOf("btnFontSize") != -1) {
        idx = btn.substr(btn.indexOf("_") + 1);
        e.applyFontSize(e.arrFontSize[parseInt(idx)][1])
    } else if (btn.indexOf("btnCustomTag") != -1) {
        idx = btn.substr(btn.indexOf("_") + 1);
        e.insertCustomTag(parseInt(idx))
    }
};

function changeHeight(v) {
    var cH = String(this.height);
    var edtObj = document.getElementById("idArea" + this.oName);
    if (cH.indexOf("%") > -1) {
        cH = edtObj.childNodes[0].offsetHeight - edtObj.rows[0].cells[0].childNodes[0].offsetHeight - (this.useTagSelector ? 20 : 0)
    }
    if (!this.minHeight) this.minHeight = parseInt(cH, 10);
    var newHeight = parseInt(cH, 10) + v;
    this.height = newHeight + "px";
    edtObj.style.height = this.height
};

function _isWordContent(str) {
    return (/msonormal/i.test(str) || /<!(?:--[\s\S]*?--\s*)?>\s*/gi.test(str) || /<\\?\?xml[^>]*>/gi.test(str) || /<\/?o:p[^>]*>/gi.test(str) || /<\/?u1:p[^>]*>/gi.test(str) || /<\/?v:[^>]*>/gi.test(str) || /<\/?o:[^>]*>/gi.test(str) || /<\/?st1:[^>]*>/gi.test(str) || /<\/?w:wrap[^>]*>/gi.test(str) || /<\/?w:anchorlock[^>]*>/gi.test(str) || /<!--\[if[^>]*>/gi.test(str) || /<!--\[endif\]-->/gi.test(str) || /<!\[endif\]-->/gi.test(str) || /<\/?meta[^>]*>/g.test(str) || /<\/?link[^>]*>/g.test(str) || /<\/?style[^>]*>/g.test(str))
};

function fixWord() {
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    var idSource = oEditor.document.getElementById("pasteArea");
    var isWord = _isWordContent(idSource.innerHTML);
    if (!isWord) {
        var cnt = idSource.innerHTML;
        idSource.parentNode.removeChild(idSource);
        return cnt
    }
    var allTags = idSource.getElementsByTagName("*");
    for (var i = 0; i < allTags.length; i++) {
        allTags[i].removeAttribute("class");
        allTags[i].removeAttribute("style")
    }
    var str = idSource.innerHTML;
    idSource.parentNode.removeChild(idSource);
    str = String(str).replace(/<\\?\?xml[^>]*>/g, "");
    str = String(str).replace(/<\/?o:p[^>]*>/g, "");
    str = String(str).replace(/<\/?u1:p[^>]*>/gi, "");
    str = String(str).replace(/<\/?v:[^>]*>/g, "");
    str = String(str).replace(/<\/?o:[^>]*>/g, "");
    str = String(str).replace(/<\/?st1:[^>]*>/g, "");
    str = String(str).replace(/<\/?w:wrap[^>]*>/gi, "");
    str = String(str).replace(/<\/?w:anchorlock[^>]*>/gi, "");
    str = String(str).replace(/<\/?SPAN[^>]*>/g, "");
    str = String(str).replace(/<\/?FONT[^>]*>/g, "");
    str = String(str).replace(/<\/?STRONG[^>]*>/g, "");
    str = String(str).replace(/<\/?H1[^>]*>/g, "");
    str = String(str).replace(/<\/?H2[^>]*>/g, "");
    str = String(str).replace(/<\/?H3[^>]*>/g, "");
    str = String(str).replace(/<\/?H4[^>]*>/g, "");
    str = String(str).replace(/<\/?H5[^>]*>/g, "");
    str = String(str).replace(/<\/?H6[^>]*>/g, "");
    str = String(str).replace(/<\/?P[^>]*><\/P>/g, "");
    str = String(str).replace(/<!--\[if[^>]*>/gi, "");
    str = String(str).replace(/<!--\[endif\]-->/gi, "");
    str = String(str).replace(/<!\[endif\]-->/gi, "");
    return str
};

function customDialogShow(s) {
    for (var j = 0; j < this.customDialog.length; j++) {
        if (this.customDialog[j][0] == s) {
            eval(this.customDialog[j][1]);
            return true
        }
    }
    return false
};

function GetEmoticons() {
    var sHtml = '';
    var arrEmoticons = [
        ["smiley.png", "smiley"],
        ["smiley-lol.png", "lol"],
        ["smiley-confuse.png", "confuse"],
        ["smiley-cool.png", "cool"],
        ["smiley-cry.png", "cry"],
        ["smiley-wink.png", "wink"],
        ["smiley-surprise.png", "surprise"],
        ["smiley-sad.png", "sad"],
        ["smiley-red.png", "red"],
        ["smiley-neutral.png", "neutral"],
        ["smiley-kiss.png", "kiss"],
        ["smiley-mad.png", "mad"],
        ["smiley-money.png", "money"],
        ["smiley-sleep.png", "sleep"],
        ["smiley-yell.png", "yell"],
        ["smiley-roll.png", "roll"],
        ["smiley-grin.png", "grin"],
        ["smiley-razz.png", "razz"],
        ["smiley-sweat.png", "sweat"],
        ["smiley-eek.png", "eek"],
        ["smiley-zipper.png", "zipper"],
        ["heart.png", "love"],
        ["heart-break.png", "heart break"],
        ["light-bulb.png", "idea"]
    ];
    var icPath = this.scriptPath.substring(0, this.scriptPath.indexOf("saf/")) + this.iconPath;
    sHtml += "<div style='width:193px;height:72px;margin:2px;'>";
    for (var j = 0; j < arrEmoticons.length; j++) {
        sHtml += "<img unselectable='on' src='" + icPath + arrEmoticons[j][0] + "' onmouseover='this.style.border=\"#aaaaaa 1px solid\"' onmouseout='this.style.border=\"#ffffff 1px solid\"' style='float:left;margin:1px;cursor:pointer;border:#ffffff 1px solid;padding:2px;' onclick='" + this.oName + ".insertEmoticon(\"" + icPath + arrEmoticons[j][0] + "\",\"" + arrEmoticons[j][1] + "\");' alt='" + arrEmoticons[j][1] + "' />"
    }
    sHtml += "</div>";
    return sHtml
}

function insertEmoticon(img, alt) {
    this.insertHTML("<img src='" + img + "' alt='" + alt + "' />");
    var box = document.getElementById("ddEmoticons" + this.oName);
    box.style.display = "none";
    this.isActive = false
}

function applyQuote() {
    var oEditor = document.getElementById("idContent" + this.oName).contentWindow;
    var oSel = oEditor.getSelection();
    var element = getSelectedElement(oSel);
    var oQuote = GetElement(element, "BLOCKQUOTE");
    if (oQuote) {
        this.applyParagraph('p')
    } else {
        this.applyParagraph('blockquote')
    }
}

function cleanWebkitUselessSpan(obj) {
    var oEditor = document.getElementById("idContent" + obj.oName).contentWindow;
    var sel = oEditor.getSelection();
    var r = sel.getRangeAt(0);
    var el = getSelectedElement(sel);
    var span = GetElement(el, "SPAN");
    if (span && span.textContent == "") {
        var parent = span.parentNode;
        span.parentNode.removeChild(span);
        if (parent.tagName == "BODY") {
            var tag = "div";
            if (obj.returnKeyMode == 2) {
                tag = "br"
            } else if (obj.returnKeyMode == 3) {
                tag = "p"
            }
            var docFrag = r.createContextualFragment("<" + tag + "></" + tag + ">");
            var cnt = docFrag.firstChild;
            if (cnt.innerHTML == "") cnt.innerHTML = "<br />";
            r.insertNode(docFrag);
            r.selectNodeContents(cnt);
            cnt.normalize();
            cnt.parentNode.normalize();
            r.collapse(true);
            sel.removeAllRanges();
            sel.addRange(r);
            realTime(this)
        }
    }
};

function getElementPosition(el, pr) {
    var tmp = el,
        y = 0;
    do {
        y += tmp.offsetTop;
        tmp = tmp.offsetParent
    } while (tmp && tmp.id != pr.id);
    return y
};

function isNodeVisible(el, edtFrame) {
    var body = edtFrame.contentWindow.document.body;
    var elPos = getElementPosition(el, body);
    var vpStyle = null;
    if (window.getComputedStyle) {
        vpStyle = window.getComputedStyle(edtFrame, "")
    } else {
        vpStyle = edtFrame.currentStyle
    }
    var vph = parseInt(vpStyle.height, 10);
    return (elPos > body.scrollTop && elPos < (body.scrollTop + vph - 20))
};

function ensureVisible(el, edtFrame) {
    if (!isNodeVisible(el, edtFrame)) {
        var body = edtFrame.contentWindow.document.body;
        var elPos = getElementPosition(el, body);
        body.scrollTop = elPos
    }
};