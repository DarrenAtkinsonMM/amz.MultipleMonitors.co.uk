<%
' ============================================================
' standSpecs.asp
' Per-stand reference content for shop/pc/viewprd-stand-v2.asp.
'
' Identity fields (name, price, image, SKU, short description)
' come from the products table. This file holds the *static*
' reference content the DB has no column for - eyebrow line,
' inclusion-strip copy, the six "specs at a glance" cards,
' "in the box" chips, micro-band trio, VESA intro, six
' dim-stats, and dimension-diagram image paths.
'
' Keyed by SKU - same shape as monitorSpecs.asp.
'
' Sparse fields per entry are fine - the page falls back to
' hardcoded defaults via mmMetaStr(key, fallback) when a key
' is missing or the SKU is unknown.
'
' Keys:
'   eyebrow                    Category line above the H1
'   pitch                      Pitch paragraph (overrides DB if set)
'   tpRating / tpCount         Trustpilot widget
'   inclMadeTitle / Sub        Inclusion-strip "UK-made" item
'   inclWarrantyTitle / Sub    Inclusion-strip warranty item
'   inclDispatchTitle / Sub    Inclusion-strip dispatch item
'   specCardN_label/value/desc/icon  N=1..6 - "Specs at a glance" cards
'   inTheBox                   Array of strings - chips
'   microAssembly / Sub        Micro-band trio bold + small lines
'   microTools / Sub
'   microWarranty / Sub
'   vesaIntro1 / vesaIntro2    VESA-section paragraphs
'   monMaxWidth / monMaxHeight Max monitor envelope (SVG labels, mm)
'   monMaxNote1 / monMaxNote2  Monitor-dim explanatory paragraphs
'   dimsLead                   Lead paragraph above dim-stats
'   dimStatN_value / label     N=1..6 - dim-stats grid
'   dimImgFront / FrontAlt     Front profile image src + alt
'   dimImgSide  / SideAlt      Side profile image src + alt
'   dimPdf                     Optional PDF download URL
' ============================================================
Function mmGetStandMeta(ByVal sku)
  Dim m : Set m = Nothing
  Dim mmKey : mmKey = UCase(Trim(sku & ""))

  Select Case mmKey
    Case "MM-S4S"  ' Quad Square Synergy Stand - 4 screens, 2x2 square
      Set m = Server.CreateObject("Scripting.Dictionary")
      m.Add "eyebrow", "Synergy Stand &middot; 4-screen square"
      m.Add "Title", "Quad Square"
      m.Add "pitch",   "Four screens in the footprint of two. All-steel, UK-designed and UK-built &mdash; a square 2&times;2 layout that hides cables behind the column and scales up to six screens later."
      m.Add "ColumnText", "One"
      m.Add "ScreenNum", "4"
      m.Add "ScreenText", "Four"

      m.Add "specCard1_label", "Layout"
      m.Add "specCard1_value", "4 screens, 2&times;2 square"
      m.Add "specCard1_desc",  "Two screens above, two below. Mount four screens in the footprint of two side-by-side monitors. Ideal for maximising display space on a smaller desk."
      m.Add "specCard1_icon",  "fa-th-large"

      m.Add "specCard2_label", "Max screen size"
      m.Add "specCard2_value", "Up to 28&Prime; per screen"
      m.Add "specCard2_desc",  "Designed with room to be able to angle outer screens inward even at the upper screen sizes for a gentle curved setup."
      m.Add "specCard2_icon",  "fa-expand"

      m.Add "specCard3_label", "Max weight"
      m.Add "specCard3_value", "8&nbsp;kg per screen"
      m.Add "specCard3_desc",  "Our all-steel construction makes handling heavier screens easy. Not something weaker stands with plastic parts or cheaper alloys can do."
      m.Add "specCard3_icon",  "fa-balance-scale"

      m.Add "specCard4_label", "VESA"
      m.Add "specCard4_value", "75&times;75 &amp; 100&times;100"
      m.Add "specCard4_desc",  "Fits every monitor we sell, and is compatible with any 'wall mountable' screen that uses VESA 75 or 100 interfaces."
      m.Add "specCard4_icon",  "fa-crosshairs"

      m.Add "specCard5_label", "Adjustability"
      m.Add "specCard5_value", "6 degrees of freedom"
      m.Add "specCard5_desc",  "Height, arm hinge, horizontal slide, pivot, tilt and 30&nbsp;mm of fine height adjust at every mount. Set once and then it locks solid."
      m.Add "specCard5_icon",  "fa-sliders"

      m.Add "specCard6_label", "Materials"
      m.Add "specCard6_value", "All-steel, no plastic"
      m.Add "specCard6_desc",  "Unlike cheaper imports we do not use plastic parts or cheap alloys. The entire frame, column and arms are powder-coated steel."
      m.Add "specCard6_icon",  "fa-cubes"

      m.Add "inTheBox", Array( _
            "Base &amp; central column", _
            "4 &times; arm assembly", _
            "4 &times; VESA plates", _
            "All fixings", _
            "Cable management ties", _
            "All assembly tools", _
            "<a href=""/synergy-assembly.pdf"" style=""color:inherit;"">Assembly guide</a>" _
          )

      m.Add "microAssembly",    "30&ndash;60 min assembly"
      m.Add "microAssemblySub", "No drilling, no wall fixings"
      m.Add "microTools",       "All tools included"
      m.Add "microToolsSub",    "Nothing to buy separately"
      m.Add "microWarranty",    "Lifetime warranty"
      m.Add "microWarrantySub", "On every steel part, forever"

      m.Add "vesaIntro1", "If you already have screens or are purchasing some new ones then look for a <b style=""color:var(--ink);"">VESA 75</b> or <b style=""color:var(--ink);"">VESA 100</b> rating, sometimes described as a '<b style=""color:var(--ink);"">Wall mount interface</b>'."
      m.Add "vesaIntro2", "This is the four screw holes in the back of a screen in a square configuration. If your monitor has them then it should be compatible with this Synergy Stand."

      m.Add "monMaxWidth",  "630"
      m.Add "monMaxHeight", "450"
      m.Add "monMaxNote1",  "These limits still leave room for a <b style=""color:var(--ink);"">gentle curve</b> across a multi-screen setup."
      m.Add "monMaxNote2",  "The <b style=""color:var(--ink);"">450&nbsp;mm</b> height assumes the VESA pattern sits roughly in the middle of the screen. If your VESA holes are towards the top on the back of your screen then the max height may be reduced somewhat."

      m.Add "dimsLead", "The Quad Square takes roughly the same desk space as a single 27&Prime; monitor on its stock stand &mdash; and gives you four screens in its place. A single central column keeps the desk surface below the screens clear for keyboard, notebook and coffee."

      m.Add "dimStat1_value", "400<span class=""u"">mm</span>"
      m.Add "dimStat1_label", "Base width"
      m.Add "dimStat2_value", "310<span class=""u"">mm</span>"
      m.Add "dimStat2_label", "Base depth"
      m.Add "dimStat3_value", "780<span class=""u"">mm</span>"
      m.Add "dimStat3_label", "Column height"
      m.Add "dimStat4_value", "360<span class=""u"">mm</span>"
      m.Add "dimStat4_label", "Arm reach (each)"
      m.Add "dimStat5_value", "8<span class=""u"">kg</span>"
      m.Add "dimStat5_label", "Max per mount"
      m.Add "dimStat6_value", "12.5<span class=""u"">kg</span>"
      m.Add "dimStat6_label", "Weight without screens"

      m.Add "dimImgFront",    "/images/stands/dim-4s.jpg"
      m.Add "dimImgFrontAlt", "Front-elevation engineering drawing of the Quad Square Synergy Stand with dimensions in millimetres"
      m.Add "dimImgSide",     "/images/stands/dim-side-tall.jpg"
      m.Add "dimImgSideAlt",  "Side-elevation engineering drawing of the Quad Square Synergy Stand with dimensions in millimetres"

      m.Add "dimPdf", "/synergy-assembly.pdf"

    Case "MM-S2H"  ' Dual Horizontal Synergy Stand 
      Set m = Server.CreateObject("Scripting.Dictionary")
      m.Add "eyebrow", "Synergy Stand &middot; 2-screen horizontal"
      m.Add "Title", "Dual Horizontal"
      m.Add "pitch",   "Two screens held securely side by side. All-steel, UK-designed and UK-built &mdash; a square 2&times;2 layout that hides cables behind the column and scales up to six screens later."
      m.Add "ColumnText", "One"
      m.Add "ScreenNum", "2"
      m.Add "ScreenText", "Two"

      m.Add "specCard1_label", "Layout"
      m.Add "specCard1_value", "2 screens, horizontal"
      m.Add "specCard1_desc",  "Two screens held securely side by side. Great for freeing up desk space and creating a professional and tidy layout."
      m.Add "specCard1_icon",  "fa-th-large"

      m.Add "specCard2_label", "Max screen size"
      m.Add "specCard2_value", "Up to 28&Prime; per screen"
      m.Add "specCard2_desc",  "Designed with room to be able to angle outer screens inward even at the upper screen sizes for a gentle curved setup."
      m.Add "specCard2_icon",  "fa-expand"

      m.Add "specCard3_label", "Max weight"
      m.Add "specCard3_value", "8&nbsp;kg per screen"
      m.Add "specCard3_desc",  "Our all-steel construction makes handling heavier screens easy. Not something weaker stands with plastic parts or cheaper alloys can do."
      m.Add "specCard3_icon",  "fa-balance-scale"

      m.Add "specCard4_label", "VESA"
      m.Add "specCard4_value", "75&times;75 &amp; 100&times;100"
      m.Add "specCard4_desc",  "Fits every monitor we sell, and is compatible with any 'wall mountable' screen that uses VESA 75 or 100 interfaces."
      m.Add "specCard4_icon",  "fa-crosshairs"

      m.Add "specCard5_label", "Adjustability"
      m.Add "specCard5_value", "6 degrees of freedom"
      m.Add "specCard5_desc",  "Height, arm hinge, horizontal slide, pivot, tilt and 30&nbsp;mm of fine height adjust at every mount. Set once and then it locks solid."
      m.Add "specCard5_icon",  "fa-sliders"

      m.Add "specCard6_label", "Materials"
      m.Add "specCard6_value", "All-steel, no plastic"
      m.Add "specCard6_desc",  "Unlike cheaper imports we do not use plastic parts or cheap alloys. The entire frame, column and arms are powder-coated steel."
      m.Add "specCard6_icon",  "fa-cubes"

      m.Add "inTheBox", Array( _
            "Base &amp; central column", _
            "2 &times; arm assembly", _
            "2 &times; VESA plates", _
            "All fixings", _
            "Cable management ties", _
            "All assembly tools", _
            "<a href=""/synergy-assembly.pdf"" style=""color:inherit;"">Assembly guide</a>" _
          )

      m.Add "microAssembly",    "20&ndash;30 min assembly"
      m.Add "microAssemblySub", "No drilling, no wall fixings"
      m.Add "microTools",       "All tools included"
      m.Add "microToolsSub",    "Nothing to buy separately"
      m.Add "microWarranty",    "Lifetime warranty"
      m.Add "microWarrantySub", "On every steel part, forever"

      m.Add "vesaIntro1", "If you already have screens or are purchasing some new ones then look for a <b style=""color:var(--ink);"">VESA 75</b> or <b style=""color:var(--ink);"">VESA 100</b> rating, sometimes described as a '<b style=""color:var(--ink);"">Wall mount interface</b>'."
      m.Add "vesaIntro2", "This is the four screw holes in the back of a screen in a square configuration. If your monitor has them then it should be compatible with this Synergy Stand."

      m.Add "monMaxWidth",  "630"
      m.Add "monMaxHeight", "650"
      m.Add "monMaxNote1",  "These limits still leave room for a <b style=""color:var(--ink);"">gentle curve</b> across a multi-screen setup."
      m.Add "monMaxNote2",  "The <b style=""color:var(--ink);"">650&nbsp;mm</b> height assumes the VESA pattern sits roughly in the middle of the screen. If your VESA holes are towards the top on the back of your screen then the max height may be reduced somewhat."

      m.Add "dimsLead", "The Quad Square takes roughly the same desk space as a single 27&Prime; monitor on its stock stand &mdash; and gives you four screens in its place. A single central column keeps the desk surface below the screens clear for keyboard, notebook and coffee."

      m.Add "dimStat1_value", "400<span class=""u"">mm</span>"
      m.Add "dimStat1_label", "Base width"
      m.Add "dimStat2_value", "310<span class=""u"">mm</span>"
      m.Add "dimStat2_label", "Base depth"
      m.Add "dimStat3_value", "400<span class=""u"">mm</span>"
      m.Add "dimStat3_label", "Column height"
      m.Add "dimStat4_value", "360<span class=""u"">mm</span>"
      m.Add "dimStat4_label", "Arm reach (each)"
      m.Add "dimStat5_value", "8<span class=""u"">kg</span>"
      m.Add "dimStat5_label", "Max per mount"
      m.Add "dimStat6_value", "12.5<span class=""u"">kg</span>"
      m.Add "dimStat6_label", "Weight without screens"

      m.Add "dimImgFront",    "/images/stands/dim-2h.jpg"
      m.Add "dimImgFrontAlt", "Front-elevation engineering drawing of the Quad Square Synergy Stand with dimensions in millimetres"
      m.Add "dimImgSide",     "/images/stands/dim-side-short.jpg"
      m.Add "dimImgSideAlt",  "Side-elevation engineering drawing of the Quad Square Synergy Stand with dimensions in millimetres"

      m.Add "dimPdf", "/synergy-assembly.pdf"
    
    Case "MM-S2V"  ' Dual Vertical Synergy Stand 
      Set m = Server.CreateObject("Scripting.Dictionary")
      m.Add "eyebrow", "Synergy Stand &middot; 2-screen vertical"
      m.Add "Title", "Dual Vertical"
      m.Add "pitch",   "Two screens held securely one over the other. All-steel, UK-designed and UK-built &mdash; a square 2&times;2 layout that hides cables behind the column and scales up to six screens later."
      m.Add "ColumnText", "One"
      m.Add "ScreenNum", "2"
      m.Add "ScreenText", "Two"

      m.Add "specCard1_label", "Layout"
      m.Add "specCard1_value", "2 screens, vertical"
      m.Add "specCard1_desc",  "Two screens held securely one on top of the other. Perfect for maximising desk space and fitting two screens into a smaller space."
      m.Add "specCard1_icon",  "fa-th-large"

      m.Add "specCard2_label", "Max screen size"
      m.Add "specCard2_value", "Up to 49&Prime; per screen"
      m.Add "specCard2_desc",  "The single pole design has no horizontal width restriction so can handle a pair of large ultrawide screens."
      m.Add "specCard2_icon",  "fa-expand"

      m.Add "specCard3_label", "Max weight"
      m.Add "specCard3_value", "15&nbsp;kg per screen"
      m.Add "specCard3_desc",  "Our all-steel construction makes handling heavier screens easy. Not something weaker stands with plastic parts or cheaper alloys can do."
      m.Add "specCard3_icon",  "fa-balance-scale"

      m.Add "specCard4_label", "VESA"
      m.Add "specCard4_value", "75&times;75 &amp; 100&times;100"
      m.Add "specCard4_desc",  "Fits every monitor we sell, and is compatible with any 'wall mountable' screen that uses VESA 75 or 100 interfaces."
      m.Add "specCard4_icon",  "fa-crosshairs"

      m.Add "specCard5_label", "Adjustability"
      m.Add "specCard5_value", "Height and tilt adjustment"
      m.Add "specCard5_desc",  "Mount screens at any height on the central column and then tilt them up and down. Position the stand anywhere on your desk with the freestanding base."
      m.Add "specCard5_icon",  "fa-sliders"

      m.Add "specCard6_label", "Materials"
      m.Add "specCard6_value", "All-steel, no plastic"
      m.Add "specCard6_desc",  "Unlike cheaper imports we do not use plastic parts or cheap alloys. The entire frame, column and arms are powder-coated steel."
      m.Add "specCard6_icon",  "fa-cubes"

      m.Add "inTheBox", Array( _
            "Base &amp; central column", _
            "2 &times; arm assembly", _
            "2 &times; VESA plates", _
            "All fixings", _
            "Cable management ties", _
            "All assembly tools", _
            "<a href=""/synergy-assembly.pdf"" style=""color:inherit;"">Assembly guide</a>" _
          )

      m.Add "microAssembly",    "20&ndash;30 min assembly"
      m.Add "microAssemblySub", "No drilling, no wall fixings"
      m.Add "microTools",       "All tools included"
      m.Add "microToolsSub",    "Nothing to buy separately"
      m.Add "microWarranty",    "Lifetime warranty"
      m.Add "microWarrantySub", "On every steel part, forever"

      m.Add "vesaIntro1", "If you already have screens or are purchasing some new ones then look for a <b style=""color:var(--ink);"">VESA 75</b> or <b style=""color:var(--ink);"">VESA 100</b> rating, sometimes described as a '<b style=""color:var(--ink);"">Wall mount interface</b>'."
      m.Add "vesaIntro2", "This is the four screw holes in the back of a screen in a square configuration. If your monitor has them then it should be compatible with this Synergy Stand."

      m.Add "monMaxWidth",  "630"
      m.Add "monMaxHeight", "650"
      m.Add "monMaxNote1",  "These limits still leave room for a <b style=""color:var(--ink);"">gentle curve</b> across a multi-screen setup."
      m.Add "monMaxNote2",  "The <b style=""color:var(--ink);"">650&nbsp;mm</b> height assumes the VESA pattern sits roughly in the middle of the screen. If your VESA holes are towards the top on the back of your screen then the max height may be reduced somewhat."

      m.Add "dimsLead", "The Quad Square takes roughly the same desk space as a single 27&Prime; monitor on its stock stand &mdash; and gives you four screens in its place. A single central column keeps the desk surface below the screens clear for keyboard, notebook and coffee."

      m.Add "dimStat1_value", "400<span class=""u"">mm</span>"
      m.Add "dimStat1_label", "Base width"
      m.Add "dimStat2_value", "310<span class=""u"">mm</span>"
      m.Add "dimStat2_label", "Base depth"
      m.Add "dimStat3_value", "400<span class=""u"">mm</span>"
      m.Add "dimStat3_label", "Column height"
      m.Add "dimStat4_value", "360<span class=""u"">mm</span>"
      m.Add "dimStat4_label", "Arm reach (each)"
      m.Add "dimStat5_value", "8<span class=""u"">kg</span>"
      m.Add "dimStat5_label", "Max per mount"
      m.Add "dimStat6_value", "12.5<span class=""u"">kg</span>"
      m.Add "dimStat6_label", "Weight without screens"

      m.Add "dimImgFront",    "/images/stands/dim-2h.jpg"
      m.Add "dimImgFrontAlt", "Front-elevation engineering drawing of the Quad Square Synergy Stand with dimensions in millimetres"
      m.Add "dimImgSide",     "/images/stands/dim-side-short.jpg"
      m.Add "dimImgSideAlt",  "Side-elevation engineering drawing of the Quad Square Synergy Stand with dimensions in millimetres"

      m.Add "dimPdf", "/synergy-assembly.pdf"

    ' ----------------------------------------------------------
    ' Placeholders for other Synergy stand SKUs.
    ' Dimension drawings already on disk in /images/stands/:
    '   dim-2h.jpg / dim-2v.jpg       2-screen horizontal / vertical
    '   dim-3h.jpg / dim-3p.jpg       3-screen horizontal / pyramid
    '   dim-4p.jpg                    4-screen pyramid
    '   dim-5p.jpg                    5-screen
    '   dim-6r.jpg                    6-screen radial
    '   dim-side-short.jpg            Side-elevation, short column
    ' Fill in matching products.sku entries when each SKU lands
    ' in the products table. Until then they fall back to the
    ' page-level hardcoded defaults.
    ' ----------------------------------------------------------
    ' Case "MM-S2H"   ' 2-screen horizontal
    ' Case "MM-S2V"   ' 2-screen vertical
    ' Case "MM-S3H"   ' 3-screen horizontal
    ' Case "MM-S3P"   ' 3-screen pyramid
    ' Case "MM-S4P"   ' 4-screen pyramid
    ' Case "MM-S5P"   ' 5-screen
    ' Case "MM-S6R"   ' 6-screen radial

  End Select

  Set mmGetStandMeta = m
End Function
%>
