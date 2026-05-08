<%
' ============================================================
' monitorSpecs.asp
' Per-monitor reference content for shop/pc/viewprd-monitor-v2.asp.
'
' Identity fields (name, price, image, SKU, short description)
' come from the products table. This file holds the *static*
' reference content the DB has no column for — eyebrow category
' line, inclusion-strip warranty/dispatch copy, and the two
' spec-sheet tables.
'
' Keyed by SKU. SKU is the human-readable, immutable product
' identifier and is already loaded by the page, so no extra DB
' lookup is needed.
'
' Sparse fields per entry — the page falls back to its own
' hardcoded defaults when a key is missing or the SKU is unknown,
' so a partial entry never breaks rendering.
'
' Keys:
'   eyebrow              Category line above the H1
'                        (e.g. "Monitor &middot; 27&Prime; Quad HD IPS")
'   inclWarrantyTitle    Inclusion-strip warranty title
'   inclWarrantySub      Inclusion-strip warranty sub-line
'   inclDispatchTitle    Inclusion-strip dispatch title
'   inclDispatchSub      Inclusion-strip dispatch sub-line
'   specDisplayHeading   <h4> above the Display spec table
'   specDisplayRows      Array of Array(label, value) pairs
'   specConnHeading      <h4> above the Connectivity spec table
'   specConnRows         Array of Array(label, value) pairs
' ============================================================
Function mmGetMonitorMeta(ByVal sku)
  Dim m : Set m = Nothing
  Dim mmKey : mmKey = UCase(Trim(sku & ""))

  Select Case mmKey
    Case "MM-MIIYAMA23"  ' 24" Iiyama Full HD IPS
      Set m = Server.CreateObject("Scripting.Dictionary")
      m.Add "eyebrow", "Monitor &middot; 24&Prime; Full HD IPS"
      m.Add "inclWarrantyTitle", "3-year warranty"
      m.Add "inclWarrantySub",   "Iiyama cover"
      m.Add "inclDispatchTitle", "2-day dispatch"
      m.Add "inclDispatchSub",   "UK courier"
      m.Add "specDisplayHeading", "24&Prime; Iiyama Widescreen"
      m.Add "specDisplayRows", Array( _
            Array("Manufacturer",  "Iiyama"), _
            Array("Size",          "24&Prime; (diagonal)"), _
            Array("Resolution",    "1920 &times; 1080 (Full HD)"), _
            Array("Refresh rate",  "100 Hz"), _
            Array("Response time", "1 ms (GtG)"), _
            Array("Pixel pitch", "0.233mm"), _
            Array("Panel type",    "IPS") _
          )
      m.Add "specConnHeading", "Ports, mounts &amp; what&rsquo;s in the box"
      m.Add "specConnRows", Array( _
            Array("Inputs",          "HDMI &middot; DisplayPort"), _
            Array("Bezel width",     "Ultra thin (three-side frameless)"), _
            Array("Dimensions",      "540mm (width) &times; 315mm (height)"), _
            Array("Weight",          "3.3Kg"), _
            Array("VESA mount",      "100 &times; 100"), _
            Array("Warranty",        "3-year manufacturer (Iiyama)"), _
            Array("Cables included", "HDMI <em>or</em> Displayport &middot; UK power") _
          )

    Case "MM-MIY27QHD"  ' 27" Iiyama Quad HD IPS
      Set m = Server.CreateObject("Scripting.Dictionary")
      m.Add "eyebrow", "Monitor &middot; 27&Prime; Quad HD IPS"
      m.Add "inclWarrantyTitle", "3-year warranty"
      m.Add "inclWarrantySub",   "Iiyama cover"
      m.Add "inclDispatchTitle", "2-day dispatch"
      m.Add "inclDispatchSub",   "UK courier"
      m.Add "specDisplayHeading", "27&Prime; Iiyama Widescreen"
      m.Add "specDisplayRows", Array( _
            Array("Manufacturer",  "Iiyama"), _
            Array("Size",          "27&Prime; (diagonal)"), _
            Array("Resolution",    "2560 &times; 1440 (Quad HD)"), _
            Array("Refresh rate",  "100 Hz"), _
            Array("Response time", "1 ms (GtG)"), _
            Array("Pixel pitch", "0.233mm"), _
            Array("Panel type",    "IPS") _
          )
      m.Add "specConnHeading", "Ports, mounts &amp; what&rsquo;s in the box"
      m.Add "specConnRows", Array( _
            Array("Inputs",          "HDMI &middot; DisplayPort"), _
            Array("Bezel width",     "Ultra thin (three-side frameless)"), _
            Array("Dimensions",      "615mm (width) &times; 370mm (height)"), _
            Array("Weight",          "3.3Kg"), _
            Array("VESA mount",      "100 &times; 100"), _
            Array("Warranty",        "3-year manufacturer (Iiyama)"), _
            Array("Cables included", "HDMI <em>or</em> Displayport &middot; UK power") _
          )

    ' ----------------------------------------------------------
    ' Placeholders for the other 5 monitors. Fill in the SKU
    ' (matching products.sku) and content when each is added to
    ' the products table. Until then, requests for those URLs
    ' won't reach this page (router 404s on missing slug).
    ' ----------------------------------------------------------

    ' Case "MM-AOC22"  ' 21.5" AOC 24B2XH — Full HD
    ' Case "MM-ACR24"  ' 24"   Acer K243Y E — Full HD
    ' Case "MM-ACR27"  ' 27"   Acer KA272 E — Full HD
    ' Case "MM-IIY24"  ' 24"   Iiyama ProLite XUB2493HS — Full HD
    ' Case "MM-AOC27"  ' 27"   AOC Q27P2Q — Quad HD

  End Select

  Set mmGetMonitorMeta = m
End Function
%>
