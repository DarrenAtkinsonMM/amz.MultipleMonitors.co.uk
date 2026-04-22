<%
' ============================================================
' Shared "talk to Darren" founder CTA block. Uses mmMachineName
' (defaulting to "machine") if the per-page ASP file sets it,
' so the copy can reference the product by name.
' ============================================================
Dim mmDarrenMachine
If IsEmpty(mmMachineName) Or mmMachineName = "" Then
  mmDarrenMachine = "machine"
Else
  mmDarrenMachine = mmMachineName
End If
%>
<section class="darren" id="darren">
  <div class="container">
    <div class="darren-grid">
      <div class="darren-photo reveal">
        <img src="/images/pages/darren.jpg" alt="Darren Atkinson, founder of Multiple Monitors Ltd">
      </div>
      <div class="reveal" style="transition-delay:.08s">
        <h5>Still deciding?</h5>
        <h2>Talk to <em>Darren</em> &mdash; the founder, not a call centre.</h2>
        <p>If you want a <%= mmDarrenMachine %> spec&rsquo;d for exactly what you do &mdash; which broker, which platform, how many charts, whether you backtest overnight &mdash; fifteen minutes on the phone usually does it. No pushy sales, no corner-cutting. Just the right machine for the workload.</p>
        <div class="darren-ctas">
          <a href="tel:03302236655" class="btn btn-primary btn-lg"><i class="fa fa-phone"></i>0330 223 66 55</a>
          <a href="/contact/" class="btn btn-ghost btn-lg"><i class="fa fa-calendar"></i>Book a 15-min call</a>
        </div>
        <div class="darren-sig">&mdash; Darren Atkinson, founder, Multiple Monitors Ltd</div>
      </div>
    </div>
  </div>
</section>
