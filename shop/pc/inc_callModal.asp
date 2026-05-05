<%
' ============================================================
' inc_callModal.asp
' Site-wide "call us" modal + its click handler.
' Included once from footer_wrapper.asp — markup ships in the
' DOM on every storefront page, the inline IIFE wires up the
' desktop tel:-intercept and the .js-book-call trigger.
' Styles live in css/mm-site.css chrome section.
' Form submission stubbed; email wiring is a follow-up plan.
' ============================================================
%>
<div class="call-modal" id="callModal" role="dialog" aria-modal="true" aria-labelledby="callModalTitle" aria-hidden="true">
  <div class="call-modal-backdrop" data-close></div>
  <div class="call-modal-dialog" role="document">
    <button type="button" class="call-modal-close" aria-label="Close" data-close>&times;</button>

    <div class="call-modal-hd">
      <p class="call-modal-eyebrow">Call us now</p>
      <a href="tel:03302236655" class="call-modal-number">0330 223 66 55</a>
      <p class="call-modal-hours">Mon&ndash;Fri, 9:00am&ndash;5:00pm</p>
    </div>

    <div class="call-modal-divider"><span>or</span></div>

    <div class="call-modal-bd">
      <h3 id="callModalTitle">Prefer we call you back?</h3>
      <p class="call-modal-sub">Drop your details below and we will get back to you.</p>
      <form class="call-modal-form" action="/shop/pc/callback_submit.asp" method="post" novalidate>
        <label class="call-modal-field">
          <span>Your name</span>
          <input type="text" name="name" required autocomplete="name" maxlength="80">
        </label>
        <label class="call-modal-field">
          <span>Phone number</span>
          <input type="tel" name="phone" required autocomplete="tel" inputmode="tel" maxlength="40">
        </label>
        <label class="call-modal-field">
          <span>When suits you?</span>
          <input type="text" name="time" required placeholder="e.g. Tomorrow afternoon, Weds 2pm" maxlength="120">
        </label>
        <label class="call-modal-field">
          <span>Email <em>(optional)</em></span>
          <input type="email" name="email" autocomplete="email" maxlength="120">
        </label>

        <!-- Anti-spam honeypot: hidden from real users, bots fill it. -->
        <div class="call-modal-hp" aria-hidden="true">
          <label>
            Website (leave blank)
            <input type="text" name="website" tabindex="-1" autocomplete="off">
          </label>
        </div>
        <input type="hidden" name="fillMs" value="">

        <button type="submit" class="call-modal-submit">
          <i class="fa fa-calendar-check-o"></i>Request a callback
        </button>
      </form>
    </div>
  </div>
</div>

<script>
(function(){
  var modal = document.getElementById('callModal');
  if (!modal) return;

  var form = modal.querySelector('.call-modal-form');
  var eyebrow = modal.querySelector('.call-modal-eyebrow');
  var bodyEl = document.body;
  var openedAt = 0;

  function isDesktop(){
    return window.matchMedia('(hover: hover) and (pointer: fine)').matches;
  }

  // Return a context-aware prompt for the modal heading based on UK
  // wall-clock time, so an overseas visitor sees the right status
  // regardless of their own timezone. Business hours are Mon-Fri
  // 9:00 to 17:00 UK time.
  function callPhrase(){
    var uk = new Date(new Date().toLocaleString('en-US', { timeZone: 'Europe/London' }));
    var day = uk.getDay();       // 0=Sun ... 6=Sat
    var hour = uk.getHours();
    var isWeekday = day >= 1 && day <= 5;

    if (isWeekday && hour >= 9 && hour < 17) return 'Call us now';
    if (isWeekday && hour < 9)               return 'Call us later today';

    // Find the next weekday.
    var daysAhead = 1;
    while (true) {
      var d = (day + daysAhead) % 7;
      if (d >= 1 && d <= 5) break;
      daysAhead++;
    }
    if (daysAhead === 1) return 'Call us tomorrow';

    var dayNames = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
    return 'Call us ' + dayNames[(day + daysAhead) % 7];
  }

  function open(){
    if (eyebrow) eyebrow.textContent = callPhrase();
    modal.classList.add('is-open');
    modal.setAttribute('aria-hidden', 'false');
    bodyEl.style.overflow = 'hidden';
    openedAt = Date.now();
    clearError();
    setTimeout(function(){
      var first = form && form.querySelector('input[name="name"]');
      if (first) first.focus();
    }, 50);
  }

  function close(){
    modal.classList.remove('is-open');
    modal.setAttribute('aria-hidden', 'true');
    bodyEl.style.overflow = '';
  }

  // Delegated open trigger — covers every tel: link on the page
  // (desktop only) and every .js-book-call button (all viewports).
  document.addEventListener('click', function(e){
    var book = e.target.closest('.js-book-call');
    if (book) { e.preventDefault(); open(); return; }

    // Don't intercept clicks originating inside the modal. Its own
    // tel: number is rendered inert on desktop via CSS pointer-events.
    if (modal.contains(e.target)) return;

    var tel = e.target.closest('a[href^="tel:"]');
    if (tel && isDesktop()) { e.preventDefault(); open(); }
  });

  // Close triggers — backdrop / x button (both marked [data-close]) and Escape.
  modal.addEventListener('click', function(e){
    if (e.target.closest('[data-close]')) close();
  });
  document.addEventListener('keydown', function(e){
    if (e.key === 'Escape' && modal.classList.contains('is-open')) close();
  });

  // -------- Form submission --------

  function showError(msg){
    clearError();
    if (!form) return;
    var div = document.createElement('div');
    div.className = 'call-modal-error';
    div.setAttribute('role', 'alert');
    div.textContent = msg;
    form.appendChild(div);
  }

  function clearError(){
    if (!form) return;
    var existing = form.querySelector('.call-modal-error');
    if (existing) existing.parentNode.removeChild(existing);
  }

  function errorMessage(code){
    switch (code) {
      case 'missing': return 'Please fill in your name, phone and preferred time.';
      case 'phone':   return 'That phone number doesn’t look right. Please check and try again.';
      case 'email':   return 'That email address doesn’t look right. You can also leave it blank.';
      case 'send':    return 'Sorry, we couldn’t send your request just now. Please try again, or give us a call.';
      default:        return 'Sorry, something went wrong. Please try again, or give us a call.';
    }
  }

  function showSuccess(){
    form.innerHTML =
      '<div class="call-modal-success">' +
        '<i class="fa fa-check-circle"></i>' +
        '<h3>Thanks, we&rsquo;ll be in touch.</h3>' +
        '<p>We will try and call you at your preferred time.</p>' +
      '</div>';
  }

  if (form) {
    form.addEventListener('submit', function(e){
      e.preventDefault();

      // Browser-native required-field check first.
      var fd = new FormData(form);
      if (!fd.get('name') || !fd.get('phone') || !fd.get('time')) {
        form.reportValidity();
        return;
      }

      // Stamp time-to-submit for the server-side anti-spam gate.
      var fillMsField = form.querySelector('input[name="fillMs"]');
      if (fillMsField) {
        fillMsField.value = String(Math.max(0, Date.now() - openedAt));
      }

      // Build a URL-encoded body. Classic ASP's Request.Form does not
      // parse multipart/form-data without a third-party component, so
      // FormData would arrive empty on the server.
      var params = new URLSearchParams();
      var inputs = form.querySelectorAll('input[name], textarea[name], select[name]');
      for (var i = 0; i < inputs.length; i++) {
        params.append(inputs[i].name, inputs[i].value || '');
      }

      var submitBtn = form.querySelector('.call-modal-submit');
      function lockUI(){ if (submitBtn) { submitBtn.disabled = true; submitBtn.classList.add('is-busy'); } }
      function unlockUI(){ if (submitBtn) { submitBtn.disabled = false; submitBtn.classList.remove('is-busy'); } }

      clearError();
      lockUI();

      fetch(form.getAttribute('action') || '/shop/pc/callback_submit.asp', {
        method: 'POST',
        body: params,
        credentials: 'same-origin'
      })
      .then(function(r){
        if (!r.ok) throw new Error('http');
        return r.json();
      })
      .then(function(data){
        unlockUI();
        if (data && data.ok) {
          showSuccess();
        } else {
          showError(errorMessage(data && data.error));
        }
      })
      .catch(function(){
        unlockUI();
        showError(errorMessage());
      });
    });
  }
})();
</script>
