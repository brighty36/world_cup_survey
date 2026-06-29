/* =====================================================================
   embed-child.js  ·  Voice of the Trade · 2026 World Cup Brand Tracker

   Loaded by each tab page (insights / explore / compare). When the page is
   shown inside the index.html shell it:
     1. reports its content height up to the shell, so the shell can size the
        tab's iframe and forward the total height to the host site; and
     2. routes in-app "[data-nav]" links through the shell (postMessage) so a
        tab can switch tabs without the link breaking out of the host iframe.

   When a tab page is opened on its own (not inside the shell) it does nothing
   and the [data-nav] links fall back to their normal href.
   ===================================================================== */
(function(){
  var inShell = window.parent && window.parent !== window;

  // ---- report content height up to the shell ----
  var last = -1, pending = false;
  function measure(){ return Math.ceil(document.documentElement.getBoundingClientRect().height); }
  function post(){
    pending = false;
    var h = measure();
    if(h === last) return;          // only post on a real change -> no resize loop
    last = h;
    try{ window.parent.postMessage({ type:"wc-embed-size", height:h }, "*"); }catch(e){}
  }
  function schedule(){ if(pending) return; pending = true; requestAnimationFrame(post); }

  if(inShell){
    window.addEventListener("resize", schedule);
    window.addEventListener("load", function(){ schedule(); setTimeout(schedule, 400); });
    new MutationObserver(schedule).observe(document.body, { attributes:true, childList:true, subtree:true });
    if(document.fonts && document.fonts.ready) document.fonts.ready.then(schedule);
    schedule();
  }

  // ---- route in-app tab links through the shell ----
  document.addEventListener('click', function(e){
    var a = e.target.closest ? e.target.closest('[data-nav]') : null;
    if(!a) return;
    if(inShell){
      e.preventDefault();
      window.parent.postMessage({ type:"wc-nav", panel:a.getAttribute('data-nav') }, "*");
    }
    // not embedded: let the href navigate to index.html#<panel> as before
  });
})();
