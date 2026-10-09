/* Ashercrest concept polish — scroll reveals + concept remount hooks */
(function () {
  function observeReveals(root) {
    var nodes = (root || document).querySelectorAll(".ac-reveal:not(.is-in)");
    if (!nodes.length) return;
    if (!("IntersectionObserver" in window)) {
      nodes.forEach(function (n) { n.classList.add("is-in"); });
      return;
    }
    var io = new IntersectionObserver(
      function (entries) {
        entries.forEach(function (entry) {
          if (entry.isIntersecting) {
            entry.target.classList.add("is-in");
            io.unobserve(entry.target);
          }
        });
      },
      { rootMargin: "0px 0px -8% 0px", threshold: 0.12 }
    );
    nodes.forEach(function (n) { io.observe(n); });
  }

  function restartHeroMotion(root) {
    var media = (root || document).querySelectorAll(".ac-hero-media");
    media.forEach(function (el) {
      el.style.animation = "none";
      // force reflow
      void el.offsetWidth;
      el.style.animation = "";
    });
    var copy = (root || document).querySelectorAll(".ac-hero-copy > *");
    copy.forEach(function (el) {
      el.style.animation = "none";
      void el.offsetWidth;
      el.style.animation = "";
    });
  }

  function boot() {
    observeReveals(document);
    restartHeroMotion(document);
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", boot);
  } else {
    boot();
  }

  // Re-run when concept switcher remounts content
  var mo = new MutationObserver(function () {
    window.clearTimeout(mo._t);
    mo._t = window.setTimeout(function () {
      observeReveals(document);
      restartHeroMotion(document);
    }, 40);
  });
  mo.observe(document.body, { childList: true, subtree: true });
})();
