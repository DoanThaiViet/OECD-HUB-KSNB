/* ============================================================
   SHELL dùng chung 7 trang: badge KSNB (SVG inline), footer,
   FAB về Hub / lên đầu, fade-up .fx theo IntersectionObserver.
   Nạp SAU data.js. Không localStorage, không phụ thuộc mạng.
   ============================================================ */
(function () {
  "use strict";
  const META = (window.OECD && window.OECD.META) || { updated: "" };

  /* ---------- Badge KSNB: SVG shield-check inline ---------- */
  function badgeSVG() {
    return '<svg viewBox="0 0 20 20" aria-hidden="true"><circle cx="10" cy="10" r="9" fill="#00A651"/><path d="M6 10.2l2.6 2.6L14 7.4" fill="none" stroke="#fff" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/></svg>';
  }
  window.ksnbBadge = function (txt) {
    return '<span class="ksnb-badge">' + badgeSVG() + '<span>' + (txt || "Ban Kiểm soát nội bộ xây dựng") + "</span></span>";
  };
  // tự gắn vào các chỗ đánh dấu sẵn
  document.querySelectorAll("[data-ksnb-badge]").forEach(function (el) {
    el.insertAdjacentHTML("beforeend", window.ksnbBadge(el.getAttribute("data-ksnb-badge") || undefined));
  });

  /* ---------- Footer chạy chung ---------- */
  if (!document.querySelector(".cfoot") && document.body.dataset.nofoot === undefined) {
    const f = document.createElement("div");
    f.className = "cfoot";
    f.innerHTML = window.ksnbBadge("Ban Kiểm soát nội bộ xây dựng và vận hành") +
      '<span>OECD Governance Command Center · PVEP</span>' +
      '<span style="margin-left:auto">Cập nhật ' + META.updated + "</span>";
    document.body.appendChild(f);
  }

  /* ---------- FAB: về Hub + lên đầu (hiện khi cuộn > 1 màn) ---------- */
  const isHub = /(^|\/)index\.html$|\/$/.test(location.pathname) || document.body.dataset.page === "hub";
  const fab = document.createElement("div");
  fab.className = "fabnav noprint";
  fab.innerHTML = (isHub ? "" : '<a href="index.html" title="Về Hub">⌂</a>') +
                  '<button type="button" title="Lên đầu trang">↑</button>';
  document.body.appendChild(fab);
  fab.querySelector("button").addEventListener("click", function () {
    scrollTo({ top: 0, behavior: "smooth" });
  });
  const sentinel = document.createElement("div");
  sentinel.style.cssText = "position:absolute;top:0;height:" + Math.round(innerHeight * 0.9) + "px;width:1px;pointer-events:none";
  document.body.prepend(sentinel);
  new IntersectionObserver(function (en) {
    fab.classList.toggle("show", !en[0].isIntersecting);
  }).observe(sentinel);

  /* ---------- fade-up .fx ---------- */
  const io = new IntersectionObserver(function (ents) {
    ents.forEach(function (e) { if (e.isIntersecting) { e.target.classList.add("inview"); io.unobserve(e.target); } });
  }, { threshold: 0.08 });
  document.querySelectorAll(".fx").forEach(function (el) { io.observe(el); });
  setTimeout(function () { document.querySelectorAll(".fx").forEach(function (el) { el.classList.add("inview"); }); }, 900);
})();

/* ============================================================
   ZOOM TRANSITION hub ↔ phân hệ (chỉ chạy trên trang theme-phN)
   - Reveal (G4): đến từ hub qua click-to-zoom → wash màu phân hệ
     tan dần, header vào trước, các khối nội dung stagger sau.
   - Back: bấm về Hub → thu về màu phân hệ rồi mới điều hướng;
     hub đọc sessionStorage "oecdReturn" để chạy pull-back.
   ============================================================ */
(function () {
  "use strict";
  var PHC = { ph1: "#0E9F6E", ph2: "#D97706", ph3: "#2563EB", ph4: "#0D9488", ph5: "#7C3AED" };
  var m = document.body.className.match(/theme-(ph\d)/);
  if (!m) return;
  var ph = m[1], col = PHC[ph] || "#00A651";
  var RM = matchMedia("(prefers-reduced-motion: reduce)").matches;
  var wash = "radial-gradient(140% 140% at 50% 46%, color-mix(in srgb," + col + " 62%, #0A1C33) 0%, #061222 80%)";

  /* ---------- G4 · DASHBOARD REVEAL ---------- */
  var z = null;
  try { z = JSON.parse(sessionStorage.getItem("oecdZoom") || "null"); } catch (e) {}
  sessionStorage.removeItem("oecdZoom");
  if (z && !RM && z.t && Date.now() - z.t < 15000) {
    var ov = document.createElement("div");
    ov.style.cssText = "position:fixed;inset:0;z-index:999;pointer-events:none;background:" + wash;
    document.body.appendChild(ov);
    var head = document.querySelector(".cbar");
    var parts = [].slice.call(document.body.children).filter(function (el) {
      return el !== ov && !/^(SCRIPT|STYLE|LINK)$/.test(el.tagName) && el.offsetHeight > 0;
    }).slice(0, 9);
    var idx = 0;
    parts.forEach(function (el) {
      var isHead = el === head;
      if (!isHead) idx++;
      el.animate(
        [{ opacity: 0, transform: "translateY(16px)" }, { opacity: 1, transform: "none" }],
        { duration: 460, delay: isHead ? 0 : 150 + idx * 85, easing: "cubic-bezier(.22,.61,.36,1)", fill: "backwards" });
    });
    ov.animate([{ opacity: 1 }, { opacity: 0 }], { duration: 580, easing: "ease-out" })
      .onfinish = function () { ov.remove(); };
    setTimeout(function () { if (ov.parentNode) ov.remove(); }, 900);   /* phòng tab bị throttle animation */
  }

  /* ---------- BACK: thu về màu phân hệ rồi về Hub ---------- */
  if (!RM) {
    document.addEventListener("click", function (e) {
      var a = e.target.closest ? e.target.closest('a[href$="index.html"]') : null;
      if (!a || e.defaultPrevented || e.metaKey || e.ctrlKey || e.shiftKey || e.button !== 0) return;
      e.preventDefault();
      try { sessionStorage.setItem("oecdReturn", JSON.stringify({ ph: ph, t: Date.now() })); } catch (err) {}
      var ov = document.createElement("div");
      ov.style.cssText = "position:fixed;inset:0;z-index:999;pointer-events:none;opacity:0;background:" + wash;
      document.body.appendChild(ov);
      document.body.animate([{ opacity: 1 }, { opacity: .4 }], { duration: 300, easing: "ease-in", fill: "forwards" });
      var gone = false;
      var nav = function () { if (gone) return; gone = true; location.href = a.href; };
      ov.animate([{ opacity: 0 }, { opacity: 1 }], { duration: 300, easing: "ease-in", fill: "forwards" }).onfinish = nav;
      setTimeout(nav, 420);   /* đảm bảo luôn điều hướng dù animation bị throttle */
    }, true);
  }
})();
