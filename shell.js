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
