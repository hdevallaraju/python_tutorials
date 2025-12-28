function setLastUpdated() {
  const el = document.getElementById("last-updated");
  if (!el) return;
  const d = new Date();
  el.textContent = d.toLocaleDateString(undefined, { year: "numeric", month: "short", day: "2-digit" });
}

function initTabs() {
  const buttons = [
    { btn: document.getElementById("tab-sales-btn"), panel: document.getElementById("tab-sales") },
    { btn: document.getElementById("tab-collab-btn"), panel: document.getElementById("tab-collab") },
    { btn: document.getElementById("tab-tech-btn"), panel: document.getElementById("tab-tech") },
    { btn: document.getElementById("tab-diagram-btn"), panel: document.getElementById("tab-diagram") },
  ].filter((x) => x.btn && x.panel);

  if (buttons.length === 0) return;

  function activate(targetBtn) {
    for (const { btn, panel } of buttons) {
      const isActive = btn === targetBtn;
      btn.classList.toggle("is-active", isActive);
      btn.setAttribute("aria-selected", isActive ? "true" : "false");
      panel.classList.toggle("is-active", isActive);
    }
  }

  for (const { btn } of buttons) {
    btn.addEventListener("click", () => activate(btn));
  }

  // Optional: allow deep-links like #diagram / #technology / #sales / #collaborations
  const hash = (window.location.hash || "").replace("#", "").toLowerCase();
  const map = {
    sales: "tab-sales-btn",
    collaborations: "tab-collab-btn",
    technology: "tab-tech-btn",
    diagram: "tab-diagram-btn",
  };
  const targetId = map[hash];
  if (targetId) {
    const t = document.getElementById(targetId);
    if (t) activate(t);
  }
}

setLastUpdated();
initTabs();

