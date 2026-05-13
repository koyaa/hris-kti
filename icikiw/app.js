const quickActions = [
  { label: "Add Contact", icon: "contact" },
  { label: "New Proposal", icon: "proposal" },
  { label: "New Template", icon: "template" },
  { label: "New Board Book", icon: "book" },
  { label: "Get Help", icon: "help" },
  { label: "User Settings", icon: "settings" },
];

const resultTabs = [
  { label: "Tasks", icon: "tasks" },
  { label: "Proposals", icon: "proposal" },
  { label: "Grantees", icon: "office" },
  { label: "Documents", icon: "document" },
  { label: "Forms", icon: "form" },
];

const results = [
  {
    label: "Farmlink Project",
    tag: "Nonprofit",
    tagTone: "red",
    icon: "home",
    keywords: ["farmlink", "project", "farm", "nonprofit"],
  },
  {
    label: "frida.jpeg",
    tag: "File",
    tagTone: "orange",
    icon: "file",
    keywords: ["frida", "jpeg", "image", "file"],
  },
  {
    label: "FY24-budget-categories-2024-05-02-13-45-00.csv",
    tag: "File",
    tagTone: "orange",
    icon: "file",
    keywords: ["fy24", "budget", "categories", "csv", "file"],
  },
  {
    label: "Final report due for Google Grant",
    tag: "Task",
    tagTone: "violet",
    icon: "taskCheck",
    keywords: ["final", "report", "google", "grant", "task"],
  },
];

// ===== Minimize State =====
// ==========================
const menuState = { minimized: false, buttonEdge: 'right', buttonTop: 200 };

const safeStorage = {
  get(key, fallback) {
    try {
      const v = localStorage.getItem(key);
      return v !== null ? JSON.parse(v) : fallback;
    } catch {
      return fallback;
    }
  },
  set(key, value) {
    try {
      localStorage.setItem(key, JSON.stringify(value));
    } catch { }
  },
};

function saveState() {
  safeStorage.set('menu-minimized', menuState.minimized);
  safeStorage.set('menu-button-pos', { edge: menuState.buttonEdge, top: menuState.buttonTop });
}

function loadState() {
  menuState.minimized = safeStorage.get('menu-minimized', false);
  const pos = safeStorage.get('menu-button-pos', { edge: 'right', top: 200 });
  menuState.buttonEdge = pos.edge;
  menuState.buttonTop = pos.top;
}

loadState();

const iconMarkup = {
  search: `
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round">
      <circle cx="11" cy="11" r="6.5"></circle>
      <path d="M16 16L21 21"></path>
    </svg>
  `,
  lightbulb: `
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">
      <path d="M9 18H15"></path>
      <path d="M10 21H14"></path>
      <path d="M8.5 14.6C7.12 13.43 6.25 11.69 6.25 9.75C6.25 6.3 8.97 3.5 12.33 3.5C15.7 3.5 18.42 6.3 18.42 9.75C18.42 11.71 17.53 13.46 16.13 14.62C15.34 15.28 15 15.77 15 16.5V17.25H9.67V16.5C9.67 15.77 9.33 15.28 8.5 14.6Z"></path>
      <path d="M12.33 1.8V0.8"></path>
      <path d="M4.63 4.1L3.92 3.39"></path>
      <path d="M20.04 4.1L20.75 3.39"></path>
      <path d="M2.75 9.75H1.75"></path>
      <path d="M22.92 9.75H21.92"></path>
    </svg>
  `,
  contact: `
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">
      <rect x="3.5" y="4" width="16.5" height="16" rx="2.5"></rect>
      <circle cx="12" cy="10" r="2.9"></circle>
      <path d="M7.8 17C8.81 14.95 10.15 14 12 14C13.85 14 15.19 14.95 16.2 17"></path>
      <path d="M20.5 7V12"></path>
      <path d="M18 9.5H23"></path>
    </svg>
  `,
  proposal: `
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">
      <path d="M7 3.8H13L18.2 9V20.2H7Z"></path>
      <path d="M12.7 3.8V9H18.2"></path>
      <path d="M10 12H15"></path>
      <path d="M10 15H15"></path>
      <path d="M16.5 14.8V19.5"></path>
      <path d="M14.1 17.15H18.85"></path>
    </svg>
  `,
  template: `
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">
      <path d="M7 3.8H13L18.2 9V20.2H7Z"></path>
      <path d="M12.7 3.8V9H18.2"></path>
      <path d="M10 12.5H14.8"></path>
      <path d="M10 15.5H13.6"></path>
      <rect x="15.7" y="14.2" width="5" height="5" rx="1"></rect>
      <path d="M18.2 12.3V15.8"></path>
      <path d="M16.45 14.05H19.95"></path>
    </svg>
  `,
  book: `
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">
      <rect x="6.2" y="4.2" width="11.6" height="15.6" rx="2.3"></rect>
      <path d="M9.2 8.3H14.8"></path>
      <path d="M9.2 11.8H14.8"></path>
      <path d="M8.5 19.8V21"></path>
      <path d="M15.5 19.8V21"></path>
    </svg>
  `,
  chevronDown: `
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
      <path d="M6 10L12 16"></path>
      <path d="M12 16L18 10"></path>
    </svg>
  `,
  help: `
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">
      <circle cx="12" cy="12" r="7.8"></circle>
      <circle cx="12" cy="12" r="2.7"></circle>
      <path d="M12 4.2V7"></path>
      <path d="M12 17V19.8"></path>
      <path d="M4.2 12H7"></path>
      <path d="M17 12H19.8"></path>
      <path d="M6.4 6.4L8.3 8.3"></path>
      <path d="M15.7 15.7L17.6 17.6"></path>
      <path d="M17.6 6.4L15.7 8.3"></path>
      <path d="M8.3 15.7L6.4 17.6"></path>
    </svg>
  `,
  settings: `
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">
      <circle cx="12" cy="6.8" r="3.1"></circle>
      <path d="M18.4 21C17.5 17.9 15.34 16.35 12 16.35C8.66 16.35 6.5 17.9 5.6 21"></path>
      <path d="M18.2 8.8L20.8 9.5"></path>
      <path d="M17.7 12.1L19.8 13.8"></path>
      <path d="M16.5 15.1L17.5 17.6"></path>
    </svg>
  `,
  tasks: `
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">
      <rect x="4.5" y="5" width="4.2" height="4.2" rx="1"></rect>
      <rect x="4.5" y="14.2" width="4.2" height="4.2" rx="1"></rect>
      <path d="M11 7.1H19.5"></path>
      <path d="M11 16.3H19.5"></path>
    </svg>
  `,
  office: `
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">
      <path d="M7 20V5.2L15 3V20"></path>
      <path d="M7 8.5H3.8V20H20.2V8.5H15"></path>
      <path d="M10.3 8.7H11.7"></path>
      <path d="M10.3 12.1H11.7"></path>
      <path d="M10.3 15.5H11.7"></path>
    </svg>
  `,
  document: `
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">
      <path d="M7 3.8H13L18.2 9V20.2H7Z"></path>
      <path d="M12.7 3.8V9H18.2"></path>
    </svg>
  `,
  form: `
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">
      <path d="M4 20H20"></path>
      <path d="M14.5 5.2L18.8 9.5"></path>
      <path d="M7.4 16.4L6.6 19.2L9.4 18.4L18.1 9.7C18.88 8.92 18.88 7.66 18.1 6.88L17.12 5.9C16.34 5.12 15.08 5.12 14.3 5.9L7.4 12.8Z"></path>
    </svg>
  `,
  home: `
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">
      <path d="M4.4 11.3L12 5L19.6 11.3"></path>
      <path d="M6.5 10.1V18.7H17.5V13.4"></path>
      <path d="M15.6 18.7V14.8H8.4V18.7"></path>
    </svg>
  `,
  gridDots: `
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">
      <circle cx="6.2" cy="6.2" r="1.4"></circle>
      <circle cx="12" cy="6.2" r="1.4"></circle>
      <circle cx="17.8" cy="6.2" r="1.4"></circle>
      <circle cx="6.2" cy="12" r="1.4"></circle>
      <circle cx="12" cy="12" r="1.4"></circle>
      <circle cx="17.8" cy="12" r="1.4"></circle>
      <circle cx="6.2" cy="17.8" r="1.4"></circle>
      <circle cx="12" cy="17.8" r="1.4"></circle>
      <circle cx="17.8" cy="17.8" r="1.4"></circle>
    </svg>
  `,
  file: `
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">
      <path d="M7 3.8H13L18.2 9V20.2H7Z"></path>
      <path d="M12.7 3.8V9H18.2"></path>
    </svg>
  `,
  taskCheck: `
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">
      <rect x="5" y="4.2" width="14" height="16" rx="2.3"></rect>
      <path d="M8.5 7.8H15.5"></path>
      <path d="M9.1 13.1L11 15L15.2 10.8"></path>
    </svg>
  `,
};

const searchInput = document.querySelector("#command-search");
const clearButton = document.querySelector(".search-clear");
const emptyState = document.querySelector("#empty-state");
const resultsState = document.querySelector("#results-state");
const commandMenu = document.querySelector(".command-menu");

const MENU_TRANSITION_MS = 260;

let floatingMenuButton = null;
let menuAnimationTimer = 0;
let buttonAnimationTimer = 0;

if (menuState.minimized) {
  commandMenu.classList.add("menu-is-hidden");
}
commandMenu.setAttribute("aria-hidden", String(menuState.minimized));
renderEmptyState();
syncView("");

if (menuState.minimized) {
  showFloatingButton();
}

searchInput.addEventListener("input", (event) => {
  syncView(event.target.value);
});

clearButton.addEventListener("click", () => {
  searchInput.value = "";
  syncView("");
  searchInput.focus();
});

document.addEventListener("pointerdown", (event) => {
  if (menuState.minimized) {
    return;
  }

  const target = event.target;
  if (commandMenu.contains(target)) {
    return;
  }

  if (floatingMenuButton && floatingMenuButton.contains(target)) {
    return;
  }

  minimizeMenu();
});

window.addEventListener("resize", () => {
  if (!menuState.minimized || !floatingMenuButton) return;
  const currentTop = parseFloat(floatingMenuButton.style.top) || menuState.buttonTop;
  const clampedTop = Math.max(0, Math.min(currentTop, window.innerHeight - 48));
  floatingMenuButton.style.top = `${clampedTop}px`;
  if (menuState.buttonEdge === "right") {
    floatingMenuButton.style.right = "16px";
    floatingMenuButton.style.left = "";
  } else {
    floatingMenuButton.style.left = "16px";
    floatingMenuButton.style.right = "";
  }
  menuState.buttonTop = clampedTop;
  saveState();
});

let dragState = { isDragging: false, startY: 0, startX: 0, startTop: 0, startLeft: 0, moved: false };

function clearAnimationTimer(timerId) {
  if (timerId) {
    window.clearTimeout(timerId);
  }
}

function positionFloatingButton(button) {
  button.style.top = `${menuState.buttonTop}px`;

  if (menuState.buttonEdge === "right") {
    button.style.right = "16px";
    button.style.left = "";
  } else {
    button.style.left = "16px";
    button.style.right = "";
  }
}

function setMenuMotionFromRect(sourceRect) {
  const menuRect = commandMenu.getBoundingClientRect();
  const deltaX = sourceRect.left + sourceRect.width / 2 - (menuRect.left + menuRect.width / 2);
  const deltaY = sourceRect.top + sourceRect.height / 2 - (menuRect.top + menuRect.height / 2);
  commandMenu.style.setProperty("--menu-origin-x", `${deltaX}px`);
  commandMenu.style.setProperty("--menu-origin-y", `${deltaY}px`);
}

function setButtonMotionFromRect(targetRect, button) {
  const buttonRect = button.getBoundingClientRect();
  const deltaX = targetRect.left + targetRect.width / 2 - (buttonRect.left + buttonRect.width / 2);
  const deltaY = targetRect.top + targetRect.height / 2 - (buttonRect.top + buttonRect.height / 2);
  button.style.setProperty("--button-travel-x", `${deltaX}px`);
  button.style.setProperty("--button-travel-y", `${deltaY}px`);
}

function createFloatingButton() {
  const button = document.createElement("button");
  button.className = "floating-menu-btn";
  button.type = "button";
  button.setAttribute("aria-label", "Open menu");
  button.tabIndex = 0;
  button.innerHTML = iconMarkup.gridDots;
  button.classList.add("button-is-hidden");
  positionFloatingButton(button);

  button.addEventListener("click", restoreMenu);
  button.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      restoreMenu();
    }
  });
  document.body.appendChild(button);
  setupDrag(button);
  floatingMenuButton = button;
  return button;
}

function setupDrag(button) {
  button.addEventListener("pointerdown", (event) => {
    dragState.startY = event.clientY;
    dragState.startX = event.clientX;
    dragState.startTop = parseFloat(button.style.top) || menuState.buttonTop;
    dragState.startLeft = button.getBoundingClientRect().left;
    dragState.moved = false;
    dragState.isDragging = true;
    button.setPointerCapture(event.pointerId);
    document.body.style.userSelect = "none";
  });

  button.addEventListener("pointermove", (event) => {
    if (!dragState.isDragging) {
      return;
    }

    event.preventDefault();
    const deltaY = event.clientY - dragState.startY;
    const deltaX = event.clientX - dragState.startX;
    const newTop = dragState.startTop + deltaY;
    const newLeft = dragState.startLeft + deltaX;
    const clampedTop = Math.max(0, Math.min(newTop, window.innerHeight - 48));
    const clampedLeft = Math.max(0, Math.min(newLeft, window.innerWidth - 48));
    button.style.top = `${clampedTop}px`;
    button.style.left = `${clampedLeft}px`;
    button.style.right = "";

    if (Math.abs(deltaY) > 5 || Math.abs(deltaX) > 5) {
      dragState.moved = true;
    }
  });

  button.addEventListener("pointerup", (event) => {
    const wasDragging = dragState.isDragging;

    dragState.isDragging = false;
    document.body.style.userSelect = "";
    button.releasePointerCapture(event.pointerId);

    if (!wasDragging || !dragState.moved) {
      return;
    }

    const rect = button.getBoundingClientRect();
    if (rect.left + rect.width / 2 < window.innerWidth / 2) {
      button.style.left = "16px";
      button.style.right = "";
      menuState.buttonEdge = "left";
    } else {
      button.style.right = "16px";
      button.style.left = "";
      menuState.buttonEdge = "right";
    }

    menuState.buttonTop = parseFloat(button.style.top);
    saveState();
  });

  button.addEventListener("pointercancel", () => {
    dragState.isDragging = false;
    document.body.style.userSelect = "";
  });

  button.addEventListener(
    "click",
    (event) => {
      if (dragState.moved) {
        event.preventDefault();
        event.stopImmediatePropagation();
      }
    },
    true,
  );
}

function showFloatingButton(sourceRect = null) {
  const button = floatingMenuButton || createFloatingButton();

  clearAnimationTimer(buttonAnimationTimer);
  positionFloatingButton(button);
  button.classList.remove("button-is-hidden", "is-animating-out");

  if (!sourceRect) {
    return;
  }

  setButtonMotionFromRect(sourceRect, button);
  void button.offsetWidth;
  button.classList.add("is-animating-in");
  buttonAnimationTimer = window.setTimeout(() => {
    button.classList.remove("is-animating-in");
  }, MENU_TRANSITION_MS);
}

function hideFloatingButton(targetRect = null) {
  if (!floatingMenuButton) {
    return;
  }

  const button = floatingMenuButton;
  clearAnimationTimer(buttonAnimationTimer);
  button.classList.remove("is-animating-in");

  if (!targetRect) {
    button.classList.add("button-is-hidden");
    return;
  }

  setButtonMotionFromRect(targetRect, button);
  void button.offsetWidth;
  button.classList.add("is-animating-out");
  buttonAnimationTimer = window.setTimeout(() => {
    button.classList.remove("is-animating-out");
    button.classList.add("button-is-hidden");
  }, MENU_TRANSITION_MS);
}

function openMenuFromButton() {
  const sourceRect = floatingMenuButton?.getBoundingClientRect();
  clearAnimationTimer(menuAnimationTimer);
  commandMenu.classList.remove("menu-is-hidden", "is-closing");
  commandMenu.setAttribute("aria-hidden", "false");

  if (!sourceRect) {
    return;
  }

  setMenuMotionFromRect(sourceRect);
  void commandMenu.offsetWidth;
  commandMenu.classList.add("is-opening");
  menuAnimationTimer = window.setTimeout(() => {
    commandMenu.classList.remove("is-opening");
  }, MENU_TRANSITION_MS);
}

function closeMenuToButton() {
  const menuRect = commandMenu.getBoundingClientRect();
  showFloatingButton(menuRect);

  const targetRect = floatingMenuButton.getBoundingClientRect();
  clearAnimationTimer(menuAnimationTimer);
  commandMenu.classList.remove("is-opening");
  setMenuMotionFromRect(targetRect);
  void commandMenu.offsetWidth;
  commandMenu.classList.add("is-closing");
  menuAnimationTimer = window.setTimeout(() => {
    commandMenu.classList.remove("is-closing");
    commandMenu.classList.add("menu-is-hidden");
    commandMenu.setAttribute("aria-hidden", "true");
  }, MENU_TRANSITION_MS);
}

function restoreMenu() {
  if (!menuState.minimized) {
    return;
  }

  menuState.minimized = false;
  openMenuFromButton();
  hideFloatingButton(commandMenu.getBoundingClientRect());
  saveState();
  window.setTimeout(() => {
    searchInput.focus();
  }, MENU_TRANSITION_MS / 2);
}

function minimizeMenu() {
  if (menuState.minimized) {
    return;
  }

  menuState.minimized = true;
  closeMenuToButton();
  saveState();
  window.setTimeout(() => {
    floatingMenuButton?.focus();
  }, MENU_TRANSITION_MS / 2);
}

function syncView(rawValue) {
  const query = rawValue.trim();
  const hasQuery = query.length > 0;
  const filteredResults = filterResults(query);
  const showResults = hasQuery && filteredResults.length > 0;

  renderEmptyState(query, showResults);

  if (showResults) {
    renderResultsState(filteredResults);
  }

  emptyState.classList.toggle("is-hidden", showResults);
  resultsState.classList.toggle("is-hidden", !showResults);
  clearButton.classList.toggle("is-hidden", !hasQuery);
}

function renderEmptyState(query = "", showResults = false) {
  const isNoResults = query.length > 0 && !showResults;
  const title = isNoResults ? "No results found" : "Get started with these Quick Actions";
  const subtitle = isNoResults
    ? `Try a different search or use one of these Quick Actions instead`
    : "Or start searching to get results";
  const eyebrow = isNoResults ? `<p class="empty-state-eyebrow">Search: ${escapeHtml(query)}</p>` : "";

  emptyState.innerHTML = `
    <div class="empty-state">
      <div class="empty-state-icon" aria-hidden="true">${iconMarkup.lightbulb}</div>
      ${eyebrow}
      <h1 class="empty-state-title">${title}</h1>
      <p class="empty-state-subtitle">${subtitle}</p>
      <div class="quick-action-grid">
        ${quickActions.map(renderQuickAction).join("")}
      </div>
    </div>
  `;
}

function renderResultsState(filteredResults) {
  resultsState.innerHTML = `
    <div class="results-panel">
      <div class="results-tabs" role="tablist" aria-label="Search categories">
        ${resultTabs.map((tab, index) => renderTab(tab, index === 0)).join("")}
      </div>
      <h2 class="results-group-title">Top Results</h2>
      <div class="results-list">
        ${filteredResults.map(renderResult).join("")}
      </div>
    </div>
  `;
}

function renderQuickAction(action) {
  return `
    <button class="quick-action" type="button">
      <span class="quick-action-icon" aria-hidden="true">${iconMarkup[action.icon]}</span>
      <span class="quick-action-label">${action.label}</span>
    </button>
  `;
}

function renderTab(tab, isActive) {
  return `
    <button class="tab-chip ${isActive ? "is-active" : ""}" type="button" role="tab" aria-selected="${String(isActive)}">
      <span aria-hidden="true">${iconMarkup[tab.icon]}</span>
      <span class="tab-chip-label">${tab.label}</span>
    </button>
  `;
}

function renderResult(result) {
  return `
    <button class="result-item" type="button">
      <span class="result-icon" aria-hidden="true">${iconMarkup[result.icon]}</span>
      <span class="result-label">${result.label}</span>
      <span class="result-tag tag-${result.tagTone}">${result.tag}</span>
    </button>
  `;
}

function filterResults(query) {
  if (!query) {
    return [];
  }

  const normalizedQuery = query.toLowerCase();

  return results.filter((result) => {
    const haystack = [result.label, result.tag, ...(result.keywords || [])].join(" ").toLowerCase();
    return haystack.includes(normalizedQuery);
  });
}

function escapeHtml(value) {
  return value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}
