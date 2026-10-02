/* Exhibition interaction driver. Student pages keep their original manual controls. */
(function () {
  "use strict";
  const states = new Map();
  let active = false;
  let timer = null;
  let statusElement = null;
  const random = (min, max) => min + Math.random() * (max - min);

  function visible(el) {
    if (!el || !el.getClientRects().length) return false;
    for (let node = el; node && node.nodeType === 1; node = node.parentElement) {
      const css = node.ownerDocument.defaultView.getComputedStyle(node);
      if (css.display === "none" || css.visibility === "hidden" || Number(css.opacity) === 0) return false;
    }
    return true;
  }

  function emit(s, type, target, x, y, buttons) {
    const options = { bubbles: true, cancelable: true, view: s.win,
      clientX: x, clientY: y, button: 0, buttons: buttons || 0 };
    const event = type.startsWith("pointer")
      ? new s.win.PointerEvent(type, { ...options, pointerId: 941, pointerType: "mouse", isPrimary: true })
      : new s.win.MouseEvent(type, options);
    target.dispatchEvent(event);
  }

  function move(s, x, y) {
    s.x = x; s.y = y;
    emit(s, "pointermove", s.target, x, y, s.down ? 1 : 0);
    emit(s, "mousemove", s.target, x, y, s.down ? 1 : 0);
  }

  function release(s, click) {
    if (!s.down) return;
    s.down = false;
    emit(s, "pointerup", s.target, s.x, s.y, 0);
    emit(s, "mouseup", s.target, s.x, s.y, 0);
    if (click) emit(s, "click", s.target, s.x, s.y, 0);
  }

  function stopState(s) {
    if (!s.win) return;
    try {
      release(s, false);
      if (s.demoStarted) s.win.galleryDemo.stop();
    } catch (_) { /* An iframe may be unloading. */ }
    s.demoStarted = false;
    s.running = false;
  }

  // Explicit entry controls only: never click arbitrary buttons, links, or forms.
  const entryControls = {
    "徐妏寧": ["#stage-1 .main-btn", "#burn-message .fire-btn"],
    "蔡佳翰": ["#startBtn"],
    "連浡崴": ["#overlay"],
    "陳子怡": ["#enter-btn"],
    "陳翰博": ["#start-screen"]
  };

  function uiStep(s, now) {
    if (now < s.nextUI) return false;
    s.nextUI = now + 3000;
    for (const selector of entryControls[s.name] || []) {
      const control = s.doc.querySelector(selector);
      if (visible(control)) { release(s, false); control.click(); return true; }
    }
    if (s.name === "連浡崴" && now >= s.nextMode) {
      const modes = ["pure", "reverie", "deep"];
      const button = s.doc.querySelector('[data-type="' + modes[s.mode++ % modes.length] + '"]');
      if (visible(button)) { button.click(); s.nextMode = now + 15000; }
    }
    if (s.name === "黃鈺倢" && visible(s.doc.getElementById("colorOk"))) {
      release(s, false);
      const input = s.doc.getElementById("colorInput");
      input.value = ["#ff7ad9", "#7dcfff", "#9ece6a", "#e6c229"][s.mode++ % 4];
      input.dispatchEvent(new s.win.Event("input", { bubbles: true }));
      s.doc.getElementById("colorOk").click();
      return true;
    }
    return false;
  }

  function startState(s, now) {
    s.win = s.frame.contentWindow;
    s.doc = s.win.document;
    if (!s.doc.body || s.doc.readyState === "loading") return false;
    if (s.doc.URL === "about:blank") return false;
    s.target = s.doc.querySelector(s.name === "楊漢軒" ? "#stage" : "canvas") || s.doc.body;
    s.running = true;
    s.nextAction = now + random(100, 1800);
    s.nextUI = 0;
    s.nextMode = now + 8000;
    s.nextWheel = now;
    s.mode = 0;
    s.count = 0;
    if (s.win.galleryDemo) {
      s.win.galleryDemo.start();
      s.demoStarted = true;
    }
    // This artwork's "add memory" button opens a prompt. Seed a visual directly instead.
    if (s.name === "翁圓舒" && !s.memorySeeded && typeof s.win.createBlob === "function") {
      s.win.createBlob(s.win.innerWidth / 2, s.win.innerHeight / 2);
      s.win.createBubbles(s.win.innerWidth / 2, s.win.innerHeight / 2);
      s.memorySeeded = true;
    }
    return true;
  }

  function step(s, now) {
    if (!s.loaded || s.frame.closest(".hidden")) { stopState(s); return; }
    if (!s.running && !startState(s, now)) return;
    if ((s.name === "楊靜慧" || s.name === "黃政文") && now >= s.nextWheel) {
      s.target.dispatchEvent(new s.win.WheelEvent("wheel", { bubbles: true, cancelable: true,
        deltaY: s.name === "楊靜慧" ? 80 : (s.count % 10 < 7 ? 140 : -140) }));
      s.nextWheel = now + (s.name === "楊靜慧" ? 100 : 1200);
    }
    if (uiStep(s, now)) { s.nextAction = now + 1600; return; }
    if (s.demoStarted) {
      if (now >= s.nextAction) { s.win.galleryDemo.tick(); s.nextAction = now + 2800; }
      return;
    }
    if (s.down) {
      const t = Math.min(1, (now - s.pressTime) / s.duration);
      move(s, s.fromX + (s.toX - s.fromX) * t, s.fromY + (s.toY - s.fromY) * t);
      if (t >= 1) { release(s, true); s.nextAction = now + random(900, 2200); }
      return;
    }
    if (now < s.nextAction) return;
    // Bubble popping is part of this artwork; alternate with its glass clicks.
    const bubble = s.name === "王欣怡" && s.count % 2 === 0 ? s.doc.querySelector(".bubble") : null;
    s.target = visible(bubble) ? bubble : (s.doc.querySelector(s.name === "楊漢軒" ? "#stage" : "canvas") || s.doc.body);
    const rect = s.target.getBoundingClientRect();
    const width = Math.min(s.win.innerWidth, rect.width || s.win.innerWidth);
    const height = Math.min(s.win.innerHeight, rect.height || s.win.innerHeight);
    s.fromX = Math.max(0, rect.left) + width * random(0.15, 0.85);
    s.fromY = Math.max(0, rect.top) + height * random(0.2, 0.85);
    s.toX = Math.max(0, rect.left) + width * random(0.15, 0.85);
    s.toY = Math.max(0, rect.top) + height * random(0.2, 0.85);
    s.duration = s.count++ % 3 === 0 ? 160 : random(600, 1400);
    if (s.name === "陳陽安") { s.fromY = s.toY = s.win.innerHeight * 0.92; s.duration = 6200; }
    if (s.name === "黃鈺倢") s.duration = 4500;
    if (s.name === "楊漢軒" && s.count % 3 === 0) {
      s.duration = 1800; s.toX = s.fromX; s.toY = s.fromY;
    }
    move(s, s.fromX, s.fromY);
    s.down = true; s.pressTime = now;
    emit(s, "pointerdown", s.target, s.x, s.y, 1);
    emit(s, "mousedown", s.target, s.x, s.y, 1);
  }

  function updateStatus() {
    if (!statusElement) return;
    if (!active) { statusElement.textContent = ""; return; }
    const current = [...states.values()].filter(s => !s.frame.closest(".hidden"));
    const ready = current.filter(s => s.running).length;
    const failed = current.filter(s => s.failed).length;
    statusElement.textContent = "自動播放 " + ready + " / " + current.length + " 件" +
      (failed ? "（" + failed + " 件無法操作）" : "");
  }

  function loop() {
    timer = null;
    if (!active || document.hidden) return;
    const now = performance.now();
    for (const s of states.values()) {
      if (s.failed) continue;
      try { step(s, now); }
      catch (_) { stopState(s); s.failed = true; }
    }
    updateStatus();
    timer = setTimeout(loop, 100);
  }

  window.GalleryAutoplay = {
    attach(frame, entry) {
      const s = { frame, name: entry.studentName, loaded: false };
      states.set(frame, s);
      frame.addEventListener("load", () => {
        stopState(s);
        s.loaded = true; s.failed = false; s.memorySeeded = false;
      });
    },
    clear() { for (const s of states.values()) stopState(s); states.clear(); },
    setActive(value) {
      active = Boolean(value);
      clearTimeout(timer); timer = null;
      if (!active) for (const s of states.values()) stopState(s);
      else for (const s of states.values()) s.failed = false;
      loop(); updateStatus();
    },
    setStatusElement(el) { statusElement = el; }
  };

  document.addEventListener("visibilitychange", () => {
    clearTimeout(timer); timer = null;
    if (document.hidden) for (const s of states.values()) stopState(s);
    else loop();
    updateStatus();
  });
})();
