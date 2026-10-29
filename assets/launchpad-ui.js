/* =====================================================================
   Finance Launchpad — shared UI script for GitHub-hosted modules.
   Load it in <head> (no defer) so the remembered theme applies before paint.
   - Theme: <button id="lp-theme-toggle" class="lp-theme"> toggles dark mode
     (html[data-theme="dark"] + body.theme-dark), remembered in localStorage.
   - Dew drops: any element with data-lp-fx gets the canvas animation.
   - Back button: onclick="showBackMessage(event, this)" shows the
     "Checking your access..." overlay, then returns to the Launchpad.
   ===================================================================== */
(function () {
  var KEY = 'fpa_color_scheme';
  try { if (localStorage.getItem(KEY) === 'dark') document.documentElement.setAttribute('data-theme', 'dark'); } catch (e) {}

  function isDark() { return document.documentElement.getAttribute('data-theme') === 'dark'; }
  function setTheme(dark) {
    if (dark) document.documentElement.setAttribute('data-theme', 'dark'); else document.documentElement.removeAttribute('data-theme');
    if (document.body) document.body.classList.toggle('theme-dark', dark);
    try { localStorage.setItem(KEY, dark ? 'dark' : 'light'); } catch (e) {}
    var b = document.getElementById('lp-theme-toggle');
    if (b) {
      b.innerHTML = '<i class="ti ' + (dark ? 'ti-sun' : 'ti-moon') + '" aria-hidden="true"></i>';
      b.title = dark ? 'Switch to light mode' : 'Switch to dark mode';
      b.setAttribute('aria-pressed', dark ? 'true' : 'false');
    }
  }

  window.showBackMessage = function (event, link) {
    if (event) event.preventDefault();
    var overlay = document.createElement('div');
    overlay.className = 'navigation-loading';
    overlay.innerHTML =
      '<div class="navigation-loading-card">' +
        '<i class="ti ti-user-search navigation-loading-icon"></i>' +
        '<strong>Checking your access...</strong>' +
        '<span>Authenticating with Google Workspace and loading your authorized modules.</span>' +
      '</div>';
    document.body.appendChild(overlay);
    setTimeout(function () { window.top.location.replace(link.href); }, 250);
  };

  function init() {
    var toggle = document.getElementById('lp-theme-toggle');
    if (toggle) toggle.addEventListener('click', function () { setTheme(!isDark()); });
    setTheme(isDark());
    Array.prototype.forEach.call(document.querySelectorAll('[data-lp-fx]'), function (el) { startHeroFx(el); });
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();
})();

/* =====================================================================
   Hero ambient animation — "dew drops" on the perspective grid.
   Small beads of light slide along the floor lines, pause at each knot
   (line crossing) with a soft shimmer, and speed up as they come closer.
   Same geometry as the floor drawn in LaunchpadCss (viewBox 1600×320,
   bottom 78% of the hero). Pure canvas, no assets. Respects
   prefers-reduced-motion and pauses when hidden or off-screen.
   ===================================================================== */
function startHeroFx(hero) {
  if (!hero || !window.requestAnimationFrame || hero.querySelector('.lp-hero-fx')) return;

  var reduce = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)');
  if (reduce && reduce.matches) return;

  var VB_W = 1600, VB_H = 320;
  var KNOTS = 15;             // horizontal lines at y = 320 * 0.8^k, k = 1..15
  var MAX_DROPS = 4;

  var canvas = document.createElement('canvas');
  canvas.className = 'lp-hero-fx';
  canvas.setAttribute('aria-hidden', 'true');
  hero.insertBefore(canvas, hero.firstChild);

  var ctx = canvas.getContext('2d');
  var W = 0, H = 0, dpr = 1;
  var drops = [], rings = [];
  var running = false, visible = true, last = 0, spawnIn = 1.2;
  var color = '159,225,203';

  function readColor() {
    var raw = getComputedStyle(document.body).getPropertyValue('--mint').trim();
    var m = raw.match(/^#?([0-9a-f]{6})$/i);
    if (m) {
      var n = parseInt(m[1], 16);
      color = ((n >> 16) & 255) + ',' + ((n >> 8) & 255) + ',' + (n & 255);
    }
  }

  function resize() {
    var r = canvas.getBoundingClientRect();
    dpr = Math.min(window.devicePixelRatio || 1, 2);
    W = r.width; H = r.height;
    canvas.width = Math.round(W * dpr);
    canvas.height = Math.round(H * dpr);
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  }

  /* depth d: 0 = front edge (y=320), 15 = horizon. y = 320 * 0.8^d */
  function yOf(d) { return VB_H * Math.pow(0.8, d); }
  function xOnRadial(i, sy) {
    var x0 = 800 + (i - 60) * 23.04, x1 = 800 + (i - 60) * 64;
    return x0 + (x1 - x0) * (sy / VB_H);
  }
  /* same fade as the floor's CSS mask (bottom solid → top transparent) */
  function maskA(sy) {
    var t = 1 - sy / VB_H;               // 0 at bottom, 1 at top
    if (t <= 0.4) return 1 - t;          // 1 → .6
    if (t >= 0.95) return 0;
    return 0.6 * (1 - (t - 0.4) / 0.55);
  }
  function toPx(sx, sy) { return [sx / VB_W * W, sy / VB_H * H]; }

  function spawn() {
    var radial = Math.random() < 0.85;
    if (radial) {
      // pick a line that is on screen near the front
      var i = 60 + Math.round((Math.random() * 2 - 1) * 26);
      drops.push({ kind: 'r', line: i, d: 5 + Math.random() * 3, seg: 0, speed: 0.75 + Math.random() * 0.35, trail: [] });
    } else {
      var k = 1 + Math.floor(Math.random() * 4);  // front horizontal lines
      var dir = Math.random() < 0.5 ? -1 : 1;
      drops.push({ kind: 'h', k: k, x: dir > 0 ? -20 : VB_W + 20, dir: dir, speed: 90 + Math.random() * 60, trail: [] });
    }
  }

  function ring(sx, sy, size) {
    rings.push({ sx: sx, sy: sy, t: 0, size: size });
  }

  function step(dt) {
    spawnIn -= dt;
    if (spawnIn <= 0 && drops.length < MAX_DROPS) {
      spawn();
      spawnIn = 2.2 + Math.random() * 2.8;
    }

    drops = drops.filter(function(p) {
      var sx, sy;

      if (p.kind === 'r') {
        // slide knot to knot: ease within each segment, short rest at the knot
        p.seg += dt * p.speed;
        if (p.seg >= 1.25) {
          p.seg = 0;
          p.d -= 1;
          var ky = yOf(Math.round(p.d));
          ring(xOnRadial(p.line, ky), ky, 1 + (ky / VB_H) * 2);
        }
        var u = Math.min(p.seg, 1);
        var e = u * u * (3 - 2 * u);
        var d = p.d - e;
        sy = yOf(d);
        sx = xOnRadial(p.line, sy);
        p.glow = p.seg > 1 ? 1 : 0.6;   // brighter while resting on a knot
        if (d < -0.4) return false;
      } else {
        p.x += p.dir * p.speed * dt * (1 + p.k * 0.15);
        sy = yOf(p.k);
        sx = p.x;
        p.glow = 0.6;
        if (p.x < -40 || p.x > VB_W + 40) return false;
      }

      p.sx = sx; p.sy = sy;
      p.trail.push([sx, sy]);
      if (p.trail.length > 10) p.trail.shift();
      return true;
    });

    rings = rings.filter(function(r) { r.t += dt; return r.t < 1.1; });
  }

  function draw() {
    ctx.clearRect(0, 0, W, H);
    ctx.globalCompositeOperation = 'lighter';

    rings.forEach(function(r) {
      var a = maskA(r.sy) * (1 - r.t / 1.1) * 0.32;
      if (a <= 0.01) return;
      var p = toPx(r.sx, r.sy);
      ctx.beginPath();
      ctx.ellipse(p[0], p[1], (3 + r.t * 16) * r.size, (1.2 + r.t * 5) * r.size, 0, 0, Math.PI * 2);
      ctx.strokeStyle = 'rgba(' + color + ',' + a.toFixed(3) + ')';
      ctx.lineWidth = 1;
      ctx.stroke();
    });

    drops.forEach(function(p) {
      var a = maskA(p.sy);
      if (a <= 0.01) return;
      var size = 1 + (p.sy / VB_H) * 1.9;

      // trail along the strand
      if (p.trail.length > 1) {
        ctx.beginPath();
        p.trail.forEach(function(pt, n) {
          var q = toPx(pt[0], pt[1]);
          if (n) ctx.lineTo(q[0], q[1]); else ctx.moveTo(q[0], q[1]);
        });
        var head = toPx(p.sx, p.sy), tail = toPx(p.trail[0][0], p.trail[0][1]);
        var g = ctx.createLinearGradient(tail[0], tail[1], head[0], head[1]);
        g.addColorStop(0, 'rgba(' + color + ',0)');
        g.addColorStop(1, 'rgba(' + color + ',' + (0.32 * a).toFixed(3) + ')');
        ctx.strokeStyle = g;
        ctx.lineWidth = size * 0.9;
        ctx.lineCap = 'round';
        ctx.stroke();
      }

      // the drop: bright core + halo
      var c = toPx(p.sx, p.sy);
      var r = size * (p.glow > 0.9 ? 5 : 3.6);
      var halo = ctx.createRadialGradient(c[0], c[1], 0, c[0], c[1], r);
      halo.addColorStop(0, 'rgba(255,255,255,' + (0.8 * a).toFixed(3) + ')');
      halo.addColorStop(0.25, 'rgba(' + color + ',' + (0.5 * a).toFixed(3) + ')');
      halo.addColorStop(1, 'rgba(' + color + ',0)');
      ctx.fillStyle = halo;
      ctx.beginPath();
      ctx.arc(c[0], c[1], r, 0, Math.PI * 2);
      ctx.fill();
    });

    ctx.globalCompositeOperation = 'source-over';
  }

  function frame(now) {
    if (!running) return;
    var dt = Math.min(0.05, (now - (last || now)) / 1000);
    last = now;
    step(dt);
    draw();
    requestAnimationFrame(frame);
  }

  function start() {
    if (running || document.hidden || !visible) return;
    running = true; last = 0;
    requestAnimationFrame(frame);
  }
  function stop() { running = false; }

  readColor();
  resize();

  if (window.ResizeObserver) new ResizeObserver(resize).observe(canvas);
  else window.addEventListener('resize', resize);

  if (window.IntersectionObserver) {
    new IntersectionObserver(function(entries) {
      visible = entries[0].isIntersecting;
      if (visible) start(); else stop();
    }).observe(hero);
  }

  document.addEventListener('visibilitychange', function() { if (document.hidden) stop(); else start(); });

  // accent follows workspace / theme changes
  new MutationObserver(readColor).observe(document.body, { attributes: true, attributeFilter: ['data-workspace', 'class'] });
  new MutationObserver(readColor).observe(document.documentElement, { attributes: true, attributeFilter: ['style'] });

  if (reduce && reduce.addEventListener) {
    reduce.addEventListener('change', function(e) { if (e.matches) { stop(); ctx.clearRect(0, 0, W, H); } else start(); });
  }

  start();
}
