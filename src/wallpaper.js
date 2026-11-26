/**
 * Samsung Galaxy S9 / S9+ "Infinity" Inspired Live Wallpaper Engine
 * Smooth 60fps Canvas silk-wave crystalline ribbons + iridescent specular ridges + drifting bokeh stardust.
 * Dynamically adapts its chromatic harmony to the user's chosen Accent Color.
 */

export class GalaxyS9Wallpaper {
  constructor(canvas) {
    this.canvas = canvas;
    this.ctx = canvas.getContext('2d', { alpha: false });
    this.width = 0;
    this.height = 0;
    this.dpr = Math.min(window.devicePixelRatio || 1, 2);

    this.time = 0;
    this.speed = 1.0;
    this.paused = false;
    this.preset = 's9-silk'; // 's9-silk' | 's9-midnight' | 's9-pearl'
    this.accentHsl = { h: 171, s: 66, l: 42 }; // Default IMG_9401 emerald teal

    // Smooth parallax mouse target
    this.pointer = { x: 0, y: 0, tx: 0, ty: 0 };

    // Floating crystalline bokeh particles
    this.particles = [];
    this.rafId = null;
    this.lastTs = 0;

    this._onResize = this.resize.bind(this);
    this._onPointerMove = this.handlePointer.bind(this);
    this._loop = this.loop.bind(this);
  }

  start() {
    this.resize();
    this.initParticles();
    window.addEventListener('resize', this._onResize, { passive: true });
    window.addEventListener('pointermove', this._onPointerMove, { passive: true });
    this.lastTs = performance.now();
    this.rafId = requestAnimationFrame(this._loop);
  }

  destroy() {
    if (this.rafId) cancelAnimationFrame(this.rafId);
    window.removeEventListener('resize', this._onResize);
    window.removeEventListener('pointermove', this._onPointerMove);
  }

  setAccentHex(hex) {
    this.accentHsl = hexToHsl(hex);
  }

  setPreset(preset) {
    this.preset = preset || 's9-silk';
  }

  setSpeed(speed) {
    this.speed = Number(speed);
  }

  setPaused(paused) {
    this.paused = Boolean(paused);
  }

  handlePointer(e) {
    const nx = (e.clientX / Math.max(window.innerWidth, 1) - 0.5) * 2;
    const ny = (e.clientY / Math.max(window.innerHeight, 1) - 0.5) * 2;
    this.pointer.tx = nx;
    this.pointer.ty = ny;
  }

  resize() {
    this.width = window.innerWidth;
    this.height = window.innerHeight;
    this.dpr = Math.min(window.devicePixelRatio || 1, 2);
    this.canvas.width = Math.floor(this.width * this.dpr);
    this.canvas.height = Math.floor(this.height * this.dpr);
    this.ctx.setTransform(this.dpr, 0, 0, this.dpr, 0, 0);
  }

  initParticles() {
    const count = 48;
    this.particles = [];
    for (let i = 0; i < count; i++) {
      this.particles.push({
        x: Math.random(),
        y: Math.random(),
        r: 0.9 + Math.random() * 2.4,
        vx: (Math.random() - 0.5) * 0.00018,
        vy: -0.00008 - Math.random() * 0.00022,
        phase: Math.random() * Math.PI * 2,
        depth: 0.4 + Math.random() * 1.1,
      });
    }
  }

  loop(ts) {
    const dt = Math.min((ts - this.lastTs) / 1000, 0.05);
    this.lastTs = ts;

    if (!this.paused) {
      this.time += dt * this.speed;
    }

    // Smooth pointer interpolation
    this.pointer.x += (this.pointer.tx - this.pointer.x) * 0.045;
    this.pointer.y += (this.pointer.ty - this.pointer.y) * 0.045;

    this.draw(dt);
    this.rafId = requestAnimationFrame(this._loop);
  }

  draw(dt) {
    const { ctx, width: w, height: h, time: t } = this;
    const { h: hue, s: sat } = this.accentHsl;

    const isPearl = this.preset === 's9-pearl';
    const isMidnight = this.preset === 's9-midnight';

    // 1. Deep Luminous Base Backdrop (S9 Infinity signature gradient)
    const bgGrad = ctx.createLinearGradient(0, 0, w, h);
    if (isPearl) {
      bgGrad.addColorStop(0, `hsl(${(hue + 195) % 360}, 28%, 87%)`);
      bgGrad.addColorStop(0.48, `hsl(${hue}, 24%, 92%)`);
      bgGrad.addColorStop(1, `hsl(${(hue + 28) % 360}, 30%, 84%)`);
    } else if (isMidnight) {
      bgGrad.addColorStop(0, `hsl(${(hue + 215) % 360}, 44%, 7%)`);
      bgGrad.addColorStop(0.5, `hsl(${(hue + 195) % 360}, 48%, 11%)`);
      bgGrad.addColorStop(1, `hsl(${hue}, 42%, 9%)`);
    } else {
      // Iconic Samsung Galaxy S9 Silk (Warm-slate / Obsidian crystalline with luminous ambient glow)
      bgGrad.addColorStop(0, `hsl(${(hue + 210) % 360}, 32%, 14%)`);
      bgGrad.addColorStop(0.45, `hsl(${(hue + 22) % 360}, 20%, 24%)`);
      bgGrad.addColorStop(1, `hsl(${(hue + 195) % 360}, 36%, 12%)`);
    }
    ctx.fillStyle = bgGrad;
    ctx.fillRect(0, 0, w, h);

    const px = this.pointer.x * 28;
    const py = this.pointer.y * 22;

    // 2. Soft Ambient Prism Bloom Orbs behind the S9 ribbons
    this.drawBloom(
      w * 0.22 + Math.sin(t * 0.35) * 90 + px * 0.7,
      h * 0.26 + Math.cos(t * 0.28) * 60 + py * 0.7,
      Math.max(w, h) * 0.45,
      isPearl
        ? `hsla(${hue}, ${Math.min(sat + 10, 82)}%, 65%, 0.38)`
        : `hsla(${hue}, ${Math.min(sat + 12, 88)}%, 50%, 0.36)`
    );

    this.drawBloom(
      w * 0.8 + Math.cos(t * 0.31) * 100 - px * 0.9,
      h * 0.68 + Math.sin(t * 0.39) * 75 - py * 0.9,
      Math.max(w, h) * 0.48,
      isPearl
        ? `hsla(${(hue + 38) % 360}, 72%, 72%, 0.36)`
        : `hsla(${(hue + 36) % 360}, 78%, 48%, 0.3)`
    );

    this.drawBloom(
      w * 0.54 + Math.sin(t * 0.22 + 1.4) * 110 + px * 0.5,
      h * 0.14 + Math.cos(t * 0.26) * 45 + py * 0.5,
      Math.max(w, h) * 0.38,
      isPearl
        ? `hsla(${(hue + 325) % 360}, 60%, 78%, 0.32)`
        : `hsla(${(hue + 320) % 360}, 68%, 54%, 0.22)`
    );

    // 3. Signature Galaxy S9 Flowing Folded Silk Ribbons (5 layered parametric waves)
    const ribbonSpecs = [
      {
        yBase: 0.72,
        amp1: 115,
        amp2: 52,
        freq: 0.0021,
        speed: 0.52,
        phase: 0.0,
        hueShift: 0,
        alpha: isPearl ? 0.26 : 0.34,
        thickness: h * 0.34,
        ridgeAlpha: isPearl ? 0.55 : 0.62,
      },
      {
        yBase: 0.56,
        amp1: 135,
        amp2: 64,
        freq: 0.0018,
        speed: -0.44,
        phase: 1.7,
        hueShift: 28,
        alpha: isPearl ? 0.22 : 0.3,
        thickness: h * 0.31,
        ridgeAlpha: isPearl ? 0.48 : 0.54,
      },
      {
        yBase: 0.4,
        amp1: 105,
        amp2: 46,
        freq: 0.0024,
        speed: 0.39,
        phase: 3.1,
        hueShift: -26,
        alpha: isPearl ? 0.2 : 0.26,
        thickness: h * 0.28,
        ridgeAlpha: isPearl ? 0.45 : 0.48,
      },
      {
        yBase: 0.24,
        amp1: 88,
        amp2: 40,
        freq: 0.0027,
        speed: -0.33,
        phase: 4.6,
        hueShift: 52,
        alpha: isPearl ? 0.17 : 0.22,
        thickness: h * 0.24,
        ridgeAlpha: isPearl ? 0.38 : 0.42,
      },
      {
        yBase: 0.86,
        amp1: 95,
        amp2: 38,
        freq: 0.0022,
        speed: 0.46,
        phase: 5.4,
        hueShift: 14,
        alpha: isPearl ? 0.2 : 0.25,
        thickness: h * 0.26,
        ridgeAlpha: isPearl ? 0.42 : 0.5,
      },
    ];

    ctx.save();
    ctx.globalCompositeOperation = isPearl ? 'source-over' : 'screen';

    for (let r = 0; r < ribbonSpecs.length; r++) {
      const spec = ribbonSpecs[r];
      const rHue = (hue + spec.hueShift + 360) % 360;
      const steps = 42;
      const dx = w / steps;
      const parallaxFactor = (r + 1) * 0.25;

      const topPoints = [];
      for (let i = 0; i <= steps; i++) {
        const x = i * dx;
        const diagonalTilt = (i / steps - 0.5) * h * 0.26;
        const y =
          h * spec.yBase +
          diagonalTilt +
          Math.sin(x * spec.freq + t * spec.speed + spec.phase) * spec.amp1 +
          Math.cos(x * spec.freq * 1.85 - t * spec.speed * 0.72 + spec.phase) * spec.amp2 +
          py * parallaxFactor;
        topPoints.push({ x: x + px * parallaxFactor * 0.5, y });
      }

      // Draw translucent folded silk body
      ctx.beginPath();
      ctx.moveTo(topPoints[0].x - 40, topPoints[0].y);
      for (let i = 1; i < topPoints.length; i++) {
        const prev = topPoints[i - 1];
        const curr = topPoints[i];
        const cx = (prev.x + curr.x) * 0.5;
        const cy = (prev.y + curr.y) * 0.5;
        ctx.quadraticCurveTo(prev.x, prev.y, cx, cy);
      }
      const last = topPoints[topPoints.length - 1];
      ctx.lineTo(last.x + 40, last.y);

      // Return along lower fold wave
      for (let i = topPoints.length - 1; i >= 0; i--) {
        const pt = topPoints[i];
        const foldOffset =
          spec.thickness *
          (0.65 + 0.35 * Math.sin((i / steps) * Math.PI * 1.4 + t * 0.45 + spec.phase));
        ctx.lineTo(pt.x, pt.y + foldOffset);
      }
      ctx.closePath();

      const ribGrad = ctx.createLinearGradient(0, h * spec.yBase - 140, w, h * spec.yBase + 220);
      if (isPearl) {
        ribGrad.addColorStop(0, `hsla(${rHue}, ${Math.min(sat + 12, 85)}%, 62%, ${spec.alpha})`);
        ribGrad.addColorStop(0.5, `hsla(${(rHue + 26) % 360}, 78%, 74%, ${spec.alpha * 1.15})`);
        ribGrad.addColorStop(1, `hsla(${(rHue + 48) % 360}, 65%, 68%, 0.02)`);
      } else {
        ribGrad.addColorStop(0, `hsla(${rHue}, ${Math.min(sat + 18, 92)}%, 52%, ${spec.alpha})`);
        ribGrad.addColorStop(0.52, `hsla(${(rHue + 24) % 360}, 86%, 60%, ${spec.alpha * 1.2})`);
        ribGrad.addColorStop(1, `hsla(${(rHue + 52) % 360}, 74%, 44%, 0.02)`);
      }

      ctx.fillStyle = ribGrad;
      ctx.fill();

      // Luminous crystalline edge ridge (hallmark of Galaxy S9 Infinity silk folds)
      ctx.beginPath();
      ctx.moveTo(topPoints[0].x, topPoints[0].y);
      for (let i = 1; i < topPoints.length; i++) {
        const prev = topPoints[i - 1];
        const curr = topPoints[i];
        const cx = (prev.x + curr.x) * 0.5;
        const cy = (prev.y + curr.y) * 0.5;
        ctx.quadraticCurveTo(prev.x, prev.y, cx, cy);
      }
      const ridgeGrad = ctx.createLinearGradient(0, 0, w, 0);
      ridgeGrad.addColorStop(0, `hsla(${rHue}, 90%, 84%, 0)`);
      ridgeGrad.addColorStop(0.3, `hsla(${rHue}, 92%, ${isPearl ? 99 : 82}%, ${spec.ridgeAlpha})`);
      ridgeGrad.addColorStop(0.7, `hsla(${(rHue + 35) % 360}, 95%, ${isPearl ? 98 : 88}%, ${spec.ridgeAlpha * 0.85})`);
      ridgeGrad.addColorStop(1, `hsla(${(rHue + 50) % 360}, 88%, 82%, 0)`);

      ctx.strokeStyle = ridgeGrad;
      ctx.lineWidth = isPearl ? 1.7 : 2.1;
      ctx.stroke();
    }

    ctx.restore();

    // 4. Floating Galaxy S9 Stardust / Crystalline Motes
    ctx.save();
    for (const p of this.particles) {
      if (!this.paused) {
        p.x += p.vx * (dt * 60) * this.speed;
        p.y += p.vy * (dt * 60) * this.speed;
        if (p.y < -0.04) {
          p.y = 1.04;
          p.x = Math.random();
        }
        if (p.x < -0.04) p.x = 1.04;
        if (p.x > 1.04) p.x = -0.04;
      }

      const sx = p.x * w + px * p.depth;
      const sy = p.y * h + py * p.depth;
      const twinkle = 0.25 + 0.75 * (0.5 + 0.5 * Math.sin(t * 1.9 + p.phase));

      ctx.beginPath();
      ctx.arc(sx, sy, p.r, 0, Math.PI * 2);
      ctx.fillStyle = isPearl
        ? `hsla(${hue}, 40%, 100%, ${twinkle * 0.75})`
        : `hsla(${(hue + 20) % 360}, 90%, 92%, ${twinkle * 0.65})`;
      ctx.fill();
    }
    ctx.restore();
  }

  drawBloom(x, y, radius, color) {
    const grad = this.ctx.createRadialGradient(x, y, 0, x, y, radius);
    grad.addColorStop(0, color);
    grad.addColorStop(1, 'rgba(0, 0, 0, 0)');
    this.ctx.fillStyle = grad;
    this.ctx.fillRect(x - radius, y - radius, radius * 2, radius * 2);
  }
}

export function hexToHsl(hex) {
  let clean = String(hex || '#1F9D8B').replace('#', '').trim();
  if (clean.length === 3) {
    clean = clean
      .split('')
      .map((c) => c + c)
      .join('');
  }
  const num = parseInt(clean, 16);
  const r = ((num >> 16) & 255) / 255;
  const g = ((num >> 8) & 255) / 255;
  const b = (num & 255) / 255;

  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  let h = 0;
  let s = 0;
  const l = (max + min) / 2;

  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    switch (max) {
      case r:
        h = (g - b) / d + (g < b ? 6 : 0);
        break;
      case g:
        h = (b - r) / d + 2;
        break;
      default:
        h = (r - g) / d + 4;
        break;
    }
    h *= 60;
  }

  return {
    h: Math.round(h),
    s: Math.round(s * 100),
    l: Math.round(l * 100),
  };
}
