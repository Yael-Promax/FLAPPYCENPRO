(() => {
  'use strict';
  let muted = false;
  let actx = null;
  /* ---------------------------------------------------------------
     Sonido (WebAudio, sin archivos externos)
  --------------------------------------------------------------- */
  function audio() {
    if (muted) return null;
    if (!actx) {
      const AC = window.AudioContext || window.webkitAudioContext;
      if (!AC) return null;
      try { actx = new AC(); } catch (e) { return null; }
    }
    if (actx.state === 'suspended') actx.resume();
    return actx;
  }

  function tone(freq, dur, type, vol, to, delay) {
    const ac = audio();
    if (!ac) return;
    const t0 = ac.currentTime + (delay || 0);
    const o = ac.createOscillator();
    const g = ac.createGain();
    o.type = type || 'sine';
    o.frequency.setValueAtTime(freq, t0);
    if (to) o.frequency.exponentialRampToValueAtTime(to, t0 + dur);
    g.gain.setValueAtTime(0.0001, t0);
    g.gain.exponentialRampToValueAtTime(vol, t0 + 0.012);
    g.gain.exponentialRampToValueAtTime(0.0001, t0 + dur);
    o.connect(g);
    g.connect(ac.destination);
    o.start(t0);
    o.stop(t0 + dur + 0.03);
  }

  const sfx = {
    flap()  { tone(560, 0.09, 'sine', 0.07, 780); tone(300, 0.05, 'triangle', 0.025); },
    score() { tone(880, 0.08, 'square', 0.035); tone(1320, 0.12, 'square', 0.028, null, 0.07); },
    hit()   { tone(170, 0.16, 'sawtooth', 0.08, 60); },
    die()   { tone(420, 0.5, 'triangle', 0.06, 80); }
  };

  window.FlappyAudio = Object.freeze({
    unlock() { audio(); },
    flap: sfx.flap,
    score: sfx.score,
    hit: sfx.hit,
    die: sfx.die,
    toggleMute() { muted = !muted; return muted; }
  });
})();
