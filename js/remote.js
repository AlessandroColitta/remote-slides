/**
 * Minimalist Remote Slides - Smartphone Controller Engine
 * Supports Gyroscope Laser Pointer & Video Playback Control
 * Optimized for Mobile Chrome on Android
 */

(function () {
  'use strict';

  // State
  let peer = null;
  let conn = null;
  let roomId = null;
  let currentSlide = 1;
  let totalSlides = 1;
  let isBlackout = false;
  let wakeLock = null;

  // Video State
  let hasVideo = false;
  let isVideoPlaying = false;
  let isVideoEnded = false;

  // Gyroscope Pointer State
  let isPointerActive = false;
  let pointerInterval = null;
  let currentAlpha = 0;
  let currentBeta = 0;
  let currentGamma = 0;
  let calibAlpha = 0;
  let calibBeta = 0;
  let calibGamma = 0;
  let pointerPressStartTime = 0;
  let isPointerToggleMode = false;

  // Timer State
  let timerInterval = null;
  let elapsedSeconds = 0;
  let isTimerRunning = false;

  // DOM Elements
  const statusEl = document.getElementById('mobile-status');
  const timerEl = document.getElementById('mobile-timer');
  const slideNumEl = document.getElementById('slide-num');
  const slideTotalEl = document.getElementById('slide-total');
  const progressFillEl = document.getElementById('progress-fill');
  const btnPrev = document.getElementById('btn-prev');
  const btnNext = document.getElementById('btn-next');
  const btnExit = document.getElementById('btn-exit');
  const btnPointer = document.getElementById('btn-pointer');
  const btnVideo = document.getElementById('btn-video');
  const btnBlackout = document.getElementById('btn-blackout');
  const btnResetTimer = document.getElementById('btn-reset-timer');

  // Extract Room ID from URL
  const params = new URLSearchParams(window.location.search);
  roomId = params.get('room');

  if (!roomId) {
    statusEl.textContent = 'Invalid Link';
    alert('No presentation room ID found in URL. Please scan the QR code again.');
    return;
  }

  // ==========================================
  // PeerJS Client Connection
  // ==========================================

  function connectToPresenter() {
    statusEl.textContent = 'Connecting...';
    statusEl.classList.remove('connected');

    const peerOptions = {
      debug: 1,
      config: {
        iceServers: [
          { urls: 'stun:stun.l.google.com:19302' },
          { urls: 'stun:stun1.l.google.com:19302' },
          { urls: 'stun:global.stun.twilio.com:3478' }
        ]
      }
    };

    try {
      peer = new Peer(peerOptions);
    } catch (e) {
      statusEl.textContent = 'PeerJS Error';
      console.error(e);
      return;
    }

    peer.on('open', (id) => {
      console.log('Mobile peer ID:', id);
      establishConnection();
    });

    peer.on('error', (err) => {
      console.warn('Peer error:', err);
      statusEl.textContent = 'Connection Error';
    });
  }

  function establishConnection() {
    conn = peer.connect(roomId, { reliable: true });

    conn.on('open', () => {
      console.log('Connected to presenter room:', roomId);
      statusEl.textContent = 'Connected';
      statusEl.classList.add('connected');
      
      sendAction('request_sync');
      startTimer();
      requestWakeLock();
    });

    conn.on('data', (data) => {
      handlePresenterSync(data);
    });

    conn.on('close', () => {
      console.log('Presenter connection closed');
      statusEl.textContent = 'Disconnected';
      statusEl.classList.remove('connected');
      setTimeout(establishConnection, 2500);
    });

    conn.on('error', (err) => {
      console.warn('Data connection error:', err);
      statusEl.textContent = 'Link Error';
    });
  }

  function handlePresenterSync(data) {
    if (!data || data.action !== 'sync') return;

    if (typeof data.current === 'number') currentSlide = data.current;
    if (typeof data.total === 'number' && data.total > 0) totalSlides = data.total;
    if (typeof data.isBlackout === 'boolean') isBlackout = data.isBlackout;

    // Update Slide UI
    slideNumEl.textContent = currentSlide;
    slideTotalEl.textContent = 'of ' + totalSlides;

    const progressPct = totalSlides > 1 ? ((currentSlide - 1) / (totalSlides - 1)) * 100 : 0;
    progressFillEl.style.width = Math.min(100, Math.max(0, progressPct)) + '%';

    // Blackout Button State
    if (isBlackout) {
      btnBlackout.textContent = 'Un-Blackout';
      btnBlackout.classList.add('active');
    } else {
      btnBlackout.textContent = 'Blackout';
      btnBlackout.classList.remove('active');
    }

    // Video Button State
    hasVideo = !!data.hasVideo;
    isVideoPlaying = !!data.isVideoPlaying;
    isVideoEnded = !!data.isVideoEnded;

    if (hasVideo) {
      btnVideo.style.display = 'block';
      if (isVideoPlaying) {
        btnVideo.textContent = 'Pause Video';
        btnVideo.classList.add('video-playing');
      } else if (isVideoEnded) {
        btnVideo.textContent = 'Replay Video';
        btnVideo.classList.remove('video-playing');
      } else {
        btnVideo.textContent = 'Play Video';
        btnVideo.classList.remove('video-playing');
      }
    } else {
      btnVideo.style.display = 'none';
      btnVideo.classList.remove('video-playing');
    }
  }

  function sendAction(action, payload) {
    if (conn && conn.open) {
      const msg = Object.assign({ action: action }, payload || {});
      try {
        conn.send(msg);
      } catch (e) {
        console.warn('Send failed:', e);
      }
    }
  }

  // ==========================================
  // Gyroscope / Orientation Laser Pointer
  // ==========================================

  window.addEventListener('deviceorientation', (e) => {
    if (e.alpha !== null) currentAlpha = e.alpha;
    if (e.beta !== null) currentBeta = e.beta;
    if (e.gamma !== null) currentGamma = e.gamma;
  }, { passive: true });

  function startPointer() {
    if (isPointerActive) return;
    isPointerActive = true;
    btnPointer.textContent = 'Pointer: ON';
    btnPointer.classList.add('pointer-active');
    triggerVibrate(25);

    // Calibrate origin when user begins pointing
    calibAlpha = currentAlpha;
    calibBeta = currentBeta;
    calibGamma = currentGamma;

    if (typeof DeviceOrientationEvent !== 'undefined' && typeof DeviceOrientationEvent.requestPermission === 'function') {
      DeviceOrientationEvent.requestPermission().catch(() => {});
    }

    // Transmit coordinates at ~40 fps
    pointerInterval = setInterval(() => {
      if (!isPointerActive) return;

      // 1. Pitch: Tilting top down decreases beta.
      // Rule: "tilting downwards the forward part of the phone makes the pointer go down, the opposite to go upwards"
      const deltaBeta = currentBeta - calibBeta;
      const sensitivityY = 0.024; // ~40 degrees covers full vertical slide
      let targetY = 0.5 - deltaBeta * sensitivityY;

      // 2. Yaw + Roll: Rule: "left and right for left and right phone rotation"
      let diffAlpha = (currentAlpha - calibAlpha + 540) % 360 - 180;
      let diffGamma = currentGamma - calibGamma;
      const deltaH = -diffAlpha * 0.7 + diffGamma * 0.45;
      const sensitivityX = 0.022; // ~45 degrees covers full horizontal slide
      let targetX = 0.5 + deltaH * sensitivityX;

      // Clamp to screen
      targetX = Math.max(0.01, Math.min(0.99, targetX));
      targetY = Math.max(0.01, Math.min(0.99, targetY));

      sendAction('pointer', { active: true, x: targetX, y: targetY });
    }, 25);
  }

  function stopPointer() {
    if (!isPointerActive) return;
    isPointerActive = false;
    isPointerToggleMode = false;
    btnPointer.textContent = 'Pointer';
    btnPointer.classList.remove('pointer-active');
    clearInterval(pointerInterval);
    triggerVibrate(15);
    sendAction('pointer', { active: false });
  }

  // Pointer Button: Supports both press-and-hold (like a real remote) & tap-to-toggle
  let isPointerPressing = false;

  btnPointer.addEventListener('pointerdown', (e) => {
    e.preventDefault();
    if (isPointerToggleMode) {
      stopPointer();
      return;
    }
    isPointerPressing = true;
    pointerPressStartTime = Date.now();
    startPointer();
  });

  window.addEventListener('pointerup', () => {
    if (isPointerPressing) {
      isPointerPressing = false;
      const pressDuration = Date.now() - pointerPressStartTime;
      if (pressDuration < 280) {
        // Quick tap: keep pointer toggled ON until next tap
        isPointerToggleMode = true;
      } else {
        // Held and released: turn off immediately
        stopPointer();
      }
    }
  });

  btnPointer.addEventListener('pointercancel', () => {
    if (isPointerPressing) {
      isPointerPressing = false;
      stopPointer();
    }
  });

  // ==========================================
  // Video Playback Control
  // ==========================================

  btnVideo.addEventListener('click', () => {
    sendAction('video_toggle');
    triggerVibrate(25);
  });

  // ==========================================
  // Haptic Feedback
  // ==========================================

  function triggerVibrate(ms) {
    if ('vibrate' in navigator) {
      try {
        navigator.vibrate(ms || 25);
      } catch (e) {}
    }
  }

  // ==========================================
  // Presentation Timer
  // ==========================================

  function startTimer() {
    if (isTimerRunning) return;
    isTimerRunning = true;
    timerInterval = setInterval(() => {
      elapsedSeconds++;
      updateTimerDisplay();
    }, 1000);
  }

  function resetTimer() {
    elapsedSeconds = 0;
    updateTimerDisplay();
    triggerVibrate(20);
  }

  function updateTimerDisplay() {
    const mins = Math.floor(elapsedSeconds / 60);
    const secs = elapsedSeconds % 60;
    timerEl.textContent = (mins < 10 ? '0' : '') + mins + ':' + (secs < 10 ? '0' : '') + secs;
  }

  // ==========================================
  // Screen Wake Lock API
  // ==========================================

  async function requestWakeLock() {
    try {
      if ('wakeLock' in navigator) {
        wakeLock = await navigator.wakeLock.request('screen');
        wakeLock.addEventListener('release', () => {
          wakeLock = null;
        });
      }
    } catch (e) {}
  }

  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'visible' && !wakeLock) {
      requestWakeLock();
    }
  });

  // ==========================================
  // Slide Navigation Handlers
  // ==========================================

  btnNext.addEventListener('click', () => {
    sendAction('next');
    triggerVibrate(25);
  });

  btnPrev.addEventListener('click', () => {
    sendAction('prev');
    triggerVibrate(25);
  });

  btnExit.addEventListener('click', () => {
    sendAction('exit');
    triggerVibrate(45);
  });

  btnBlackout.addEventListener('click', () => {
    sendAction('blackout');
    triggerVibrate(25);
  });

  btnResetTimer.addEventListener('click', resetTimer);

  // ==========================================
  // Swipe Gestures Support
  // ==========================================

  let touchStartX = 0;
  let touchStartY = 0;

  window.addEventListener('touchstart', (e) => {
    if (e.touches.length === 1) {
      touchStartX = e.touches[0].clientX;
      touchStartY = e.touches[0].clientY;
    }
  }, { passive: true });

  window.addEventListener('touchend', (e) => {
    if (!e.changedTouches || e.changedTouches.length === 0) return;
    // Don't swipe while pointer is active
    if (isPointerActive) return;

    const diffX = e.changedTouches[0].clientX - touchStartX;
    const diffY = e.changedTouches[0].clientY - touchStartY;

    if (Math.abs(diffX) > 60 && Math.abs(diffY) < 50) {
      if (diffX < 0) {
        sendAction('next');
        triggerVibrate(25);
      } else {
        sendAction('prev');
        triggerVibrate(25);
      }
    }
  }, { passive: true });

  // Initialize
  connectToPresenter();
})();
