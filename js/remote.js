/**
 * Minimalist Remote Slides - Smartphone Controller Engine
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
      
      // Request current slide sync
      sendAction('request_sync');

      // Start elapsed timer
      startTimer();

      // Acquire screen wake lock so phone doesn't sleep while speaking
      requestWakeLock();
    });

    conn.on('data', (data) => {
      handlePresenterSync(data);
    });

    conn.on('close', () => {
      console.log('Presenter connection closed');
      statusEl.textContent = 'Disconnected';
      statusEl.classList.remove('connected');
      // Retry connection after brief backoff
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

    // Update UI
    slideNumEl.textContent = currentSlide;
    slideTotalEl.textContent = 'of ' + totalSlides;

    const progressPct = totalSlides > 1 ? ((currentSlide - 1) / (totalSlides - 1)) * 100 : 0;
    progressFillEl.style.width = Math.min(100, Math.max(0, progressPct)) + '%';

    // Blackout button state
    if (isBlackout) {
      btnBlackout.textContent = 'Un-Blackout';
      btnBlackout.classList.add('active');
    } else {
      btnBlackout.textContent = 'Blackout';
      btnBlackout.classList.remove('active');
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
    } catch (e) {
      // Wake lock can be denied under low-battery mode
    }
  }

  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'visible' && !wakeLock) {
      requestWakeLock();
    }
  });

  // ==========================================
  // Button Event Handlers
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
    const diffX = e.changedTouches[0].clientX - touchStartX;
    const diffY = e.changedTouches[0].clientY - touchStartY;

    // Detect horizontal swipe with minimum threshold and low vertical drift
    if (Math.abs(diffX) > 60 && Math.abs(diffY) < 50) {
      if (diffX < 0) {
        // Swipe left -> Next slide
        sendAction('next');
        triggerVibrate(25);
      } else {
        // Swipe right -> Prev slide
        sendAction('prev');
        triggerVibrate(25);
      }
    }
  }, { passive: true });

  // Initialize
  connectToPresenter();
})();
