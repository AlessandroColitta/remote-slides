/**
 * Minimalist Remote Slides - Presenter Engine (PC View)
 * Supports Microsoft PowerPoint (.pptx) & PDF files
 * Works in Google Chrome & Microsoft Edge
 */

(function () {
  'use strict';

  // Deck State
  let engineType = null; // 'pptx' | 'pdf'
  let pptxViewer = null;
  let pdfDoc = null;
  let currentPage = 1;
  let totalPages = 0;
  let isRendering = false;
  let pendingPage = null;
  let isPresenting = false;
  let isBlackout = false;

  // WebRTC PeerJS State
  let peer = null;
  let activeConnections = [];
  let roomId = '';

  // DOM Elements
  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('file-input');
  const btnDemo = document.getElementById('btn-demo');
  const btnStartPresentation = document.getElementById('btn-start-presentation');
  const fileInfo = document.getElementById('file-info');
  const fileNameEl = document.getElementById('file-name');
  const filePagesEl = document.getElementById('file-pages');
  const qrContainer = document.getElementById('qr-box');
  const roomCodeTag = document.getElementById('room-code-tag');
  const connectionStatusEl = document.getElementById('connection-status');
  const mobileStatusTag = document.getElementById('mobile-status-tag');
  
  // Presentation Elements
  const presentationView = document.getElementById('presentation-view');
  const slideCanvas = document.getElementById('slide-canvas');
  const canvasCtx = slideCanvas.getContext('2d');
  const blackoutOverlay = document.getElementById('blackout-overlay');
  const presentationBar = document.getElementById('presentation-bar');
  const presentationCounter = document.getElementById('presentation-counter');
  const btnBarPrev = document.getElementById('btn-bar-prev');
  const btnBarNext = document.getElementById('btn-bar-next');
  const btnBarExit = document.getElementById('btn-bar-exit');

  let hideBarTimer = null;
  let currentPdfRenderTask = null;

  // Configure PDF.js worker if available
  if (window.pdfjsLib) {
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'vendor/pdf.worker.min.js';
  }

  // ==========================================
  // PeerJS Connection Management
  // ==========================================

  function generateRoomId() {
    const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
    let code = '';
    for (let i = 0; i < 6; i++) {
      code += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    return 'RS-' + code;
  }

  function initPeer() {
    roomId = generateRoomId();
    roomCodeTag.textContent = roomId;

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
      peer = new Peer(roomId, peerOptions);
    } catch (e) {
      console.error('PeerJS init error:', e);
      return;
    }

    peer.on('open', (id) => {
      console.log('Presenter peer opened with ID:', id);
      renderQRCode(id);
    });

    peer.on('connection', (conn) => {
      console.log('Incoming phone connection:', conn.peer);
      
      conn.on('open', () => {
        activeConnections.push(conn);
        updateConnectionUI(true);
        sendSync(conn);
      });

      conn.on('data', (data) => {
        handleRemoteCommand(data);
      });

      conn.on('close', () => {
        activeConnections = activeConnections.filter(c => c !== conn);
        if (activeConnections.length === 0) {
          updateConnectionUI(false);
        }
      });

      conn.on('error', (err) => {
        console.warn('Connection error:', err);
      });
    });

    peer.on('error', (err) => {
      console.warn('Peer error:', err);
      if (err.type === 'unavailable-id') {
        initPeer();
      }
    });
  }

  function renderQRCode(id) {
    qrContainer.innerHTML = '';
    const remoteUrl = new URL('remote.html', window.location.href);
    remoteUrl.searchParams.set('room', id);

    if (window.QRCode) {
      new QRCode(qrContainer, {
        text: remoteUrl.toString(),
        width: 170,
        height: 170,
        colorDark: '#000000',
        colorLight: '#ffffff',
        correctLevel: QRCode.CorrectLevel.M
      });
    }

    roomCodeTag.style.cursor = 'pointer';
    roomCodeTag.title = 'Click to copy remote link';
    roomCodeTag.onclick = () => {
      navigator.clipboard.writeText(remoteUrl.toString()).then(() => {
        const orig = roomCodeTag.textContent;
        roomCodeTag.textContent = 'LINK COPIED';
        setTimeout(() => roomCodeTag.textContent = orig, 1500);
      });
    };
  }

  function updateConnectionUI(isConnected) {
    if (isConnected) {
      connectionStatusEl.textContent = 'Smartphone connected';
      connectionStatusEl.classList.add('connected');
      if (mobileStatusTag) {
        mobileStatusTag.textContent = 'Connected (' + activeConnections.length + ')';
        mobileStatusTag.classList.add('connected');
      }
    } else {
      connectionStatusEl.textContent = 'Waiting for smartphone...';
      connectionStatusEl.classList.remove('connected');
      if (mobileStatusTag) {
        mobileStatusTag.textContent = 'Waiting for connection';
        mobileStatusTag.classList.remove('connected');
      }
    }
  }

  function broadcast(data) {
    activeConnections.forEach(conn => {
      if (conn && conn.open) {
        try {
          conn.send(data);
        } catch (e) {
          console.warn('Failed to send to conn:', e);
        }
      }
    });
  }

  function sendSync(targetConn) {
    const payload = {
      action: 'sync',
      current: currentPage,
      total: totalPages,
      isPresenting: isPresenting,
      isBlackout: isBlackout
    };
    if (targetConn && targetConn.open) {
      targetConn.send(payload);
    } else {
      broadcast(payload);
    }
  }

  function handleRemoteCommand(data) {
    if (!data || !data.action) return;
    
    switch (data.action) {
      case 'next':
        nextPage();
        break;
      case 'prev':
        prevPage();
        break;
      case 'goto':
        if (typeof data.page === 'number') {
          goToPage(data.page);
        }
        break;
      case 'blackout':
        toggleBlackout();
        break;
      case 'exit':
        exitPresentation();
        break;
      case 'request_sync':
        sendSync();
        break;
    }
  }

  // ==========================================
  // PPTX & PDF Presentation Engine
  // ==========================================

  function calculateTargetCanvasSize() {
    const availW = isPresenting ? window.innerWidth : 800;
    const availH = isPresenting ? window.innerHeight : 450;
    const aspect = 16 / 9;

    let w, h;
    if (availW / availH > aspect) {
      h = availH;
      w = Math.floor(h * aspect);
    } else {
      w = availW;
      h = Math.floor(w / aspect);
    }
    return { width: w, height: h };
  }

  async function loadPptx(fileOrBuffer, fileName) {
    if (!window.PptxViewJS || !window.PptxViewJS.PPTXViewer) {
      alert('PowerPoint viewer library not loaded');
      return;
    }

    try {
      pptxViewer = new PptxViewJS.PPTXViewer({ canvas: slideCanvas });
      
      if (typeof fileOrBuffer === 'string') {
        await pptxViewer.loadFromUrl(fileOrBuffer);
      } else {
        await pptxViewer.loadFile(fileOrBuffer);
      }

      engineType = 'pptx';
      pdfDoc = null;
      totalPages = pptxViewer.getSlideCount();
      currentPage = 1;

      // Update UI
      fileNameEl.textContent = fileName || 'presentation.pptx';
      filePagesEl.textContent = totalPages + ' slides';
      fileInfo.style.display = 'block';
      btnStartPresentation.disabled = false;

      await renderCurrentSlide();
      sendSync();
    } catch (err) {
      console.error('Failed to load PPTX:', err);
      alert('Could not render this PowerPoint file: ' + err.message);
    }
  }

  async function loadPdf(arrayBuffer, fileName) {
    if (!window.pdfjsLib) {
      alert('PDF library not loaded');
      return;
    }

    try {
      const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer });
      pdfDoc = await loadingTask.promise;
      engineType = 'pdf';
      pptxViewer = null;
      totalPages = pdfDoc.numPages;
      currentPage = 1;

      // Update UI
      fileNameEl.textContent = fileName || 'presentation.pdf';
      filePagesEl.textContent = totalPages + ' slides';
      fileInfo.style.display = 'block';
      btnStartPresentation.disabled = false;

      await renderCurrentSlide();
      sendSync();
    } catch (err) {
      console.error('Failed to load PDF:', err);
      alert('Could not render this PDF: ' + err.message);
    }
  }

  function handleLoadedFile(file) {
    const name = file.name.toLowerCase();
    if (name.endsWith('.pptx')) {
      loadPptx(file, file.name);
    } else if (name.endsWith('.pdf')) {
      const reader = new FileReader();
      reader.onload = (e) => loadPdf(e.target.result, file.name);
      reader.readAsArrayBuffer(file);
    } else {
      alert('Please select a PowerPoint (.pptx) or PDF (.pdf) file.');
    }
  }

  async function renderCurrentSlide() {
    if (!engineType) return;
    if (isRendering) {
      pendingPage = currentPage;
      return;
    }
    isRendering = true;

    const { width, height } = calculateTargetCanvasSize();

    if (engineType === 'pptx' && pptxViewer) {
      try {
        slideCanvas.width = width;
        slideCanvas.height = height;
        slideCanvas.style.width = width + 'px';
        slideCanvas.style.height = height + 'px';

        // PptxViewJS uses 0-based slide indexing
        await pptxViewer.render(slideCanvas, { slideIndex: currentPage - 1 });
      } catch (err) {
        console.error('PPTX render error:', err);
      }
    } else if (engineType === 'pdf' && pdfDoc) {
      try {
        if (currentPdfRenderTask) {
          currentPdfRenderTask.cancel();
          currentPdfRenderTask = null;
        }

        const page = await pdfDoc.getPage(currentPage);
        const unscaled = page.getViewport({ scale: 1.0 });
        const scale = Math.min(width / unscaled.width, height / unscaled.height);
        const dpr = window.devicePixelRatio || 1;
        const viewport = page.getViewport({ scale: scale });

        slideCanvas.width = Math.floor(viewport.width * dpr);
        slideCanvas.height = Math.floor(viewport.height * dpr);
        slideCanvas.style.width = Math.floor(viewport.width) + 'px';
        slideCanvas.style.height = Math.floor(viewport.height) + 'px';

        currentPdfRenderTask = page.render({
          canvasContext: canvasCtx,
          transform: [dpr, 0, 0, dpr, 0, 0],
          viewport: viewport
        });
        await currentPdfRenderTask.promise;
      } catch (err) {
        if (err && err.name !== 'RenderingCancelledException') {
          console.error('PDF render error:', err);
        }
      }
    }

    isRendering = false;
    presentationCounter.textContent = currentPage + ' / ' + totalPages;
    sendSync();

    if (pendingPage !== null) {
      const target = pendingPage;
      pendingPage = null;
      currentPage = target;
      renderCurrentSlide();
    }
  }

  function loadSampleDeck() {
    btnDemo.disabled = true;
    btnDemo.textContent = 'Loading...';
    // Load sample-presentation.pptx
    loadPptx('sample-presentation.pptx', 'sample-presentation.pptx')
      .then(() => {
        btnDemo.textContent = 'PPTX Loaded';
      })
      .catch(() => {
        btnDemo.textContent = 'Try Demo Deck';
        btnDemo.disabled = false;
      });
  }

  function nextPage() {
    if (!engineType || currentPage >= totalPages) return;
    currentPage++;
    renderCurrentSlide();
  }

  function prevPage() {
    if (!engineType || currentPage <= 1) return;
    currentPage--;
    renderCurrentSlide();
  }

  function goToPage(num) {
    if (!engineType) return;
    if (num >= 1 && num <= totalPages) {
      currentPage = num;
      renderCurrentSlide();
    }
  }

  // ==========================================
  // Presentation Mode Management
  // ==========================================

  function enterPresentation() {
    if (!engineType) return;
    isPresenting = true;
    presentationView.classList.add('active');

    // Request browser fullscreen (requires user gesture on PC)
    if (document.documentElement.requestFullscreen) {
      document.documentElement.requestFullscreen().catch(() => {});
    }

    renderCurrentSlide();
    showPresentationBarBriefly();
    sendSync();
  }

  function exitPresentation() {
    isPresenting = false;
    presentationView.classList.remove('active');

    if (document.fullscreenElement && document.exitFullscreen) {
      document.exitFullscreen().catch(() => {});
    }

    renderCurrentSlide();
    sendSync();
  }

  function toggleBlackout() {
    isBlackout = !isBlackout;
    if (isBlackout) {
      blackoutOverlay.classList.add('active');
    } else {
      blackoutOverlay.classList.remove('active');
    }
    sendSync();
  }

  function showPresentationBarBriefly() {
    presentationBar.classList.add('visible');
    clearTimeout(hideBarTimer);
    hideBarTimer = setTimeout(() => {
      presentationBar.classList.remove('visible');
    }, 2800);
  }

  document.addEventListener('fullscreenchange', () => {
    if (!document.fullscreenElement && isPresenting) {
      exitPresentation();
    }
  });

  presentationView.addEventListener('mousemove', () => {
    showPresentationBarBriefly();
  });

  slideCanvas.addEventListener('click', (e) => {
    const rect = slideCanvas.getBoundingClientRect();
    const x = e.clientX - rect.left;
    if (x > rect.width * 0.5) {
      nextPage();
    } else {
      prevPage();
    }
  });

  window.addEventListener('resize', () => {
    if (engineType) {
      renderCurrentSlide();
    }
  });

  // Keyboard navigation on PC
  window.addEventListener('keydown', (e) => {
    if (['ArrowRight', 'ArrowLeft', 'ArrowUp', 'ArrowDown', 'Space', 'PageDown', 'PageUp'].includes(e.code)) {
      e.preventDefault();
    }

    switch (e.code) {
      case 'ArrowRight':
      case 'ArrowDown':
      case 'Space':
      case 'PageDown':
      case 'Enter':
        nextPage();
        break;
      case 'ArrowLeft':
      case 'ArrowUp':
      case 'PageUp':
      case 'Backspace':
        prevPage();
        break;
      case 'KeyF':
        if (isPresenting) {
          exitPresentation();
        } else {
          enterPresentation();
        }
        break;
      case 'KeyB':
      case 'Period':
        toggleBlackout();
        break;
      case 'Escape':
        if (isPresenting) {
          exitPresentation();
        }
        break;
    }
  });

  // ==========================================
  // File Upload & Drag-and-Drop
  // ==========================================

  dropzone.addEventListener('click', () => fileInput.click());

  fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) {
      handleLoadedFile(file);
    }
  });

  ['dragenter', 'dragover'].forEach(eventName => {
    dropzone.addEventListener(eventName, (e) => {
      e.preventDefault();
      dropzone.classList.add('dragover');
    });
  });

  ['dragleave', 'drop'].forEach(eventName => {
    dropzone.addEventListener(eventName, (e) => {
      e.preventDefault();
      dropzone.classList.remove('dragover');
    });
  });

  dropzone.addEventListener('drop', (e) => {
    const file = e.dataTransfer.files[0];
    if (file) {
      handleLoadedFile(file);
    }
  });

  window.addEventListener('dragover', (e) => e.preventDefault());
  window.addEventListener('drop', (e) => e.preventDefault());

  // Buttons in UI
  btnDemo.addEventListener('click', loadSampleDeck);
  btnStartPresentation.addEventListener('click', enterPresentation);
  btnBarPrev.addEventListener('click', prevPage);
  btnBarNext.addEventListener('click', nextPage);
  btnBarExit.addEventListener('click', exitPresentation);

  // Initialize WebRTC
  initPeer();
})();
