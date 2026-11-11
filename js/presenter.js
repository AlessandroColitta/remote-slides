/**
 * Minimalist Remote Slides - Presenter Engine (PC View)
 * Supports Microsoft PowerPoint (.pptx) & PDF files
 * Embedded Video Reproduction & Smartphone Gyroscope Laser Pointer
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

  // Video Reproduction State
  let currentZip = null;
  let slideVideos = {}; // 0-based slideIndex -> [{ targetPath, xPct, yPct, wPct, hPct, isEllipse, playOnClick, trimStartSec, trimEndSec, hasPoster }]
  let currentVideoEl = null;
  let currentVideoBlobUrl = null;
  let currentVideoInfo = null;
  let videoBlobCache = {};
  let mountedSlideIndex = -1;
  let mountingSlideIndex = -1;
  let videoMountSeq = 0;
  let clickPlayTriggered = false;
  let pendingClickPlay = false;

  // Laser Pointer State
  let isLaserActive = false;
  let laserTargetX = 0.5;
  let laserTargetY = 0.5;
  let laserCurrentX = 0.5;
  let laserCurrentY = 0.5;
  let laserAnimId = null;

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
  const videoLayer = document.getElementById('presentation-video-layer');
  const laserPointerEl = document.getElementById('laser-pointer');
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
    const hasVid = !!currentVideoEl;
    const isEnded = hasVid && !!(currentVideoEl.ended || currentVideoEl._trimEnded);
    const isPlaying = hasVid && !currentVideoEl.paused && !isEnded;

    const payload = {
      action: 'sync',
      current: currentPage,
      total: totalPages,
      isPresenting: isPresenting,
      isBlackout: isBlackout,
      hasVideo: hasVid,
      isVideoPlaying: isPlaying,
      isVideoEnded: isEnded
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
      case 'pointer':
        setLaserState(data.active, data.x, data.y);
        break;
      case 'video_toggle':
        toggleVideoPlayback();
        break;
      case 'video_play':
        if (currentVideoEl && (currentVideoEl.paused || currentVideoEl.ended || currentVideoEl._trimEnded)) {
          playCurrentVideo();
        }
        break;
      case 'video_pause':
        if (currentVideoEl && !currentVideoEl.paused) {
          currentVideoEl.pause();
          sendSync();
        }
        break;
      case 'request_sync':
        sendSync();
        break;
    }
  }

  // ==========================================
  // Laser Pointer Engine
  // ==========================================

  function renderLaserLoop() {
    if (!isLaserActive) return;
    laserCurrentX += (laserTargetX - laserCurrentX) * 0.4;
    laserCurrentY += (laserTargetY - laserCurrentY) * 0.4;

    const w = window.innerWidth;
    const h = window.innerHeight;
    laserPointerEl.style.left = Math.round(laserCurrentX * w) + 'px';
    laserPointerEl.style.top = Math.round(laserCurrentY * h) + 'px';

    laserAnimId = requestAnimationFrame(renderLaserLoop);
  }

  function setLaserState(active, normX, normY) {
    isLaserActive = !!active;
    if (isLaserActive) {
      laserPointerEl.classList.add('active');
      if (typeof normX === 'number') laserTargetX = normX;
      if (typeof normY === 'number') laserTargetY = normY;
      cancelAnimationFrame(laserAnimId);
      laserAnimId = requestAnimationFrame(renderLaserLoop);
    } else {
      laserPointerEl.classList.remove('active');
      cancelAnimationFrame(laserAnimId);
    }
  }

  // ==========================================
  // Video Reproduction Engine
  // ==========================================

  function findFirstNS(root, localName) {
    if (!root) return null;
    const list = root.getElementsByTagNameNS('*', localName);
    return list && list.length > 0 ? list[0] : null;
  }

  function findAllNS(root, localName) {
    if (!root) return [];
    return Array.from(root.getElementsByTagNameNS('*', localName));
  }

  function getRelAttr(el, attrName) {
    if (!el) return null;
    return (
      el.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', attrName) ||
      el.getAttribute('r:' + attrName) ||
      el.getAttribute(attrName) ||
      null
    );
  }

  function parseClickSequenceVideoIds(slideDoc) {
    const clickSpIds = new Set();
    const cTns = findAllNS(slideDoc, 'cTn');
    const mainSeq = cTns.find(node => node.getAttribute('nodeType') === 'mainSeq');
    if (!mainSeq) return clickSpIds;

    // Look for click-triggered media playback actions inside mainSeq (ignore interactiveSeq hover/direct click pauses)
    findAllNS(mainSeq, 'cTn').forEach(node => {
      if (node.getAttribute('nodeType') !== 'clickEffect') return;
      const presetClass = node.getAttribute('presetClass') || '';
      const presetId = node.getAttribute('presetID') || '';
      const cmds = findAllNS(node, 'cmd');

      let isMediaPlay = false;
      if (cmds.length > 0) {
        isMediaPlay = cmds.some(c => {
          const cmdStr = (c.getAttribute('cmd') || '').toLowerCase();
          return cmdStr.startsWith('play') || (presetClass === 'mediacall' && !cmdStr.includes('pause') && !cmdStr.includes('stop'));
        });
      } else if (presetClass === 'mediacall' && presetId === '1') {
        isMediaPlay = true;
      }

      if (isMediaPlay) {
        findAllNS(node, 'spTgt').forEach(tgt => {
          const spid = tgt.getAttribute('spid');
          if (spid) clickSpIds.add(spid);
        });
      }
    });

    return clickSpIds;
  }

  async function scanPptxVideos(zip) {
    slideVideos = {};
    try {
      const presXmlStr = await zip.file('ppt/presentation.xml')?.async('text');
      if (!presXmlStr) return;
      const parser = new DOMParser();
      const presDoc = parser.parseFromString(presXmlStr, 'application/xml');
      const sldSz = findFirstNS(presDoc, 'sldSz');
      const sldW = parseInt(sldSz?.getAttribute('cx')) || 12192000;
      const sldH = parseInt(sldSz?.getAttribute('cy')) || 6858000;

      for (let i = 1; i <= 250; i++) {
        const slidePath = `ppt/slides/slide${i}.xml`;
        const relsPath = `ppt/slides/_rels/slide${i}.xml.rels`;
        const slideFile = zip.file(slidePath);
        const relsFile = zip.file(relsPath);
        if (!slideFile || !relsFile) continue;

        const relsXml = await relsFile.async('text');
        const relsDoc = parser.parseFromString(relsXml, 'application/xml');
        const vidMap = {};

        relsDoc.querySelectorAll('Relationship').forEach(rel => {
          const type = rel.getAttribute('Type') || '';
          const targetMode = (rel.getAttribute('TargetMode') || '').toLowerCase();
          const rawTarget = rel.getAttribute('Target') || '';
          if (targetMode === 'external' || !rawTarget || rawTarget.toUpperCase() === 'NULL') {
            return;
          }
          if (type.includes('video') || type.includes('media')) {
            const id = rel.getAttribute('Id');
            let target = rawTarget.replace(/\\/g, '/').replace(/^\/+/, '');
            if (target.startsWith('../')) {
              target = target.replace(/^(\.\.\/)+/, 'ppt/');
            } else if (!target.startsWith('ppt/')) {
              target = 'ppt/slides/' + target;
            }
            if (zip.file(target)) {
              vidMap[id] = target;
            }
          }
        });

        if (Object.keys(vidMap).length === 0) continue;

        const slideXml = await slideFile.async('text');
        const slideDoc = parser.parseFromString(slideXml, 'application/xml');
        const clickVideoSpIds = parseClickSequenceVideoIds(slideDoc);
        const pics = findAllNS(slideDoc, 'pic');
        const vids = [];

        pics.forEach(pic => {
          const vidLink = findFirstNS(pic, 'videoFile') || findFirstNS(pic, 'quickTimeFile');
          const mediaEmbed = findFirstNS(pic, 'media');
          const rEmbed = getRelAttr(mediaEmbed, 'embed') || getRelAttr(mediaEmbed, 'link');
          const rLink = getRelAttr(vidLink, 'link') || getRelAttr(vidLink, 'embed');
          const rId = (rEmbed && vidMap[rEmbed]) ? rEmbed : ((rLink && vidMap[rLink]) ? rLink : null);

          if (rId && vidMap[rId]) {
            const cNvPr = findFirstNS(pic, 'cNvPr');
            const spId = cNvPr?.getAttribute('id') || '';

            // Locate xfrm (transform) specifically inside shape properties
            const xfrm = findFirstNS(pic, 'xfrm');
            const off = findFirstNS(xfrm || pic, 'off');

            // Find ext element that has cx/cy attributes (avoiding a:extLst/a:ext metadata)
            let ext = xfrm ? findFirstNS(xfrm, 'ext') : null;
            if (!ext || !ext.getAttribute('cx')) {
              const allExts = findAllNS(pic, 'ext');
              ext = allExts.find(e => e.getAttribute('cx')) || null;
            }

            const geom = findFirstNS(pic, 'prstGeom');
            const isEllipse = geom?.getAttribute('prst') === 'ellipse';
            const blip = findFirstNS(pic, 'blip');
            const hasPoster = !!(blip && (getRelAttr(blip, 'embed') || getRelAttr(blip, 'link')));

            const trimEl = mediaEmbed ? findFirstNS(mediaEmbed, 'trim') : findFirstNS(pic, 'trim');
            const trimStMs = parseFloat(trimEl?.getAttribute('st'));
            const trimEndMs = parseFloat(trimEl?.getAttribute('end'));
            const trimStartSec = (!isNaN(trimStMs) && trimStMs > 0) ? trimStMs / 1000 : 0;
            const trimEndSec = (!isNaN(trimEndMs) && trimEndMs > 0) ? trimEndMs / 1000 : 0;

            const x = parseInt(off?.getAttribute('x')) || 0;
            const y = parseInt(off?.getAttribute('y')) || 0;
            const cx = parseInt(ext?.getAttribute('cx')) || 0;
            const cy = parseInt(ext?.getAttribute('cy')) || 0;

            let xPct = x / sldW;
            let yPct = y / sldH;
            let wPct = cx / sldW;
            let hPct = cy / sldH;

            // Fallback to full slide if dimensions are missing or 0
            if (wPct <= 0 || hPct <= 0) {
              xPct = 0;
              yPct = 0;
              wPct = 1.0;
              hPct = 1.0;
            }

            vids.push({
              targetPath: vidMap[rId],
              xPct,
              yPct,
              wPct,
              hPct,
              isEllipse,
              hasPoster,
              playOnClick: spId ? clickVideoSpIds.has(spId) : false,
              trimStartSec,
              trimEndSec
            });
          }
        });

        if (vids.length > 0) {
          slideVideos[i - 1] = vids;
        }
      }
    } catch (err) {
      console.warn('Video scanner notice:', err);
    }
  }

  function clearVideoBlobCache() {
    Object.values(videoBlobCache).forEach(url => URL.revokeObjectURL(url));
    videoBlobCache = {};
  }

  function teardownCurrentVideo() {
    videoMountSeq++;
    mountingSlideIndex = -1;
    pendingClickPlay = false;
    if (currentVideoEl) {
      currentVideoEl.pause();
      currentVideoEl.remove();
      currentVideoEl = null;
    }
    videoLayer.innerHTML = '';
    currentVideoBlobUrl = null;
    currentVideoInfo = null;
    mountedSlideIndex = -1;
    clickPlayTriggered = false;
  }

  async function mountSlideVideos() {
    const slideIndex = currentPage - 1;

    // Preserve active video when resizing or toggling fullscreen on the same slide
    if (currentVideoEl && mountedSlideIndex === slideIndex) {
      positionVideoElement(currentVideoEl, currentVideoInfo);
      sendSync();
      return;
    }
    if (mountingSlideIndex === slideIndex) {
      return;
    }

    teardownCurrentVideo();

    const vids = slideVideos[slideIndex];
    if (vids && vids.length > 0 && currentZip) {
      // Prefer modern playable formats (mp4, webm) over legacy avi
      const vidInfo = vids.find(v => !v.targetPath.toLowerCase().endsWith('.avi')) || vids[0];
      const vidFile = currentZip.file(vidInfo.targetPath);
      if (vidFile) {
        const myMountSeq = ++videoMountSeq;
        mountingSlideIndex = slideIndex;
        try {
          let blobUrl = videoBlobCache[vidInfo.targetPath];
          if (!blobUrl) {
            const arrayBuffer = await vidFile.async('arraybuffer');
            if (myMountSeq !== videoMountSeq || currentPage - 1 !== slideIndex) {
              return;
            }
            const targetLower = vidInfo.targetPath.toLowerCase();
            let mime = 'video/mp4';
            if (targetLower.endsWith('.webm')) mime = 'video/webm';
            else if (targetLower.endsWith('.ogv') || targetLower.endsWith('.ogg')) mime = 'video/ogg';
            else if (targetLower.endsWith('.mov')) mime = 'video/quicktime';

            const blob = new Blob([arrayBuffer], { type: mime });
            blobUrl = URL.createObjectURL(blob);
            videoBlobCache[vidInfo.targetPath] = blobUrl;
          }

          currentVideoBlobUrl = blobUrl;
          const videoEl = document.createElement('video');
          videoEl.src = currentVideoBlobUrl;
          videoEl.playsInline = true;
          videoEl.className = 'presentation-video';
          videoEl.preload = 'auto';
          if (vidInfo.hasPoster) {
            videoEl.style.opacity = '0';
          }

          currentVideoEl = videoEl;
          currentVideoInfo = vidInfo;
          mountedSlideIndex = slideIndex;

          positionVideoElement(videoEl, vidInfo);

          videoEl.addEventListener('loadedmetadata', () => {
            if (currentVideoEl !== videoEl || videoEl._trimInitialized) return;
            videoEl._trimInitialized = true;
            const trimStart = vidInfo.trimStartSec || 0;
            const dur = videoEl.duration;
            if (vidInfo.trimEndSec > 0 && isFinite(dur)) {
              videoEl._stopTime = Math.max(trimStart, dur - vidInfo.trimEndSec);
            } else {
              videoEl._stopTime = Infinity;
            }
            if (trimStart > 0 && isFinite(dur) && trimStart < dur && videoEl.currentTime < trimStart - 0.1) {
              videoEl.currentTime = trimStart;
            } else if (!videoEl.paused) {
              videoEl.style.opacity = '1';
            }
          });

          videoEl.addEventListener('seeked', () => {
            if (currentVideoEl !== videoEl) return;
            if (!videoEl.paused && !videoEl._trimEnded) {
              videoEl.style.opacity = '1';
            }
          });

          videoEl.addEventListener('timeupdate', () => {
            if (currentVideoEl !== videoEl) return;
            if (videoEl._stopTime && isFinite(videoEl._stopTime) && videoEl.currentTime >= videoEl._stopTime) {
              videoEl.pause();
              videoEl._trimEnded = true;
              sendSync();
            }
          });

          videoEl.addEventListener('play', () => {
            if (currentVideoEl !== videoEl) return;
            videoEl._trimEnded = false;
            const trimStart = vidInfo.trimStartSec || 0;
            if (!vidInfo.hasPoster || trimStart === 0 || videoEl.currentTime >= trimStart - 0.25) {
              videoEl.style.opacity = '1';
            }
            sendSync();
          });
          videoEl.addEventListener('pause', () => sendSync());
          videoEl.addEventListener('ended', () => sendSync());
          videoEl.addEventListener('click', (e) => {
            e.stopPropagation();
            toggleVideoPlayback();
          });

          videoLayer.innerHTML = '';
          videoLayer.appendChild(videoEl);
          if (mountingSlideIndex === slideIndex) {
            mountingSlideIndex = -1;
          }
          if (pendingClickPlay && vidInfo.playOnClick && !clickPlayTriggered) {
            playCurrentVideo();
          }
        } catch (e) {
          console.warn('Could not mount video:', e);
        } finally {
          if (mountingSlideIndex === slideIndex) {
            mountingSlideIndex = -1;
          }
        }
      }
    }

    sendSync();
  }

  function positionVideoElement(videoEl, vidInfo) {
    if (!videoEl || !vidInfo) return;
    const canvasRect = slideCanvas.getBoundingClientRect();
    const presRect = presentationView.getBoundingClientRect();

    if (canvasRect.width === 0 || canvasRect.height === 0) return;

    const left = (canvasRect.left - presRect.left) + vidInfo.xPct * canvasRect.width;
    const top = (canvasRect.top - presRect.top) + vidInfo.yPct * canvasRect.height;
    const width = vidInfo.wPct * canvasRect.width;
    const height = vidInfo.hPct * canvasRect.height;

    videoEl.style.left = Math.round(left) + 'px';
    videoEl.style.top = Math.round(top) + 'px';
    videoEl.style.width = Math.round(width) + 'px';
    videoEl.style.height = Math.round(height) + 'px';

    if (vidInfo.isEllipse) {
      videoEl.style.borderRadius = '50%';
      videoEl.style.objectFit = 'cover';
    } else {
      videoEl.style.borderRadius = '4px';
      videoEl.style.objectFit = 'contain';
    }
  }

  function repositionMountedVideos() {
    if (currentVideoEl && currentVideoInfo) {
      positionVideoElement(currentVideoEl, currentVideoInfo);
    }
  }

  function playCurrentVideo() {
    if (!currentVideoEl) return;
    if (!isPresenting) {
      enterPresentation();
    }
    clickPlayTriggered = true;
    pendingClickPlay = false;
    const trimStart = currentVideoInfo?.trimStartSec || 0;
    if (!currentVideoInfo?.hasPoster || trimStart === 0 || currentVideoEl.readyState >= 1) {
      currentVideoEl.style.opacity = '1';
    }
    const isAtEnd = currentVideoEl.ended || currentVideoEl._trimEnded ||
      (currentVideoEl._stopTime && isFinite(currentVideoEl._stopTime) && currentVideoEl.currentTime >= currentVideoEl._stopTime - 0.05);
    if (isAtEnd) {
      currentVideoEl._trimEnded = false;
      currentVideoEl.currentTime = trimStart;
    } else if (trimStart > 0 && currentVideoEl.readyState >= 1 && currentVideoEl.currentTime < trimStart - 0.1) {
      currentVideoEl.currentTime = trimStart;
    }
    const playPromise = currentVideoEl.play();
    if (playPromise && typeof playPromise.catch === 'function') {
      playPromise.catch(err => {
        console.warn('Video play blocked:', err);
        if (currentVideoEl && !currentVideoEl.muted) {
          currentVideoEl.muted = true;
          currentVideoEl.play().catch(e2 => console.warn('Muted video play blocked:', e2));
        }
      });
    }
    sendSync();
  }

  function toggleVideoPlayback() {
    if (!currentVideoEl) return;
    if (!isPresenting) {
      enterPresentation();
    }
    if (currentVideoEl.paused || currentVideoEl.ended || currentVideoEl._trimEnded) {
      playCurrentVideo();
    } else {
      currentVideoEl.pause();
      sendSync();
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
      teardownCurrentVideo();
      clearVideoBlobCache();
      pptxViewer = new PptxViewJS.PPTXViewer({
        canvas: slideCanvas,
        maxFileSize: 1024 * 1024 * 1024,
        timeout: 180000
      });

      // Prepare ArrayBuffer if passed a File or Blob
      let buf = fileOrBuffer;
      if (typeof fileOrBuffer === 'string') {
        const res = await fetch(fileOrBuffer);
        buf = await res.arrayBuffer();
      } else if (typeof File !== 'undefined' && fileOrBuffer instanceof File) {
        buf = await fileOrBuffer.arrayBuffer();
      } else if (typeof Blob !== 'undefined' && fileOrBuffer instanceof Blob) {
        buf = await fileOrBuffer.arrayBuffer();
      }

      // Parse zip structure for embedded videos
      if (window.JSZip) {
        currentZip = await JSZip.loadAsync(buf);
        await pptxViewer.loadFile(buf);
        await scanPptxVideos(currentZip);
      } else {
        await pptxViewer.loadFile(buf);
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

  async function loadPdf(fileOrBuffer, fileName) {
    if (!window.pdfjsLib) {
      alert('PDF library not loaded');
      return;
    }

    try {
      teardownCurrentVideo();
      clearVideoBlobCache();
      let arrayBuffer = fileOrBuffer;
      if (typeof File !== 'undefined' && fileOrBuffer instanceof File) {
        arrayBuffer = await fileOrBuffer.arrayBuffer();
      } else if (typeof Blob !== 'undefined' && fileOrBuffer instanceof Blob) {
        arrayBuffer = await fileOrBuffer.arrayBuffer();
      }

      const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer });
      pdfDoc = await loadingTask.promise;
      engineType = 'pdf';
      pptxViewer = null;
      currentZip = null;
      slideVideos = {};
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

  async function handleLoadedFile(file) {
    const name = file.name.toLowerCase();
    fileNameEl.textContent = file.name;
    filePagesEl.textContent = 'Loading presentation...';
    fileInfo.style.display = 'block';
    btnStartPresentation.disabled = true;

    try {
      const buffer = await file.arrayBuffer();
      if (name.endsWith('.pptx')) {
        await loadPptx(buffer, file.name);
      } else if (name.endsWith('.pdf')) {
        await loadPdf(buffer, file.name);
      } else {
        fileInfo.style.display = 'none';
        alert('Please select a PowerPoint (.pptx) or PDF (.pdf) file.');
      }
    } catch (err) {
      console.error('Error reading presentation file:', err);
      alert('Error reading presentation file: ' + err.message);
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

    // Mount any videos embedded on this slide
    await mountSlideVideos();

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
    if (!engineType) return;
    // Start click-sequence video first before advancing to next slide
    const slideIdx = currentPage - 1;
    const slideVids = slideVideos[slideIdx];
    if (!clickPlayTriggered && slideVids && slideVids.some(v => v.playOnClick)) {
      if (currentVideoEl && currentVideoInfo && currentVideoInfo.playOnClick) {
        playCurrentVideo();
        return;
      }
      if (isRendering || mountingSlideIndex === slideIdx) {
        pendingClickPlay = true;
        return;
      }
    }
    if (currentPage >= totalPages) return;
    pendingClickPlay = false;
    currentPage++;
    renderCurrentSlide();
  }

  function prevPage() {
    if (!engineType || currentPage <= 1) return;
    pendingClickPlay = false;
    currentPage--;
    renderCurrentSlide();
  }

  function goToPage(num) {
    if (!engineType) return;
    if (num >= 1 && num <= totalPages) {
      pendingClickPlay = false;
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

    setLaserState(false);
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
    } else if (engineType) {
      renderCurrentSlide();
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
      case 'PageDown':
      case 'Enter':
        nextPage();
        break;
      case 'Space':
        // If a video is present, Space toggles video play/pause; otherwise advances slide
        if (currentVideoEl) {
          toggleVideoPlayback();
        } else {
          nextPage();
        }
        break;
      case 'KeyV':
        toggleVideoPlayback();
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
