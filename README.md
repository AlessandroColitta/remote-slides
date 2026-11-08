# Remote Slides

A zero-installation, web-based presentation deck that reads **Microsoft PowerPoint (`.pptx`)** presentations directly in the browser and turns any Android smartphone into a remote clicker with a **gyroscope laser pointer** and **video playback control** via Chrome.

Designed with a clean, distraction-free aesthetic:
- **PC Presenter Deck:** Pure minimalist typography, zero icons, zero clutter.
- **Smartphone Controller:** Large tactile touch zones, haptic feedback, elapsed speaking timer, and minimalist vector icons strictly for changing slides and exiting presentation mode.

---

## Key Features

- **Native PowerPoint (`.pptx`) Support:** Directly opens and renders `.pptx` files client-side in the browser (using `PptxViewJS`, `JSZip`, and `Chart.js`). No Microsoft Office, no PowerPoint installation, and no server conversions required.
- **Embedded Video Reproduction:** Extracts and plays embedded `.mp4` presentation videos directly on the slide at their exact coordinates. Supports click-to-play on the PC and remote Play/Pause control directly from the smartphone.
- **Smartphone Gyroscope Laser Pointer:** Turns your phone into a physical presentation pointer using device orientation:
  - Tilting the forward part of the phone downwards moves the pointer down.
  - Tilting the forward part upwards moves the pointer up.
  - Rotating the phone left or right moves the pointer left or right.
  - Supports both clicker press-and-hold and tap-to-toggle modes with motion smoothing.
- **Fallback PDF Support:** Also accepts standard `.pdf` presentations.
- **Zero Software Installation:** Works instantly on modern browsers (**Google Chrome** and **Microsoft Edge**) on any podium or conference PC.
- **Instant QR Code Pairing:** Generates a secure session code on the screen. Scan with your Android camera to connect.
- **Sub-millisecond Latency:** Direct peer-to-peer communication powered by WebRTC (`PeerJS`). No centralized servers recording your slides.
- **Screen Wake Lock:** Keeps your phone screen awake throughout your talk so it never goes to sleep while speaking.
- **Haptic Feedback:** Vibrates subtly upon changing slides, engaging the pointer, or toggling videos so you feel the transition without looking down.
- **Swipe Gestures:** Swipe left or right anywhere on your phone screen to navigate.
- **Blackout Support:** Temporarily black the screen during Q&A or audience discussions.

---

## Presentation Shortcuts on PC

| Key | Action |
| :--- | :--- |
| **Right Arrow / Space / PageDown / Enter** | Next slide |
| **Space** (when video is present) | Toggle video play / pause |
| **V** | Toggle video play / pause |
| **Left Arrow / PageUp / Backspace** | Previous slide |
| **F** | Enter / Exit presentation mode |
| **B** or **.** | Blackout screen |
| **Escape** | Exit presentation mode |
