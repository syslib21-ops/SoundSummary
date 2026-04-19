/**
 * 음성 파일: @xenova/transformers (Whisper)
 * 마이크 실시간: Web Speech API
 */

const $ = (id) => document.getElementById(id);

/** 브라우저용 ONNX Whisper (tiny는 WER이 커서 기본은 base) */
const WHISPER_MODELS = {
  tiny: "Xenova/whisper-tiny",
  base: "Xenova/whisper-base",
  small: "Xenova/whisper-small",
};

const els = {
  tabs: document.querySelectorAll(".tab"),
  panelFile: $("panel-file"),
  panelRecord: $("panel-record"),
  panelMic: $("panel-mic"),
  dropzone: $("dropzone"),
  fileInput: $("fileInput"),
  fileName: $("fileName"),
  modelSelect: $("modelSelect"),
  speedSelect: $("speedSelect"),
  btnTranscribeFile: $("btnTranscribeFile"),
  btnClearFile: $("btnClearFile"),
  btnMicStart: $("btnMicStart"),
  btnMicStop: $("btnMicStop"),
  micStatus: $("micStatus"),
  btnRecordStart: $("btnRecordStart"),
  btnRecordStop: $("btnRecordStop"),
  btnSaveRecording: $("btnSaveRecording"),
  btnRecordTranscribe: $("btnRecordTranscribe"),
  recordStatus: $("recordStatus"),
  progressWrap: $("progressWrap"),
  progressLabel: $("progressLabel"),
  progressFill: $("progressFill"),
  progressBar: document.querySelector(".progress-bar"),
  output: $("output"),
  summaryOutput: $("summaryOutput"),
  btnCopy: $("btnCopy"),
  btnSaveFull: $("btnSaveFull"),
  btnCopySummary: $("btnCopySummary"),
  btnSaveSummary: $("btnSaveSummary"),
  btnRefreshSummary: $("btnRefreshSummary"),
  summaryCoreOutput: $("summaryCoreOutput"),
  btnCopySummaryCore: $("btnCopySummaryCore"),
  btnRefreshSummaryCore: $("btnRefreshSummaryCore"),
  btnSaveSummaryCore: $("btnSaveSummaryCore"),
  error: $("error"),
  themeDark: $("themeDark"),
  themeLight: $("themeLight"),
  saveDialog: $("saveDialog"),
  saveDialogTitle: $("saveDialogTitle"),
  saveDialogNote: $("saveDialogNote"),
  saveDialogFileName: $("saveDialogFileName"),
  saveDialogCancel: $("saveDialogCancel"),
  saveDialogConfirm: $("saveDialogConfirm"),
};

const THEME_STORAGE_KEY = "stt-theme";
/** 1차: 변환 전체 → 핵심 줄 수 */
const SUMMARY_PRIMARY_MAX = 6;
/** 2차: 핵심 요약 → 더 짧게 */
const SUMMARY_CORE_MAX = 3;

let selectedFile = null;
/** @type {Worker | null} */
let transcribeWorker = null;
let micRecognition = null;
/** 마이크 실시간: 확정 슬롯은 한 번만 이어 붙임(results 재스캔 시 깜빡임·중복 완화) */
let micCommittedRaw = "";
let micLastFinalIdx = -1;

/** @type {MediaStream | null} */
let recordStream = null;
/** @type {MediaRecorder | null} */
let recordMediaRecorder = null;
/** @type {BlobPart[]} */
let recordChunks = [];
/** 탭 전환 등으로 녹음을 버릴 때 onstop에서 변환하지 않음 */
let recordDiscardOnStop = false;

/** @type {Blob | null} */
let lastRecordedAudioBlob = null;
let lastRecordedAudioExt = "webm";

/** 마지막으로 선택된 입력 탭(`mic` | `record` | `file`) — 탭 전환 시에만 변환 결과를 비웁니다 */
let activeInputTab = "file";

function showError(msg) {
  els.error.textContent = msg;
  els.error.hidden = !msg;
}

/** 1·2차 저장 버튼과 동일 조건으로 복사 가능 여부를 맞춤 (textarea 값과 어긋남 방지) */
function syncSummaryCopyButtons() {
  if (els.btnCopySummary && els.btnSaveSummary) {
    els.btnCopySummary.disabled = els.btnSaveSummary.disabled;
  }
  if (els.btnCopySummaryCore && els.btnSaveSummaryCore) {
    els.btnCopySummaryCore.disabled = els.btnSaveSummaryCore.disabled;
  }
}

function syncSummaryCopyButtonsSoon() {
  syncSummaryCopyButtons();
  queueMicrotask(() => syncSummaryCopyButtons());
}

/**
 * @param {string} text
 * @param {HTMLButtonElement | null | undefined} btn
 * @param {HTMLTextAreaElement | null | undefined} textareaFallback
 */
async function copyTextWithFeedback(text, btn, textareaFallback) {
  const t = typeof text === "string" ? text : "";
  if (!t.trim() || !btn) return;
  const orig = btn.textContent;
  try {
    await navigator.clipboard.writeText(t);
    btn.textContent = "복사됨";
    setTimeout(() => {
      btn.textContent = orig;
    }, 1500);
  } catch {
    if (textareaFallback) {
      try {
        textareaFallback.focus();
        textareaFallback.select();
        document.execCommand("copy");
      } catch {
        /* ignore */
      }
    }
    btn.textContent = "복사됨";
    setTimeout(() => {
      btn.textContent = orig;
    }, 1500);
  }
}

function setProgress(visible, label = "", percent = 0) {
  els.progressWrap.hidden = !visible;
  els.progressLabel.textContent = label;
  const pct = Math.min(100, Math.max(0, percent));
  els.progressFill.style.width = `${pct}%`;
  if (els.progressBar) {
    els.progressBar.setAttribute("aria-valuenow", String(Math.round(pct)));
  }
}

function setOutput(text, opts = {}) {
  const skipSummary = opts.skipSummary === true;
  const full = typeof text === "string" ? text : "";
  els.output.value = full;
  const has = full.trim().length > 0;
  els.btnCopy.disabled = !has;
  els.btnSaveFull.disabled = !has;
  els.btnRefreshSummary.disabled = !has;
  if (!has) {
    els.summaryOutput.value = "";
    els.summaryCoreOutput.value = "";
    els.btnSaveSummary.disabled = true;
    els.btnSaveSummaryCore.disabled = true;
    els.btnRefreshSummaryCore.disabled = true;
    syncSummaryCopyButtons();
  } else if (!skipSummary) {
    els.btnSaveSummary.disabled = true;
    scheduleSummaryGeneration(full);
  } else {
    els.btnSaveSummary.disabled = !els.summaryOutput.value.trim();
    const hasCore = els.summaryCoreOutput.value.trim().length > 0;
    els.btnSaveSummaryCore.disabled = !hasCore;
    els.btnRefreshSummaryCore.disabled = !els.summaryOutput.value.trim();
    syncSummaryCopyButtonsSoon();
  }
}

/** 변환 결과를 직접 고칠 때: 상단 버튼 상태를 맞추고, 본문이 비면 요약 영역도 비웁니다. */
function syncUiAfterOutputEdit() {
  const has = els.output.value.trim().length > 0;
  els.btnCopy.disabled = !has;
  els.btnSaveFull.disabled = !has;
  els.btnRefreshSummary.disabled = !has;
  if (!has) {
    els.summaryOutput.value = "";
    clearSummaryCoreOutput();
    els.btnSaveSummary.disabled = true;
    els.btnRefreshSummaryCore.disabled = true;
    syncSummaryCopyButtons();
  }
}

function tokenizeForKeywords(text) {
  return text
    .toLowerCase()
    .replace(/[^a-z0-9가-힣\s]/gi, " ")
    .split(/\s+/)
    .filter((w) => w.length >= 2);
}

function dedupeRough(sentences) {
  const seen = new Set();
  const out = [];
  for (const s of sentences) {
    const key = s.slice(0, 48);
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(s);
  }
  return out;
}

function splitIntoSentences(text) {
  const chunks = text
    .split(/(?<=[.!?…]|다\.|니다\.|요\.|죠\.|까\?|입니까\.)\s+/)
    .map((x) => x.trim())
    .filter((x) => x.length >= 12);
  const merged = chunks.length
    ? chunks
    : text
        .split(/\n+/)
        .map((x) => x.trim())
        .filter((x) => x.length >= 8);
  return dedupeRough(merged);
}

/** 키워드 빈도·문장 길이·앞쪽 가중으로 상위 문장을 뽑는 추출식 요약(서버 없음). */
function buildExtractiveSummary(text, maxBullets = 6) {
  const t = text.replace(/\s+/g, " ").trim();
  if (!t) return "";

  const tokens = tokenizeForKeywords(t);
  if (tokens.length === 0) {
    const one = t.slice(0, 400);
    return one.length < t.length ? `${one}…` : one;
  }

  const freq = new Map();
  for (const w of tokens) {
    freq.set(w, (freq.get(w) || 0) + 1);
  }
  const important = new Set(
    [...freq.entries()]
      .filter(([, c]) => c >= 2)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 28)
      .map(([w]) => w)
  );

  const sentences = splitIntoSentences(t);
  if (sentences.length === 0) {
    const one = t.slice(0, 400);
    return one.length < t.length ? `${one}…` : one;
  }
  if (sentences.length <= maxBullets) {
    return sentences.map((s, i) => `${i + 1}. ${s}`).join("\n");
  }

  const scored = sentences.map((s, idx) => {
    const st = tokenizeForKeywords(s);
    let score = 0;
    for (const w of st) {
      if (important.has(w)) score += freq.get(w) || 0;
    }
    score += Math.min(s.length / 100, 0.4);
    score += (1 - idx / Math.max(1, sentences.length - 1)) * 0.2;
    return { s, idx, score };
  });

  scored.sort((a, b) => b.score - a.score);
  const picked = scored.slice(0, maxBullets).sort((a, b) => a.idx - b.idx);
  return picked.map(({ s }, i) => `${i + 1}. ${s}`).join("\n");
}

function clearSummaryCoreOutput() {
  els.summaryCoreOutput.value = "";
  els.btnSaveSummaryCore.disabled = true;
  syncSummaryCopyButtons();
}

/** 텍스트 클리어가 화면에 반영된 뒤 이어서 작업하기 위한 짧은 대기 */
function flushPaint() {
  return new Promise((resolve) => {
    requestAnimationFrame(() => {
      requestAnimationFrame(() => resolve());
    });
  });
}

/** 변환 전체 → 1차 요약 → 2차 요약(1차 기준) */
function applyPrimaryAndCoreSummary(fullText) {
  const primary = buildExtractiveSummary(fullText, SUMMARY_PRIMARY_MAX).trim();
  els.summaryOutput.value = primary;
  els.btnSaveSummary.disabled = !primary;
  if (!primary) {
    clearSummaryCoreOutput();
    els.btnRefreshSummaryCore.disabled = true;
    syncSummaryCopyButtonsSoon();
    return;
  }
  setProgress(true, "2차 요약(핵심의 핵심)…", 72);
  const core = buildExtractiveSummary(primary, SUMMARY_CORE_MAX).trim();
  els.summaryCoreOutput.value = core;
  els.btnSaveSummaryCore.disabled = !core;
  els.btnRefreshSummaryCore.disabled = false;
  syncSummaryCopyButtonsSoon();
}

function scheduleSummaryGeneration(fullText) {
  try {
    setProgress(true, "1차 핵심 요약…", 35);
    applyPrimaryAndCoreSummary(fullText);
  } catch (err) {
    console.error(err);
    showError("요약에 실패했습니다.");
    els.summaryOutput.value = "";
    clearSummaryCoreOutput();
    els.btnSaveSummary.disabled = true;
    els.btnRefreshSummaryCore.disabled = true;
    syncSummaryCopyButtonsSoon();
  } finally {
    setTimeout(() => setProgress(false), 250);
  }
}

async function refreshSummaryFromOutput() {
  const full = els.output.value.trim();
  if (!full) return;
  els.btnRefreshSummary.disabled = true;
  showError("");
  els.summaryOutput.value = "";
  clearSummaryCoreOutput();
  els.btnSaveSummary.disabled = true;
  els.btnRefreshSummaryCore.disabled = true;
  syncSummaryCopyButtons();
  await flushPaint();
  try {
    setProgress(true, "1차 핵심 요약…", 35);
    applyPrimaryAndCoreSummary(full);
  } catch (err) {
    console.error(err);
    showError("요약에 실패했습니다.");
    syncSummaryCopyButtonsSoon();
  } finally {
    els.btnRefreshSummary.disabled = false;
    setTimeout(() => setProgress(false), 250);
  }
}

/** 현재 1차 핵심 요약만으로 2차만 다시 계산 */
async function refreshSummaryCoreFromSummary() {
  const primary = els.summaryOutput.value.trim();
  if (!primary) return;
  els.btnRefreshSummaryCore.disabled = true;
  showError("");
  els.summaryCoreOutput.value = "";
  els.btnSaveSummaryCore.disabled = true;
  syncSummaryCopyButtons();
  await flushPaint();
  try {
    setProgress(true, "2차 요약 다시…", 55);
    const core = buildExtractiveSummary(primary, SUMMARY_CORE_MAX).trim();
    els.summaryCoreOutput.value = core;
    els.btnSaveSummaryCore.disabled = !core;
    syncSummaryCopyButtonsSoon();
  } catch (err) {
    console.error(err);
    showError("2차 요약에 실패했습니다.");
    syncSummaryCopyButtonsSoon();
  } finally {
    els.btnRefreshSummaryCore.disabled = false;
    setTimeout(() => setProgress(false), 250);
  }
}

function saveTextFile(filename, content) {
  const blob = new Blob([content], { type: "text/plain;charset=utf-8" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.rel = "noopener";
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(a.href);
}

function saveBlobAsDownload(filename, blob) {
  const a = document.createElement("a");
  const url = URL.createObjectURL(blob);
  a.href = url;
  a.download = filename;
  a.rel = "noopener";
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60_000);
}

function extFromAudioMime(mime) {
  const m = String(mime || "").toLowerCase();
  if (m.includes("webm")) return "webm";
  if (m.includes("ogg")) return "ogg";
  if (m.includes("mp4") || m.includes("mpeg")) return "mp4";
  if (m.includes("wav")) return "wav";
  return "webm";
}

/** 녹음 등 바이너리 파일명 (.mp3·.webm 등) */
function sanitizeAudioDownloadFileName(name, extFallback) {
  const ext = extFallback || "webm";
  let s = basenameFromPath(String(name || "").trim());
  if (!s) s = `음성녹음_${defaultDownloadBasename()}.${ext}`;
  s = s.replace(/[\\/:*?"<>|]/g, "_").replace(/\s+/g, " ").trim();
  const low = s.toLowerCase();
  const ok = [".mp3", ".wav", ".webm", ".ogg", ".opus", ".mp4", ".m4a"];
  if (!ok.some((a) => low.endsWith(a))) {
    const base = s.replace(/\.[^/.]+$/, "");
    s = `${base || `음성녹음_${defaultDownloadBasename()}`}.${ext}`;
  }
  if (s.length > 180) s = `${s.slice(0, 172)}.${ext}`;
  return s;
}

/**
 * @param {string | undefined} extOverride 저장 파일 확장자(예: mp3). 없으면 `lastRecordedAudioExt`.
 * @returns {Promise<"saved"|"aborted"|"noapi"|"failed">}
 */
async function trySaveBlobWithNativePicker(blob, suggestedName, extOverride) {
  const picker = window.showSaveFilePicker;
  if (typeof picker !== "function") return "noapi";
  const ext = extOverride ?? lastRecordedAudioExt ?? "mp3";
  try {
    const handle = await picker.call(window, {
      suggestedName: sanitizeAudioDownloadFileName(suggestedName, ext),
      types: [
        {
          description: "MP3",
          accept: {
            "audio/mpeg": [".mp3"],
          },
        },
        {
          description: "기타 오디오",
          accept: {
            "audio/wav": [".wav"],
            "audio/webm": [".webm"],
            "audio/ogg": [".ogg", ".opus"],
            "audio/mp4": [".m4a", ".mp4"],
          },
        },
      ],
    });
    const writable = await handle.createWritable();
    await writable.write(blob);
    await writable.close();
    return "saved";
  } catch (e) {
    if (e?.name === "AbortError") return "aborted";
    console.warn(e);
    return "failed";
  }
}

/** 성공 시에만 캐시 — 한 번 import 실패 시 Promise가 영구 reject 되는 문제 방지 */
let lamejsResolvedModule = null;

async function loadLameJsModule() {
  if (lamejsResolvedModule) return lamejsResolvedModule;
  const localCandidates = [new URL("./vendor/lamejs.js", import.meta.url)];
  for (const url of localCandidates) {
    try {
      lamejsResolvedModule = await import(url.href);
      return lamejsResolvedModule;
    } catch (err) {
      console.warn("lamejs 로컬 로드 실패:", url.href, err);
    }
  }
  try {
    lamejsResolvedModule = await import(
      /* webpackIgnore: true */
      "https://cdn.jsdelivr.net/npm/lamejs@1.2.1/+esm"
    );
    return lamejsResolvedModule;
  } catch (err) {
    console.warn("lamejs jsDelivr 실패:", err);
  }
  lamejsResolvedModule = await import(
    /* webpackIgnore: true */
    "https://esm.sh/lamejs@1.2.1"
  );
  return lamejsResolvedModule;
}

/** MP3 인코딩용 샘플레이트(lamejs 권장) */
const MP3_ENCODE_SAMPLE_RATE = 44100;

function floatTo16BitPcmMono(float32) {
  const n = float32.length;
  const out = new Int16Array(n);
  for (let i = 0; i < n; i += 1) {
    const x = Math.max(-1, Math.min(1, float32[i]));
    out[i] = x < 0 ? x * 0x8000 : x * 0x7fff;
  }
  return out;
}

/**
 * 녹음 Blob(WebM 등) → 지정 샘플레이트 모노 Float32
 * @param {Blob} blob
 * @param {number} sampleRate
 */
async function decodeRecordingBlobToMonoPcm(blob, sampleRate) {
  const AC = window.AudioContext || window.webkitAudioContext;
  if (!AC) throw new Error("AudioContext를 사용할 수 없습니다.");
  const ctx = new AC();
  try {
    await ctx.resume().catch(() => {});
    const ab = await blob.arrayBuffer();
    const decoded = await ctx.decodeAudioData(ab.slice(0));
    const len = decoded.length;
    const nCh = decoded.numberOfChannels;
    const monoAtSource = new Float32Array(len);
    if (nCh === 1) {
      monoAtSource.set(decoded.getChannelData(0));
    } else {
      for (let i = 0; i < len; i += 1) {
        let s = 0;
        for (let c = 0; c < nCh; c += 1) s += decoded.getChannelData(c)[i];
        monoAtSource[i] = s / nCh;
      }
    }
    const monoBuf = ctx.createBuffer(1, len, decoded.sampleRate);
    monoBuf.copyToChannel(monoAtSource, 0);
    const duration = decoded.duration;
    const outFrames = Math.max(1, Math.ceil(duration * sampleRate));
    const offline = new OfflineAudioContext(1, outFrames, sampleRate);
    const src = offline.createBufferSource();
    src.buffer = monoBuf;
    src.connect(offline.destination);
    src.start(0);
    const rendered = await offline.startRendering();
    return rendered.getChannelData(0);
  } finally {
    await ctx.close();
  }
}

/**
 * MediaRecorder 원본 Blob을 브라우저에서 MP3로 인코딩합니다.
 * @param {Blob} blob
 */
async function recordingBlobToMp3Blob(blob) {
  const mod = await loadLameJsModule();
  const Mp3Encoder =
    mod.Mp3Encoder ||
    (mod.default && mod.default.Mp3Encoder) ||
    (typeof mod.default === "function" ? mod.default : null);
  if (typeof Mp3Encoder !== "function") {
    throw new Error("MP3 인코더(lamejs)를 불러오지 못했습니다.");
  }

  const pcmF = await decodeRecordingBlobToMonoPcm(blob, MP3_ENCODE_SAMPLE_RATE);
  if (!pcmF.length) throw new Error("디코딩된 오디오가 비어 있습니다.");

  const pcm16 = floatTo16BitPcmMono(pcmF);
  const enc = new Mp3Encoder(1, MP3_ENCODE_SAMPLE_RATE, 128);
  const block = 1152;
  const parts = [];
  let i = 0;
  for (; i + block <= pcm16.length; i += block) {
    const buf = enc.encodeBuffer(pcm16.subarray(i, i + block));
    if (buf?.length > 0) parts.push(new Uint8Array(buf));
  }
  if (i < pcm16.length) {
    let rest = pcm16.subarray(i);
    if (rest.length > 0 && rest.length < block) {
      const padded = new Int16Array(block);
      padded.set(rest);
      rest = padded;
    }
    if (rest.length > 0) {
      const buf = enc.encodeBuffer(rest);
      if (buf?.length > 0) parts.push(new Uint8Array(buf));
    }
  }
  const end = enc.flush();
  if (end?.length > 0) parts.push(new Uint8Array(end));
  if (!parts.length) {
    throw new Error("MP3 인코딩 결과가 비어 있습니다.");
  }
  return new Blob(parts, { type: "audio/mpeg" });
}

async function saveLastRecordedAudio() {
  if (!lastRecordedAudioBlob || !lastRecordedAudioBlob.size) return;
  if (els.btnSaveRecording) els.btnSaveRecording.disabled = true;
  showError("");
  if (els.recordStatus) {
    els.recordStatus.textContent = "MP3로 변환하는 중… 잠시만 기다려 주세요.";
  }
  try {
    const mp3Blob = await recordingBlobToMp3Blob(lastRecordedAudioBlob);
    if (!mp3Blob?.size) {
      throw new Error("MP3 변환 결과가 비어 있습니다.");
    }
    const suggested = `음성녹음_${defaultDownloadBasename()}.mp3`;
    const native = await trySaveBlobWithNativePicker(mp3Blob, suggested, "mp3");
    if (native === "saved" || native === "aborted") {
      if (els.recordStatus && native === "saved") {
        els.recordStatus.textContent = "MP3 파일을 저장했습니다.";
      }
      return;
    }
    saveBlobAsDownload(
      sanitizeAudioDownloadFileName(suggested, "mp3"),
      mp3Blob
    );
    if (els.recordStatus) {
      els.recordStatus.textContent = "MP3 파일을 다운로드했습니다.";
    }
  } catch (e) {
    console.error(e);
    if (els.recordStatus) {
      els.recordStatus.textContent =
        "MP3로 저장할 수 없어 원본 녹음 형식으로 저장을 시도합니다.";
    }
    try {
      const ext = lastRecordedAudioExt || "webm";
      const suggested = `음성녹음_${defaultDownloadBasename()}.${ext}`;
      const native = await trySaveBlobWithNativePicker(
        lastRecordedAudioBlob,
        suggested,
        ext
      );
      if (native === "saved") {
        showError("");
        if (els.recordStatus) {
          els.recordStatus.textContent = `MP3는 생략하고 원본(.${ext})으로 저장했습니다.`;
        }
        return;
      }
      if (native === "aborted") {
        showError("");
        if (els.recordStatus) els.recordStatus.textContent = "저장을 취소했습니다.";
        return;
      }
      saveBlobAsDownload(
        sanitizeAudioDownloadFileName(suggested, ext),
        lastRecordedAudioBlob
      );
      showError("");
      if (els.recordStatus) {
        els.recordStatus.textContent = `MP3는 생략하고 원본(.${ext})으로 다운로드했습니다.`;
      }
    } catch (e2) {
      console.error(e2);
      showError(
        e2?.message
          ? `파일 저장 실패: ${e2.message}`
          : "파일을 저장하지 못했습니다."
      );
      if (els.recordStatus) {
        els.recordStatus.textContent = "저장에 실패했습니다.";
      }
    }
  } finally {
    if (els.btnSaveRecording) els.btnSaveRecording.disabled = false;
  }
}

function rememberLastRecordedAudio(blob, mime) {
  lastRecordedAudioBlob = blob;
  lastRecordedAudioExt = extFromAudioMime(mime);
  if (els.btnSaveRecording) els.btnSaveRecording.disabled = false;
  if (els.btnRecordTranscribe) els.btnRecordTranscribe.disabled = false;
}

function clearLastRecordedAudio() {
  lastRecordedAudioBlob = null;
  lastRecordedAudioExt = "webm";
  if (els.btnSaveRecording) els.btnSaveRecording.disabled = true;
  if (els.btnRecordTranscribe) els.btnRecordTranscribe.disabled = true;
}

function basenameFromPath(name) {
  return String(name || "")
    .trim()
    .replace(/^.*[/\\]/, "");
}

/** 사용자 입력·경로에서 파일명만 안전하게 .txt 로 정리 */
function sanitizeDownloadFileName(name) {
  let s = basenameFromPath(name);
  if (!s) return `download_${defaultDownloadBasename()}.txt`;
  s = s.replace(/[\\/:*?"<>|]/g, "_").replace(/\s+/g, " ").trim();
  if (!s.toLowerCase().endsWith(".txt")) s += ".txt";
  if (s.length > 180) s = `${s.slice(0, 176)}.txt`;
  return s;
}

/**
 * Chrome/Edge 등: 저장 대화상자에서 폴더·파일명 선택(File System Access API).
 * @returns {Promise<"saved"|"aborted"|"noapi"|"failed">}
 */
async function trySaveWithNativePicker(content, suggestedName) {
  const picker = window.showSaveFilePicker;
  if (typeof picker !== "function") return "noapi";
  try {
    const handle = await picker.call(window, {
      suggestedName: sanitizeDownloadFileName(suggestedName),
      types: [
        {
          description: "텍스트 파일",
          accept: { "text/plain": [".txt"] },
        },
      ],
    });
    const writable = await handle.createWritable();
    await writable.write(
      new Blob([content], { type: "text/plain;charset=utf-8" })
    );
    await writable.close();
    return "saved";
  } catch (e) {
    if (e?.name === "AbortError") return "aborted";
    console.warn(e);
    return "failed";
  }
}

/** 네이티브 저장 미지원 시 모달로 파일명만 받고 다운로드 */
let saveDialogPending = { content: "", suggested: "" };

async function saveTextWithLocationAndName(content, suggestedStem, dialogTitle) {
  const suggested = `${suggestedStem}_${defaultDownloadBasename()}.txt`;
  const native = await trySaveWithNativePicker(content, suggested);
  if (native === "saved" || native === "aborted") return;

  saveDialogPending = { content, suggested: sanitizeDownloadFileName(suggested) };
  if (els.saveDialogTitle) els.saveDialogTitle.textContent = dialogTitle;
  if (els.saveDialogNote) {
    els.saveDialogNote.textContent =
      "이 브라우저는 저장 폴더를 고르는 창을 제공하지 않습니다. 파일명을 확인한 뒤 저장하면 보통 「다운로드」폴더에 저장됩니다. 폴더와 이름을 한 번에 고르려면 Chrome 또는 Edge에서 https로 열어 주세요.";
  }
  if (els.saveDialogFileName) els.saveDialogFileName.value = saveDialogPending.suggested;
  els.saveDialog?.showModal();
  queueMicrotask(() => els.saveDialogFileName?.select());
}

function setupSaveDialog() {
  els.saveDialogCancel?.addEventListener("click", () => {
    els.saveDialog?.close();
    saveDialogPending = { content: "", suggested: "" };
  });
  els.saveDialogConfirm?.addEventListener("click", () => {
    const body = saveDialogPending.content;
    if (!body) {
      els.saveDialog?.close();
      return;
    }
    const name = sanitizeDownloadFileName(els.saveDialogFileName?.value || "");
    saveTextFile(name, body);
    els.saveDialog?.close();
    saveDialogPending = { content: "", suggested: "" };
  });
  els.saveDialogFileName?.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      els.saveDialogConfirm?.click();
    }
  });
}

function defaultDownloadBasename() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
}

/**
 * Whisper가 침묵·말끝 뒤에서 같은 짧은 구절을 수십 번 반복하는 환각을 줄입니다.
 * (짧은 구문 + 공백으로 구분된 동일 구문이 5회 이상 이어질 때 1회로 압축)
 * @param {string} text
 */
function collapseAdjacentPhraseRepeats(text) {
  let s = String(text || "").replace(/\s+/g, " ").trim();
  if (s.length < 24) return s;
  for (let iter = 0; iter < 100; iter += 1) {
    const re = /(.{5,45}?)(\s+\1){4,}/gu;
    re.lastIndex = 0;
    const m = re.exec(s);
    if (!m) break;
    s = `${s.slice(0, m.index)}${m[1]}${s.slice(m.index + m[0].length)}`
      .replace(/\s+/g, " ")
      .trim();
  }
  return s;
}

/**
 * Whisper 파이프라인 반환값을 한 덩어리 문자열로 정리합니다.
 * 긴 오디오는 chunks만 채이거나 text와 함께 올 수 있어, 짧은 쪽만 쓰면 누락이 납니다.
 */
function normalizeAsrResult(result) {
  if (result == null) return "";
  if (typeof result === "string") {
    return collapseAdjacentPhraseRepeats(result.trim());
  }
  if (Array.isArray(result)) {
    return collapseAdjacentPhraseRepeats(
      result.map(normalizeAsrResult).filter(Boolean).join(" ").trim()
    );
  }
  const top = typeof result.text === "string" ? result.text.trim() : "";
  const chunks = result.chunks;
  if (Array.isArray(chunks) && chunks.length) {
    const fromChunks = chunks
      .map((c) => (typeof c === "string" ? c : (c && c.text) || ""))
      .map((t) => t.replace(/\s+/g, " ").trim())
      .filter(Boolean)
      .join(" ")
      .trim();
    if (fromChunks && top) {
      return collapseAdjacentPhraseRepeats(
        fromChunks.length >= top.length ? fromChunks : top
      );
    }
    if (fromChunks) return collapseAdjacentPhraseRepeats(fromChunks);
  }
  return collapseAdjacentPhraseRepeats(top);
}

function getSelectedModelId() {
  const key = els.modelSelect?.value || "base";
  return WHISPER_MODELS[key] ? key : "base";
}

/**
 * 짧은 클립은 한 번에 디코딩하는 편이 덜 깨지고, 긴 파일은 청크가 필요합니다.
 * @param {number | null} durationSec 실제 PCM 기준 길이(초) 권장
 */
function getTranscribeOptions(durationSec) {
  const speedPreset = els.speedSelect?.value || "balanced";
  /** @type {Record<string, unknown>} */
  const opts = {
    task: "transcribe",
    language: "korean",
    // 침묵·저신뢰 구간 반복 환각 완화(transformers가 무시하면 영향 없음)
    no_speech_threshold: 0.68,
    compression_ratio_threshold: 2.0,
    logprob_threshold: -0.85,
  };
  const longForm =
    durationSec == null || !Number.isFinite(durationSec) || durationSec > 28;
  if (longForm) {
    // 30s 청크는 일부 환경에서 타임스탬프/경계 이슈 보고 → 29s 권장에 맞춤
    opts.chunk_length_s = 29;
    // stride_length_s = 청크 간 겹침(초). 클수록 경계에서 단어 누락·반복이 줄고 연산은 늘어남.
    if (speedPreset === "fast") opts.stride_length_s = 2;
    else if (speedPreset === "quality") opts.stride_length_s = 10;
    else opts.stride_length_s = 6;
  }
  return opts;
}

const WHISPER_SAMPLE_RATE = 16000;

/**
 * Worker에는 AudioContext가 없으므로, 여기(메인)에서 파일을 PCM으로 디코딩합니다.
 * Whisper는 16kHz 모노 Float32Array 입력을 기대합니다.
 * @param {ArrayBuffer} arrayBuffer
 * @returns {Promise<{ samples: Float32Array, durationSec: number }>}
 */
async function decodeFileToWhisperSamples(arrayBuffer) {
  const AC = window.AudioContext || window.webkitAudioContext;
  if (!AC) {
    throw new Error("이 브라우저에서 AudioContext를 사용할 수 없습니다.");
  }
  const ctx = new AC();
  try {
    const decoded = await ctx.decodeAudioData(arrayBuffer.slice(0));

    const nCh = decoded.numberOfChannels;
    const len = decoded.length;
    const monoAtSource = new Float32Array(len);
    if (nCh === 1) {
      monoAtSource.set(decoded.getChannelData(0));
    } else {
      for (let i = 0; i < len; i++) {
        let s = 0;
        for (let c = 0; c < nCh; c++) s += decoded.getChannelData(c)[i];
        monoAtSource[i] = s / nCh;
      }
    }

    const monoBuf = ctx.createBuffer(1, len, decoded.sampleRate);
    monoBuf.copyToChannel(monoAtSource, 0);

    const duration = decoded.duration;
    const outFrames = Math.max(1, Math.ceil(duration * WHISPER_SAMPLE_RATE));
    const offline = new OfflineAudioContext(1, outFrames, WHISPER_SAMPLE_RATE);
    const src = offline.createBufferSource();
    src.buffer = monoBuf;
    src.connect(offline.destination);
    src.start(0);
    const rendered = await offline.startRendering();
    const ch0 = rendered.getChannelData(0);
    const core = new Float32Array(ch0);
    const durationSec =
      Number.isFinite(duration) && duration > 0
        ? duration
        : core.length / WHISPER_SAMPLE_RATE;
    // 마지막 음절/어미가 잘리는 현상 완화(무음 꼬리). 청크 길이 판단은 원본 durationSec 기준.
    const tailPadSec = 0.35;
    const pad = Math.round(tailPadSec * WHISPER_SAMPLE_RATE);
    const samples = new Float32Array(core.length + pad);
    samples.set(core);
    return { samples, durationSec };
  } finally {
    await ctx.close();
  }
}

function getTranscribeWorker() {
  if (!transcribeWorker) {
    try {
      transcribeWorker = new Worker(
        new URL("./transcribe-worker.js", import.meta.url),
        { type: "module" }
      );
    } catch (e) {
      console.error(e);
      throw new Error(
        "백그라운드 Worker를 만들 수 없습니다. 페이지를 http(s)로 여는지 확인해 주세요."
      );
    }
    transcribeWorker.addEventListener("error", (e) => {
      console.error(e);
      showError(
        "음성 변환 모듈 오류가 발생했습니다. 페이지를 새로고침한 뒤 다시 시도해 주세요."
      );
    });
  }
  return transcribeWorker;
}

/**
 * Worker에서 모델 로드 + 추론(메인 스레드 블로킹 없음)
 * @param {File} file
 */
function transcribeWithWorker(file) {
  const worker = getTranscribeWorker();
  const modelKey = getSelectedModelId();
  const modelId = WHISPER_MODELS[modelKey];
  const jobId = crypto.randomUUID();

  return new Promise((resolve, reject) => {
    const onMessage = (ev) => {
      const d = ev.data;
      if (!d) return;

      if (d.type === "loadProgress") {
        if (d.status === "progress" && d.file != null) {
          const pct = Math.round(d.progress || 0);
          setProgress(true, `모델 로딩… ${d.file}`, pct);
        } else if (d.status === "done") {
          setProgress(true, "모델 준비 완료", 100);
        }
        return;
      }
      if (d.jobId !== jobId) return;

      if (d.type === "inferStart") {
        setProgress(
          true,
          "음성 분석 중…(백그라운드에서 처리 중이니 탭을 닫지 마세요)",
          45
        );
        return;
      }
      if (d.type === "inferDone") {
        worker.removeEventListener("message", onMessage);
        resolve(d.result);
        return;
      }
      if (d.type === "error") {
        worker.removeEventListener("message", onMessage);
        reject(new Error(d.message || "Worker 오류"));
      }
    };

    worker.addEventListener("message", onMessage);

    (async () => {
      try {
        const arrayBuffer = await file.arrayBuffer();
        setProgress(true, "오디오를 PCM(16kHz)으로 변환 중…", 18);
        const { samples, durationSec } = await decodeFileToWhisperSamples(
          arrayBuffer
        );
        const options = getTranscribeOptions(durationSec);
        worker.postMessage(
          {
            type: "transcribe",
            jobId,
            modelKey,
            modelId,
            options,
            raw: samples,
          },
          [samples.buffer]
        );
      } catch (err) {
        worker.removeEventListener("message", onMessage);
        reject(err);
      }
    })();
  });
}

async function transcribeFile(file) {
  showError("");
  setOutput("", { skipSummary: true });
  setProgress(true, "오디오·모델 준비 중…", 10);
  try {
    const result = await transcribeWithWorker(file);
    setProgress(true, "완료", 100);
    setOutput(normalizeAsrResult(result));
  } finally {
    setTimeout(() => setProgress(false), 400);
  }
}

function stopMicRecognitionIfAny() {
  if (micRecognition) {
    try {
      micRecognition.stop();
    } catch {
      /* ignore */
    }
  }
}

function stopRecordStreamTracks() {
  if (recordStream) {
    try {
      recordStream.getTracks().forEach((t) => t.stop());
    } catch {
      /* ignore */
    }
    recordStream = null;
  }
}

/** 녹음 UI만 초기화(스트림은 이미 정리된 뒤 호출 가능) */
function resetRecordUiIdle() {
  recordMediaRecorder = null;
  recordChunks = [];
  if (els.btnRecordStart) els.btnRecordStart.disabled = false;
  if (els.btnRecordStop) els.btnRecordStop.disabled = true;
}

/**
 * 녹음 중이면 중단하고 변환은 하지 않습니다(다른 탭으로 나갈 때).
 */
function abortRecordSession() {
  if (recordMediaRecorder && recordMediaRecorder.state === "recording") {
    recordDiscardOnStop = true;
    try {
      recordMediaRecorder.stop();
    } catch {
      recordDiscardOnStop = false;
      stopRecordStreamTracks();
      resetRecordUiIdle();
    }
  } else {
    stopRecordStreamTracks();
    resetRecordUiIdle();
  }
}

function setupTabs() {
  els.tabs.forEach((tab) => {
    tab.addEventListener("click", () => {
      const name = tab.dataset.tab;
      if (
        name &&
        (name === "mic" || name === "record" || name === "file") &&
        name !== activeInputTab
      ) {
        setOutput("", { skipSummary: true });
        activeInputTab = name;
      }
      els.tabs.forEach((t) => {
        const on = t === tab;
        t.classList.toggle("is-active", on);
        t.setAttribute("aria-selected", on);
      });
      const fileMode = name === "file";
      const recordMode = name === "record";
      const micMode = name === "mic";

      if (!recordMode) {
        abortRecordSession();
      }
      if (!micMode) {
        stopMicRecognitionIfAny();
      }

      els.panelFile.classList.toggle("is-visible", fileMode);
      els.panelFile.hidden = !fileMode;
      els.panelRecord.classList.toggle("is-visible", recordMode);
      els.panelRecord.hidden = !recordMode;
      els.panelMic.classList.toggle("is-visible", micMode);
      els.panelMic.hidden = !micMode;
    });
  });
}

function setupFileUi() {
  const setFile = (file) => {
    if (
      !file ||
      (!file.type.startsWith("audio/") &&
        !/\.(webm|mp3|wav|ogg|m4a|flac)$/i.test(file.name))
    ) {
      showError("오디오 파일을 선택해 주세요.");
      return;
    }
    setOutput("", { skipSummary: true });
    selectedFile = file;
    els.fileName.textContent = `선택됨: ${file.name}`;
    els.fileName.hidden = false;
    els.btnTranscribeFile.disabled = false;
    showError("");
  };

  /* label이 파일 입력과 연결되어 있어 추가 click()을 넣으면 파일 선택창이 두 번 뜹니다. */
  els.fileInput.addEventListener("change", () => {
    const f = els.fileInput.files?.[0];
    if (f) setFile(f);
  });

  ["dragenter", "dragover"].forEach((ev) => {
    els.dropzone.addEventListener(ev, (e) => {
      e.preventDefault();
      els.dropzone.classList.add("is-dragover");
    });
  });
  ["dragleave", "drop"].forEach((ev) => {
    els.dropzone.addEventListener(ev, () => {
      els.dropzone.classList.remove("is-dragover");
    });
  });
  els.dropzone.addEventListener("drop", (e) => {
    e.preventDefault();
    const f = e.dataTransfer?.files?.[0];
    if (f) setFile(f);
  });

  els.btnTranscribeFile.addEventListener("click", async () => {
    if (!selectedFile) return;
    els.btnTranscribeFile.disabled = true;
    try {
      await transcribeFile(selectedFile);
    } catch (err) {
      console.error(err);
      showError(
        err?.message
          ? `변환 실패: ${err.message}`
          : "변환에 실패했습니다. 네트워크와 파일 형식을 확인한 뒤 다시 시도해 주세요."
      );
    } finally {
      els.btnTranscribeFile.disabled = false;
    }
  });

  els.btnClearFile.addEventListener("click", () => {
    selectedFile = null;
    els.fileInput.value = "";
    els.fileName.hidden = true;
    els.btnTranscribeFile.disabled = true;
    setOutput("", {});
    showError("");
  });
}

function setupMic() {
  const SpeechRecognition =
    window.SpeechRecognition || window.webkitSpeechRecognition;
  if (!SpeechRecognition) {
    els.btnMicStart.disabled = true;
    els.micStatus.textContent =
      "이 브라우저는 실시간 음성 인식을 지원하지 않습니다. Chrome 또는 Edge를 사용해 주세요.";
    return;
  }

  els.btnMicStart.addEventListener("click", () => {
    showError("");
    if (micRecognition) {
      try {
        micRecognition.stop();
      } catch {
        /* ignore */
      }
    }
    setOutput("", { skipSummary: true });
    micCommittedRaw = "";
    micLastFinalIdx = -1;
    micRecognition = new SpeechRecognition();
    micRecognition.lang = "ko-KR";
    micRecognition.continuous = true;
    micRecognition.interimResults = true;
    micRecognition.maxAlternatives = 1;

    micRecognition.onresult = (event) => {
      const results = event.results;
      if (!results?.length) return;

      let i = micLastFinalIdx + 1;
      while (i < results.length) {
        const r = results[i];
        if (!r?.[0]) {
          i += 1;
          continue;
        }
        if (r.isFinal) {
          micCommittedRaw += r[0].transcript || "";
          micLastFinalIdx = i;
          i += 1;
        } else {
          break;
        }
      }

      let interimRaw = "";
      for (let j = micLastFinalIdx + 1; j < results.length; j += 1) {
        const r = results[j];
        if (!r?.[0] || r.isFinal) continue;
        interimRaw += r[0].transcript || "";
      }

      const committed = micCommittedRaw.replace(/\s+/g, " ").trimEnd();
      const interim = interimRaw.replace(/\s+/g, " ").trim();
      const display =
        interim && committed
          ? `${committed} ${interim}`
          : interim || committed;

      if (els.output.value !== display) {
        setOutput(display, { skipSummary: true });
      }
    };

    micRecognition.onerror = (e) => {
      els.micStatus.textContent = `오류: ${e.error}`;
      els.btnMicStart.disabled = false;
      els.btnMicStop.disabled = true;
    };

    micRecognition.onend = () => {
      els.btnMicStart.disabled = false;
      els.btnMicStop.disabled = true;
      els.micStatus.textContent = "중지됨.";
    };

    try {
      micRecognition.start();
      els.btnMicStart.disabled = true;
      els.btnMicStop.disabled = false;
      els.micStatus.textContent = "듣는 중… 말씀해 주세요.";
    } catch (e) {
      els.micStatus.textContent = "마이크를 시작할 수 없습니다.";
      showError(e?.message || String(e));
    }
  });

  els.btnMicStop.addEventListener("click", () => {
    if (micRecognition) {
      try {
        micRecognition.stop();
      } catch {
        /* ignore */
      }
    }
  });
}

function setupRecordUi() {
  if (typeof MediaRecorder === "undefined") {
    if (els.btnRecordStart) els.btnRecordStart.disabled = true;
    if (els.btnRecordTranscribe) els.btnRecordTranscribe.disabled = true;
    if (els.recordStatus) {
      els.recordStatus.textContent =
        "이 브라우저는 MediaRecorder 녹음을 지원하지 않습니다. Chrome 또는 Edge를 사용해 주세요.";
    }
    return;
  }

  els.btnRecordStart?.addEventListener("click", async () => {
    if (recordMediaRecorder && recordMediaRecorder.state === "recording") return;
    showError("");
    recordDiscardOnStop = false;
    recordChunks = [];
    setOutput("", { skipSummary: true });
    clearLastRecordedAudio();
    try {
      recordStream = await navigator.mediaDevices.getUserMedia({
        audio: {
          channelCount: { ideal: 1 },
          echoCancellation: { ideal: true },
          noiseSuppression: { ideal: true },
          autoGainControl: { ideal: true },
        },
      });
    } catch (e) {
      console.error(e);
      showError("마이크 권한이 필요합니다. 브라우저 설정을 확인해 주세요.");
      if (els.recordStatus) els.recordStatus.textContent = "마이크를 사용할 수 없습니다.";
      return;
    }

    let recorder;
    const candidates = ["audio/webm;codecs=opus", "audio/webm"];
    for (const mimeType of candidates) {
      if (MediaRecorder.isTypeSupported(mimeType)) {
        try {
          recorder = new MediaRecorder(recordStream, {
            mimeType,
            audioBitsPerSecond: 128000,
          });
          break;
        } catch {
          try {
            recorder = new MediaRecorder(recordStream, { mimeType });
            break;
          } catch {
            /* try next mime */
          }
        }
      }
    }
    if (!recorder) {
      try {
        recorder = new MediaRecorder(recordStream, {
          audioBitsPerSecond: 128000,
        });
      } catch {
        try {
          recorder = new MediaRecorder(recordStream);
        } catch (e) {
          console.error(e);
          stopRecordStreamTracks();
          showError("녹음기를 시작할 수 없습니다.");
          if (els.recordStatus)
            els.recordStatus.textContent = "녹음을 시작할 수 없습니다.";
          return;
        }
      }
    }

    recordMediaRecorder = recorder;
    recordMediaRecorder.ondataavailable = (ev) => {
      if (ev.data && ev.data.size > 0) recordChunks.push(ev.data);
    };
    recordMediaRecorder.onerror = (ev) => {
      console.error(ev.error || ev);
      showError("녹음 중 오류가 발생했습니다.");
    };
    recordMediaRecorder.onstop = async () => {
      stopRecordStreamTracks();
      const mr = recordMediaRecorder;
      const mime = mr?.mimeType || "audio/webm";
      recordMediaRecorder = null;

      if (recordDiscardOnStop) {
        recordDiscardOnStop = false;
        recordChunks = [];
        resetRecordUiIdle();
        clearLastRecordedAudio();
        if (els.recordStatus) els.recordStatus.textContent = "녹음을 취소했습니다.";
        return;
      }

      const blob = new Blob(recordChunks, { type: mime });
      recordChunks = [];
      resetRecordUiIdle();

      if (!blob.size) {
        clearLastRecordedAudio();
        if (els.recordStatus) els.recordStatus.textContent = "녹음 데이터가 없습니다.";
        showError("녹음이 너무 짧거나 비어 있습니다.");
        return;
      }

      rememberLastRecordedAudio(blob, mime);

      if (els.recordStatus) {
        els.recordStatus.textContent =
          "녹음이 준비되었습니다. 저장이 필요하면 「녹음 파일 저장」, 텍스트는 「텍스트 변환」을 눌러 주세요.";
      }
    };

    try {
      recordMediaRecorder.start(250);
    } catch (e) {
      console.error(e);
      stopRecordStreamTracks();
      recordMediaRecorder = null;
      recordChunks = [];
      resetRecordUiIdle();
      showError("녹음을 시작할 수 없습니다.");
      if (els.recordStatus) els.recordStatus.textContent = "녹음을 시작할 수 없습니다.";
      return;
    }

    if (els.btnRecordStart) els.btnRecordStart.disabled = true;
    if (els.btnRecordStop) els.btnRecordStop.disabled = false;
    if (els.recordStatus) {
      els.recordStatus.textContent = "녹음 중… 종료하면 메모리에 보관됩니다.";
    }
  });

  els.btnRecordTranscribe?.addEventListener("click", async () => {
    if (!lastRecordedAudioBlob?.size) return;
    showError("");
    setOutput("", { skipSummary: true });
    const mime = lastRecordedAudioBlob.type || "audio/webm";
    const nameStem = `recording_${defaultDownloadBasename()}`;
    const ext = lastRecordedAudioExt || extFromAudioMime(mime);
    const file = new File(
      [lastRecordedAudioBlob],
      `${nameStem}.${ext}`,
      { type: mime }
    );
    if (els.btnRecordTranscribe) els.btnRecordTranscribe.disabled = true;
    if (els.btnRecordStart) els.btnRecordStart.disabled = true;
    try {
      await transcribeFile(file);
      if (els.recordStatus) {
        els.recordStatus.textContent = "변환 완료. 아래 변환 결과·요약을 확인하세요.";
      }
    } catch (err) {
      console.error(err);
      showError(
        err?.message
          ? `변환 실패: ${err.message}`
          : "변환에 실패했습니다. 네트워크와 녹음 형식을 확인해 주세요."
      );
      if (els.recordStatus) els.recordStatus.textContent = "변환에 실패했습니다.";
    } finally {
      if (els.btnRecordTranscribe) els.btnRecordTranscribe.disabled = false;
      if (els.btnRecordStart) els.btnRecordStart.disabled = false;
    }
  });

  els.btnRecordStop?.addEventListener("click", () => {
    if (!recordMediaRecorder || recordMediaRecorder.state !== "recording") return;
    recordDiscardOnStop = false;
    if (els.recordStatus) els.recordStatus.textContent = "녹음을 마무리하는 중…";
    if (els.btnRecordStop) els.btnRecordStop.disabled = true;
    try {
      recordMediaRecorder.stop();
    } catch (e) {
      console.error(e);
      stopRecordStreamTracks();
      resetRecordUiIdle();
      clearLastRecordedAudio();
      showError("녹음 종료에 실패했습니다.");
      if (els.recordStatus) els.recordStatus.textContent = "녹음 종료에 실패했습니다.";
    }
  });

  els.btnSaveRecording?.addEventListener("click", () => {
    void saveLastRecordedAudio();
  });
}

els.btnCopy.addEventListener("click", async () => {
  const t = els.output.value;
  if (!t) return;
  try {
    await navigator.clipboard.writeText(t);
    els.btnCopy.textContent = "복사됨";
    setTimeout(() => {
      els.btnCopy.textContent = "복사";
    }, 1500);
  } catch {
    els.output.select();
    document.execCommand("copy");
  }
});

els.output?.addEventListener("input", () => {
  syncUiAfterOutputEdit();
});

els.btnCopySummary?.addEventListener("click", async () => {
  await copyTextWithFeedback(
    els.summaryOutput.value,
    els.btnCopySummary,
    els.summaryOutput
  );
});

els.btnCopySummaryCore?.addEventListener("click", async () => {
  await copyTextWithFeedback(
    els.summaryCoreOutput.value,
    els.btnCopySummaryCore,
    els.summaryCoreOutput
  );
});

els.btnSaveFull.addEventListener("click", async () => {
  const t = els.output.value.trim();
  if (!t) return;
  await saveTextWithLocationAndName(t, "음성변환_전체", "전체 변환 결과 저장");
});

els.btnSaveSummary.addEventListener("click", async () => {
  const t = els.summaryOutput.value.trim();
  if (!t) return;
  await saveTextWithLocationAndName(t, "음성변환_요약", "요약 결과 저장");
});

els.btnSaveSummaryCore.addEventListener("click", async () => {
  const t = els.summaryCoreOutput.value.trim();
  if (!t) return;
  await saveTextWithLocationAndName(t, "음성변환_핵심의핵심", "2차 요약 저장");
});

els.btnRefreshSummary.addEventListener("click", () => {
  void refreshSummaryFromOutput();
});

els.btnRefreshSummaryCore.addEventListener("click", () => {
  void refreshSummaryCoreFromSummary();
});

function readStoredTheme() {
  try {
    const v = localStorage.getItem(THEME_STORAGE_KEY);
    if (v === "light" || v === "dark") return v;
  } catch {
    /* ignore */
  }
  return null;
}

function getDefaultTheme() {
  try {
    if (window.matchMedia("(prefers-color-scheme: light)").matches) return "light";
  } catch {
    /* ignore */
  }
  return "dark";
}

function syncThemeToggleUi() {
  const t = document.documentElement.dataset.theme === "light" ? "light" : "dark";
  if (els.themeDark && els.themeLight) {
    els.themeDark.classList.toggle("is-active", t === "dark");
    els.themeLight.classList.toggle("is-active", t === "light");
    els.themeDark.setAttribute("aria-pressed", String(t === "dark"));
    els.themeLight.setAttribute("aria-pressed", String(t === "light"));
  }
}

/** 사용자가 고른 테마만 localStorage에 저장합니다. */
function applyTheme(theme) {
  const t = theme === "light" ? "light" : "dark";
  document.documentElement.dataset.theme = t;
  try {
    localStorage.setItem(THEME_STORAGE_KEY, t);
  } catch {
    /* ignore */
  }
  syncThemeToggleUi();
}

function initThemeFromDocument() {
  if (!document.documentElement.dataset.theme) {
    document.documentElement.dataset.theme =
      readStoredTheme() ?? getDefaultTheme();
  }
  syncThemeToggleUi();
}

function setupThemeToggle() {
  els.themeDark?.addEventListener("click", () => applyTheme("dark"));
  els.themeLight?.addEventListener("click", () => applyTheme("light"));
}

function warnIfFileProtocol() {
  if (window.location.protocol === "file:") {
    showError(
      "이 페이지가 file://로 열려 있습니다. Web Worker가 필요하므로 로컬 서버로 여세요. 예: 터미널에서 폴더로 이동 후 npx --yes serve . 실행 후 표시된 주소로 접속"
    );
  }
}

setupTabs();
setupFileUi();
setupRecordUi();
setupMic();
setupSaveDialog();
setupThemeToggle();
initThemeFromDocument();
warnIfFileProtocol();
syncSummaryCopyButtonsSoon();
