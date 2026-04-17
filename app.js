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
  panelMic: $("panel-mic"),
  dropzone: $("dropzone"),
  fileInput: $("fileInput"),
  fileName: $("fileName"),
  modelSelect: $("modelSelect"),
  languageSelect: $("languageSelect"),
  speedSelect: $("speedSelect"),
  btnTranscribeFile: $("btnTranscribeFile"),
  btnClearFile: $("btnClearFile"),
  btnMicStart: $("btnMicStart"),
  btnMicStop: $("btnMicStop"),
  micStatus: $("micStatus"),
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

function showError(msg) {
  els.error.textContent = msg;
  els.error.hidden = !msg;
}

function syncSummaryCopyButtons() {
  if (els.btnCopySummary) {
    els.btnCopySummary.disabled = !els.summaryOutput.value.trim();
  }
  if (els.btnCopySummaryCore) {
    els.btnCopySummaryCore.disabled = !els.summaryCoreOutput.value.trim();
  }
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
    els.btnCopySummary.disabled = true;
    els.btnCopySummaryCore.disabled = true;
    els.btnRefreshSummaryCore.disabled = true;
  } else if (!skipSummary) {
    els.btnSaveSummary.disabled = true;
    scheduleSummaryGeneration(full);
  } else {
    els.btnSaveSummary.disabled = !els.summaryOutput.value.trim();
    const hasCore = els.summaryCoreOutput.value.trim().length > 0;
    els.btnSaveSummaryCore.disabled = !hasCore;
    els.btnRefreshSummaryCore.disabled = !els.summaryOutput.value.trim();
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
  if (els.btnCopySummaryCore) els.btnCopySummaryCore.disabled = true;
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
    syncSummaryCopyButtons();
    return;
  }
  setProgress(true, "2차 요약(핵심의 핵심)…", 72);
  const core = buildExtractiveSummary(primary, SUMMARY_CORE_MAX).trim();
  els.summaryCoreOutput.value = core;
  els.btnSaveSummaryCore.disabled = !core;
  els.btnRefreshSummaryCore.disabled = false;
  syncSummaryCopyButtons();
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
    syncSummaryCopyButtons();
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
  if (els.btnCopySummary) els.btnCopySummary.disabled = true;
  if (els.btnCopySummaryCore) els.btnCopySummaryCore.disabled = true;
  await flushPaint();
  try {
    setProgress(true, "1차 핵심 요약…", 35);
    applyPrimaryAndCoreSummary(full);
  } catch (err) {
    console.error(err);
    showError("요약에 실패했습니다.");
    syncSummaryCopyButtons();
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
  if (els.btnCopySummaryCore) els.btnCopySummaryCore.disabled = true;
  await flushPaint();
  try {
    setProgress(true, "2차 요약 다시…", 55);
    const core = buildExtractiveSummary(primary, SUMMARY_CORE_MAX).trim();
    els.summaryCoreOutput.value = core;
    els.btnSaveSummaryCore.disabled = !core;
    syncSummaryCopyButtons();
  } catch (err) {
    console.error(err);
    showError("2차 요약에 실패했습니다.");
    syncSummaryCopyButtons();
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
 * Whisper 파이프라인 반환값을 한 덩어리 문자열로 정리합니다.
 * (청크·타임스탬프 모드에서 chunks만 채워지는 경우가 있어 누락을 막습니다.)
 */
function normalizeAsrResult(result) {
  if (result == null) return "";
  if (typeof result === "string") return result.trim();
  if (Array.isArray(result)) {
    return result.map(normalizeAsrResult).filter(Boolean).join(" ").trim();
  }
  const chunks = result.chunks;
  if (Array.isArray(chunks) && chunks.length) {
    const fromChunks = chunks
      .map((c) => (typeof c === "string" ? c : (c && c.text) || ""))
      .join(" ")
      .trim();
    if (fromChunks) return fromChunks;
  }
  const top = typeof result.text === "string" ? result.text.trim() : "";
  return top;
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
  const lang = els.languageSelect?.value || "auto";
  const speedPreset = els.speedSelect?.value || "balanced";
  /** @type {Record<string, unknown>} */
  const opts = {
    task: "transcribe",
  };
  const longForm =
    durationSec == null || !Number.isFinite(durationSec) || durationSec > 28;
  if (longForm) {
    opts.chunk_length_s = 30;
    // 겹침이 클수록 경계 품질↑·연산↑. 속도 우선은 겹침만 줄임(경계 문장 품질은 약간↓ 가능).
    opts.stride_length_s = speedPreset === "fast" ? 2 : 5;
  }
  if (lang === "korean") opts.language = "korean";
  else if (lang === "english") opts.language = "english";
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
    const samples = new Float32Array(ch0);
    const durationSec = Number.isFinite(duration)
      ? duration
      : samples.length / WHISPER_SAMPLE_RATE;
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
  setProgress(true, "오디오·모델 준비 중…", 10);
  try {
    const result = await transcribeWithWorker(file);
    setProgress(true, "완료", 100);
    setOutput(normalizeAsrResult(result));
  } finally {
    setTimeout(() => setProgress(false), 400);
  }
}

function setupTabs() {
  els.tabs.forEach((tab) => {
    tab.addEventListener("click", () => {
      const name = tab.dataset.tab;
      els.tabs.forEach((t) => {
        const on = t === tab;
        t.classList.toggle("is-active", on);
        t.setAttribute("aria-selected", on);
      });
      const fileMode = name === "file";
      els.panelFile.classList.toggle("is-visible", fileMode);
      els.panelFile.hidden = !fileMode;
      els.panelMic.classList.toggle("is-visible", !fileMode);
      els.panelMic.hidden = fileMode;
      if (fileMode && micRecognition) {
        try {
          micRecognition.stop();
        } catch {
          /* ignore */
        }
      }
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
    micRecognition = new SpeechRecognition();
    micRecognition.lang = "ko-KR";
    micRecognition.continuous = true;
    micRecognition.interimResults = true;

    let finalText = els.output.value.trim();

    micRecognition.onresult = (event) => {
      let interim = "";
      for (let i = event.resultIndex; i < event.results.length; i++) {
        const piece = event.results[i][0].transcript;
        if (event.results[i].isFinal) finalText += (finalText ? " " : "") + piece;
        else interim += piece;
      }
      setOutput(finalText + (interim ? (finalText ? " " : "") + interim : ""), {
        skipSummary: true,
      });
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
setupMic();
setupSaveDialog();
setupThemeToggle();
initThemeFromDocument();
warnIfFileProtocol();
