/**
 * Whisper 추론 전용 Worker — 메인 스레드(탭)가 멈추지 않도록 합니다.
 */
import { pipeline, env } from "https://cdn.jsdelivr.net/npm/@xenova/transformers@2.17.2/dist/transformers.min.js";

env.allowLocalModels = false;
env.useBrowserCache = true;

/* SharedArrayBuffer + COOP/COEP(crossOriginIsolated)일 때만 멀티스레드 WASM 사용 가능 → 추론 단축 */
try {
  const wasm = env.backends?.onnx?.wasm;
  if (wasm && typeof self !== "undefined" && self.crossOriginIsolated) {
    const hw =
      typeof navigator !== "undefined" ? navigator.hardwareConcurrency : 2;
    wasm.numThreads = Math.min(4, Math.max(2, hw || 2));
  }
} catch {
  /* 기본 스레드 수 유지 */
}

let pipelinePromise = null;
let loadedModelKey = null;

function getTranscriber(modelKey, modelId) {
  if (loadedModelKey !== modelKey || !pipelinePromise) {
    loadedModelKey = modelKey;
    pipelinePromise = pipeline("automatic-speech-recognition", modelId, {
      progress_callback: (e) => {
        self.postMessage({
          type: "loadProgress",
          status: e.status,
          file: e.file,
          progress: e.progress,
        });
      },
    });
  }
  return pipelinePromise;
}

self.onmessage = async (ev) => {
  const msg = ev.data;
  if (msg?.type !== "transcribe") return;

  /** 메인 스레드에서 만든 16kHz 모노 Float32Array (Worker에는 AudioContext 없음) */
  const { modelKey, modelId, options, raw, jobId } = msg;

  try {
    const transcriber = await getTranscriber(modelKey, modelId);
    self.postMessage({ type: "inferStart", jobId });

    const pcm = raw instanceof Float32Array ? raw : new Float32Array(raw);
    // 16kHz 모노 Float32Array — Whisper 기본 입력 (Worker에는 URL/AudioContext 사용 불가)
    const result = await transcriber(pcm, options ?? {});

    self.postMessage({ type: "inferDone", jobId, result });
  } catch (err) {
    self.postMessage({
      type: "error",
      jobId,
      message: err?.message || String(err),
    });
  }
};
