# LM-Wars — Specification (SPEC.md)

## 1. Product Vision & Overview
**LM-Wars** is a premier local-first Language Model Benchmark and Arena Battleground. Designed to operate seamlessly with local LLM runtimes—primarily **LM Studio** (via OpenAI-compatible endpoints)—LM-Wars allows developers, AI researchers, and hobbyists to pit local models against each other in real-time head-to-head battles, automated standard evaluation suites (reasoning, coding, math, instruction following, creative generation), and custom prompt challenges.

LM-Wars eliminates the guesswork in choosing which quantized model, parameter size, or system prompt performs best on your local machine by delivering transparent, reproducible leaderboards, throughput telemetry (tokens/sec, time-to-first-token, latency), and qualitative scoring.

---

## 2. Target Audience & Personas
- **Local AI Enthusiasts & Power Users**: Running local models with LM Studio on workstation GPUs, needing empirical data on which GGUF/model checkpoint excels at specific workloads.
- **AI Application Engineers & Prompt Engineers**: Evaluating prompt templates, system instructions, and temperature/top-p settings across local LLMs before deploying to production.
- **Hardware Benchmarkers**: Comparing token generation speed (TPS), memory efficiency, and response quality across different hardware configs (Apple Silicon, NVIDIA RTX, AMD ROCm).

---

## 3. Core Capabilities & User Stories

### US-1: Local LM Studio Connection & Model Discovery
- **Story**: As a user with LM Studio running, I want LM-Wars to automatically detect my active LM Studio instance and list loaded/available models so that I can configure battles without manual setup.
- **Acceptance Criteria**:
  - Connects to `http://127.0.0.1:1234/v1` (with configurable port/host and fallback).
  - Queries `/v1/models` and displays current model status (loaded model, context length, engine).
  - Health check ping indicator in the header (Connected / Offline / Reconnecting).

### US-2: Head-to-Head Arena ("Battle Mode")
- **Story**: As a researcher, I want to send a prompt simultaneously to Model A and Model B, view their streaming answers side-by-side with live latency and TPS, and vote or auto-judge the winner.
- **Acceptance Criteria**:
  - Dual-panel synchronized streaming responses.
  - Metrics recorded per model: TTFT (Time to First Token), Total Duration, Tokens per Second (eval TPS), Character count.
  - Blind / Named battle mode options with user rating or automated LLM-as-judge scoring.

### US-3: Automated Benchmark Suites
- **Story**: As a developer, I want to run standardized test categories (Code Generation, Logic & Deduction, Multi-turn QA, Constraint Following, Fast Summarization) against my model.
- **Acceptance Criteria**:
  - Pre-packaged eval categories with deterministic test cases and rubric assertions.
  - Batch execution runner with progress bar and stop/pause controls.
  - Aggregate scores computed across accuracy, speed, and instruction obedience.

### US-4: Dynamic Leaderboard & Match History
- **Story**: As a user, I want a persistent leaderboard showing model rankings, win rates, ELO ratings, and historical performance charts.
- **Acceptance Criteria**:
  - ELO rating calculations based on arena battle results.
  - Filterable leaderboard by category (Coding, Reasoning, Speed, Creative).
  - Complete history log with prompt, model parameters, answers, and metrics.

---

## 4. Non-Functional Requirements
- **Privacy & Local Sovereignty**: 100% offline capable; zero outbound telemetry or cloud dependencies required.
- **Latency & Responsiveness**: Streaming UI with low render latency using React 19 / Next.js and fast web sockets or Server-Sent Events (SSE).
- **Extensibility**: Modular architecture supporting additional endpoints (Ollama, vLLM, Aphrodite, LocalAI, OpenRouter) via unified OpenAI-compatible client adapter.
