# AI Assistant Guidelines for LM-Wars

## Project Overview
**LM-Wars** is an AI-first local Language Model Benchmark & Arena Battleground designed specifically for local runtimes like **LM Studio** (`http://127.0.0.1:1234/v1`). It provides real-time head-to-head battles, streaming telemetry (TTFT, TPS), automated evaluation suites, and an ELO leaderboard.

---

## Quick Reference Commands
- **Dev Server**: `bun run dev` (or `pnpm dev` / `npm run dev`) -> runs on `http://localhost:3000`
- **Build**: `bun run build`
- **Lint**: `bun run lint`
- **Package Manager**: Prefer `bun` for package management and scripts.

---

## Architecture Reference
Before proposing architectural shifts or adding major modules, consult:
- [`docs/SPEC.md`](file:///D:/projects/lm-wars/docs/SPEC.md) — Product vision, user stories, requirements
- [`docs/ARCHITECTURE.md`](file:///D:/projects/lm-wars/docs/ARCHITECTURE.md) — Data flow, schemas, telemetry protocol, directory conventions
- [`docs/ROADMAP.md`](file:///D:/projects/lm-wars/docs/ROADMAP.md) — Phased task breakdown

---

## Task Tracking Rule
- AI assistants working on LM-Wars must consult [`docs/ROADMAP.md`](file:///D:/projects/lm-wars/docs/ROADMAP.md) at the start of any feature implementation.
- As tasks are completed, tested, and verified, the assistant must update `docs/ROADMAP.md` by marking items `[x]` and recording any important notes or adjustments.

---

## Code Style & Conventions
- **TypeScript**: Strict type checking. Avoid `any`; use well-defined interfaces in `src/types/` or co-located schemas.
- **Next.js App Router**: Keep Client Components (`"use client"`) at the leaves where user interaction and streaming are needed. Keep server routes clean and defensively handle offline LM Studio endpoints.
- **Styling**: Tailwind CSS v4 utility classes. Prefer a modern, high-contrast dark arena aesthetic suitable for AI benchmarks.
- **Error Handling**: Gracefully handle connection drops to `http://127.0.0.1:1234`. The UI should inform the user when LM Studio is paused or not running rather than crashing.
