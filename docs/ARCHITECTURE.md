# LM-Wars — Architecture Blueprint (ARCHITECTURE.md)

## 1. System Overview

```
+-------------------------------------------------------------+
|                      LM-Wars Web App                        |
|   (Next.js App Router, React 19, Tailwind CSS, Lucide)     |
+-------------------------------------------------------------+
| UI Layers:                                                  |
|  - Arena Battle View (Side-by-side streaming & metrics)     |
|  - Automated Benchmark Runner (Test suite executor)         |
|  - ELO Leaderboard & Performance Analytics                  |
|  - Connection Manager (LM Studio status & model selector)   |
+-----------------------------+-------------------------------+
                              |
                     API Routes / Services
       (/api/battle, /api/benchmark, /api/models, /api/storage)
                              |
            +-----------------+-----------------+
            |                                   |
            v                                   v
+-----------------------+           +-----------------------+
|   LM Studio Server    |           | Local JSON / SQLite   |
| (http://127.0.0.1:    |           |   Leaderboard State   |
|      1234/v1)         |           | (Battles, ELO, Logs)  |
| - /v1/models          |           +-----------------------+
| - /v1/chat/completions|
+-----------------------+
```

---

## 2. Tech Stack Selection Rationale
- **Frontend & Server Framework**: Next.js 16+ (App Router) + React 19 + TypeScript.
  - *Rationale*: Unified server-client architecture. Server routes allow direct local network access to LM Studio avoiding browser CORS issues, while Server-Sent Events (SSE) provide real-time token streaming to the UI.
- **Styling**: Tailwind CSS v4.
  - *Rationale*: Rapid, utility-first UI styling with high contrast dark/cyberpunk/arena aesthetic fitting for model battles.
- **Model Runtime Protocol**: OpenAI-compatible REST & SSE specification.
  - *Rationale*: LM Studio natively exposes an OpenAI-compatible server at `http://127.0.0.1:1234/v1`. This standard protocol also unlocks interoperability with Ollama (`11434`), vLLM, and OpenRouter with zero code redesign.
- **Storage Strategy**: Local file-based JSON store with transactional write safety, easily extensible to SQLite.
  - *Rationale*: Zero-config local persistence that requires no external database daemons to run.

---

## 3. Core Modules & Data Contracts

### 3.1 Model Connection & Client (`src/lib/lmstudio/client.ts`)
```typescript
export interface ModelInfo {
  id: string;
  object: string;
  owned_by?: string;
  context_length?: number;
}

export interface StreamChunkTelemetry {
  token: string;
  timestamp: number;
}

export interface ExecutionMetrics {
  ttftMs: number;         // Time to first token
  totalDurationMs: number;// Total generation duration
  tokensGenerated: number;// Approximate or reported token count
  tokensPerSec: number;   // Evaluation speed (TPS)
}
```

### 3.2 Battle Engine & ELO Rating (`src/lib/battle/elo.ts`)
- Standard ELO implementation:
  - Initial rating: $1200$
  - K-factor: $32$ (for responsive updates during local matches)
  - Expected score: $E_A = \frac{1}{1 + 10^{(R_B - R_A)/400}}$
  - Rating update: $R'_A = R_A + K \cdot (S_A - E_A)$ where $S \in \{1.0, 0.5, 0.0\}$.

### 3.3 Benchmark Test Suites (`src/lib/benchmark/suites.ts`)
- Standardized test suites:
  - **Reasoning & Logic**: Multi-step puzzles, syllogisms, fallacy spotting.
  - **Code Generation**: Algorithmic problems, TypeScript typing challenges, syntax compliance.
  - **Instruction & Constraint Following**: Strict word counts, specific formatting (JSON only, no markdown), reverse sentence orders.
  - **Speed & Latency**: Standard token burst test to measure raw tokens/sec and TTFT.

---

## 4. Directory Conventions
```
lm-wars/
├── docs/                      # Architectural docs & specifications
│   ├── SPEC.md
│   ├── ARCHITECTURE.md
│   └── ROADMAP.md
├── src/
│   ├── app/                   # Next.js App Router pages & API routes
│   │   ├── api/               # Server-side API proxies & runners
│   │   │   ├── models/        # LM Studio model listing proxy
│   │   │   ├── chat/          # Streaming completions proxy with timing
│   │   │   ├── leaderboard/   # Persistent leaderboard CRUD
│   │   │   └── benchmark/     # Benchmark suite execution
│   │   ├── arena/             # Head-to-head battle interface
│   │   ├── benchmarks/        # Automated benchmark suites interface
│   │   ├── leaderboard/       # ELO ratings & historical statistics
│   │   ├── layout.tsx
│   │   └── page.tsx           # Home dashboard
│   ├── components/            # Reusable UI components
│   │   ├── arena/             # Side-by-side battle panels, prompt input
│   │   ├── leaderboard/       # Ranking tables, ELO badge, stats
│   │   ├── layout/            # Navbar, ConnectionIndicator, Theme
│   │   └── ui/                # Buttons, cards, badges, modal dialogs
│   ├── lib/                   # Business logic, engines, client wrappers
│   │   ├── lmstudio/          # LM Studio API client & telemetry
│   │   ├── battle/            # Battle state & ELO calculation
│   │   ├── benchmark/         # Test suites & evaluators
│   │   └── storage/           # Local persistence engine
│   └── types/                 # Shared TypeScript interfaces
└── AGENTS.md                  # AI Assistant instructions
```
