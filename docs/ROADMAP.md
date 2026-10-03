# LM-Wars — Development Roadmap (ROADMAP.md)

## Phased Execution Plan

### Phase 1: MVP / Core Foundation
- [ ] **Task 1.1: LM Studio Client & Telemetry Engine**
  - Implement `src/lib/lmstudio/client.ts` with connection verification (`/v1/models`), OpenAI-compatible streaming completion handler, and real-time telemetry (TTFT, duration, token estimation, TPS).
- [ ] **Task 1.2: Server API Endpoints**
  - Create `/api/models` to probe LM Studio status and loaded models.
  - Create `/api/chat` to stream responses with chunk-level timing metadata.
  - Create `/api/leaderboard` for saving battle results and reading rankings.
- [ ] **Task 1.3: Head-to-Head Arena Interface (Battle Mode)**
  - Dual-panel synchronized prompt execution against two selected local models (or two configurations/temperatures of the same model).
  - Side-by-side real-time streaming displays with token speed meters (TPS) and latency timers (TTFT).
  - Blind battle mode option (Model A / Model B identities concealed until vote is cast).
  - Voting buttons: Model A Wins, Model B Wins, Draw / Both Bad, Both Good.
- [ ] **Task 1.4: ELO Rating Engine & Storage**
  - Implement ELO calculation logic (`src/lib/battle/elo.ts`).
  - Implement local storage engine (`src/lib/storage/store.ts`) for persisting battle logs and calculated rankings.
- [ ] **Task 1.5: Dynamic Leaderboard View**
  - Create `/leaderboard` page displaying current model rankings, win/loss records, average TPS, and recent battle replay log.

---

### Phase 2: Automated Benchmark Suites
- [ ] **Task 2.1: Pre-configured Benchmark Test Suites**
  - Define evaluation sets for:
    - Code Generation (Python, TypeScript, SQL)
    - Logical Reasoning & Math
    - Instruction / Negative Constraint Following
    - Context Recall & Summarization
- [ ] **Task 2.2: Automated Batch Runner & Progress Visualizer**
  - Build UI and runner to queue and execute test sets against one or more models automatically.
  - Display real-time progress bars, pass/fail status, and live output stream.
- [ ] **Task 2.3: Automated LLM-as-Judge & Heuristic Scorer**
  - Deterministic evaluation (regex, JSON parsing, exact match, constraint verification).
  - Model-as-judge prompt evaluation using a designated judge model.

---

### Phase 3: Analytics, Polish & Multi-Runtime Extensions
- [ ] **Task 3.1: Hardware & Speed Performance Graphs**
  - Charts showing TPS across context lengths, temperature distributions, and prompt complexity.
- [ ] **Task 3.2: Multi-Runtime Support (Ollama & Custom Endpoints)**
  - Add quick-switch presets for Ollama (`http://localhost:11434/v1`), vLLM, and custom API endpoints.
- [ ] **Task 3.3: Export & Sharing**
  - Export benchmark results and battle cards as JSON, CSV, and markdown images/cards.
