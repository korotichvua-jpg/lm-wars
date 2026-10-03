# ⚔️ LM-Wars

**LM-Wars** is a local-first Language Model Benchmark and Arena Battleground. Designed to connect to your local **LM Studio** server, it enables live head-to-head model battles, automated multi-category benchmarking, real-time token throughput metrics (TPS & TTFT), and an interactive ELO leaderboard.

---

## ✨ Features
- 🥊 **Side-by-Side Arena Battles**: Stream responses from two local models (or different parameters/temperatures) concurrently.
- ⚡ **Real-Time Telemetry**: Measure Time-to-First-Token (TTFT), tokens/sec (TPS), character count, and latency on every turn.
- 🏆 **Dynamic ELO Leaderboard**: Community-standard ELO rating updates based on human voting or automated judging.
- 🎯 **Automated Benchmark Suites**: Evaluate models across reasoning, coding, instruction adherence, and speed with reproducible rubrics.
- 🔒 **100% Local & Private**: Direct connection to LM Studio (`http://127.0.0.1:1234/v1`). No cloud API keys or external data transmission required.

---

## 🚀 Quick Start

### 1. Start LM Studio
Ensure **LM Studio** is running with its local server enabled:
- Port: `1234` (Default URL: `http://127.0.0.1:1234/v1`)
- Load one or more models in LM Studio.

### 2. Launch LM-Wars
```bash
# Install dependencies (if not already done)
bun install

# Start the development server
bun run dev
```

Open [http://localhost:3000](http://localhost:3000) in your browser.

---

## 📚 Documentation
- [**Product Specification (`docs/SPEC.md`)**](docs/SPEC.md) — Personas, user stories, and acceptance criteria.
- [**Architecture Blueprint (`docs/ARCHITECTURE.md`)**](docs/ARCHITECTURE.md) — System flow, data models, and protocols.
- [**Roadmap & Tasks (`docs/ROADMAP.md`)**](docs/ROADMAP.md) — Phased execution tracking and feature backlog.
- [**AI Guidelines (`AGENTS.md`)**](AGENTS.md) — Coding conventions and assistant rules.

---

## 🛠️ Tech Stack
- **Framework**: [Next.js 16](https://nextjs.org/) (App Router, React 19)
- **Runtime**: [Bun](https://bun.sh/)
- **Styling**: [Tailwind CSS v4](https://tailwindcss.com/)
- **Icons**: [Lucide React](https://lucide.dev/)
- **LLM Interface**: OpenAI-compatible REST & Server-Sent Events (SSE)
