import os
import sys
import json
import time
import urllib.request

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

PROMPT = (
    "what is total construction cost for building in texas 2 gigawath ai most advanced and secure "
    "data center (full expenses, investments, all materials, detailed econimic model in advanced top tier "
    "investment top tier firm form, all source prices, calculations in excel)"
)

BASE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "benchmark_results", "texas_2gw_datacenter")
SUMMARY_MD = os.path.join(BASE_DIR, "processing_times_summary.md")
SUMMARY_CSV = os.path.join(BASE_DIR, "processing_times_summary.csv")
SUMMARY_JSON = os.path.join(BASE_DIR, "processing_times_summary.json")

def sanitize_folder_name(name: str) -> str:
    return name.replace("/", "_").replace("\\", "_").replace(":", "_").replace(" ", "_")

def update_summary(model_name, record):
    records = []
    if os.path.exists(SUMMARY_JSON):
        try:
            with open(SUMMARY_JSON, "r", encoding="utf-8") as f:
                records = json.load(f)
        except Exception:
            records = []
    
    # Update or append
    found = False
    for i, r in enumerate(records):
        if r.get("model") == model_name:
            records[i] = record
            found = True
            break
    if not found:
        records.append(record)

    with open(SUMMARY_JSON, "w", encoding="utf-8") as f:
        json.dump(records, f, indent=2)

    with open(SUMMARY_CSV, "w", encoding="utf-8") as f:
        f.write("Model,Status,TTFT (s),Generation Time (s),Total Time (s),Tokens,Tokens/sec,Chars,Load Time (s),Error\n")
        for r in records:
            err = (r.get("error") or "").replace(",", " ")
            f.write(f'"{r["model"]}","{r["status"]}",{r.get("ttft_sec",0)},{r.get("gen_duration_sec",0)},{r.get("total_duration_sec",0)},{r.get("token_count",0)},{r.get("tokens_per_sec",0)},{r.get("char_count",0)},{r.get("load_time_sec",0)},"{err}"\n')

    with open(SUMMARY_MD, "w", encoding="utf-8") as f:
        f.write("# 📊 Benchmark Processing Times & Performance Summary\n\n")
        f.write("**Evaluation Task**: Construction Cost & Advanced Economic Model for a 2 Gigawatt AI Data Center in Texas\n\n")
        f.write(f"**Updated**: {time.strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        f.write("| Model | Status | Load Time | TTFT | Gen Time | Total Time | Output Tokens | Speed (TPS) |\n")
        f.write("| :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: |\n")
        for r in records:
            status_icon = "✅ Success" if r["status"] == "SUCCESS" else f"❌ {r['status']}"
            f.write(
                f"| **`{r['model']}`** | {status_icon} | {r.get('load_time_sec', 0)}s | {r.get('ttft_sec', 0)}s | "
                f"{r.get('gen_duration_sec', 0)}s | {r.get('total_duration_sec', 0)}s | {r.get('token_count', 0)} | **{r.get('tokens_per_sec', 0)} tps** |\n"
            )
        f.write("\n---\n\n")
        f.write("### Sub-folder Results Directory\n")
        for r in records:
            if r.get("folder"):
                f.write(f"- [`{r['model']}`](./{os.path.basename(r['folder'])}/response.md)\n")

def run_single(model_name, load_time_sec=29.29):
    print(f"Running inference for loaded model: {model_name}")
    model_folder = os.path.join(BASE_DIR, sanitize_folder_name(model_name))
    os.makedirs(model_folder, exist_ok=True)

    url = "http://127.0.0.1:1234/v1/chat/completions"
    payload = {
        "model": model_name,
        "messages": [
            {
                "role": "system",
                "content": "You are a senior infrastructure investment director, quantitative financial modeler, and mission-critical data center engineering expert."
            },
            {
                "role": "user",
                "content": PROMPT
            }
        ],
        "temperature": 0.3,
        "max_tokens": 2048,
        "stream": True
    }
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(url, data=data, headers={"Content-Type": "application/json"})

    t_start = time.time()
    t_first_token = None
    chunks_text = []
    token_count = 0

    try:
        with urllib.request.urlopen(req, timeout=3600) as resp:
            for raw_line in resp:
                line = raw_line.decode("utf-8").strip()
                if not line or not line.startswith("data:"):
                    continue
                data_part = line[5:].strip()
                if data_part == "[DONE]":
                    break
                try:
                    chunk_obj = json.loads(data_part)
                    delta = chunk_obj.get("choices", [{}])[0].get("delta", {})
                    content = delta.get("content") or delta.get("reasoning_content")
                    if content:
                        if t_first_token is None:
                            t_first_token = time.time()
                            print(f"  -> First token received in {t_first_token - t_start:.2f}s", flush=True)
                        chunks_text.append(content)
                        token_count += 1
                        if token_count % 50 == 0:
                            elapsed = time.time() - t_first_token
                            cur_tps = token_count / elapsed if elapsed > 0 else 0
                            print(f"  -> Streamed {token_count} tokens ({cur_tps:.2f} TPS)...", flush=True)
                except Exception:
                    continue

        t_end = time.time()
        if t_first_token is None:
            t_first_token = t_end

        total_text = "".join(chunks_text)
        total_duration = t_end - t_start
        ttft = t_first_token - t_start
        gen_duration = t_end - t_first_token if t_end > t_first_token else 0.001
        tps = token_count / gen_duration if gen_duration > 0 else 0.0

        # Save response
        response_file = os.path.join(model_folder, "response.md")
        with open(response_file, "w", encoding="utf-8") as f:
            f.write(f"# Economic Analysis & Construction Cost: 2GW Texas AI Data Center\n\n")
            f.write(f"**Model**: `{model_name}`  \n")
            f.write(f"**Generated**: {time.strftime('%Y-%m-%d %H:%M:%S')}  \n")
            f.write(f"**TTFT**: {ttft:.3f}s | **Total Duration**: {total_duration:.3f}s | **Tokens**: {token_count} | **Speed**: {tps:.2f} TPS  \n\n")
            f.write(f"### Prompt\n> {PROMPT}\n\n---\n\n")
            f.write(total_text)

        # Save metrics
        metrics_file = os.path.join(model_folder, "metrics.json")
        res_data = {
            "model": model_name,
            "prompt": PROMPT,
            "timestamp": time.time(),
            "load_time_sec": round(load_time_sec, 2),
            "ttft_sec": round(ttft, 3),
            "gen_duration_sec": round(gen_duration, 3),
            "total_duration_sec": round(total_duration, 3),
            "token_count": token_count,
            "char_count": len(total_text),
            "tokens_per_sec": round(tps, 2),
            "success": True,
            "error": None
        }
        with open(metrics_file, "w", encoding="utf-8") as f:
            json.dump(res_data, f, indent=2)

        record = {
            "model": model_name,
            "status": "SUCCESS",
            "load_time_sec": round(load_time_sec, 2),
            "ttft_sec": round(ttft, 3),
            "gen_duration_sec": round(gen_duration, 3),
            "total_duration_sec": round(total_duration, 3),
            "token_count": token_count,
            "char_count": len(total_text),
            "tokens_per_sec": round(tps, 2),
            "error": None,
            "folder": model_folder
        }
        update_summary(model_name, record)
        print(f"[OK] Completed {model_name}: {token_count} tokens in {gen_duration:.2f}s ({tps:.2f} TPS)")

    except Exception as e:
        t_end = time.time()
        print(f"[!] Error during inference: {e}")
        record = {
            "model": model_name,
            "status": "INFERENCE_FAILED",
            "load_time_sec": round(load_time_sec, 2),
            "ttft_sec": 0,
            "gen_duration_sec": 0,
            "total_duration_sec": round(t_end - t_start, 3),
            "token_count": 0,
            "char_count": 0,
            "tokens_per_sec": 0,
            "error": str(e),
            "folder": model_folder
        }
        update_summary(model_name, record)

if __name__ == "__main__":
    model = sys.argv[1] if len(sys.argv) > 1 else "qwen2.5-32b-instruct"
    run_single(model)
