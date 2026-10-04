import os
import sys
import json
import time
import subprocess
import urllib.request
import urllib.error

# Force UTF-8 stdout and stderr for Windows console
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

PROMPT = (
    "what is total construction cost for building in texas 2 gigawath ai most advanced and secure "
    "data center (full expenses, investments, all materials, detailed econimic model in advanced top tier "
    "investment top tier firm form, all source prices, calculations in excel)"
)

MODELS = [
    "cerberus-4b-v2-abliterated",
    "google/gemma-4-e4b",
    "deepseek-r1-distill-qwen-14b",
    "qwen2.5-coder-14b-instruct",
    "mistral-small-24b-instruct-2501",
    "qwen2.5-32b-instruct",
    "llama-3.3-70b-instruct",
    "cerberus-v0.1",
]

BASE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "benchmark_results", "texas_2gw_datacenter")
SUMMARY_MD = os.path.join(BASE_DIR, "processing_times_summary.md")
SUMMARY_CSV = os.path.join(BASE_DIR, "processing_times_summary.csv")
SUMMARY_JSON = os.path.join(BASE_DIR, "processing_times_summary.json")

def sanitize_folder_name(name: str) -> str:
    return name.replace("/", "_").replace("\\", "_").replace(":", "_").replace(" ", "_")

def run_cmd(cmd_list, timeout=180):
    try:
        res = subprocess.run(
            cmd_list,
            capture_output=True,
            text=True,
            timeout=timeout,
            shell=True
        )
        return res.returncode == 0, res.stdout, res.stderr
    except subprocess.TimeoutExpired:
        return False, "", "Command timed out"
    except Exception as e:
        return False, "", str(e)

def unload_models():
    print("  -> Unloading all models from LM Studio memory...")
    run_cmd(["lms", "unload", "--all"])
    time.sleep(2)

def load_model(model_name: str):
    print(f"  -> Loading model '{model_name}'...")
    t0 = time.time()
    success, stdout, stderr = run_cmd(["lms", "load", model_name, "-y"], timeout=240)
    load_time = time.time() - t0
    if not success:
        print(f"     [!] Failed to load {model_name}: {stderr}")
        return False, load_time, stderr
    print(f"     [OK] Model loaded in {load_time:.2f}s")
    return True, load_time, ""

def query_model_streaming(model_name: str, prompt: str):
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
                "content": prompt
            }
        ],
        "temperature": 0.3,
        "max_tokens": 8192,
        "stream": True
    }
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=data,
        headers={"Content-Type": "application/json"}
    )

    t_start = time.time()
    t_first_token = None
    chunks_text = []
    reasoning_text = []
    token_count = 0

    try:
        with urllib.request.urlopen(req, timeout=900) as resp:
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
                    
                    # Capture reasoning content (e.g. DeepSeek R1)
                    r_content = delta.get("reasoning_content") or delta.get("reasoning")
                    if r_content:
                        if t_first_token is None:
                            t_first_token = time.time()
                        reasoning_text.append(r_content)
                        token_count += 1

                    # Capture actual message content
                    content = delta.get("content")
                    if content:
                        if t_first_token is None:
                            t_first_token = time.time()
                        chunks_text.append(content)
                        token_count += 1
                except Exception:
                    continue

        t_end = time.time()
        if t_first_token is None:
            t_first_token = t_end

        total_content = "".join(chunks_text)
        total_reasoning = "".join(reasoning_text)
        
        full_text = total_content
        if total_reasoning:
            full_text = f"<think>\n{total_reasoning}\n</think>\n\n" + total_content

        total_duration = t_end - t_start
        ttft = t_first_token - t_start
        gen_duration = t_end - t_first_token if t_end > t_first_token else 0.001
        tps = token_count / gen_duration if gen_duration > 0 else 0.0

        return {
            "success": True,
            "error": None,
            "response_text": full_text,
            "token_count": token_count,
            "char_count": len(full_text),
            "ttft_sec": round(ttft, 3),
            "gen_duration_sec": round(gen_duration, 3),
            "total_duration_sec": round(total_duration, 3),
            "tokens_per_sec": round(tps, 2)
        }

    except Exception as e:
        t_end = time.time()
        return {
            "success": False,
            "error": str(e),
            "response_text": "",
            "token_count": 0,
            "char_count": 0,
            "ttft_sec": 0,
            "gen_duration_sec": 0,
            "total_duration_sec": round(t_end - t_start, 3),
            "tokens_per_sec": 0
        }

def save_summary_files(all_records):
    os.makedirs(BASE_DIR, exist_ok=True)
    
    # Save JSON summary
    with open(SUMMARY_JSON, "w", encoding="utf-8") as f:
        json.dump(all_records, f, indent=2)

    # Save CSV summary
    with open(SUMMARY_CSV, "w", encoding="utf-8") as f:
        f.write("Model,Status,TTFT (s),Generation Time (s),Total Time (s),Tokens,Tokens/sec,Chars,Load Time (s),Error\n")
        for r in all_records:
            err = (r.get("error") or "").replace(",", " ")
            f.write(f'"{r["model"]}","{r["status"]}",{r["ttft_sec"]},{r["gen_duration_sec"]},{r["total_duration_sec"]},{r["token_count"]},{r["tokens_per_sec"]},{r["char_count"]},{r["load_time_sec"]},"{err}"\n')

    # Save Markdown summary
    with open(SUMMARY_MD, "w", encoding="utf-8") as f:
        f.write("# 📊 Benchmark Processing Times & Performance Summary\n\n")
        f.write("**Evaluation Task**: Construction Cost & Advanced Economic Model for a 2 Gigawatt AI Data Center in Texas\n\n")
        f.write(f"**Date**: {time.strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        f.write("| Model | Status | Load Time | TTFT | Gen Time | Total Time | Output Tokens | Speed (TPS) |\n")
        f.write("| :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: |\n")
        for r in all_records:
            status_icon = "✅ Success" if r["status"] == "SUCCESS" else f"❌ {r['status']}"
            f.write(
                f"| **`{r['model']}`** | {status_icon} | {r['load_time_sec']}s | {r['ttft_sec']}s | "
                f"{r['gen_duration_sec']}s | {r['total_duration_sec']}s | {r['token_count']} | **{r['tokens_per_sec']} tps** |\n"
            )
        f.write("\n---\n\n")
        f.write("### Sub-folder Results Directory\n")
        for r in all_records:
            if r.get("folder"):
                f.write(f"- [`{r['model']}`](./{os.path.basename(r['folder'])}/response.md)\n")

def main():
    os.makedirs(BASE_DIR, exist_ok=True)
    print("=" * 70)
    print("LM-WARS BENCHMARK RUNNER: 2GW TEXAS AI DATA CENTER ECONOMIC MODEL")
    print(f"Results Directory: {BASE_DIR}")
    print("=" * 70)

    all_records = []

    for idx, model_name in enumerate(MODELS, 1):
        print(f"\n[{idx}/{len(MODELS)}] Processing: {model_name}")
        model_folder = os.path.join(BASE_DIR, sanitize_folder_name(model_name))
        os.makedirs(model_folder, exist_ok=True)

        unload_models()

        loaded, load_time, load_err = load_model(model_name)
        if not loaded:
            record = {
                "model": model_name,
                "status": "LOAD_FAILED",
                "load_time_sec": round(load_time, 2),
                "ttft_sec": 0,
                "gen_duration_sec": 0,
                "total_duration_sec": round(load_time, 2),
                "token_count": 0,
                "char_count": 0,
                "tokens_per_sec": 0,
                "error": load_err,
                "folder": model_folder
            }
            all_records.append(record)
            save_summary_files(all_records)
            continue

        print("  -> Sending prompt and streaming response...")
        res = query_model_streaming(model_name, PROMPT)

        if res["success"]:
            print(f"     [OK] Completed: {res['token_count']} tokens in {res['gen_duration_sec']}s ({res['tokens_per_sec']} TPS)")
            # Save response.md
            response_file = os.path.join(model_folder, "response.md")
            with open(response_file, "w", encoding="utf-8") as f:
                f.write(f"# Economic Analysis & Construction Cost: 2GW Texas AI Data Center\n\n")
                f.write(f"**Model**: `{model_name}`  \n")
                f.write(f"**Generated**: {time.strftime('%Y-%m-%d %H:%M:%S')}  \n")
                f.write(f"**TTFT**: {res['ttft_sec']}s | **Total Duration**: {res['total_duration_sec']}s | **Tokens**: {res['token_count']} | **Speed**: {res['tokens_per_sec']} TPS  \n\n")
                f.write(f"### Prompt\n> {PROMPT}\n\n---\n\n")
                f.write(res["response_text"])

            # Save metrics.json
            metrics_file = os.path.join(model_folder, "metrics.json")
            with open(metrics_file, "w", encoding="utf-8") as f:
                json.dump({
                    "model": model_name,
                    "prompt": PROMPT,
                    "timestamp": time.time(),
                    "load_time_sec": round(load_time, 2),
                    **res
                }, f, indent=2)

            record = {
                "model": model_name,
                "status": "SUCCESS",
                "load_time_sec": round(load_time, 2),
                "ttft_sec": res["ttft_sec"],
                "gen_duration_sec": res["gen_duration_sec"],
                "total_duration_sec": res["total_duration_sec"],
                "token_count": res["token_count"],
                "char_count": res["char_count"],
                "tokens_per_sec": res["tokens_per_sec"],
                "error": None,
                "folder": model_folder
            }
        else:
            print(f"     [!] Generation error: {res['error']}")
            record = {
                "model": model_name,
                "status": "INFERENCE_FAILED",
                "load_time_sec": round(load_time, 2),
                "ttft_sec": 0,
                "gen_duration_sec": 0,
                "total_duration_sec": res["total_duration_sec"],
                "token_count": 0,
                "char_count": 0,
                "tokens_per_sec": 0,
                "error": res["error"],
                "folder": model_folder
            }

        all_records.append(record)
        save_summary_files(all_records)
        unload_models()

    print("\n" + "=" * 70)
    print("BENCHMARK COMPLETED FOR ALL MODELS!")
    print(f"Summary written to: {SUMMARY_MD}")
    print("=" * 70)

if __name__ == "__main__":
    main()
