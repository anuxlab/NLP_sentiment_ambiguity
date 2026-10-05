#!/usr/bin/env bash
# Reproduce the whole study on CPU (~40-90 min depending on the machine).
#   bash run_all.sh            # CLEAN run: deletes old results/models first (recommended)
#   bash run_all.sh --resume   # keep existing results and only run what is missing
set -e
cd "$(dirname "$0")/src"
if [ "$1" != "--resume" ]; then
  echo "[run_all] clean start: removing old results, checkpoints and predictions"
  rm -rf ../results/models ../results/preds ../results/results.jsonl ../results/efficiency.json \
         ../results/all_runs.csv ../results/summary.md ../results/families.jsonl ../results/ambiguity.json ../results/logs/*.jsonl ../results/logs/*.log
  mkdir -p ../results/models ../results/preds ../results/logs
fi
python pretrain_mlm.py --steps 5000 --lr 1e-3 --dropout 0.0 --save_steps 0,100,300,1000,2500,5000  # 1. MLM pre-training
python run_experiments.py --stage all                                                             # 2. LR pilot, curves, ckpt ablation, model families
python efficiency.py                                                                              # 3. size / latency / INT8
python make_assets.py                                                                             # 4. figures, tables, ambiguity analysis, summary.md
# Optional: two more domains with real (not proxy) ambiguity signals + cross-domain transfer matrix
# python run_experiments.py --stage all --domain finance
# python run_experiments.py --stage all --domain airline
# python backfill_cross_domain.py
# Optional (needs internet -> huggingface.co): real pre-trained compact models
# python hf_benchmark.py --models mini,small,minilm
