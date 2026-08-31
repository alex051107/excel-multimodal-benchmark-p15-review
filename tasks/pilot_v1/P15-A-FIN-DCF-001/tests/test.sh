#!/usr/bin/env sh
set -u

# Harbor's separate verifier requires an authoritative numeric reward file.
# Initialize it to zero on every run so evaluator or JSON failures cannot reuse
# a stale/agent-authored reward, then promote only a freshly parsed score.
if [ -f /tests/evaluate.py ]; then
  reward_dir=/logs/verifier
  result_path="$reward_dir/judge-result.json"
  stderr_path="$reward_dir/judge-stderr.txt"
  reward_tmp="$reward_dir/reward.txt.tmp"
  mkdir -p "$reward_dir"
  printf '0\n' > "$reward_dir/reward.txt"
  : > "$stderr_path"

  if /usr/local/bin/python /tests/evaluate.py /app/output/answer.xlsx --split dev > "$result_path" 2> "$stderr_path"; then
    cat "$result_path"
    if /usr/local/bin/python -c 'import json, math, pathlib, sys; value=float(json.loads(pathlib.Path(sys.argv[1]).read_text())["normalized_score"]); assert math.isfinite(value) and 0.0 <= value <= 1.0; print(value)' "$result_path" > "$reward_tmp" 2>> "$stderr_path"; then
      mv "$reward_tmp" "$reward_dir/reward.txt"
    fi
  else
    cat "$stderr_path" >&2
  fi
  exit 0
fi

exec python3 "$(dirname "$0")/evaluate.py" "$(dirname "$0")/../solution/reference.xlsx" --split dev
