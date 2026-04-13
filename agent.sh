#!/bin/bash
# ─────────────────────────────────────────────────────────────
#  Claude Local Agent  –  macOS
#  Iterates over domain directories, combines instructions.md
#  + work.md, runs through Claude Code, archives completed work.
# ─────────────────────────────────────────────────────────────
#
echo "Try Python program"
exit

BASE_DIR="$(cd "$(dirname "$0")" && pwd)"
LOG_FILE="$BASE_DIR/agent.log"
POLL_EVERY=300   # seconds between scans

# Domain directories to process
DOMAINS=("MarketResearch" "SystemDesign" "Programming-DSA" "General")

# Allowed tools: Read, Write, Bash, WebFetch, WebSearch
TOOLS="Read,Write,Bash,WebFetch,WebSearch"

# ── Helpers ──────────────────────────────────────────────────

log() {
  local msg="[$(date '+%Y-%m-%d %H:%M:%S')] $1"
  echo "$msg"
  echo "$msg" >> "$LOG_FILE"
}

check_dependencies() {
  if ! command -v claude &> /dev/null; then
    echo "ERROR: Claude Code is not installed."
    echo ""
    echo "Install it with one of:"
    echo "  curl -fsSL https://claude.ai/install.sh | bash   (recommended, no Node needed)"
    echo "  npm install -g @anthropic-ai/claude-code          (requires Node 18+)"
    echo ""
    echo "Then run:  claude   (to log in once)"
    exit 1
  fi
}

# Parse a rate-limit message like:
#   "You've hit your limit · resets 7pm (America/Los_Angeles)"
# Echoes the number of seconds to sleep until the reset time (+60s buffer).
# Returns non-zero if the message can't be parsed.
compute_reset_sleep() {
  local text="$1"
  local time_str tz normalized target_epoch now_epoch sleep_seconds

  time_str=$(echo "$text" | grep -oiE 'resets [0-9]{1,2}(:[0-9]{2})?[[:space:]]*(am|pm)' | head -1 | sed -E 's/^resets[[:space:]]+//; s/[[:space:]]+//g' | tr '[:upper:]' '[:lower:]')
  [ -z "$time_str" ] && return 1

  tz=$(echo "$text" | grep -oE '\([A-Za-z_]+/[A-Za-z_]+\)' | head -1 | tr -d '()')
  [ -z "$tz" ] && tz="$(date +%Z)"

  # Normalize "7pm" → "7:00pm"
  if [[ "$time_str" != *:* ]]; then
    normalized=$(echo "$time_str" | sed -E 's/^([0-9]+)(am|pm)$/\1:00\2/')
  else
    normalized="$time_str"
  fi

  # Parse today's date + target time in the given timezone
  local today
  today=$(TZ="$tz" date '+%Y-%m-%d')
  target_epoch=$(TZ="$tz" date -j -f '%Y-%m-%d %I:%M%p' "${today} ${normalized}" '+%s' 2>/dev/null)
  [ -z "$target_epoch" ] && return 1

  now_epoch=$(date '+%s')
  if [ "$target_epoch" -le "$now_epoch" ]; then
    target_epoch=$((target_epoch + 86400))
  fi

  sleep_seconds=$((target_epoch - now_epoch + 60))
  echo "$sleep_seconds"
}

# ── Setup ─────────────────────────────────────────────────────

source ~/set-nvm.sh

check_dependencies

for domain in "${DOMAINS[@]}"; do
  mkdir -p "$BASE_DIR/$domain/tmp" "$BASE_DIR/$domain/output"
done
touch "$LOG_FILE"

log "Claude Agent started"
log "  Base dir : $BASE_DIR"
log "  Domains  : ${DOMAINS[*]}"
log "  Log file : $LOG_FILE"
log "  Polling every ${POLL_EVERY}s  |  Ctrl-C to stop"
echo ""

# ── Main loop ─────────────────────────────────────────────────

while true; do

  for domain in "${DOMAINS[@]}"; do
    domain_dir="$BASE_DIR/$domain"
    instructions="$domain_dir/instructions.md"
    work="$domain_dir/work.md"

    # Skip if no work.md or no instructions.md
    [ -f "$instructions" ] || continue
    [ -f "$work" ] || continue

    # Skip if work.md has no active task (contains "No active task")
    if grep -qi "no active task" "$work"; then
      continue
    fi

    log "══════════════════════════════════════"
    log "Domain: $domain"
    log "──────────────────────────────────────"

    # Combine instructions + work into a single prompt
    prompt="$(cat <<PROMPT
You are working in the directory: $domain_dir

## Instructions
$(cat "$instructions")

## Current Work
$(cat "$work")

## Important
- Save final output files (docx, xlsx, pdf) to: $domain_dir/output/
- Save intermediate/temporary files (python scripts, etc.) to: $domain_dir/tmp/
- Work within the scope defined above.
- Be thorough and follow the instructions precisely.
PROMPT
)"

    # Run Claude Code non-interactively. Capture combined output so we
    # can detect rate-limit messages before resetting work.md.
    claude_output=$(echo "$prompt" | claude \
        --print \
        --allowedTools "$TOOLS" \
        --effort max \
        --dangerously-skip-permissions \
        2>&1)
    claude_exit=$?
    echo "$claude_output" >> "$LOG_FILE"

    # Detect rate-limit: don't reset work.md, sleep until the reset time.
    if echo "$claude_output" | grep -qiE "hit your limit|usage limit|rate limit"; then
      log "Rate limit detected for $domain — preserving work.md"
      sleep_seconds=$(compute_reset_sleep "$claude_output")
      if [ -n "$sleep_seconds" ] && [ "$sleep_seconds" -gt 0 ]; then
        wake_time=$(date -r "$(( $(date +%s) + sleep_seconds ))" '+%Y-%m-%d %H:%M:%S')
        log "Sleeping ${sleep_seconds}s until ${wake_time} (rate limit reset)"
        sleep "$sleep_seconds"
      else
        log "Could not parse reset time; sleeping 1h as fallback"
        sleep 3600
      fi
      # Skip remaining domains this cycle; restart the outer loop.
      break
    fi

    if [ "$claude_exit" -eq 0 ]; then
      log "Success: $domain"
    else
      log "WARNING: Claude exited with errors for $domain (exit=$claude_exit)"
    fi

    # Reset work.md to empty state (only on normal completion)
    cat > "$work" <<'EOF'
(No active task)
EOF
    log "Reset work.md for $domain"

  done

  log "Sleeping ${POLL_EVERY}s..."
  sleep "$POLL_EVERY"
done
