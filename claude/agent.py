#!/usr/bin/env python3
"""
Claude Local Agent

Iterates over domain directories, combines instructions.md + work.md into
a prompt, runs it through the `claude` CLI, handles rate limits gracefully,
and resets work.md on successful completion.

Usage:
    ./agent.py                  # run forever
    ./agent.py --once           # process all domains once and exit
    ./agent.py --domain MarketResearch  # run a single domain once
"""

from __future__ import annotations

import argparse
import os
import re
import shutil
import subprocess
import sys
import threading
import time
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError


# ─── Configuration ─────────────────────────────────────────────────────────

BASE_DIR = Path(__file__).resolve().parent
LOG_FILE = BASE_DIR / "agent.log"

DOMAINS = ["MarketResearch", "SystemDesign", "Programming-DSA", "General"]

# Allowed tools passed to `claude --allowedTools`
TOOLS = "Read,Write,Bash,WebFetch,WebSearch"

# Seconds between full scans of all domain directories
POLL_EVERY = 300

# Seconds to sleep if rate limit is hit but reset time cannot be parsed
RATE_LIMIT_FALLBACK_SLEEP = 3600

# Extra buffer added to a parsed rate-limit reset time
RATE_LIMIT_BUFFER_SECONDS = 60

# Seconds of silence from the claude process before logging an idle warning
IDLE_WARN_SECONDS = 120

# Patterns that indicate a Claude rate-limit response
RATE_LIMIT_PATTERNS = [
    r"hit your limit",
    r"usage limit",
    r"rate limit",
]

# Patterns that indicate a transient/retriable error (preserve work.md, retry next cycle)
TRANSIENT_ERROR_PATTERNS = [
    r"stream idle timeout",
    r"partial response received",
    r"connection (reset|refused|timed? ?out)",
    r"api error",
    r"network error",
    r"503\b",
    r"502\b",
    r"529\b",
]

# Regex for extracting "resets 7pm" / "resets 10:30am" from error text
RESET_TIME_RE = re.compile(
    r"resets\s+(\d{1,2}(?::\d{2})?)\s*(am|pm)",
    re.IGNORECASE,
)

# Regex for extracting an IANA timezone like "(America/Los_Angeles)"
TIMEZONE_RE = re.compile(r"\(([A-Za-z_]+/[A-Za-z_]+)\)")

NO_ACTIVE_TASK_RE = re.compile(r"no active task", re.IGNORECASE)


# ─── Logging ───────────────────────────────────────────────────────────────


def log(msg: str) -> None:
    """Log a timestamped message to stdout and the log file."""
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{stamp}] {msg}"
    print(line, flush=True)
    try:
        with LOG_FILE.open("a") as f:
            f.write(line + "\n")
    except OSError as e:
        print(f"[{stamp}] WARN: failed to write log file: {e}", file=sys.stderr)


# ─── Rate limit parsing ────────────────────────────────────────────────────


def is_rate_limited(text: str) -> bool:
    return any(re.search(p, text, re.IGNORECASE) for p in RATE_LIMIT_PATTERNS)


def is_transient_error(text: str) -> bool:
    return any(re.search(p, text, re.IGNORECASE) for p in TRANSIENT_ERROR_PATTERNS)


def compute_reset_sleep(text: str) -> Optional[int]:
    """
    Parse a rate-limit message and return seconds to sleep until the reset.

    Handles messages like:
        "You've hit your limit · resets 7pm (America/Los_Angeles)"
        "You've hit your limit · resets 10:30am (America/New_York)"

    Returns None if the message cannot be parsed.
    """
    match = RESET_TIME_RE.search(text)
    if not match:
        return None

    time_str, meridiem = match.group(1), match.group(2).lower()
    if ":" not in time_str:
        time_str = f"{time_str}:00"

    # Resolve timezone (fall back to system local)
    tz_match = TIMEZONE_RE.search(text)
    tz: Optional[ZoneInfo]
    if tz_match:
        try:
            tz = ZoneInfo(tz_match.group(1))
        except ZoneInfoNotFoundError:
            tz = None
    else:
        tz = None

    now = datetime.now(tz) if tz else datetime.now().astimezone()

    try:
        parsed = datetime.strptime(f"{time_str}{meridiem}", "%I:%M%p").time()
    except ValueError:
        return None

    target = now.replace(
        hour=parsed.hour, minute=parsed.minute, second=0, microsecond=0
    )
    if target <= now:
        target += timedelta(days=1)

    seconds = int((target - now).total_seconds()) + RATE_LIMIT_BUFFER_SECONDS
    return seconds


# ─── Domain processing ─────────────────────────────────────────────────────


@dataclass
class DomainResult:
    ok: bool
    rate_limited: bool = False
    transient_error: bool = False
    sleep_seconds: Optional[int] = None
    exit_code: int = 0
    output: str = ""


@dataclass
class Domain:
    name: str
    base_dir: Path = field(default_factory=lambda: BASE_DIR)

    @property
    def dir(self) -> Path:
        return self.base_dir / self.name

    @property
    def instructions_path(self) -> Path:
        return self.dir / "INSTRUCTIONS.md"

    @property
    def work_path(self) -> Path:
        return self.dir / "WORK.md"

    @property
    def claude_md_path(self) -> Path:
        return self.dir / "CLAUDE.md"

    @property
    def knowledgebase_path(self) -> Path:
        return self.dir / "KNOWLEDGEBASE.md"

    @property
    def tmp_dir(self) -> Path:
        return self.dir / "tmp"

    @property
    def output_dir(self) -> Path:
        return self.dir / "output"

    def ensure_dirs(self) -> None:
        self.tmp_dir.mkdir(parents=True, exist_ok=True)
        self.output_dir.mkdir(parents=True, exist_ok=True)

    def has_active_task(self) -> bool:
        if not self.instructions_path.is_file() or not self.work_path.is_file():
            return False
        content = self.work_path.read_text().strip()
        if not content:
            return False
        return not NO_ACTIVE_TASK_RE.search(content)

    def build_prompt(self) -> str:
        parts: list[str] = [f"You are working in the directory: {self.dir}"]

        if self.claude_md_path.is_file():
            parts.append(f"## Domain Guidance\n{self.claude_md_path.read_text()}")

        if self.knowledgebase_path.is_file():
            parts.append(f"## Knowledge Base\n{self.knowledgebase_path.read_text()}")

        parts.append(f"## Instructions\n{self.instructions_path.read_text()}")
        parts.append(f"## Current Work\n{self.work_path.read_text()}")
        parts.append(
            f"## Important\n"
            f"- Save final output files (docx, xlsx, pdf) to: {self.output_dir}/\n"
            f"- Save intermediate/temporary files (python scripts, etc.) to: {self.tmp_dir}/\n"
            f"- Work within the scope defined above.\n"
            f"- Be thorough and follow the instructions precisely.\n"
        )

        return "\n\n".join(parts)

    def reset_work(self) -> None:
        self.work_path.write_text("(No active task)\n")


# ─── Claude invocation ─────────────────────────────────────────────────────


def run_claude(prompt: str) -> tuple[int, str]:
    """
    Run the `claude` CLI with the given prompt.

    Streams stdout+stderr line-by-line to the terminal and log file in real
    time.  A background watchdog thread logs a warning if the process produces
    no output for IDLE_WARN_SECONDS seconds, so silent hangs are visible.

    Returns (exit_code, combined_output).
    """
    cmd = [
        "claude",
        "--print",
        "--allowedTools", TOOLS,
        "--effort", "max",
        "--dangerously-skip-permissions",
    ]
    try:
        proc = subprocess.Popen(
            cmd,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
        )
    except FileNotFoundError:
        return 127, "ERROR: `claude` CLI not found in PATH."

    assert proc.stdin is not None
    proc.stdin.write(prompt)
    proc.stdin.close()

    lines: list[str] = []
    last_output_at = time.monotonic()
    idle_warned = False

    def _watchdog() -> None:
        nonlocal idle_warned
        while proc.poll() is None:
            time.sleep(15)
            idle = time.monotonic() - last_output_at
            if idle >= IDLE_WARN_SECONDS and not idle_warned:
                log(
                    f"  [idle {int(idle)}s] claude is still running (PID {proc.pid}) "
                    f"but has produced no output — may be waiting on a tool call or API"
                )
                idle_warned = True

    watchdog = threading.Thread(target=_watchdog, daemon=True)
    watchdog.start()

    assert proc.stdout is not None
    for raw_line in proc.stdout:
        last_output_at = time.monotonic()
        idle_warned = False
        lines.append(raw_line)
        # Stream to terminal and log file immediately
        print(raw_line, end="", flush=True)
        try:
            with LOG_FILE.open("a") as fh:
                fh.write(raw_line)
        except OSError:
            pass

    proc.wait()
    watchdog.join(timeout=0)
    return proc.returncode, "".join(lines)


def check_dependencies() -> None:
    if shutil.which("claude") is None:
        print("ERROR: Claude Code is not installed.", file=sys.stderr)
        print("", file=sys.stderr)
        print("Install it with one of:", file=sys.stderr)
        print("  curl -fsSL https://claude.ai/install.sh | bash", file=sys.stderr)
        print("  npm install -g @anthropic-ai/claude-code", file=sys.stderr)
        print("", file=sys.stderr)
        print("Then run:  claude   (to log in once)", file=sys.stderr)
        sys.exit(1)


# ─── Main processing ───────────────────────────────────────────────────────


def process_domain(domain: Domain) -> DomainResult:
    log("══════════════════════════════════════")
    log(f"Domain: {domain.name}")
    log("──────────────────────────────────────")

    prompt = domain.build_prompt()
    exit_code, output = run_claude(prompt)

    if is_rate_limited(output):
        log(f"Rate limit detected for {domain.name} — preserving work.md")
        sleep_seconds = compute_reset_sleep(output)
        return DomainResult(
            ok=False,
            rate_limited=True,
            sleep_seconds=sleep_seconds,
            exit_code=exit_code,
            output=output,
        )

    if exit_code != 0 and is_transient_error(output):
        log(f"Transient error for {domain.name} (exit={exit_code}) — preserving work.md for retry")
        return DomainResult(
            ok=False,
            transient_error=True,
            exit_code=exit_code,
            output=output,
        )

    if exit_code == 0:
        log(f"Success: {domain.name}")
    else:
        log(f"WARNING: Claude exited with errors for {domain.name} (exit={exit_code})")

    domain.reset_work()
    log(f"Reset work.md for {domain.name}")

    return DomainResult(ok=(exit_code == 0), exit_code=exit_code, output=output)


def handle_rate_limit(result: DomainResult) -> None:
    sleep_seconds = result.sleep_seconds
    if sleep_seconds and sleep_seconds > 0:
        wake = datetime.now() + timedelta(seconds=sleep_seconds)
        log(
            f"Sleeping {sleep_seconds}s until "
            f"{wake.strftime('%Y-%m-%d %H:%M:%S')} (rate limit reset)"
        )
        time.sleep(sleep_seconds)
    else:
        log(f"Could not parse reset time; sleeping {RATE_LIMIT_FALLBACK_SLEEP}s as fallback")
        time.sleep(RATE_LIMIT_FALLBACK_SLEEP)


def run_cycle(domains: list[Domain]) -> None:
    """Process all active domains once. Stops early on rate limit."""
    for domain in domains:
        if not domain.has_active_task():
            continue
        result = process_domain(domain)
        if result.rate_limited:
            handle_rate_limit(result)
            return  # Restart the outer loop fresh


def main() -> None:
    parser = argparse.ArgumentParser(description="Claude local agent")
    parser.add_argument(
        "--once",
        action="store_true",
        help="Process all domains once and exit",
    )
    parser.add_argument(
        "--domain",
        choices=DOMAINS,
        help="Run only a single domain once, then exit",
    )
    parser.add_argument(
        "--poll",
        type=int,
        default=POLL_EVERY,
        help=f"Seconds between scans (default: {POLL_EVERY})",
    )
    args = parser.parse_args()

    check_dependencies()

    domains = [Domain(name) for name in DOMAINS]
    for d in domains:
        d.ensure_dirs()
    LOG_FILE.touch(exist_ok=True)

    log("Claude Agent started")
    log(f"  Base dir : {BASE_DIR}")
    log(f"  Domains  : {', '.join(DOMAINS)}")
    log(f"  Log file : {LOG_FILE}")
    log(f"  Polling every {args.poll}s  |  Ctrl-C to stop")
    print()

    if args.domain:
        target = next(d for d in domains if d.name == args.domain)
        if not target.has_active_task():
            log(f"{target.name} has no active task. Nothing to do.")
            return
        result = process_domain(target)
        if result.rate_limited:
            handle_rate_limit(result)
        return

    if args.once:
        run_cycle(domains)
        return

    try:
        while True:
            run_cycle(domains)
            log(f"Sleeping {args.poll}s...")
            time.sleep(args.poll)
    except KeyboardInterrupt:
        log("Interrupted — exiting.")


if __name__ == "__main__":
    main()
