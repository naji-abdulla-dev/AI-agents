#!/usr/bin/env python3
"""
Claude Local Agent

Supports two modes per domain:

  Single-phase (default): one session does research, writes scripts, AND runs them.
  Two-phase (opt-in):     Phase 1 = research + write scripts only.
                          Phase 2 = run scripts + verify output.

Two-phase mode activates when CLAUDE_RUN.md exists in the domain directory.
Phase 1 writes work_run.md on success; Phase 2 reads it, executes, then clears it.

After every attempt, the agent logs: status, wall-clock duration, and token usage
(input / output / cache_read / cost) extracted from the stream-json result event.

Usage:
    ./agent.py                  # run forever
    ./agent.py --once           # process all domains once and exit
    ./agent.py --domain MarketResearch  # run a single domain once (both phases)
"""

from __future__ import annotations

import argparse
import json
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

# Load a .env file from the agent directory if present.
# Supports ANTHROPIC_API_KEY so the agent can use the Anthropic API directly
# instead of the claude.ai subscription, bypassing claude.ai rate limits.
_env_file = BASE_DIR / ".env"
if _env_file.is_file():
    for _line in _env_file.read_text().splitlines():
        _line = _line.strip()
        if _line and not _line.startswith("#") and "=" in _line:
            _k, _, _v = _line.partition("=")
            os.environ.setdefault(_k.strip(), _v.strip())

DOMAINS = ["MarketResearch", "SystemDesign", "Programming-DSA", "General"]

TOOLS = "Read,Write,Bash,WebFetch,WebSearch"

# Model to use — override with CLAUDE_AGENT_MODEL env var.
# "sonnet" and "opus" resolve to the latest versions automatically.
MODEL = os.getenv("CLAUDE_AGENT_MODEL", "sonnet")

POLL_EVERY = 300          # seconds between full domain scans
IDLE_WARN_SECONDS = 120   # log a warning after this many seconds of silence
IDLE_KILL_SECONDS = 600   # kill the process after this many seconds of silence
                          # 300s was too short: stream-json buffers the entire tool_use
                          # payload before emitting it, so writing a large script
                          # (~1000 lines) can produce 5+ minutes of legitimate silence
MAX_RUN_SECONDS = 1800     # hard cap per invocation (30 min)
MAX_TURNS_RESEARCH = 25   # research sessions: web fetches only, low turn budget
MAX_TURNS_SCRIPT = 60     # script-writing sessions: needs iterations for large files
MAX_TURNS = 60            # default for single-phase domains

RATE_LIMIT_FALLBACK_SLEEP = 3600
RATE_LIMIT_BUFFER_SECONDS = 60

TRANSIENT_BASE_SLEEP = 60    # first retry after a transient error
TRANSIENT_MAX_SLEEP = 600    # cap for exponential backoff

RATE_LIMIT_PATTERNS = [
    r"hit your limit",
    r"usage limit",
    r"rate limit",
]

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

RESET_TIME_RE = re.compile(r"resets\s+(\d{1,2}(?::\d{2})?)\s*(am|pm)", re.IGNORECASE)
TIMEZONE_RE = re.compile(r"\(([A-Za-z_]+/[A-Za-z_]+)\)")
NO_ACTIVE_TASK_RE = re.compile(r"no active task", re.IGNORECASE)
NO_PENDING_RUN_RE = re.compile(r"no pending run", re.IGNORECASE)


# ─── Logging ───────────────────────────────────────────────────────────────


def log(msg: str) -> None:
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{stamp}] {msg}"
    print(line, flush=True)
    try:
        with LOG_FILE.open("a") as f:
            f.write(line + "\n")
    except OSError as e:
        print(f"[{stamp}] WARN: failed to write log: {e}", file=sys.stderr)


# ─── Error classification ──────────────────────────────────────────────────


def is_rate_limited(text: str) -> bool:
    return any(re.search(p, text, re.IGNORECASE) for p in RATE_LIMIT_PATTERNS)


def is_transient_error(text: str) -> bool:
    return any(re.search(p, text, re.IGNORECASE) for p in TRANSIENT_ERROR_PATTERNS)


def compute_reset_sleep(text: str) -> Optional[int]:
    match = RESET_TIME_RE.search(text)
    if not match:
        return None
    time_str, meridiem = match.group(1), match.group(2).lower()
    if ":" not in time_str:
        time_str = f"{time_str}:00"
    tz_match = TIMEZONE_RE.search(text)
    tz: Optional[ZoneInfo] = None
    if tz_match:
        try:
            tz = ZoneInfo(tz_match.group(1))
        except ZoneInfoNotFoundError:
            pass
    now = datetime.now(tz) if tz else datetime.now().astimezone()
    try:
        parsed = datetime.strptime(f"{time_str}{meridiem}", "%I:%M%p").time()
    except ValueError:
        return None
    target = now.replace(hour=parsed.hour, minute=parsed.minute, second=0, microsecond=0)
    if target <= now:
        target += timedelta(days=1)
    return int((target - now).total_seconds()) + RATE_LIMIT_BUFFER_SECONDS


def compute_transient_sleep(retries: int) -> int:
    return min(TRANSIENT_BASE_SLEEP * (2 ** retries), TRANSIENT_MAX_SLEEP)


# ─── Run metrics ───────────────────────────────────────────────────────────


@dataclass
class RunMetrics:
    """Token usage and timing captured from the stream-json result event."""
    duration_s: float = 0.0
    input_tokens: int = 0
    output_tokens: int = 0
    cache_read_tokens: int = 0
    cache_creation_tokens: int = 0
    cost_usd: float = 0.0
    num_turns: int = 0

    @property
    def has_usage(self) -> bool:
        return self.input_tokens > 0 or self.output_tokens > 0


def _handle_stream_event(event: dict, metrics: RunMetrics) -> None:
    """
    Process one stream-json event.

    - Prints human-readable content (assistant text and tool calls) in real time.
    - Populates metrics from the final "result" event.
    """
    etype = event.get("type")

    if etype == "assistant":
        for block in event.get("message", {}).get("content", []):
            btype = block.get("type")
            if btype == "text":
                text = block.get("text", "")
                if text:
                    print(text, end="", flush=True)
            elif btype == "tool_use":
                tool_name = block.get("name", "?")
                inp = block.get("input", {})
                # Pick the most descriptive argument for each tool type
                if tool_name in ("WebSearch",):
                    arg = inp.get("query", "")
                elif tool_name in ("WebFetch",):
                    arg = inp.get("url", "")
                elif tool_name in ("Read", "Write"):
                    arg = inp.get("file_path", "")
                elif tool_name == "Bash":
                    arg = inp.get("command", "")
                else:
                    arg = next(iter(inp.values()), "") if inp else ""
                arg_str = str(arg)[:80]
                print(f"\n[{tool_name}: {arg_str!r}]", flush=True)

    elif etype == "result":
        usage = event.get("usage", {})
        metrics.input_tokens = usage.get("input_tokens", 0)
        metrics.output_tokens = usage.get("output_tokens", 0)
        metrics.cache_read_tokens = usage.get("cache_read_input_tokens", 0)
        metrics.cache_creation_tokens = usage.get("cache_creation_input_tokens", 0)
        metrics.cost_usd = event.get("total_cost_usd", 0.0)
        metrics.num_turns = event.get("num_turns", 0)

    # system / user / other event types: no display needed


def _log_metrics(label: str, status: str, metrics: RunMetrics, killed: bool) -> None:
    mins, secs = divmod(int(metrics.duration_s), 60)
    duration_str = f"{mins}m {secs}s" if mins else f"{secs}s"

    if killed or not metrics.has_usage:
        log(f"  [{label}] {status} | {duration_str} | tokens=unavailable (process killed/timed out)")
        return

    cost_str = f"${metrics.cost_usd:.4f}" if metrics.cost_usd else "n/a"
    cache_str = (
        f" cache_read={metrics.cache_read_tokens:,}"
        if metrics.cache_read_tokens else ""
    )
    log(
        f"  [{label}] {status} | {duration_str} | "
        f"in={metrics.input_tokens:,} out={metrics.output_tokens:,}"
        f"{cache_str} | cost={cost_str} | turns={metrics.num_turns}"
    )


# ─── Domain ────────────────────────────────────────────────────────────────


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
    def work_path(self) -> Path:
        return self.dir / "work.md"

    @property
    def work_run_path(self) -> Path:
        """Phase 2 task list — populated by Phase 1 on success."""
        return self.dir / "work_run.md"

    @property
    def claude_md_path(self) -> Path:
        return self.dir / "CLAUDE.md"

    @property
    def claude_run_md_path(self) -> Path:
        """Phase 2 instructions — its presence activates two-phase mode."""
        return self.dir / "CLAUDE_RUN.md"

    @property
    def instructions_path(self) -> Path:
        return self.dir / "INSTRUCTIONS.md"

    @property
    def knowledgebase_path(self) -> Path:
        return self.dir / "KNOWLEDGEBASE.md"

    @property
    def templates_dir(self) -> Path:
        """Curated reference scripts — Claude reads one as a structural starting point."""
        return self.dir / "templates"

    @property
    def tmp_dir(self) -> Path:
        return self.dir / "tmp"

    @property
    def output_dir(self) -> Path:
        return self.dir / "output"

    def ensure_dirs(self) -> None:
        self.tmp_dir.mkdir(parents=True, exist_ok=True)
        self.output_dir.mkdir(parents=True, exist_ok=True)

    def is_two_phase(self) -> bool:
        return self.claude_run_md_path.is_file()

    def has_active_task(self) -> bool:
        if not self.work_path.is_file():
            return False
        content = self.work_path.read_text().strip()
        return bool(content) and not NO_ACTIVE_TASK_RE.search(content)

    def has_pending_run(self) -> bool:
        if not self.work_run_path.is_file():
            return False
        content = self.work_run_path.read_text().strip()
        return bool(content) and not NO_PENDING_RUN_RE.search(content)

    def reset_work(self) -> None:
        self.work_path.write_text("(No active task)\n")

    def reset_run(self) -> None:
        self.work_run_path.write_text("(No pending runs)\n")

    def _tickers(self) -> list[str]:
        """Parse ticker symbols from work.md."""
        return [
            t.strip("()").upper()
            for t in self.work_path.read_text().split()
            if t.strip("()") and not NO_ACTIVE_TASK_RE.search(t)
        ]

    def _checkpoint_files(self) -> list[Path]:
        """Return existing research checkpoint files for the current tickers."""
        if not self.tmp_dir.exists():
            return []
        return [
            self.tmp_dir / f"{t.lower()}_data.md"
            for t in self._tickers()
            if (self.tmp_dir / f"{t.lower()}_data.md").exists()
        ]

    def _is_script_mode(self) -> bool:
        """
        True when every pending ticker already has a research checkpoint saved.
        In this mode the session skips all web research and only writes the script.
        """
        tickers = self._tickers()
        if not tickers:
            return False
        checkpointed = {f.stem.replace("_data", "").upper() for f in self._checkpoint_files()}
        return all(t in checkpointed for t in tickers)

    def _template_hint(self) -> str:
        if not self.templates_dir.exists():
            return ""
        templates = sorted(self.templates_dir.glob("*.py"))
        if not templates:
            return ""
        names = ", ".join(t.name for t in templates)
        return (
            f"- Read the template in {self.templates_dir}/ ({names}) "
            f"and use it as the structural starting point — copy the helpers "
            f"and tab scaffold, then fill in the data.\n"
        )

    def build_prompt(self) -> str:
        """
        Phase 1 prompt.

        Automatically selects one of two modes based on filesystem state:

        Research mode  — no checkpoint exists yet.
                         Do web research, save tmp/<ticker>_data.md, then stop.
                         The script is written in the next session.

        Script mode    — checkpoint exists for every pending ticker.
                         Read the checkpoint, write the script. No web fetches.
        """
        if self._is_script_mode():
            return self._build_script_prompt()
        return self._build_research_prompt()

    def _build_research_prompt(self) -> str:
        """Collect financial data and save a checkpoint. Do NOT write the script yet."""
        parts: list[str] = [f"You are working in the directory: {self.dir}"]

        if self.claude_md_path.is_file():
            parts.append(f"## Domain Guidance\n{self.claude_md_path.read_text()}")

        ref_files: list[str] = []
        if self.knowledgebase_path.is_file():
            ref_files.append(f"  - Knowledge base   : {self.knowledgebase_path}")
        if self.instructions_path.is_file():
            ref_files.append(f"  - Full instructions: {self.instructions_path}")
        if ref_files:
            parts.append(
                "## Reference Files (read via Read tool if needed)\n"
                + "\n".join(ref_files)
            )

        parts.append(f"## Current Work\n{self.work_path.read_text()}")

        parts.append(
            "## Important\n"
            f"- Save Python scripts to: {self.tmp_dir}/\n"
            f"- Save output files to: {self.output_dir}/\n"
            + "- THIS SESSION: research only. Do NOT write the analysis script — "
            "a separate session handles that.\n"
            + "- Web fetch budget: maximum 6 WebSearch and 4 WebFetch calls. "
            "Prioritise official filings and reliable aggregators "
            "(stockanalysis.com, macrotrends.net, SEC EDGAR) over many sources.\n"
            + "- Save to tmp/<ticker>_data.md INCREMENTALLY — write a partial "
            "checkpoint after collecting each major section (income statement, "
            "balance sheet, management, valuation). Do NOT wait until all research "
            "is done; partial data saved early survives a rate-limit interruption.\n"
            + "- The checkpoint must contain every figure needed to write the full "
            "analysis script without any further web access.\n"
        )

        return "\n\n".join(parts)

    def _build_script_prompt(self) -> str:
        """Write the analysis script from saved checkpoints. No web access needed."""
        checkpoints = self._checkpoint_files()
        parts: list[str] = [f"You are working in the directory: {self.dir}"]

        if self.claude_md_path.is_file():
            parts.append(f"## Domain Guidance\n{self.claude_md_path.read_text()}")

        parts.append(
            "## Research Checkpoints\n"
            "All financial data has already been collected and saved:\n"
            + "\n".join(f"  - {f}" for f in checkpoints)
            + "\n\nRead these files — they contain everything needed to write the script."
        )

        parts.append(f"## Current Work\n{self.work_path.read_text()}")

        phase_note = (
            "- PHASE 1 OF 2: Write the script only. "
            "Do NOT run it — a separate session handles execution.\n"
            if self.is_two_phase() else ""
        )

        parts.append(
            "## Important\n"
            f"- Save Python scripts to: {self.tmp_dir}/\n"
            f"- Save output files to: {self.output_dir}/\n"
            + self._template_hint()
            + phase_note
            + "- Do NOT do any web searches or fetches — all data is in the checkpoint file(s).\n"
            + "- Write the complete analysis script using only the checkpoint data.\n"
        )

        return "\n\n".join(parts)

    def build_run_prompt(self) -> str:
        """Phase 2 prompt: run scripts and verify output. No research."""
        parts: list[str] = [f"You are working in the directory: {self.dir}"]

        if self.claude_run_md_path.is_file():
            parts.append(f"## Run Instructions\n{self.claude_run_md_path.read_text()}")

        parts.append(f"## Scripts to Execute\n{self.work_run_path.read_text()}")

        parts.append(
            "## Important\n"
            f"- Scripts are in: {self.tmp_dir}/\n"
            f"- Output files go to: {self.output_dir}/\n"
            + "- PHASE 2 OF 2: Run the scripts and verify output only. Do NOT do new research.\n"
            + "- Fix runtime errors in the script if needed, but do not change the underlying data.\n"
        )

        return "\n\n".join(parts)


# ─── Claude invocation ─────────────────────────────────────────────────────


def run_claude(prompt: str, api_key: str | None = None, max_turns: int = MAX_TURNS) -> tuple[int, str, bool, RunMetrics]:
    """
    Run the `claude` CLI with the given prompt using stream-json output format.

    Streams events in real time:
      - Assistant text is printed as received.
      - Tool calls are printed as one-line summaries.
      - The final "result" event populates RunMetrics (tokens, cost, turns).

    A watchdog thread kills the process after IDLE_KILL_SECONDS of silence
    or when total runtime exceeds MAX_RUN_SECONDS.

    api_key: when None, ANTHROPIC_API_KEY is stripped from the subprocess
             environment so claude.ai account mode is forced. When provided,
             the key is injected to use the Anthropic API directly.

    Returns (exit_code, raw_output, killed_by_watchdog, metrics).
    """
    cmd = [
        "claude",
        "--print",
        "--output-format", "stream-json",
        "--verbose",
        "--model", MODEL,
        "--allowedTools", TOOLS,
        "--max-turns", str(max_turns),
        "--dangerously-skip-permissions",
    ]

    # Build subprocess environment.
    # Default: strip ANTHROPIC_API_KEY so the claude.ai account is always tried first.
    # Fallback: inject the key to route directly to the Anthropic API.
    if api_key:
        env = {**os.environ, "ANTHROPIC_API_KEY": api_key}
    else:
        env = {k: v for k, v in os.environ.items() if k != "ANTHROPIC_API_KEY"}

    try:
        proc = subprocess.Popen(
            cmd,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
            env=env,
        )
    except FileNotFoundError:
        return 127, "ERROR: `claude` CLI not found in PATH.", False, RunMetrics()

    assert proc.stdin is not None
    proc.stdin.write(prompt)
    proc.stdin.close()

    # Prevent macOS idle sleep while Claude is running so the streaming API
    # connection isn't dropped when the screen times out. caffeinate -w exits
    # automatically when the watched PID exits, so no explicit cleanup is needed.
    # Between polling cycles the Mac can sleep normally.
    _caffeinate: Optional[subprocess.Popen] = None
    if shutil.which("caffeinate"):
        try:
            _caffeinate = subprocess.Popen(
                ["caffeinate", "-i", "-w", str(proc.pid)],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
        except OSError:
            pass

    lines: list[str] = []
    last_output_at = time.monotonic()
    start_time = time.monotonic()
    idle_warned = False
    killed_by_watchdog = False
    metrics = RunMetrics()

    def _watchdog() -> None:
        nonlocal idle_warned, killed_by_watchdog
        while proc.poll() is None:
            time.sleep(15)
            now = time.monotonic()
            idle = now - last_output_at
            total = now - start_time

            if total >= MAX_RUN_SECONDS:
                log(
                    f"  [timeout {int(total)}s] — killing process that exceeded "
                    f"max runtime of {MAX_RUN_SECONDS}s (PID {proc.pid})"
                )
                killed_by_watchdog = True
                proc.terminate()
                break

            if idle >= IDLE_KILL_SECONDS:
                log(f"  [idle {int(idle)}s] — killing hung process (PID {proc.pid})")
                killed_by_watchdog = True
                proc.terminate()
                break

            if idle >= IDLE_WARN_SECONDS and not idle_warned:
                log(
                    f"  [idle {int(idle)}s] claude is still running (PID {proc.pid}) "
                    "but has produced no output — may be waiting on a tool call or API"
                )
                idle_warned = True

    watchdog = threading.Thread(target=_watchdog, daemon=True)
    watchdog.start()

    assert proc.stdout is not None
    for raw_line in proc.stdout:
        last_output_at = time.monotonic()
        idle_warned = False
        lines.append(raw_line)

        stripped = raw_line.strip()
        if stripped:
            try:
                event = json.loads(stripped)
                _handle_stream_event(event, metrics)
            except (json.JSONDecodeError, KeyError, TypeError):
                # Not a JSON event (unexpected) — print raw so nothing is lost
                print(raw_line, end="", flush=True)

        try:
            with LOG_FILE.open("a") as fh:
                fh.write(raw_line)
        except OSError:
            pass

    proc.wait()
    metrics.duration_s = time.monotonic() - start_time
    watchdog.join(timeout=0)

    # caffeinate should have exited via -w, but terminate defensively in case.
    if _caffeinate and _caffeinate.poll() is None:
        _caffeinate.terminate()

    return proc.returncode, "".join(lines), killed_by_watchdog, metrics


def check_dependencies() -> None:
    if shutil.which("claude") is None:
        print("ERROR: Claude Code is not installed.", file=sys.stderr)
        print("Install: npm install -g @anthropic-ai/claude-code", file=sys.stderr)
        sys.exit(1)


def _run_with_fallback(
    prompt: str, label: str, max_turns: int = MAX_TURNS
) -> tuple[int, str, bool, RunMetrics]:
    """
    Run using the claude.ai account first (no API key).
    If that attempt hits a rate limit and ANTHROPIC_API_KEY is configured,
    retry immediately using the API key — no sleep, no work preserved.
    """
    exit_code, output, killed, metrics = run_claude(prompt, max_turns=max_turns)

    if is_rate_limited(output):
        api_key = os.environ.get("ANTHROPIC_API_KEY", "")
        if api_key:
            log(f"[{label}] Rate limited on claude.ai — retrying with API key")
            exit_code, output, killed, metrics = run_claude(prompt, api_key=api_key, max_turns=max_turns)
        # If no key is available the rate-limited result falls through as before.

    return exit_code, output, killed, metrics


# ─── Domain processing ─────────────────────────────────────────────────────


def _classify(exit_code: int, output: str, killed: bool) -> DomainResult:
    if is_rate_limited(output):
        return DomainResult(
            ok=False,
            rate_limited=True,
            sleep_seconds=compute_reset_sleep(output),
            exit_code=exit_code,
            output=output,
        )
    if killed or (exit_code != 0 and is_transient_error(output)):
        return DomainResult(ok=False, transient_error=True, exit_code=exit_code, output=output)
    return DomainResult(ok=(exit_code == 0), exit_code=exit_code, output=output)


def process_domain(domain: Domain) -> DomainResult:
    """Phase 1: research (saves checkpoint) or script writing (reads checkpoint)."""
    mode = "SCRIPT" if domain._is_script_mode() else "RESEARCH"
    max_turns = MAX_TURNS_SCRIPT if domain._is_script_mode() else MAX_TURNS_RESEARCH
    log("══════════════════════════════════════")
    log(f"Domain: {domain.name} [{mode}] (max_turns={max_turns})")
    log("──────────────────────────────────────")

    exit_code, output, killed, metrics = _run_with_fallback(
        domain.build_prompt(), domain.name, max_turns=max_turns
    )
    result = _classify(exit_code, output, killed)

    if result.rate_limited:
        status = "RATE_LIMITED"
        _log_metrics(domain.name, status, metrics, killed)
        log(f"Rate limit detected for {domain.name} — preserving work.md")
        return result

    if result.transient_error:
        status = "TRANSIENT_ERROR"
        _log_metrics(domain.name, status, metrics, killed)
        log(f"Transient error for {domain.name} (exit={exit_code}) — preserving work.md")
        return result

    status = "SUCCESS" if result.ok else f"FAILED (exit={exit_code})"
    _log_metrics(domain.name, status, metrics, killed)

    if result.ok:
        log(f"Success: {domain.name}")
        if domain.is_two_phase():
            domain.work_run_path.write_text(domain.work_path.read_text())
            log(f"Promoted to run phase: {domain.name}")
    else:
        log(f"WARNING: Claude exited with errors for {domain.name} (exit={exit_code})")

    domain.reset_work()
    log(f"Reset work.md for {domain.name}")
    return result


def process_run(domain: Domain) -> DomainResult:
    """Phase 2: run scripts and verify output."""
    log("══════════════════════════════════════")
    log(f"Domain: {domain.name} [RUN]")
    log("──────────────────────────────────────")

    exit_code, output, killed, metrics = _run_with_fallback(domain.build_run_prompt(), f"{domain.name}:run")
    result = _classify(exit_code, output, killed)

    if result.rate_limited:
        _log_metrics(f"{domain.name}:run", "RATE_LIMITED", metrics, killed)
        log(f"Rate limit detected for {domain.name} [run] — preserving work_run.md")
        return result

    if result.transient_error:
        _log_metrics(f"{domain.name}:run", "TRANSIENT_ERROR", metrics, killed)
        log(f"Transient error for {domain.name} [run] (exit={exit_code}) — preserving work_run.md")
        return result

    status = "SUCCESS" if result.ok else f"FAILED (exit={exit_code})"
    _log_metrics(f"{domain.name}:run", status, metrics, killed)

    if result.ok:
        log(f"Success: {domain.name} [run]")
    else:
        log(f"WARNING: {domain.name} [run] exited with errors (exit={exit_code})")

    domain.reset_run()
    log(f"Reset work_run.md for {domain.name}")
    return result


# ─── Cycle coordination ────────────────────────────────────────────────────


def handle_rate_limit(result: DomainResult) -> None:
    sleep_seconds = result.sleep_seconds
    if sleep_seconds and sleep_seconds > 0:
        wake = datetime.now() + timedelta(seconds=sleep_seconds)
        log(f"Sleeping {sleep_seconds}s until {wake.strftime('%Y-%m-%d %H:%M:%S')} (rate limit reset)")
        time.sleep(sleep_seconds)
    else:
        log(f"Could not parse reset time; sleeping {RATE_LIMIT_FALLBACK_SLEEP}s as fallback")
        time.sleep(RATE_LIMIT_FALLBACK_SLEEP)


def _handle_transient(key: str, transient_retries: dict[str, int]) -> None:
    retries = transient_retries.get(key, 0)
    sleep_s = compute_transient_sleep(retries)
    transient_retries[key] = retries + 1
    log(f"Transient error #{retries + 1} for {key} — retrying in {sleep_s}s (backoff)")
    time.sleep(sleep_s)


def run_cycle(domains: list[Domain], transient_retries: dict[str, int]) -> None:
    """Process all active domains once. Stops early on rate limit."""
    for domain in domains:
        # ── Phase 1: research + write scripts ──────────────────────────────
        if domain.has_active_task():
            result = process_domain(domain)
            if result.rate_limited:
                handle_rate_limit(result)
                return
            if result.transient_error:
                _handle_transient(domain.name, transient_retries)
            else:
                transient_retries.pop(domain.name, None)

        # ── Phase 2: run scripts (two-phase domains only) ───────────────────
        if domain.has_pending_run():
            run_key = f"{domain.name}:run"
            result = process_run(domain)
            if result.rate_limited:
                handle_rate_limit(result)
                return
            if result.transient_error:
                _handle_transient(run_key, transient_retries)
            else:
                transient_retries.pop(run_key, None)


# ─── Entry point ───────────────────────────────────────────────────────────


def main() -> None:
    parser = argparse.ArgumentParser(description="Claude local agent")
    parser.add_argument("--once", action="store_true", help="Process all domains once and exit")
    parser.add_argument("--domain", choices=DOMAINS, help="Run only a single domain once")
    parser.add_argument(
        "--poll", type=int, default=POLL_EVERY,
        help=f"Seconds between scans (default: {POLL_EVERY})",
    )
    args = parser.parse_args()

    check_dependencies()

    domains = [Domain(name) for name in DOMAINS]
    for d in domains:
        d.ensure_dirs()
    LOG_FILE.touch(exist_ok=True)

    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if api_key:
        auth_mode = f"claude.ai first, API key fallback (...{api_key[-6:]})"
    else:
        auth_mode = "claude.ai only (set ANTHROPIC_API_KEY to enable fallback)"

    log("Claude Agent started")
    log(f"  Base dir : {BASE_DIR}")
    log(f"  Domains  : {', '.join(DOMAINS)}")
    log(f"  Log file : {LOG_FILE}")
    log(f"  Model    : {MODEL}")
    log(f"  Auth     : {auth_mode}")
    log(f"  Polling every {args.poll}s  |  Ctrl-C to stop")
    print()

    transient_retries: dict[str, int] = {}

    if args.domain:
        target = next(d for d in domains if d.name == args.domain)
        ran = False
        if target.has_active_task():
            ran = True
            result = process_domain(target)
            if result.rate_limited:
                handle_rate_limit(result)
                return
        if target.has_pending_run():
            ran = True
            result = process_run(target)
            if result.rate_limited:
                handle_rate_limit(result)
        if not ran:
            log(f"{target.name} has no active task or pending run. Nothing to do.")
        return

    if args.once:
        run_cycle(domains, transient_retries)
        return

    try:
        while True:
            run_cycle(domains, transient_retries)
            log(f"Sleeping {args.poll}s...")
            time.sleep(args.poll)
    except KeyboardInterrupt:
        log("Interrupted — exiting.")


if __name__ == "__main__":
    main()
