"""
upload_gate.py
--------------
Shared helper that gives every SAP engine (LX02, LX09, LT23, MB51, MB5T) a
consistent way to:

  1. Register a `--no-upload` command-line flag (also accepts --skip-dataverse).
  2. Decide at runtime whether the Dataverse upload should run.
  3. Emit a clear, uniform console log line for BOTH modes so the run log
     always shows exactly what happened per engine.

Place this file next to the other shared modules, e.g.:
    C:\\Users\\ngsoyemiol\\Documents\\Automation\\All_Locations\\_Shared\\Pys\\upload_gate.py

Design notes:
- Zero third-party dependencies (argparse + logging only).
- The flag is OPT-IN to skipping: default behaviour is UNCHANGED (upload still
  runs) so nothing breaks if an engine forgets to pass the flag through.
- An env var override (NSP_NO_UPLOAD=1) is also honoured so the orchestrator can
  flip every engine at once without editing per-engine command lines if desired.
- Each engine passes should_skip_upload(None) at its central upload choke point,
  so the decision is read from the env var (which main() sets from the CLI flag,
  and which the orchestrator's -NoUpload sets directly).
"""

from __future__ import annotations

import os
import logging


# --------------------------------------------------------------------------- #
# 1) Argparse wiring
# --------------------------------------------------------------------------- #
def add_no_upload_argument(parser):
    """
    Attach the --no-upload flag to an existing argparse.ArgumentParser.

    Call this in each engine right after you build its parser, e.g.:

        parser = argparse.ArgumentParser(...)
        parser.add_argument("--unit", ...)
        parser.add_argument("--all", action="store_true", ...)
        add_no_upload_argument(parser)      # <-- add this line
        args = parser.parse_args()
    """
    parser.add_argument(
        "--no-upload", "--skip-dataverse",
        dest="no_upload",
        action="store_true",
        default=False,
        help="Run SAP extract + post-processing only; skip the Dataverse upload step.",
    )
    return parser


# --------------------------------------------------------------------------- #
# 2) Runtime decision
# --------------------------------------------------------------------------- #
def should_skip_upload(args=None) -> bool:
    """
    Return True when the Dataverse upload should be SKIPPED.

    Precedence (any one is enough to skip):
      - CLI flag:  --no-upload / --skip-dataverse   (when args is provided)
      - Env var :  NSP_NO_UPLOAD in {"1","true","yes","on"}  (case-insensitive)

    Engines call should_skip_upload(None) at the upload choke point so the
    decision is driven purely by the env var. main() sets NSP_NO_UPLOAD=1 from
    the CLI flag, and the orchestrator's -NoUpload sets it directly -- so both
    routes funnel through the same env var here.
    """
    # CLI flag (only when an args namespace is supplied)
    cli_skip = bool(getattr(args, "no_upload", False)) if args is not None else False

    # Env-var override (lets the orchestrator flip all engines at once)
    env_val = os.environ.get("NSP_NO_UPLOAD", "").strip().lower()
    env_skip = env_val in {"1", "true", "yes", "on"}

    return cli_skip or env_skip


# --------------------------------------------------------------------------- #
# 3) Uniform console/log line
# --------------------------------------------------------------------------- #
def log_upload_mode(logger: logging.Logger, engine_key: str, skip: bool) -> None:
    """
    Print a consistent one-line banner for the run log, for BOTH modes.

    Example output:
        LX02 | Dataverse upload: SKIPPED (--no-upload)
        LX02 | Dataverse upload: ENABLED
    """
    if skip:
        logger.info("%s | Dataverse upload: SKIPPED (--no-upload)", engine_key)
    else:
        logger.info("%s | Dataverse upload: ENABLED", engine_key)
