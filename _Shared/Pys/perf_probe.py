"""
perf_probe.py -- Diagnostic timing shim for the LX02 SNAPSHOT-load 58s-gap
investigation (2026-08-22).

WHAT THIS DOES
---------------
Non-invasively monkey-patches timing wrappers around the functions most
likely to explain the dead time observed between:
    "LX02_loadonly | [OK] upserted 800 row(s)"        (07:30:14)
and
    "LX02_loadonly | [BATCH] 500 delete(s): removed=500..."   (07:31:12)

It does NOT modify dataverse_core.py or dataverse_delta.py on disk. Import
this module (see USAGE below) BEFORE run_delta() is actually CALLED and it
patches itself into both modules' namespaces for that process only.

INSTRUMENTED CALLS
-------------------
1. dataverse_core.get_token        -> isolates the ACTUAL MSAL/AAD network
                                       round trip (silent token refresh).
2. dataverse_core._bearer          -> total time incl. lock-wait + cache
   dataverse_delta._bearer            check (patched in BOTH namespaces,
                                       since dataverse_delta did
                                       `from dataverse_core import _bearer`,
                                       which creates a SEPARATE local
                                       binding that patching dataverse_core
                                       alone would NOT catch).
3. dataverse_core.DataverseLoader.load -> brackets the upsert-via-loader
                                       call, if this is the path
                                       _run_snapshot uses for adds.
4. dataverse_core._upsert_multiple -> brackets the chunked ADD writes
                                       directly, if called without going
                                       through DataverseLoader.load().
5. dataverse_delta._delete_by_keys -> brackets the chunked REMOVE writes
                                       (defined directly in dataverse_delta,
                                       no import-binding issue).
6. dataverse_delta._run_snapshot   -> brackets the WHOLE snapshot driver,
                                       giving you one absolute start/end so
                                       you can see where the gap sits
                                       relative to steps 1-5 above.

Any target that doesn't exist under the name expected is skipped with a
WARNING (not an error) -- so this is safe to leave installed even if a
name doesn't match your exact file (e.g. if DataverseLoader's method is
called something other than `load`).

USAGE -- Option A: one real scheduled run (recommended)
--------------------------------------------------------
Drop this file into:
    C:\\Users\\ngsoyemiol\\Documents\\Automation\\All_Locations\\_Shared\\Pys\\perf_probe.py

Then add ONE line near the top of lx02_engine.py (anywhere after the
sys.path bootstrap block, e.g. right after the `import config` line):

    if os.environ.get("NSP_PERF_PROBE") == "1":
        import perf_probe  # noqa: F401  (diagnostic only -- installs timing wraps)

Then run your normal launch_all_automation.py with the env var set for
just that one run, e.g. in PowerShell:

    $env:NSP_PERF_PROBE = "1"
    python launch_all_automation.py --upload
    Remove-Item Env:\\NSP_PERF_PROBE

This captures the REAL gap under REAL concurrent load (LT23/MB51/MB5T all
hammering Dataverse at the same time), which is when the 58s was observed.

USAGE -- Option B: quick isolated test (faster, but no cross-process
contention from the other four transactions)
----------------------------------------------------------------------
    python diag_load_lx02.py

(see the companion diag_load_lx02.py file)

READING THE OUTPUT
-------------------
Look for a PERF-tagged block like:

    07:30:14 | PERF | <<< EXIT  LOADER DataverseLoader.load()        12.10s
    07:30:14 | PERF | >>> ENTER TOKEN get_token()
    07:31:12 | PERF | <<< EXIT  TOKEN get_token()                    57.90s   <- smoking gun if this fires
    07:31:12 | PERF | >>> ENTER DELETE _delete_by_keys()

If get_token()'s own EXIT timestamp shows ~58s, the gap is a stale/slow
MSAL silent-refresh (very plausible given your prior VPN-latency issues
reaching org6f0703be.crm.dynamics.com). If get_token() is fast but
_run_snapshot()'s internal gap (between the LOADER exit and DELETE enter)
is still ~58s with NO probe line firing in between, the delay is happening
in code this shim didn't reach (e.g. the remove-key diff computation) --
in that case, paste me the exact _run_snapshot() body and I'll add
targeted line-level logging there instead.

REMOVE AFTER DIAGNOSIS -- this is a diagnostic aid only, not meant to run
in production long-term (it only adds logging, no behavior change, but
there is no reason to keep the overhead once you have your answer).
"""
import time
import functools
import logging
import sys

log = logging.getLogger("perf_probe")
if not log.handlers:
    log.setLevel(logging.INFO)
    _h = logging.StreamHandler(sys.stdout)
    _h.setFormatter(logging.Formatter(
        "%(asctime)s | PERF | %(message)s", datefmt="%Y-%m-%d %H:%M:%S"))
    log.addHandler(_h)
    log.propagate = False


def _wrap(owner, func_name, label):
    """
    Replace owner.func_name with a timed version, keyed by `label`.
    `owner` can be a module OR a class (for bound/unbound methods like
    DataverseLoader.load). Silently skips (with a warning) if the target
    doesn't exist, so this is safe to over-instrument.
    """
    orig = getattr(owner, func_name, None)
    if orig is None:
        log.warning("skip: %r has no attribute %r", owner, func_name)
        return

    @functools.wraps(orig)
    def timed(*args, **kwargs):
        t0 = time.perf_counter()
        log.info(">>> ENTER %-32s", label)
        try:
            return orig(*args, **kwargs)
        finally:
            dt = time.perf_counter() - t0
            log.info("<<< EXIT  %-32s  %.2fs", label, dt)

    setattr(owner, func_name, timed)
    log.info("instrumented %s.%s as %r", getattr(owner, "__name__", owner), func_name, label)


def install():
    import dataverse_core as dc
    import dataverse_delta as dd

    # 1) The actual network round trip for a silent MSAL token refresh.
    _wrap(dc, "get_token", "TOKEN get_token()")

    # 2) The cache-check wrapper around get_token -- patch BOTH namespaces
    #    because dataverse_delta did `from dataverse_core import _bearer`,
    #    which copies the reference into its OWN globals at import time.
    _wrap(dc, "_bearer", "BEARER dataverse_core._bearer()")
    _wrap(dd, "_bearer", "BEARER dataverse_delta._bearer()")

    # 3) Upsert path -- try both possible call sites.
    _wrap(dc.DataverseLoader, "load", "LOADER DataverseLoader.load()")
    _wrap(dc, "_upsert_multiple", "UPSERT _upsert_multiple()")

    # 4) Delete path -- defined directly in dataverse_delta, no import gotcha.
    _wrap(dd, "_delete_by_keys", "DELETE _delete_by_keys()")

    # 5) Whole-snapshot bracket, so you see the gap's absolute position.
    _wrap(dd, "_run_snapshot", "SNAPSHOT _run_snapshot()")

    log.info("perf_probe installed -- run your load and watch for PERF lines.")


install()
