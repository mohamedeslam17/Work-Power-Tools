"""Shared presentation shell for the materials tools."""
import streamlit as st

# Keeps custom CSS in sync with Streamlit's *actual* active theme (light or
# dark), not just the OS preference. Streamlit persists the user's choice —
# including "System" — in localStorage; this reads it (resolving "System" via
# prefers-color-scheme, same as Streamlit itself does) and stamps it onto
# <html data-theme="..."> so plain CSS attribute selectors can target it. The
# component iframe is same-origin, so window.parent.document is reachable
# without extra plumbing. Polling covers the settings-menu switch, which
# doesn't fire a same-tab storage event.
_THEME_SYNC = """
<script>
(function() {
  function apply() {
    let raw = localStorage.getItem('stActiveTheme-/-v2');
    let choice = 'System';
    try { choice = raw ? JSON.parse(raw) : 'System'; } catch (e) {}
    let effective = choice;
    if (choice === 'System' || !choice) {
      effective = window.matchMedia('(prefers-color-scheme: dark)').matches ? 'Dark' : 'Light';
    }
    const doc = window.parent.document;
    doc.documentElement.setAttribute('data-theme', String(effective).toLowerCase());
  }
  apply();
  try {
    window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', apply);
  } catch (e) {}
  setInterval(apply, 800);
})();
</script>
"""

# Tempered-steel blue on warm graphite — a palette that belongs to materials
# inspection, not a borrowed developer-tool green. Colors are declared once
# per token as CSS custom properties, light values at :root and dark values
# under [data-theme="dark"], so every rule below reads the token, never a
# hardcoded hex.
_CSS = """
<style>
:root {
  --aeg-card-bg: #ffffff;
  --aeg-card-border: #dcd8cf;
  --aeg-muted: #6b7280;
  --aeg-ink: #1e2328;
  --aeg-critical-line: #b3261e;
  --aeg-critical-bg: #fdf4f3;
  --aeg-warning-line: #a8641a;
  --aeg-warning-bg: #fcf6ee;
  --aeg-info-line: #2f6690;
  --aeg-info-bg: #f2f7fa;
  --aeg-pass-line: #2f7a4f;
  --aeg-pass-bg: #f2f8f4;
  --aeg-notlocated-line: #8a6d1f;
  --aeg-notlocated-bg: #fbf6e6;
  --aeg-badge-ink: #ffffff;
  --aeg-shadow: 0 1px 2px rgba(30, 35, 40, .06);
}
html[data-theme="dark"] {
  --aeg-card-bg: #1b2126;
  --aeg-card-border: #2b3238;
  --aeg-muted: #9aa2a9;
  --aeg-ink: #e9e6df;
  --aeg-critical-line: #ef6a62;
  --aeg-critical-bg: #2a1918;
  --aeg-warning-line: #e0a24a;
  --aeg-warning-bg: #2a2318;
  --aeg-info-line: #7fb3d1;
  --aeg-info-bg: #172329;
  --aeg-pass-line: #6bbf8e;
  --aeg-pass-bg: #16241c;
  --aeg-notlocated-line: #e0c24a;
  --aeg-notlocated-bg: #2a2718;
  /* The number badge is filled with --aeg-ink, which is LIGHT in dark mode,
     so its text has to flip to the dark background colour or the number
     disappears into the pill. */
  --aeg-badge-ink: #12161a;
  --aeg-shadow: 0 1px 3px rgba(0, 0, 0, .4);
}

/* Wider than the old 1100px: the Lab screen is a two-pane workspace (report
   beside findings), and at 1100 the report pane was too narrow to read a
   spreadsheet page in. */
.block-container { padding-top: 1.4rem; max-width: 1500px; }

/* ── Report identity — one compact line, not a section ─────────────────── */
.aeg-idbar {
  display: flex; align-items: baseline; gap: .6rem; flex-wrap: wrap;
  margin: 0 0 .15rem;
}
.aeg-idbar-name {
  font-weight: 750; font-size: 1rem; color: var(--aeg-ink); word-break: break-word;
}
.aeg-idbar-tag {
  font-size: .68rem; font-weight: 700; letter-spacing: .06em; text-transform: uppercase;
  color: var(--aeg-muted); border: 1px solid var(--aeg-card-border);
  border-radius: 999px; padding: .05rem .45rem; white-space: nowrap;
}
[data-testid="stDecoration"] { display: none; }
[data-testid="stSidebar"] [data-testid="stCaptionContainer"] { padding: 0 0 .5rem .1rem; }

.aeg-page-kicker {
  color: var(--aeg-muted); font-size: .72rem; font-weight: 700; letter-spacing: .12em;
  margin: 0 0 .45rem .15rem;
}

/* ── Verdict banner — the one line at the top of a report ────────────── */
.aeg-verdict {
  display: flex; align-items: center; gap: .7rem; border-radius: 12px;
  border: 1px solid var(--aeg-card-border); padding: .85rem 1.05rem;
  margin: 0 0 .85rem; box-shadow: var(--aeg-shadow);
}
.aeg-verdict-hold { border-left: 5px solid var(--aeg-critical-line); background: var(--aeg-critical-bg); }
.aeg-verdict-correction { border-left: 5px solid var(--aeg-warning-line); background: var(--aeg-warning-bg); }
.aeg-verdict-release { border-left: 5px solid var(--aeg-pass-line); background: var(--aeg-pass-bg); }
.aeg-verdict-label {
  font-size: .95rem; font-weight: 800; letter-spacing: .04em; text-transform: uppercase;
  white-space: nowrap; color: var(--aeg-ink);
}
.aeg-verdict-hold .aeg-verdict-label { color: var(--aeg-critical-line); }
.aeg-verdict-correction .aeg-verdict-label { color: var(--aeg-warning-line); }
.aeg-verdict-release .aeg-verdict-label { color: var(--aeg-pass-line); }
.aeg-verdict-reason { color: var(--aeg-ink); font-size: .92rem; line-height: 1.4; }

/* ── Issue / finding cards ─────────────────────────────────────────────── */
.aeg-issue-card, .aeg-clear-card {
  border: 1px solid var(--aeg-card-border); border-left-width: 4px; border-radius: 10px;
  background: var(--aeg-card-bg); padding: .78rem .82rem; margin: 0 0 .7rem;
  box-shadow: var(--aeg-shadow);
}
.aeg-critical { border-left-color: var(--aeg-critical-line); background: var(--aeg-critical-bg); }
.aeg-warning { border-left-color: var(--aeg-warning-line); background: var(--aeg-warning-bg); }
.aeg-info { border-left-color: var(--aeg-info-line); background: var(--aeg-info-bg); }
.aeg-pass { border-left-color: var(--aeg-pass-line); background: var(--aeg-pass-bg); }
.aeg-issue-head { display: flex; align-items: center; gap: .45rem; margin-bottom: .35rem; }
.aeg-issue-number {
  display: inline-flex; width: 1.45rem; height: 1.45rem; border-radius: 999px;
  align-items: center; justify-content: center; background: var(--aeg-ink); color: var(--aeg-badge-ink);
  font-size: .72rem; font-weight: 800; flex-shrink: 0;
}
.aeg-issue-label { color: var(--aeg-ink); font-size: .76rem; font-weight: 800; text-transform: uppercase; }
.aeg-issue-meta { color: var(--aeg-muted); font-size: .72rem; font-weight: 650; margin-bottom: .32rem; }
.aeg-issue-copy { color: var(--aeg-ink); font-size: .86rem; line-height: 1.42; }
.aeg-clear-card { border-left-color: var(--aeg-pass-line); background: var(--aeg-pass-bg); }
.aeg-clear-title { color: var(--aeg-pass-line); font-weight: 750; margin-bottom: .2rem; }
.aeg-clear-copy { color: var(--aeg-muted); font-size: .84rem; }

.aeg-issue-card.aeg-dismissed {
  opacity: .55; filter: grayscale(.35);
}
.aeg-issue-card.aeg-focused {
  outline: 2px solid var(--aeg-info-line); outline-offset: 1px;
}

/* ── Extraction status — a not_located field must not read as "empty" ──── */
.aeg-field-status {
  display: inline-block; font-size: .68rem; font-weight: 750; letter-spacing: .02em;
  padding: .05rem .4rem; border-radius: 999px; margin-left: .4rem; vertical-align: middle;
}
.aeg-field-notlocated {
  color: var(--aeg-notlocated-line); background: var(--aeg-notlocated-bg);
  border: 1px solid var(--aeg-notlocated-line);
}
.aeg-field-empty { color: var(--aeg-muted); background: transparent; border: 1px dashed var(--aeg-card-border); }

/* ── Template-scoped findings — said once, not per report ──────────────── */
.aeg-template-note {
  color: var(--aeg-muted); font-size: .82rem; margin-bottom: .6rem;
}

/* ── The zoomable report page ──────────────────────────────────────────────
   st.image scales an image to fit its column and offers nothing beyond that,
   which is why the annotated report could not be zoomed. The page is inlined
   into this box instead: it scrolls in both directions, so the <img> inside can
   be drawn wider than the column and panned. resize:vertical lets a reviewer
   give the page more screen without touching the layout. */
.aeg-zoom-pane {
  overflow: auto; max-height: 84vh; resize: vertical;
  border: 1px solid var(--aeg-card-border); border-radius: 10px;
  background: var(--aeg-card-bg); box-shadow: var(--aeg-shadow);
  padding: 2px;
}
.aeg-zoom-pane img { border-radius: 6px; }

/* ── Reviewer comments — the reviewer's own words, not a finding ─────────── */
.aeg-comment-card {
  border: 1px solid var(--aeg-card-border); border-left: 4px solid #5b4cb0;
  border-radius: 10px; background: var(--aeg-card-bg); box-shadow: var(--aeg-shadow);
  padding: .7rem .8rem; margin: 0 0 .55rem;
}
html[data-theme="dark"] .aeg-comment-card { border-left-color: #9b8ce8; }
.aeg-comment-head {
  display: flex; align-items: center; gap: .5rem; margin-bottom: .3rem;
  font-size: .78rem; font-weight: 800; color: #5b4cb0; letter-spacing: .03em;
}
html[data-theme="dark"] .aeg-comment-head { color: #9b8ce8; }
.aeg-comment-meta { color: var(--aeg-muted); font-size: .72rem; font-weight: 650; }
/* Lines a form's submit button up with the inputs beside it. */
.aeg-form-spacer { height: 1.85rem; }
</style>
"""


def inject():
    # A <script> tag inserted via st.markdown's innerHTML never executes —
    # that's the browser's own behavior, not a Streamlit restriction — so the
    # sync script needs an iframe (st.markdown handles the CSS below).
    st.iframe(_THEME_SYNC, height=1)
    st.markdown(_CSS, unsafe_allow_html=True)
