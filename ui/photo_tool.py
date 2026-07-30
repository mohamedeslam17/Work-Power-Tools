"""Photo Library — 'Show me micrographs of alloy X.'"""
import streamlit as st

from photo_lib import alloy_counts, photos_for, get_image_bytes, backend_name

try:
    from lab_review import HT_ORDER
except ImportError:   # tolerate a stale/cached lab_review module on redeploy
    HT_ORDER = ['As-received', 'Post-solution', 'Post stress-relief', 'Post-ageing', 'Unspecified']


@st.cache_data(show_spinner=False, ttl=120)
def _cached_counts():
    return alloy_counts()


@st.cache_data(show_spinner=False, ttl=120)
def _cached_photos(alloy):
    return photos_for(alloy)


@st.cache_data(show_spinner=False)
def _cached_image(path, drive_id):
    return get_image_bytes({"path": path, "drive_id": drive_id})


def invalidate():
    """Called by Lab Report Review after adding new micrographs."""
    _cached_counts.clear()
    _cached_photos.clear()


def render():
    st.caption(f"Storage backend: **{backend_name()}**")
    try:
        counts = _cached_counts()
    except Exception as e:
        st.error(f"Couldn't reach the photo-library backend (**{backend_name()}**) right now — "
                 "the library is unavailable, but the other tools are unaffected.")
        with st.expander("Error details"):
            st.exception(e)
        if st.button("↻ Retry"):
            invalidate()
            st.rerun()
        return

    total = sum(counts.values())
    if not total:
        st.info("The library is empty. Review a report in **Lab Report Review** and use "
                "**Actions → Add to photo library**.")
        return

    alloys = sorted(counts)
    default_alloy = max(counts, key=counts.get)
    c1, c2 = st.columns([2, 2])
    alloy = c1.selectbox("Alloy", alloys, index=alloys.index(default_alloy),
                         format_func=lambda a: f"{a} ({counts[a]})")
    dim = c2.segmented_control("Group by", ["Heat treatment", "Etchant"], default="Heat treatment")
    dim = dim or "Heat treatment"

    recs = _cached_photos(alloy)
    key = "ht" if dim == "Heat treatment" else "etchant"
    other = "etchant" if key == "ht" else "ht"
    icon = "🔥" if key == "ht" else "🧪"

    groups = {}
    for r in recs:
        groups.setdefault(r.get(key) or "Unspecified", []).append(r)
    if key == "ht":          # process sequence for HT, alphabetical for etchant
        order = [g for g in HT_ORDER if g in groups] + sorted(g for g in groups if g not in HT_ORDER)
    else:
        order = sorted(groups)

    for g in order:
        st.divider()
        st.subheader(f"{icon} {g}  ·  {len(groups[g])}")
        cols = st.columns(3)
        for i, r in enumerate(groups[g]):
            c = cols[i % 3]
            img = _cached_image(r.get("path"), r.get("drive_id"))
            contrast = "etched" if r.get("etched") else "low-contrast"
            cap = f"Job {r.get('job') or '—'} · {r.get('mag') or '?'} · {contrast}"
            if img:
                c.image(img, caption=cap, width="stretch")
            else:
                c.warning(f"missing: {r.get('path') or r.get('drive_id')}")
            bits = [f"set: {(r.get('source') or '')[:30]}"]
            if r.get(other):
                bits.append(str(r.get(other)))
            meas = r.get("measurements") or []
            if meas:
                bits.append(", ".join(f"{m}µm" for m in meas))
            c.caption(" · ".join(bits))
