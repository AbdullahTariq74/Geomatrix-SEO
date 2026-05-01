"""
GeoMatrix SEO Generator
Streamlit UI — enter a website, business name, radius, and optional extra details
to generate a full GEOMATRIX Excel workbook + 1-page visual PDF pitch deck.
"""

import streamlit as st
import os, sys
from dotenv import load_dotenv

load_dotenv()
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "tools"))

CLAUDE_API_KEY = os.getenv("ANTHROPIC_API_KEY", "")

st.set_page_config(
    page_title="GeoMatrix SEO Generator",
    page_icon="🗺️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Light, clean theme ────────────────────────────────────────────────────────
st.markdown("""
<style>
  /* App background */
  .stApp, .main { background-color: #f0f4ff !important; }
  .block-container { padding-top: 2rem; max-width: 1100px; }

  /* Headings */
  h1 { color: #1a3a6b !important; font-size: 2rem !important; }
  h2, h3, h4 { color: #1e3a5f !important; }
  p, li { color: #334155; }

  /* Form inputs */
  .stTextInput > div > div > input,
  .stNumberInput > div > div > input,
  .stTextArea > div > div > textarea {
    background-color: #ffffff !important;
    color: #1e293b !important;
    border: 1.5px solid #c7d7f5 !important;
    border-radius: 7px !important;
    font-size: 14px !important;
  }
  .stTextInput > div > div > input:focus,
  .stTextArea > div > div > textarea:focus {
    border-color: #2563eb !important;
    box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.1) !important;
  }

  /* Labels */
  label, .stSlider label { color: #334155 !important; font-weight: 600 !important; }

  /* Primary button */
  .stButton > button[kind="primary"],
  .stFormSubmitButton > button {
    background: linear-gradient(135deg, #1a3a6b 0%, #2563eb 100%) !important;
    color: #ffffff !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 700 !important;
    font-size: 15px !important;
    padding: 12px 28px !important;
    letter-spacing: 0.3px;
    box-shadow: 0 4px 12px rgba(37, 99, 235, 0.25) !important;
  }
  .stButton > button[kind="primary"]:hover,
  .stFormSubmitButton > button:hover {
    opacity: 0.92 !important;
    box-shadow: 0 6px 18px rgba(37, 99, 235, 0.35) !important;
  }

  /* Download buttons */
  .stDownloadButton > button {
    background-color: #ffffff !important;
    color: #1a3a6b !important;
    border: 2px solid #2563eb !important;
    border-radius: 8px !important;
    font-weight: 700 !important;
    font-size: 14px !important;
  }
  .stDownloadButton > button:hover {
    background-color: #eff6ff !important;
  }

  /* Metric cards */
  div[data-testid="metric-container"] {
    background-color: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    padding: 14px 18px;
    box-shadow: 0 2px 8px rgba(30, 58, 95, 0.06);
  }

  /* Expander */
  .stExpander {
    border: 1px solid #dbeafe !important;
    border-radius: 8px !important;
    background-color: #ffffff !important;
  }

  /* Dataframe */
  .stDataFrame { border-radius: 8px; }

  /* Info / success / warning boxes */
  .stAlert { border-radius: 8px !important; }

  /* Section divider */
  hr { border-color: #dbeafe; margin: 1.5rem 0; }

  /* Sidebar (if used) */
  section[data-testid="stSidebar"] { background-color: #f0f4ff; }
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
col_logo, col_title = st.columns([1, 10])
with col_logo:
    st.markdown("""
    <div style="background:linear-gradient(135deg,#1a3a6b,#2563eb);
         width:52px;height:52px;border-radius:12px;display:flex;
         align-items:center;justify-content:center;margin-top:4px;font-size:26px;">
      🗺️
    </div>""", unsafe_allow_html=True)
with col_title:
    st.title("GeoMatrix SEO Generator")
    st.markdown("<p style='margin-top:-8px;color:#64748b;font-size:15px;'>"
                "Enter a business website and radius → get a complete local SEO keyword strategy "
                "and a visual pitch deck in minutes.</p>", unsafe_allow_html=True)

st.markdown("---")

# ── Input form ────────────────────────────────────────────────────────────────
with st.form("main_form"):
    col1, col2 = st.columns([3, 2])

    with col1:
        st.markdown("**Business Details**")
        website = st.text_input(
            "Website URL",
            placeholder="https://example.com",
            help="The client's website — we'll crawl multiple pages automatically."
        )
        business_name = st.text_input(
            "Business Name",
            placeholder="e.g. Westside Family Dentistry"
        )
        extra_info = st.text_area(
            "Additional Business Details (optional)",
            placeholder=(
                "Add anything the website might not mention:\n"
                "• Extra services offered\n"
                "• Specialisations or certifications\n"
                "• Target audience or unique differentiators\n"
                "• Service areas not shown on the site"
            ),
            height=130,
            help="This info is fed directly into the AI alongside the website content to make the strategy more accurate."
        )

    with col2:
        st.markdown("**Phase Configuration**")

        ph1_radius = st.slider(
            "Phase 1 — Core radius (miles)",
            min_value=5, max_value=100, value=25, step=5,
            help="Primary coverage area, typically 10–30 miles from the office."
        )
        ph2_radius = st.number_input(
            "Phase 2 — Expansion radius (miles, 0 = skip)",
            min_value=0, max_value=200, value=0, step=5,
            help="Secondary ring. Set to 0 to run Phase 1 only."
        )
        ph3_radius = st.number_input(
            "Phase 3 — Extended radius (miles, 0 = skip)",
            min_value=0, max_value=300, value=0, step=5,
            help="Outermost ring. Leave 0 if using a custom area description below."
        )
        ph3_custom = st.text_input(
            "Phase 3 — Custom area (overrides radius if filled)",
            placeholder="e.g. High-income Hamptons area, NY  |  Westchester County, NY",
            help="Describe a specific geographic area for Phase 3 — AI will find real locations matching it."
        )

        st.markdown(
            "<div style='background:#eff6ff;border:1px solid #bfdbfe;"
            "border-radius:8px;padding:10px 14px;margin-top:6px;'>"
            "<b style='color:#1a3a6b;font-size:13px;'>📊 What you'll get</b><br>"
            "<span style='color:#475569;font-size:12px;'>"
            "✓ GeoMatrix Excel (all keyword data)<br>"
            "✓ 1-page visual pitch deck PDF + ROI model<br>"
            "✓ Longtail keyword bank<br>"
            "✓ Interactive service area map"
            "</span></div>",
            unsafe_allow_html=True
        )

    st.markdown("")
    submitted = st.form_submit_button(
        "🚀  Generate GeoMatrix Report",
        use_container_width=True,
        type="primary"
    )

# ── Run pipeline only when form is freshly submitted ─────────────────────────
if submitted:
    errors = []
    if not website:       errors.append("Website URL is required.")
    if not business_name: errors.append("Business Name is required.")
    if not CLAUDE_API_KEY:
        errors.append("No API key found. Make sure the .env file contains ANTHROPIC_API_KEY.")

    # Build phase configuration from form inputs
    _ph3_custom = (ph3_custom or "").strip()
    _ph3_r      = int(ph3_radius) if ph3_radius else 0
    phase_configs = [{"phase": "Phase 1", "market_type": "Core",
                      "radius": int(ph1_radius), "custom": None}]
    if ph2_radius and int(ph2_radius) > 0:
        phase_configs.append({"phase": "Phase 2", "market_type": "Expansion",
                               "radius": int(ph2_radius), "custom": None})
    if _ph3_custom or _ph3_r > 0:
        phase_configs.append({"phase": "Phase 3", "market_type": "Extended",
                               "radius": None if _ph3_custom else _ph3_r,
                               "custom": _ph3_custom or None})
    max_radius = max((pc.get("radius") or 0) for pc in phase_configs) or int(ph1_radius)

    if errors:
        for e in errors:
            st.error(e)
        st.stop()

    from geomatrix_engine import GeomatrixEngine
    engine   = GeomatrixEngine(CLAUDE_API_KEY)
    progress = st.progress(0)
    status   = st.empty()

    try:
        status.info("🌐  Scraping website (homepage + service pages)…")
        content = engine.scrape_website(website)
        progress.progress(12)

        status.info("🤖  Analysing business & extracting all services with AI…")
        biz = engine.extract_business_info(content, business_name, website,
                                            extra_info=extra_info)
        progress.progress(28)

        status.info("📋  Structuring service library…")
        services_raw    = biz.get("services", ["Professional Services"])
        service_library = engine.build_service_library(
            services_raw, biz.get("industry", "Services"), business_name,
            extra_info=extra_info
        )
        progress.progress(40)

        status.info("📍  Locating business on map…")
        lat, lon = engine.geocode_address(
            biz.get("city", ""), biz.get("state", ""), biz.get("address", "")
        )
        if not lat:
            st.error(
                "Could not determine the business location. "
                "Check that the website has an address, or add the city/state in "
                "the 'Additional Business Details' field."
            )
            st.stop()
        progress.progress(50)

        n_phases = len(phase_configs)
        phases_desc = " + ".join(pc["phase"] for pc in phase_configs)
        status.info(f"🗺️  Finding locations for {phases_desc}…")
        locations = engine.find_locations_by_phases(
            lat, lon, biz.get("city", ""), biz.get("state", ""), phase_configs
        )
        progress.progress(62)

        status.info("📊  Building GeoMatrix keyword data…")
        geomatrix_df = engine.build_geomatrix(
            service_library, locations, business_name, biz.get("state_abbr", "")
        )
        progress.progress(72)

        status.info("🔑  Building longtail keyword bank…")
        industry_key = engine._detect_industry(biz.get("industry", ""))
        longtail_df  = engine.build_longtail_bank(
            service_library, locations, biz.get("state_abbr", ""), industry_key
        )
        page_plan_df = engine.build_page_plan(service_library, geomatrix_df, max_radius)
        progress.progress(80)

        status.info("🗺️  Rendering service area map…")
        map_image_buf = engine.generate_map_image(lat, lon, locations, max_radius, business_name)
        folium_html   = engine.generate_folium_map(lat, lon, locations, max_radius, business_name)
        progress.progress(88)

        status.info("📄  Generating Excel workbook and PDF pitch deck…")
        excel_bytes = engine.export_excel(
            geomatrix_df, longtail_df, service_library, page_plan_df,
            business_name, website, biz
        )
        pdf_bytes = engine.export_pdf(
            business_name, biz, geomatrix_df, service_library,
            map_image_buf, phase_configs, locations
        )
        progress.progress(100)
        status.success("✅  Report generated successfully!")

        # Persist everything so download-button re-runs don't lose the results
        st.session_state["results"] = dict(
            geomatrix_df    = geomatrix_df,
            longtail_df     = longtail_df,
            service_library = service_library,
            page_plan_df    = page_plan_df,
            biz             = biz,
            locations       = locations,
            folium_html     = folium_html,
            excel_bytes     = excel_bytes,
            pdf_bytes       = pdf_bytes,
            business_name   = business_name,
            website         = website,
            phase_configs   = phase_configs,
            max_radius      = max_radius,
        )

    except RuntimeError as e:
        st.error(str(e))
        st.stop()
    except Exception as e:
        st.error(f"Something went wrong: {e}")
        import traceback
        st.code(traceback.format_exc())
        st.stop()

# Nothing generated yet — stop before results section
if "results" not in st.session_state:
    st.stop()

# ── Unpack persisted results ──────────────────────────────────────────────────
_r             = st.session_state["results"]
geomatrix_df   = _r["geomatrix_df"]
longtail_df    = _r["longtail_df"]
service_library= _r["service_library"]
page_plan_df   = _r["page_plan_df"]
biz            = _r["biz"]
locations      = _r["locations"]
folium_html    = _r["folium_html"]
excel_bytes    = _r["excel_bytes"]
pdf_bytes      = _r["pdf_bytes"]
business_name  = _r["business_name"]
website        = _r["website"]
phase_configs  = _r["phase_configs"]
max_radius     = _r["max_radius"]

# ── Results ───────────────────────────────────────────────────────────────────
st.markdown("---")

col_map, col_stats = st.columns([3, 2])

with col_map:
    st.subheader("📍 Service Area Map")
    st.components.v1.html(folium_html, height=430)

with col_stats:
    st.subheader("📊 Summary")
    n_loc = geomatrix_df["Location"].nunique()     if "Location"         in geomatrix_df.columns else 0
    n_svc = geomatrix_df["Service Category"].nunique() if "Service Category" in geomatrix_df.columns else 0
    m1, m2 = st.columns(2)
    with m1:
        st.metric("GeoMatrix Rows",     f"{len(geomatrix_df):,}")
        st.metric("Locations Covered",  n_loc)
    with m2:
        st.metric("Longtail Keywords",  f"{len(longtail_df):,}")
        st.metric("Services",           n_svc)

    st.markdown(f"**Industry:** {biz.get('industry', '—')}")
    st.markdown(f"**Base:** {biz.get('city', '—')}, {biz.get('state_abbr', '—')}")
    if biz.get("phone"):
        st.markdown(f"**Phone:** {biz['phone']}")

st.markdown("---")

# ── Downloads ─────────────────────────────────────────────────────────────────
safe_name = business_name.replace(" ", "_")
dl1, dl2 = st.columns(2)

with dl1:
    st.download_button(
        label="📊  Download GeoMatrix Excel",
        data=excel_bytes,
        file_name=f"{safe_name}_geomatrix.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
with dl2:
    st.download_button(
        label="📄  Download Pitch Deck PDF",
        data=pdf_bytes,
        file_name=f"{safe_name}_pitch_deck.pdf",
        mime="application/pdf",
        use_container_width=True,
    )

st.markdown("---")

# ── Previews ──────────────────────────────────────────────────────────────────
st.subheader("📋 GeoMatrix Keywords Preview")
st.dataframe(geomatrix_df, use_container_width=True, height=380, hide_index=True)

tab1, tab2, tab3 = st.tabs(["🔑 Longtail Keyword Bank", "📋 Service Library", "📅 Page Build Plan"])
with tab1:
    st.caption(f"{len(longtail_df):,} total longtail keywords")
    st.dataframe(longtail_df.head(200), use_container_width=True, hide_index=True)
with tab2:
    import pandas as pd
    st.dataframe(pd.DataFrame(service_library), use_container_width=True, hide_index=True)
with tab3:
    st.dataframe(page_plan_df, use_container_width=True, hide_index=True)
