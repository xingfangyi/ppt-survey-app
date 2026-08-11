import tempfile
import uuid
from pathlib import Path

import pandas as pd
import streamlit as st

from survey_parser.extract_a import parse_survey_ppt
from survey_parser.score_rules import classify_metrics
from survey_parser.action_mapper import generate_actions
from ppt_generator.build_ppt import generate_ppt


st.set_page_config(
    page_title="Survey PPT Generator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
.block-container {
    padding-top: 2rem;
    padding-bottom: 2rem;
    max-width: 1200px;
}

.top-banner {
    background: linear-gradient(90deg, #FF000F 0%, #6764F6 100%);
    padding: 24px 28px;
    border-radius: 18px;
    color: white;
    margin-bottom: 20px;
}
.top-banner h1 {
    margin: 0;
    font-size: 32px;
    font-weight: 700;
}
.top-banner p {
    margin-top: 8px;
    margin-bottom: 0;
    font-size: 15px;
    opacity: 0.95;
}

.white-card {
    background: white;
    padding: 18px 20px;
    border-radius: 16px;
    border: 1px solid #EAEAEA;
    box-shadow: 0 2px 10px rgba(0,0,0,0.04);
    margin-bottom: 16px;
}

.section-title {
    font-size: 20px;
    font-weight: 700;
    color: #111111;
    margin-bottom: 10px;
    margin-top: 10px;
}

.helper-text {
    color: #666666;
    font-size: 14px;
}

.stButton > button {
    background-color: #FF000F;
    color: white;
    border-radius: 10px;
    border: none;
    padding: 0.65rem 1.2rem;
    font-weight: 600;
}
.stButton > button:hover {
    background-color: #d9000d;
    color: white;
}

[data-testid="stFileUploader"] {
    background: #FFFFFF;
    border: 1px solid #EAEAEA;
    border-radius: 14px;
    padding: 10px;
}

.metric-box {
    background: #F7F7F9;
    border-radius: 14px;
    padding: 16px;
    border: 1px solid #ECECEC;
    margin-bottom: 12px;
}
.metric-label {
    font-size: 13px;
    color: #666666;
}
.metric-value {
    font-size: 28px;
    font-weight: 700;
    color: #111111;
    margin-top: 4px;
}

.small-gap {
    height: 6px;
}
</style>
""", unsafe_allow_html=True)


def build_key_messages(overview, strengths, opportunities, bottom3):
    messages = []

    engagement = overview.get("engagement_score")
    company = overview.get("company_score")

    if engagement is not None and company is not None:
        gap = engagement - company
        if gap > 0:
            messages.append(f"Engagement score is {engagement}, above company by {gap} points.")
        elif gap < 0:
            messages.append(f"Engagement score is {engagement}, below company by {abs(gap)} points.")
        else:
            messages.append(f"Engagement score is {engagement}, in line with company.")

    if overview.get("response_rate") is not None:
        messages.append(f"Response rate reached {overview['response_rate']}%.")

    if strengths:
        top = strengths[0]
        messages.append(f"Top strength is {top['name']} ({top['score']}).")

    if opportunities:
        opp = opportunities[0]
        messages.append(f"Priority opportunity is {opp['name']} ({opp['score']}).")
    elif bottom3:
        opp = bottom3[0]
        messages.append(f"Lowest score is {opp['name']} ({opp['score']}).")

    if not messages:
        messages.append("Survey file was parsed successfully.")
        messages.append("Please review extracted metrics before distribution.")

    return messages[:4]


def metrics_to_dataframe(metrics):
    if not metrics:
        return pd.DataFrame(columns=["Metric", "Score"])
    return pd.DataFrame(
        [{"Metric": m["name"], "Score": m["score"]} for m in metrics]
    ).sort_values(by="Score", ascending=False)


st.markdown("""
<div class="top-banner">
    <h1>Survey PPT Generator</h1>
    <p>Transforming engagement survey findings into structured organizational development actions.</p>
</div>
""", unsafe_allow_html=True)

st.markdown('<div class="section-title">User Guide</div>', unsafe_allow_html=True)
st.markdown("""
<div class="white-card">
    <div class="helper-text">
        1. Upload your survey source file (PPT A)<br>
        2. Optional: upload a different template PPT B<br>
        3. Click <b>Generate PPT</b><br>
        4. Review and download the generated result
    </div>
</div>
""", unsafe_allow_html=True)

st.markdown('<div class="section-title">Input Files</div>', unsafe_allow_html=True)

col_upload1, col_upload2 = st.columns(2)

with col_upload1:
    st.markdown('<div class="white-card">', unsafe_allow_html=True)
    uploaded_a = st.file_uploader("Upload PPT A", type=["pptx"])
    st.markdown(
        '<div class="helper-text">Required. This is the survey result PPT.</div>',
        unsafe_allow_html=True,
    )
    st.markdown('</div>', unsafe_allow_html=True)

with col_upload2:
    st.markdown('<div class="white-card">', unsafe_allow_html=True)
    uploaded_template = st.file_uploader("Optional: upload template PPT B", type=["pptx"])
    st.markdown(
        '<div class="helper-text">Optional. If empty, the built-in template will be used.</div>',
        unsafe_allow_html=True,
    )
    st.markdown('</div>', unsafe_allow_html=True)

generate_btn = st.button("Generate PPT")

if generate_btn:
    if uploaded_a is None:
        st.error("Please upload PPT A first.")
        st.stop()

    repo_template = Path("templates/base_template.pptx")

    if uploaded_template is None and not repo_template.exists():
        st.error("Template file not found. Please upload template B or add templates/base_template.pptx.")
        st.stop()

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_dir = Path(tmp_dir)

        a_path = tmp_dir / f"input_a_{uuid.uuid4().hex}.pptx"
        a_path.write_bytes(uploaded_a.getbuffer())

        if uploaded_template is not None:
            template_path = tmp_dir / f"template_b_{uuid.uuid4().hex}.pptx"
            template_path.write_bytes(uploaded_template.getbuffer())
        else:
            template_path = repo_template

        try:
            with st.spinner("Parsing survey PPT and generating output..."):
                parsed = parse_survey_ppt(str(a_path))

                metrics = parsed["metrics"]
                strengths, opportunities, bottom10, top3, bottom3 = classify_metrics(metrics)
                actions = generate_actions(opportunities, bottom3)

                report_data = {
                    "overview": parsed["overview"],
                    "metrics": metrics,
                    "strengths": strengths,
                    "opportunities": opportunities,
                    "bottom10": bottom10,
                    "top3": top3,
                    "bottom3": bottom3,
                    "actions": actions,
                    "key_messages": build_key_messages(
                        parsed["overview"], strengths, opportunities, bottom3
                    ),
                }

                output_path = tmp_dir / f"generated_output_{uuid.uuid4().hex}.pptx"
                generate_ppt(str(template_path), str(output_path), report_data)

            st.success("PPT generated successfully.")

            overview = report_data["overview"]

            m1, m2, m3 = st.columns(3)

            with m1:
                st.markdown(f"""
                <div class="metric-box">
                    <div class="metric-label">Engagement Score</div>
                    <div class="metric-value">{overview.get('engagement_score', '-')}</div>
                </div>
                """, unsafe_allow_html=True)

            with m2:
                st.markdown(f"""
                <div class="metric-box">
                    <div class="metric-label">Company Score</div>
                    <div class="metric-value">{overview.get('company_score', '-')}</div>
                </div>
                """, unsafe_allow_html=True)

            with m3:
                response_rate = overview.get("response_rate")
                response_rate_text = f"{response_rate}%" if response_rate is not None else "-"
                st.markdown(f"""
                <div class="metric-box">
                    <div class="metric-label">Response Rate</div>
                    <div class="metric-value">{response_rate_text}</div>
                </div>
                """, unsafe_allow_html=True)

            left, right = st.columns([1, 1])

            with left:
                st.markdown('<div class="section-title">Overview</div>', unsafe_allow_html=True)
                st.markdown('<div class="white-card">', unsafe_allow_html=True)
                st.json(report_data["overview"])
                st.markdown('</div>', unsafe_allow_html=True)

            with right:
                st.markdown('<div class="section-title">Extracted Metrics</div>', unsafe_allow_html=True)
                st.markdown('<div class="white-card">', unsafe_allow_html=True)
                st.dataframe(metrics_to_dataframe(metrics), use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)

            c1, c2 = st.columns(2)

            with c1:
                st.markdown('<div class="section-title">Detected Strengths</div>', unsafe_allow_html=True)
                st.markdown('<div class="white-card">', unsafe_allow_html=True)
                if strengths:
                    for x in strengths:
                        st.write(f"✅ {x['name']} ({x['score']})")
                else:
                    st.write("No strengths detected.")
                st.markdown('</div>', unsafe_allow_html=True)

            with c2:
                st.markdown('<div class="section-title">Detected Opportunities</div>', unsafe_allow_html=True)
                st.markdown('<div class="white-card">', unsafe_allow_html=True)
                if opportunities:
                    for x in opportunities:
                        st.write(f"⚠️ {x['name']} ({x['score']})")
                else:
                    st.write("No opportunities detected.")
                st.markdown('</div>', unsafe_allow_html=True)

            st.markdown('<div class="section-title">Action Themes</div>', unsafe_allow_html=True)
            st.markdown('<div class="white-card">', unsafe_allow_html=True)
            for a in actions:
                st.markdown(f"**{a['theme']}**")
                for b in a["bullets"]:
                    st.write(f"- {b}")
            st.markdown('</div>', unsafe_allow_html=True)

            st.download_button(
                label="Download Final PPT",
                data=output_path.read_bytes(),
                file_name="generated_survey_followup.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )

        except Exception as e:
            st.error("Generation failed.")
            st.exception(e)

st.markdown("---")
st.caption("Internal productivity tool | Survey PPT automation")
