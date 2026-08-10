import tempfile
import uuid
from pathlib import Path

import pandas as pd
import streamlit as st

from survey_parser.extract_a import parse_survey_ppt
from survey_parser.score_rules import classify_metrics
from survey_parser.action_mapper import generate_actions
from ppt_generator.build_ppt import generate_ppt


st.set_page_config(page_title="PPT Survey Auto Generator", layout="wide")


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


st.title("PPT Survey Auto Generator")
st.caption("Upload PPT A, then generate a new PPT based on your fixed template B.")

with st.expander("How to use", expanded=True):
    st.markdown(
        """
1. Upload your **PPT A** survey file  
2. Optional: upload a different **template B**  
3. Click **Generate PPT**  
4. Download the generated result
        """
    )

uploaded_a = st.file_uploader("Upload PPT A", type=["pptx"])
uploaded_template = st.file_uploader(
    "Optional: upload template PPT B", type=["pptx"]
)

generate_btn = st.button("Generate PPT", type="primary")

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

            col1, col2 = st.columns([1, 1])

            with col1:
                st.subheader("Overview")
                st.json(report_data["overview"])

            with col2:
                st.subheader("Extracted Metrics")
                st.dataframe(metrics_to_dataframe(metrics), use_container_width=True)

            st.subheader("Detected Strengths")
            if strengths:
                st.write([f"{x['name']} ({x['score']})" for x in strengths])
            else:
                st.write("No strengths detected.")

            st.subheader("Detected Opportunities")
            if opportunities:
                st.write([f"{x['name']} ({x['score']})" for x in opportunities])
            else:
                st.write("No opportunities detected.")

            st.subheader("Action Themes")
            for a in actions:
                st.markdown(f"**{a['theme']}**")
                for b in a["bullets"]:
                    st.write(f"- {b}")

            st.download_button(
                label="Download generated PPT",
                data=output_path.read_bytes(),
                file_name="generated_survey_followup.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )

        except Exception as e:
            st.error("Generation failed.")
            st.exception(e)
