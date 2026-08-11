from pathlib import Path

import pandas as pd
from plotnine import (
    aes,
    coord_flip,
    element_blank,
    element_text,
    geom_col,
    ggplot,
    labs,
    scale_y_continuous,
    theme,
    theme_bw,
)


ABB_RED = "#FF000F"
ABB_LILAC = "#6764F6"


def _safe_metrics_df(metrics, fallback_name="No data", fallback_value=0):
    if not metrics:
        return pd.DataFrame({"name": [fallback_name], "score": [fallback_value]})
    return pd.DataFrame(metrics)


def _save_horizontal_bar(df, title, path, color=ABB_RED):
    df = df.copy()
    if "name" not in df.columns or "score" not in df.columns:
        raise ValueError("Dataframe must contain 'name' and 'score' columns")

    df = df.sort_values("score", ascending=True)
    df["name"] = pd.Categorical(df["name"], categories=df["name"].tolist(), ordered=True)

    plot = (
        ggplot(df, aes(x="name", y="score"))
        + geom_col(fill=color)
        + coord_flip()
        + scale_y_continuous(limits=[0, 100])
        + labs(title=title, x="", y="Score")
        + theme_bw()
        + theme(
            figure_size=(8, 5),
            legend_position="none",
            panel_grid_minor=element_blank(),
            panel_grid_major_y=element_blank(),
            plot_title=element_text(size=14, weight="bold"),
            axis_title_x=element_text(size=11),
            axis_text_y=element_text(size=10),
            axis_text_x=element_text(size=9),
        )
    )
    plot.save(path, dpi=300, width=8, height=5, verbose=False)
    return str(path)


def _save_vertical_bar(df, title, path, color=ABB_RED, ymax=100):
    df = df.copy()
    if "name" not in df.columns or "score" not in df.columns:
        raise ValueError("Dataframe must contain 'name' and 'score' columns")

    plot = (
        ggplot(df, aes(x="name", y="score"))
        + geom_col(fill=color)
        + scale_y_continuous(limits=[0, ymax])
        + labs(title=title, x="", y="Score")
        + theme_bw()
        + theme(
            figure_size=(8, 5),
            legend_position="none",
            panel_grid_minor=element_blank(),
            panel_grid_major_x=element_blank(),
            plot_title=element_text(size=14, weight="bold"),
            axis_title_x=element_blank(),
            axis_text_x=element_text(size=9, rotation=15, ha="right"),
            axis_text_y=element_text(size=9),
        )
    )
    plot.save(path, dpi=300, width=8, height=5, verbose=False)
    return str(path)


def _avg(values):
    vals = [v for v in values if v is not None]
    if not vals:
        return None
    return round(sum(vals) / len(vals), 1)


def _build_overview_compare_df(overview):
    rows = []
    if overview.get("engagement_score") is not None:
        rows.append({"name": "Engagement", "score": overview["engagement_score"]})
    if overview.get("company_score") is not None:
        rows.append({"name": "Company", "score": overview["company_score"]})
    if overview.get("response_rate") is not None:
        rows.append({"name": "Response Rate", "score": overview["response_rate"]})

    if not rows:
        rows = [{"name": "No data", "score": 0}]
    return pd.DataFrame(rows)


def _build_score_band_df(metrics):
    scores = [m["score"] for m in metrics] if metrics else []
    bands = [
        {"name": "90+", "score": sum(1 for s in scores if s >= 90)},
        {"name": "80-89", "score": sum(1 for s in scores if 80 <= s < 90)},
        {"name": "70-79", "score": sum(1 for s in scores if 70 <= s < 80)},
        {"name": "<70", "score": sum(1 for s in scores if s < 70)},
    ]
    return pd.DataFrame(bands)


def _build_focus_summary_df(metrics, top3, bottom3):
    rows = []
    all_avg = _avg([m["score"] for m in metrics])
    top_avg = _avg([m["score"] for m in top3])
    bottom_avg = _avg([m["score"] for m in bottom3])

    if top_avg is not None:
        rows.append({"name": "Top 3 Avg", "score": top_avg})
    if bottom_avg is not None:
        rows.append({"name": "Bottom 3 Avg", "score": bottom_avg})
    if all_avg is not None:
        rows.append({"name": "All Metrics Avg", "score": all_avg})

    if not rows:
        rows = [{"name": "No data", "score": 0}]
    return pd.DataFrame(rows)


def _build_action_priority_df(actions, opportunities, bottom3):
    source_scores = [x["score"] for x in opportunities[:2]]
    if len(source_scores) < 2:
        source_scores += [x["score"] for x in bottom3[: 2 - len(source_scores)]]

    rows = []
    for i, action in enumerate(actions[:2]):
        score = source_scores[i] if i < len(source_scores) else 75 - i * 5
        rows.append({"name": action["theme"], "score": score})

    if not rows:
        rows = [{"name": "No action", "score": 0}]
    return pd.DataFrame(rows)


def generate_all_charts(report_data, output_dir):
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    overview = report_data["overview"]
    metrics = report_data["metrics"]
    strengths = report_data["strengths"]
    opportunities = report_data["opportunities"]
    bottom10 = report_data["bottom10"]
    top3 = report_data["top3"]
    bottom3 = report_data["bottom3"]
    actions = report_data["actions"]

    chart_paths = {}

    # Slide 2
    df_overview = _build_overview_compare_df(overview)
    chart_paths["overview_compare"] = _save_vertical_bar(
        df_overview,
        "Survey Snapshot",
        output_dir / "overview_compare.png",
        color=ABB_RED,
        ymax=100,
    )

    df_bands = _build_score_band_df(metrics)
    max_band = max(int(df_bands["score"].max()), 1) + 1
    chart_paths["score_bands"] = _save_vertical_bar(
        df_bands,
        "Score Band Distribution",
        output_dir / "score_bands.png",
        color=ABB_LILAC,
        ymax=max_band,
    )

    # Slide 3
    df_strengths = _safe_metrics_df(strengths[:5])
    chart_paths["strengths"] = _save_horizontal_bar(
        df_strengths,
        "Top Strengths",
        output_dir / "strengths.png",
        color=ABB_RED,
    )

    df_opps = _safe_metrics_df(opportunities[:5])
    chart_paths["opportunities"] = _save_horizontal_bar(
        df_opps,
        "Top Opportunities",
        output_dir / "opportunities.png",
        color=ABB_LILAC,
    )

    # Slide 4
    df_bottom_a = _safe_metrics_df(bottom10[:5])
    chart_paths["bottom10_a"] = _save_horizontal_bar(
        df_bottom_a,
        "Bottom 10 (1-5)",
        output_dir / "bottom10_a.png",
        color=ABB_RED,
    )

    df_bottom_b = _safe_metrics_df(bottom10[5:10])
    chart_paths["bottom10_b"] = _save_horizontal_bar(
        df_bottom_b,
        "Bottom 10 (6-10)",
        output_dir / "bottom10_b.png",
        color=ABB_LILAC,
    )

    # Slide 5
    df_top3 = _safe_metrics_df(top3[:3])
    chart_paths["top3"] = _save_horizontal_bar(
        df_top3,
        "Top 3 Scores",
        output_dir / "top3.png",
        color=ABB_RED,
    )

    df_summary = _build_focus_summary_df(metrics, top3, bottom3)
    chart_paths["focus_summary"] = _save_vertical_bar(
        df_summary,
        "Score Summary",
        output_dir / "focus_summary.png",
        color=ABB_LILAC,
        ymax=100,
    )

    # Slide 6
    df_focus = _safe_metrics_df(opportunities[:5])
    chart_paths["focus_topics"] = _save_horizontal_bar(
        df_focus,
        "Improvement Topics",
        output_dir / "focus_topics.png",
        color=ABB_RED,
    )

    df_actions = _build_action_priority_df(actions, opportunities, bottom3)
    chart_paths["action_themes"] = _save_vertical_bar(
        df_actions,
        "Action Themes",
        output_dir / "action_themes.png",
        color=ABB_LILAC,
        ymax=100,
    )

    # Slide 7
    df_top = _safe_metrics_df(top3[:3])
    chart_paths["top3_repeat"] = _save_horizontal_bar(
        df_top,
        "Top 3",
        output_dir / "top3_repeat.png",
        color=ABB_RED,
    )

    df_bottom = _safe_metrics_df(bottom3[:3])
    chart_paths["bottom3"] = _save_horizontal_bar(
        df_bottom,
        "Bottom 3",
        output_dir / "bottom3.png",
        color=ABB_LILAC,
    )

    return chart_paths
