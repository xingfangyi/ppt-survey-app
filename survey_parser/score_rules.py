def _deduplicate(metrics):
    best = {}
    for m in metrics:
        name = m["name"]
        score = m["score"]
        if name not in best or score > best[name]["score"]:
            best[name] = {"name": name, "score": score}
    return list(best.values())


METRIC_PATTERNS = {
    "Manager Trust": ["manager trust", "trust my manager", "i trust my manager"],
    "Safety Climate": ["safety climate", "safety"],
    "Risk Taking": ["risk taking"],
    "Performance Feedback": ["performance feedback", "feedback environment"],
    "Inclusion": ["inclusion", "inclusive environment", "valuing perspectives"],
    "Development Discussions": ["development discussions", "development discussion"],
    "Accountability": ["accountability"],
    "Role Clarity": ["role clarity"],
    "Integrity": ["integrity"],
    "Growth Opportunities": ["growth opportunities", "growth"],
    "Pride": ["pride"],
    "No-Fear Culture": ["no fear", "no-fear culture"],
    "Empowerment": ["empowerment", "empowered"],
    "Well-Being": ["well-being", "well being"],
    "Company Direction": ["company direction", "direction"],
    "Recognition": ["recognition", "recognized"],
    "Fair Treatment": ["fair treatment"],
    "Work-Life Balance": ["work-life balance", "work life balance", "balance my work"],
    "Rewards": ["rewards"],
    "Rewards Fairness": ["rewards fairness", "reward fairness", "compensation fairness", "fair pay"],
    "Challenge to Status Quo": ["challenge to status quo", "status quo"],
    "Living Company Values": ["living company values", "company values"],
    "Customer Focus": ["customer focus"],
    "Collaboration": ["collaboration"],
    "Belonging": ["belonging", "sense of belong", "sense of belonging"],
    "Resources": ["resources"],
    "Sustainability": ["sustainability"],
    "Communication Flow": ["communication flow"],
    "Initiative": ["initiative"],
}


def classify_metrics(metrics):
    metrics = _deduplicate(metrics)
    metrics_sorted_desc = sorted(metrics, key=lambda x: x["score"], reverse=True)
    metrics_sorted_asc = sorted(metrics, key=lambda x: x["score"])

    strengths = [m for m in metrics_sorted_desc if m["score"] >= 90][:5]
    opportunities = [m for m in metrics_sorted_asc if m["score"] <= 80][:5]
    bottom10 = metrics_sorted_asc[:10]
    top3 = metrics_sorted_desc[:3]
    bottom3 = metrics_sorted_asc[:3]

    return strengths, opportunities, bottom10, top3, bottom3
