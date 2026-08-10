ACTION_LIBRARY = [
    {
        "keywords": ["work-life balance", "well-being", "well being"],
        "theme": "Work-Life Balance",
        "quote": "I am able to successfully balance my work and personal life.",
        "bullets": [
            "Assess workload and maintenance windows earlier to reduce overtime.",
            "Review team scheduling and align tasks more evenly across resources.",
            "Encourage managers to monitor leave usage and recovery time.",
            "Use customer communication earlier to avoid last-minute peaks.",
        ],
    },
    {
        "keywords": ["recognition", "fair treatment"],
        "theme": "Recognition",
        "quote": "I feel recognized for the work I do.",
        "bullets": [
            "Launch a monthly or quarterly recognition rhythm.",
            "Make recognition more timely and visible across the team.",
            "Clarify recognition criteria to improve fairness perception.",
            "Share success stories and good practices more often.",
        ],
    },
    {
        "keywords": ["empowerment", "initiative", "status quo"],
        "theme": "Empowerment",
        "quote": "I have the authority and support to take action.",
        "bullets": [
            "Clarify decision boundaries for frontline employees.",
            "Encourage local problem-solving before escalation.",
            "Use manager coaching to support faster decisions.",
            "Reward improvement ideas and practical experimentation.",
        ],
    },
    {
        "keywords": ["rewards", "rewards fairness", "fair pay", "compensation fairness"],
        "theme": "Rewards Fairness",
        "quote": "Rewards are fair and clearly understood.",
        "bullets": [
            "Explain reward principles and performance links more clearly.",
            "Increase transparency on reward criteria where possible.",
            "Use manager communication to answer employee concerns early.",
            "Review local feedback on fairness perception regularly.",
        ],
    },
    {
        "keywords": ["belonging", "inclusion"],
        "theme": "Belonging",
        "quote": "I feel a sense of belonging at ABB.",
        "bullets": [
            "Organize regular team connection activities.",
            "Use site visits and skip-level discussions to hear employee feedback.",
            "Recognize contributions from different roles and locations.",
            "Share team achievements to strengthen visibility and inclusion.",
        ],
    },
    {
        "keywords": ["company direction", "communication flow"],
        "theme": "Communication",
        "quote": "I understand where the organization is going and why.",
        "bullets": [
            "Translate strategy into team-level priorities more clearly.",
            "Repeat key messages through regular team meetings.",
            "Create a short follow-up loop after important updates.",
            "Use examples to connect strategy with daily work.",
        ],
    },
]


def _match_action(metric_name):
    name = metric_name.lower()
    for item in ACTION_LIBRARY:
        if any(k in name for k in item["keywords"]):
            return item
    return None


def generate_actions(opportunities, bottom3):
    selected = []
    used_themes = set()

    candidates = opportunities + bottom3

    for metric in candidates:
        action = _match_action(metric["name"])
        if action and action["theme"] not in used_themes:
            selected.append(action)
            used_themes.add(action["theme"])
        if len(selected) >= 2:
            break

    if len(selected) < 2:
        for fallback in ACTION_LIBRARY:
            if fallback["theme"] not in used_themes:
                selected.append(fallback)
                used_themes.add(fallback["theme"])
            if len(selected) >= 2:
                break

    return selected[:2]
