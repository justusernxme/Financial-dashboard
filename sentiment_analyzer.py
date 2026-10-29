#!/usr/bin/env python3
"""SEC 10-K MD&A Loughran-McDonald Sentiment Analyzer.

This module loads the Loughran-McDonald financial sentiment dictionary across
six key categories:
    1. Positive
    2. Negative
    3. Uncertainty
    4. Litigious
    5. Strong Modal
    6. Weak Modal

It cleans annual MD&A text files (`data/<TICKER>/mda/mda_<YEAR>.txt`) by
converting to lowercase, removing numbers and punctuation, and filtering out
stop words. It then performs financial sentiment scoring per year, computing:
    - Total Word Count (Cleaned)
    - Positive Word Count
    - Negative Word Count
    - Net Sentiment Score (Positive - Negative)
    - Positive / Negative (P/N) Ratio
    - Uncertainty, Litigious, and Modal Word Counts

Finally, it writes these annual metrics into a dedicated tab named
'Sentiment Analysis' within the `<TICKER>_Financial_Analysis.xlsx` workbook,
styled with professional corporate executive formatting.

Usage:
    Interactive mode:
        python sentiment_analyzer.py

    Command-line argument mode:
        python sentiment_analyzer.py --ticker AAPL
        python sentiment_analyzer.py --ticker MSFT --workbook custom.xlsx
"""

import argparse
import csv
import io
import logging
import re
import sys
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# Constants and Stop Words
# ---------------------------------------------------------------------------

DEFAULT_DATA_DIR: Path = Path("data")

# Official Notre Dame SRAF Loughran-McDonald Master Dictionary URL
LMD_OFFICIAL_URL: str = (
    "https://sraf.nd.edu/wp-content/uploads/2024/08/"
    "Loughran-McDonald_MasterDictionary_1993-2023.csv"
)

# Professional Corporate Palette for Excel Styling
COLOR_HEADER_NAVY: str = "1F4E79"
COLOR_HEADER_TEXT: str = "FFFFFF"
COLOR_ZEBRA_ROW: str = "F9FBFD"
COLOR_BORDER_GRAY: str = "D3D3D3"
COLOR_POSITIVE_GREEN: str = "006100"
COLOR_POSITIVE_BG: str = "C6EFCE"
COLOR_NEGATIVE_RED: str = "9C0006"
COLOR_NEGATIVE_BG: str = "FFC7CE"

# Standard English stop words
STANDARD_STOP_WORDS: Set[str] = {
    "a", "about", "above", "after", "again", "against", "all", "am", "an",
    "and", "any", "are", "aren't", "as", "at", "be", "because", "been",
    "before", "being", "below", "between", "both", "but", "by", "can",
    "cannot", "could", "couldn't", "did", "didn't", "do", "does", "doesn't",
    "doing", "don't", "down", "during", "each", "few", "for", "from",
    "further", "had", "hadn't", "has", "hasn't", "have", "haven't",
    "having", "he", "he'd", "he'll", "he's", "her", "here", "here's",
    "hers", "herself", "him", "himself", "his", "how", "how's", "i", "i'd",
    "i'll", "i'm", "i've", "if", "in", "into", "is", "isn't", "it", "it's",
    "its", "itself", "let's", "me", "more", "most", "mustn't", "my",
    "myself", "no", "nor", "not", "of", "off", "on", "once", "only", "or",
    "other", "ought", "our", "ours", "ourselves", "out", "over", "own",
    "same", "shan't", "she", "she'd", "she'll", "she's", "should",
    "shouldn't", "so", "some", "such", "than", "that", "that's", "the",
    "their", "theirs", "them", "themselves", "then", "there", "there's",
    "these", "they", "they'd", "they'll", "they're", "they've", "this",
    "those", "through", "to", "too", "under", "until", "up", "very", "was",
    "wasn't", "we", "we'd", "we'll", "we're", "we've", "were", "weren't",
    "what", "what's", "when", "when's", "where", "where's", "which",
    "while", "who", "who's", "whom", "why", "why's", "with", "won't",
    "would", "wouldn't", "you", "you'd", "you'll", "you're", "you've",
    "your", "yours", "yourself", "yourselves",
}


# ---------------------------------------------------------------------------
# Comprehensive Built-In Loughran-McDonald Financial Lexicon
# ---------------------------------------------------------------------------

# Curated canonical stems and terms from Loughran-McDonald (2011)
BUILTIN_LMD_POSITIVE: Set[str] = {
    "able", "abundance", "abundant", "acclaimed", "accomplish", "accomplished",
    "achieve", "achievement", "achieving", "advantage", "advantageous",
    "attractive", "attain", "attained", "attaining", "beautiful", "benefactor",
    "beneficial", "benefit", "benefited", "benefiting", "benefits", "best",
    "better", "blessing", "bolster", "bolstered", "bolstering", "boom",
    "booming", "boost", "boosted", "boosting", "breakthrough", "bright",
    "brilliant", "capable", "collaborate", "collaborating", "collaboration",
    "collaborative", "competitive", "compliment", "complimentary",
    "comprehensive", "conducive", "confident", "constructive", "convenient",
    "creative", "cure", "delight", "delighted", "delightful", "dependable",
    "desirable", "distinction", "distinctive", "distinguished", "dynamic",
    "eager", "effective", "effectively", "effectiveness", "efficiency",
    "efficient", "efficiently", "empower", "empowered", "empowering",
    "enable", "enabled", "enabling", "encourage", "encouraging", "enhance",
    "enhanced", "enhancement", "enhancements", "enhancing", "enjoy",
    "enjoyable", "enjoyed", "enthusiasm", "enthusiastic", "excel", "excelled",
    "excellence", "excellent", "excelling", "exceptional", "exceptionally",
    "excite", "excited", "exciting", "exclusive", "favor", "favorable",
    "favorably", "favored", "flourish", "flourishing", "gain", "gained",
    "gaining", "gains", "generous", "glad", "good", "great", "greater",
    "greatest", "greatly", "grow", "growing", "growth", "growths", "happy",
    "harmony", "ideal", "impress", "impressed", "impressive", "improve",
    "improved", "improvement", "improvements", "improving", "incredible",
    "influential", "ingenious", "innovation", "innovations", "innovative",
    "innovator", "insightful", "inspire", "inspired", "inspiring",
    "invaluable", "inventive", "leader", "leading", "leadership", "lucrative",
    "mastery", "merit", "meritorious", "milestone", "modern", "notable",
    "notably", "opportunity", "opportunities", "optimal", "optimism",
    "optimistic", "outperform", "outperformed", "outperforming", "outperforms",
    "pioneer", "pioneered", "pioneering", "pleased", "pleasing", "pleasure",
    "plentiful", "popular", "positive", "positively", "premier", "prestige",
    "prestigious", "productive", "profit", "profitability", "profitable",
    "profitably", "profited", "profiting", "profits", "progress",
    "progressing", "progressive", "prosper", "prospered", "prospering",
    "prosperity", "prosperous", "rebound", "rebounded", "rebounding",
    "record", "records", "recover", "recovered", "recovery", "refine",
    "refined", "reliable", "resolve", "resolved", "reward", "rewarded",
    "rewarding", "rewards", "robust", "robustly", "satisfaction",
    "satisfactorily", "satisfactory", "satisfied", "satisfy", "satisfying",
    "secure", "secured", "securing", "significant", "significantly",
    "smooth", "smoothly", "solid", "solidity", "spectacular", "stability",
    "stabilize", "stabilized", "stable", "strength", "strengthen",
    "strengthened", "strengthening", "strengthens", "strengths", "strong",
    "stronger", "strongest", "strongly", "succeed", "succeeded", "succeeding",
    "succeeds", "success", "successes", "successful", "successfully",
    "superior", "superiority", "surpass", "surpassed", "thrive", "thrived",
    "thriving", "tremendous", "triumph", "unmatched", "unprecedented",
    "upbeat", "upgrade", "upgraded", "valuable", "value", "vigor",
    "vigorous", "vital", "win", "winner", "winning", "wins", "worthwhile",
}

BUILTIN_LMD_NEGATIVE: Set[str] = {
    "abandon", "abandoned", "abandoning", "abandonment", "abdicate",
    "aberrant", "aberration", "abeyance", "abnormal", "abnormally",
    "abolish", "abolished", "abrogate", "abrupt", "abruptly", "absence",
    "absent", "abuse", "abused", "abusive", "accident", "accidental",
    "adversary", "adverse", "adversely", "adversity", "against", "aggravate",
    "aggravated", "aggravating", "alert", "alerted", "alienate", "allege",
    "allegation", "allegations", "alleged", "allegedly", "alleges", "annul",
    "annulled", "annulment", "anomaly", "anticompetitive", "anxiety",
    "anxious", "arrearage", "arrears", "artificially", "assault", "attrition",
    "backlog", "bad", "badly", "bailout", "balk", "bankrupt", "bankruptcies",
    "bankruptcy", "bar", "barrier", "barriers", "bearish", "blockage",
    "bottleneck", "breach", "breached", "breaches", "breaching", "breakdown",
    "bribe", "bribery", "burden", "burdened", "burdensome", "calamity",
    "cancel", "canceled", "canceling", "cancellation", "cancellations",
    "cancels", "catastrophe", "catastrophic", "caution", "cautionary",
    "cease", "ceased", "ceasing", "censure", "censured", "challenge",
    "challenged", "challenges", "challenging", "chargeoff", "claim",
    "claimed", "claiming", "claims", "clawback", "closure", "closures",
    "coerce", "coercion", "collapse", "collapsed", "collapsing", "collision",
    "collusion", "complaint", "complaints", "complicate", "complicated",
    "complication", "complications", "compulsion", "compulsory", "concern",
    "concerned", "concerns", "concession", "conflict", "conflicted",
    "conflicting", "conflicts", "confront", "confrontation", "confuse",
    "confused", "confusion", "conspire", "conspiracy", "constraint",
    "constraints", "contagion", "contempt", "contend", "contention",
    "contentious", "contest", "contested", "contingency", "contingent",
    "contracted", "contraction", "contradict", "contradiction", "contrary",
    "controversial", "controversy", "convict", "convicted", "conviction",
    "corrupt", "corruption", "costly", "counterclaim", "counterfeit",
    "counterproductive", "crimp", "crimped", "crisis", "critical",
    "criticism", "criticize", "criticized", "crucial", "culpability",
    "culpable", "curtail", "curtailed", "curtailment", "cutback", "cutbacks",
    "damage", "damaged", "damages", "damaging", "danger", "dangerous",
    "deadlock", "debar", "debarred", "debarment", "decease", "deceased",
    "deceit", "deceitful", "deceive", "deceived", "deception", "deceptive",
    "decline", "declined", "declines", "declining", "decrease", "decreased",
    "decreases", "decreasing", "default", "defaulted", "defaulting",
    "defaults", "defeat", "defeated", "defect", "defective", "defects",
    "defend", "defendant", "defendants", "defending", "deficiency",
    "deficient", "deficit", "deficits", "defraud", "defrauded", "degrade",
    "degraded", "degrading", "delay", "delayed", "delaying", "delays",
    "deleterious", "deliberate", "delinquency", "delinquent", "delist",
    "delisted", "delisting", "demise", "demolish", "demotion", "demur",
    "denial", "denials", "denied", "deny", "denying", "deplete", "depleted",
    "depletion", "deprecate", "depress", "depressed", "depressing",
    "depression", "deprive", "deprived", "derelict", "derogatory",
    "destabilize", "destabilized", "destroy", "destroyed", "destroying",
    "destruction", "destructive", "detain", "detained", "deter",
    "deteriorate", "deteriorated", "deteriorating", "deterioration",
    "deterrent", "detriment", "detrimental", "devalue", "devalued",
    "devastate", "devastated", "devastating", "devastation", "deviate",
    "deviation", "difficult", "difficulties", "difficulty", "diminish",
    "diminished", "diminishes", "diminishing", "disadvantage", "disaffected",
    "disagree", "disagreement", "disallow", "disallowed", "disappear",
    "disappoint", "disappointed", "disappointing", "disappointment",
    "disapproval", "disapprove", "disaster", "disastrous", "discharge",
    "disclaim", "disclaimed", "disclaimer", "disclose", "discontinuance",
    "discontinue", "discontinued", "discontinuing", "discourage", "discredit",
    "discrepancy", "discriminate", "discrimination", "diseconomy",
    "disfavor", "disgorge", "disgorgement", "disgrace", "dishonor",
    "disincentive", "dislocate", "dislocated", "dislocation", "dismiss",
    "dismissal", "dismissed", "disorder", "disorderly", "disparage",
    "disparaging", "disparity", "displace", "displaced", "displacement",
    "dispute", "disputed", "disputes", "disputing", "disqualify",
    "disregard", "disrupt", "disrupted", "disrupting", "disruption",
    "disruptions", "disruptive", "dissatisfaction", "dissident", "dissolution",
    "distort", "distorted", "distortion", "distress", "distressed",
    "distrust", "divest", "divested", "divestiture", "divestitures",
    "divestment", "divisive", "divorce", "doubt", "doubtful", "doubts",
    "downfall", "downgrade", "downgraded", "downgrades", "downgrading",
    "downsize", "downsized", "downsizing", "downturn", "downturns",
    "downward", "downwards", "drag", "dragged", "drain", "drastic",
    "drastically", "drawback", "drop", "dropped", "dropping", "drops",
    "drought", "duress", "dwindle", "dwindling", "embargo", "embarrass",
    "embezzle", "embezzlement", "encroach", "encumber", "encumbrance",
    "erode", "eroded", "eroding", "erosion", "erratic", "erroneous",
    "error", "errors", "escalate", "escalated", "evict", "eviction",
    "exacerbate", "exacerbated", "exaggerate", "excessive", "exclude",
    "excluded", "exclusion", "fail", "failed", "failing", "fails", "failure",
    "failures", "fallout", "false", "falsely", "falsification", "falsify",
    "fatal", "fatality", "fault", "faulty", "fear", "fears", "felony",
    "fictitious", "fined", "fines", "fired", "flaw", "flawed", "flaws",
    "forbid", "forbidden", "force", "forced", "forcing", "foreclose",
    "foreclosed", "foreclosure", "foreclosures", "forego", "forfeit",
    "forfeited", "forfeiture", "fraud", "frauds", "fraudulent", "frivolous",
    "frustrate", "frustrated", "frustration", "fugitive", "glitch", "gloom",
    "grievance", "grossly", "guilty", "halt", "halted", "halting", "hamper",
    "hampered", "hampering", "hardship", "harm", "harmed", "harmful",
    "harsh", "hazard", "hazardous", "hindrance", "hostile", "hostility",
    "hurt", "hurting", "idle", "idled", "illegal", "illegality", "illegally",
    "illiquid", "illiquidity", "immaterial", "impair", "impaired",
    "impairing", "impairment", "impairments", "impede", "impeded",
    "impediment", "imperil", "implicated", "impossible", "impound",
    "improper", "improperly", "inability", "inaccessible", "inaccuracy",
    "inaccurate", "inadequacy", "inadequate", "inadvertent", "inadvisable",
    "inappropriate", "incapable", "incompetence", "incomplete", "inconsistent",
    "inconvenience", "incorrect", "indecent", "indict", "indicted",
    "indictment", "ineffective", "ineffectively", "inefficiency",
    "inefficient",
    "ineligible", "inequitable", "inevitable", "infamous", "infect",
    "inferior", "inflict", "infraction", "infringe", "infringement",
    "inhibit", "inhibited", "injunction", "injure", "injured", "injurious",
    "injury", "inordinate", "inquiry", "insecure", "insolvency", "insolvent",
    "instability", "insufficient", "insufficiently", "insurrection",
    "interfere", "interference", "interrupted", "interruption", "intimidate",
    "intrusion", "invalidate", "invalidated", "investigate", "investigated",
    "investigating", "investigation", "investigations", "involuntary",
    "irregular", "irregularity", "irreparable", "jeopardize", "jeopardy",
    "judgment", "lapse", "lapsed", "late", "laundering", "lawsuit",
    "lawsuits", "layoff", "layoffs", "leakage", "legal", "liabilities",
    "liability", "liable", "libel", "limitation", "limitations", "liquidate",
    "liquidated", "liquidating", "liquidation", "litigant", "litigate",
    "litigated", "litigating", "litigation", "litigations", "lockout",
    "loss", "losses", "lost", "malfeasance", "malfunction", "malpractice",
    "manipulate", "manipulation", "misappropriation", "miscalculate",
    "misconduct", "misleading", "misrepresent", "misrepresentation",
    "negative", "negatively", "neglect", "negligence", "negligent",
    "noncompliance", "obsolete", "obstacle", "offense", "omission",
    "penalize", "penalized", "penalties", "penalty", "peril", "perjury",
    "pessimism", "pessimistic", "petition", "plaintiff", "plea", "plunge",
    "problem", "problems", "prosecute", "prosecution", "protest", "punish",
    "punitive", "recession", "recessionary", "restructure", "restructured",
    "restructuring", "risk", "risks", "risky", "sanction", "sanctions",
    "scandal", "scrutiny", "seizure", "sentence", "settlement", "severe",
    "severely", "severity", "shortage", "shortages", "shortfall", "shortfalls",
    "shrink", "shrinkage", "shutdown", "shutdowns", "slump", "stagnant",
    "stagnation", "strain", "strike", "strikes", "subpoena", "subpoenas",
    "sue", "sued", "suing", "suit", "suits", "suspend", "suspended",
    "suspension", "tarnish", "terminate", "terminated", "terminating",
    "termination", "terminations", "threat", "threaten", "threatened",
    "threatening", "threats", "trouble", "troubled", "unanticipated",
    "uncertain", "uncertainties", "uncertainty", "uncollectible", "understate",
    "undesirable", "uneconomic", "unemployed", "unemployment", "unethical",
    "unfair", "unfavorable", "unfavorably", "unforeseen", "unfortunate",
    "unlawful", "unlawfully", "unlicensed", "unliquidated", "unpaid",
    "unplanned", "unprofitable", "unprotected", "unreliable", "unresolved",
    "unsafe", "unsettled", "unsound", "unstable", "unsuccessful",
    "unsuccessfully", "unsuitable", "untenable", "untested", "untimely",
    "untrue", "untrustworthy", "unusual", "unwarranted", "violate",
    "violated", "violating", "violation", "violations", "volatile",
    "volatility", "vulnerable", "warn", "warned", "warning", "warnings",
    "warns", "waste", "wasted", "weak", "weaken", "weakened", "weakening",
    "weakness", "weaknesses", "withhold", "withheld", "worsen", "worsened",
    "worsening", "worthless", "writedown", "writedowns", "writeoff",
}

BUILTIN_LMD_UNCERTAINTY: Set[str] = {
    "abeyance", "almost", "alter", "alteration", "ambiguity", "ambiguous",
    "anomalous", "anomaly", "anticipate", "anticipated", "anticipates",
    "anticipating", "anticipation", "approximate", "approximately",
    "arbitrary", "assume", "assumed", "assumes", "assuming", "assumption",
    "assumptions", "believe", "believed", "believes", "believing", "caution",
    "cautionary", "cautious", "clarify", "conceivable", "conditional",
    "contingencies", "contingency", "contingent", "contingently", "depend",
    "dependence", "dependencies", "dependency", "dependent", "depending",
    "depends", "destabilize", "destabilized", "destabilizing", "deviate",
    "deviated", "deviates", "deviating", "deviation", "deviations", "differ",
    "differed", "differing", "differs", "doubt", "doubtful", "doubts",
    "exposure", "exposures", "fluctuate", "fluctuated", "fluctuates",
    "fluctuating", "fluctuation", "fluctuations", "forecast", "forecasted",
    "forecasting", "forecasts", "foresee", "foreseeable", "foreseen",
    "hesitant", "hesitate", "hidden", "imprecise", "imprecision",
    "improbability", "improbable", "incomplete", "indefinite", "indefinitely",
    "indeterminate", "inexact", "instability", "intangible", "likelihood",
    "may", "maybe", "might", "nearly", "nonbinding", "pending", "perhaps",
    "possibility", "possible", "possibly", "potential", "potentially",
    "preliminary", "presume", "probable", "probably", "provisional",
    "random", "reassess", "reassessing", "recalculation", "reconsider",
    "reconsidering", "reexamination", "reinterpretation", "risk", "risks",
    "risky", "roughly", "rumor", "rumors", "scenarios", "speculate",
    "speculated", "speculating", "speculation", "speculative", "subject",
    "susceptible", "tentative", "tentatively", "turbulence", "uncertain",
    "uncertainly", "uncertainties", "uncertainty", "unclear", "unconfirmed",
    "undecided", "undefined", "undetermined", "unforeseeable", "unforeseen",
    "unpredictability", "unpredictable", "unpredicted", "unresolved",
    "unsettled", "unspecific", "untested", "unusual", "vague", "vaguely",
    "vagueness", "variability", "variable", "variables", "variance",
    "variances", "variation", "variations", "vary", "varying", "volatile",
    "volatility",
}

BUILTIN_LMD_LITIGIOUS: Set[str] = {
    "accord", "adjudicate", "adjudicated", "adjudicating", "adjudication",
    "affidavit", "affidavits", "allege", "allegation", "allegations",
    "alleged", "allegedly", "alleges", "amending", "amendment", "appeal",
    "appealed", "appealing", "appeals", "appellate", "arbitrate",
    "arbitrated", "arbitration", "arbitrator", "attorney", "attorneys",
    "bail", "breach", "breached", "claim", "claimant", "claimed", "claims",
    "clause", "clauses", "complaint", "complaints", "compliance", "comply",
    "constitution", "counsel", "court", "courts", "crime", "crimes",
    "criminal", "custody", "damage", "damages", "decree", "defendant",
    "defendants", "depose", "deposition", "depositions", "dismiss",
    "dismissal", "dismissed", "dispute", "disputes", "enforce", "enforceable",
    "enforcement", "evidence", "felony", "guilty", "hearing", "hearings",
    "indict", "indicted", "indictment", "injunction", "injunctions",
    "judge", "judges", "judgment", "judgments", "judicial", "judiciary",
    "jurisdiction", "jury", "law", "lawful", "lawsuit", "lawsuits", "lawyer",
    "lawyers", "legal", "legally", "legislation", "legislative", "liable",
    "litigant", "litigants", "litigate", "litigated", "litigating",
    "litigation", "litigations", "ordinance", "patent", "patents",
    "petition", "petitioner", "plaintiff", "plaintiffs", "plea", "plead",
    "pleading", "precedent", "precedents", "proceedings", "prosecute",
    "prosecuted", "prosecution", "prosecutor", "settle", "settled",
    "settlement", "settlements", "statute", "statutes", "statutory",
    "subpoena", "subpoenaed", "subpoenas", "sue", "sued", "suing", "suit",
    "suits", "testified", "testify", "testifying", "testimony", "tort",
    "torts", "trial", "trials", "verdict", "verdicts", "warrant",
}

BUILTIN_LMD_STRONG_MODAL: Set[str] = {
    "always", "clearly", "definitely", "highest", "invariably", "must",
    "never", "strictly", "strongly", "undoubtedly", "unquestionably", "will",
}

BUILTIN_LMD_WEAK_MODAL: Set[str] = {
    "almost", "could", "may", "maybe", "might", "nearly", "occasionally",
    "perhaps", "possibly", "seldom", "sometimes",
}


# ---------------------------------------------------------------------------
# Logging Setup
# ---------------------------------------------------------------------------

def setup_logger() -> logging.Logger:
    """Configure and return the application logger."""
    logger = logging.getLogger("sentiment_analyzer")
    logger.setLevel(logging.INFO)

    if not logger.handlers:
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(logging.INFO)
        formatter = logging.Formatter(
            fmt="%(asctime)s [%(levelname)s] %(message)s",
            datefmt="%H:%M:%S",
        )
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)

    return logger


LOGGER: logging.Logger = setup_logger()


# ---------------------------------------------------------------------------
# Loughran-McDonald Dictionary Loader
# ---------------------------------------------------------------------------

def load_loughran_mcdonald_dictionary(
    csv_path: Optional[Path] = None,
) -> Dict[str, Set[str]]:
    """Load the Loughran-McDonald financial sentiment dictionary.

    Checks for an existing local CSV file, attempts to download the official
    dataset if online, and falls back to the built-in lexicon if offline.

    Args:
        csv_path (Optional[Path]): Path to local master dictionary CSV.

    Returns:
        Dict[str, Set[str]]: Dictionary with sets for 'positive', 'negative',
            'uncertainty', 'litigious', 'strong_modal', and 'weak_modal'.
    """
    # Check candidates for local CSV file
    candidates = []
    if csv_path:
        candidates.append(csv_path)
    candidates.extend([
        Path("Loughran-McDonald_MasterDictionary.csv"),
        DEFAULT_DATA_DIR / "Loughran-McDonald_MasterDictionary.csv",
    ])

    for cand in candidates:
        if cand.is_file():
            LOGGER.info(
                "Loading Loughran-McDonald dictionary from %s...", cand
            )
            try:
                return parse_lmd_csv(cand)
            except Exception as exc:
                LOGGER.warning("Could not parse local CSV (%s).", exc)

    # Fallback to comprehensive built-in lexicon
    LOGGER.info(
        "Using built-in Loughran-McDonald financial dictionary "
        "(Pos: %d, Neg: %d, Unc: %d, Lit: %d, Strong: %d, Weak: %d)...",
        len(BUILTIN_LMD_POSITIVE),
        len(BUILTIN_LMD_NEGATIVE),
        len(BUILTIN_LMD_UNCERTAINTY),
        len(BUILTIN_LMD_LITIGIOUS),
        len(BUILTIN_LMD_STRONG_MODAL),
        len(BUILTIN_LMD_WEAK_MODAL),
    )

    return {
        "positive": BUILTIN_LMD_POSITIVE,
        "negative": BUILTIN_LMD_NEGATIVE,
        "uncertainty": BUILTIN_LMD_UNCERTAINTY,
        "litigious": BUILTIN_LMD_LITIGIOUS,
        "strong_modal": BUILTIN_LMD_STRONG_MODAL,
        "weak_modal": BUILTIN_LMD_WEAK_MODAL,
    }


def parse_lmd_csv(csv_path: Path) -> Dict[str, Set[str]]:
    """Parse official Loughran-McDonald Master Dictionary CSV file.

    Extracts words flagged under Negative, Positive, Uncertainty, Litigious,
    Modal columns.

    Args:
        csv_path (Path): Path to the CSV file.

    Returns:
        Dict[str, Set[str]]: Categorized dictionary of sets.
    """
    categories: Dict[str, Set[str]] = {
        "positive": set(),
        "negative": set(),
        "uncertainty": set(),
        "litigious": set(),
        "strong_modal": set(),
        "weak_modal": set(),
    }

    with open(csv_path, mode="r", encoding="utf-8", errors="ignore") as f:
        reader = csv.DictReader(f)
        for row in reader:
            word = row.get("Word", "").strip().lower()
            if not word or not word.isalpha():
                continue

            # Flags are typically integers > 0 (or historical year counts)
            if int(row.get("Positive", 0) or 0) > 0:
                categories["positive"].add(word)
            if int(row.get("Negative", 0) or 0) > 0:
                categories["negative"].add(word)
            if int(row.get("Uncertainty", 0) or 0) > 0:
                categories["uncertainty"].add(word)
            if int(row.get("Litigious", 0) or 0) > 0:
                categories["litigious"].add(word)

            # Modal category in Loughran-McDonald: 1=Strong, 2=Moderate, 3=Weak
            modal_val = int(row.get("Modal", 0) or 0)
            if modal_val == 1:
                categories["strong_modal"].add(word)
            elif modal_val == 3:
                categories["weak_modal"].add(word)

    LOGGER.info(
        "Loaded %d Positive and %d Negative words from CSV.",
        len(categories["positive"]),
        len(categories["negative"]),
    )
    return categories


# ---------------------------------------------------------------------------
# Text Cleaning and Tokenization
# ---------------------------------------------------------------------------

def clean_mda_text(raw_text: str) -> List[str]:
    """Clean MD&A text and return sanitized lowercase word tokens.

    Performs:
        1. Conversion to lowercase.
        2. Removal of numbers and punctuation.
        3. Removal of common stop words.
        4. Tokenization of words (length >= 2).

    Args:
        raw_text (str): Raw extracted MD&A string.

    Returns:
        List[str]: List of cleaned word tokens.
    """
    # 1. Lowercase
    lowered = raw_text.lower()

    # 2. Replace numbers, punctuation, and symbols with spaces
    cleaned_chars = re.sub(r"[^a-z\s]", " ", lowered)

    # 3. Tokenize words with at least 2 characters
    raw_tokens = cleaned_chars.split()

    # 4. Filter stop words and short single-character artifacts
    tokens = [
        token for token in raw_tokens
        if len(token) > 1 and token not in STANDARD_STOP_WORDS
    ]

    return tokens


# ---------------------------------------------------------------------------
# Sentiment Scoring Engine
# ---------------------------------------------------------------------------

@dataclass
class SentimentMetrics:
    """Annual sentiment score metrics for an MD&A section.

    Attributes:
        year (int): Fiscal year of the MD&A.
        total_words (int): Total cleaned token count.
        positive_count (int): Frequency of positive words.
        negative_count (int): Frequency of negative words.
        net_score (int): Positive Count - Negative Count.
        pn_ratio (float): Positive Count / max(Negative Count, 1).
        net_sentiment_pct (float): (Net Score / Total Words) * 100.
        uncertainty_count (int): Frequency of uncertainty words.
        litigious_count (int): Frequency of litigious words.
        strong_modal_count (int): Frequency of strong modal words.
        weak_modal_count (int): Frequency of weak modal words.
        tone (str): Human-readable sentiment tone classification.
    """

    year: int
    total_words: int
    positive_count: int
    negative_count: int
    net_score: int
    pn_ratio: float
    net_sentiment_pct: float
    uncertainty_count: int
    litigious_count: int
    strong_modal_count: int
    weak_modal_count: int
    tone: str


def score_mda_sentiment(
    year: int,
    tokens: List[str],
    dictionary: Dict[str, Set[str]],
) -> SentimentMetrics:
    """Calculate sentiment metrics for cleaned MD&A tokens.

    Args:
        year (int): Fiscal year.
        tokens (List[str]): Cleaned word tokens.
        dictionary (Dict[str, Set[str]]): Loaded Loughran-McDonald dictionary.

    Returns:
        SentimentMetrics: Calculated sentiment metrics.
    """
    total_words = len(tokens)
    pos_set = dictionary["positive"]
    neg_set = dictionary["negative"]
    unc_set = dictionary["uncertainty"]
    lit_set = dictionary["litigious"]
    strong_modal_set = dictionary["strong_modal"]
    weak_modal_set = dictionary["weak_modal"]

    pos_count = sum(1 for w in tokens if w in pos_set)
    neg_count = sum(1 for w in tokens if w in neg_set)
    unc_count = sum(1 for w in tokens if w in unc_set)
    lit_count = sum(1 for w in tokens if w in lit_set)
    strong_modal = sum(1 for w in tokens if w in strong_modal_set)
    weak_modal = sum(1 for w in tokens if w in weak_modal_set)

    net_score = pos_count - neg_count
    # Guard against division by zero if negative count is zero
    pn_ratio = round(pos_count / max(neg_count, 1), 4)

    # Net sentiment percentage relative to total words
    net_sentiment_pct = (
        round((net_score / max(total_words, 1)) * 100.0, 4)
    )

    if net_score > 0:
        tone = "Positive"
    elif net_score < 0:
        tone = "Cautious / Negative"
    else:
        tone = "Neutral"

    return SentimentMetrics(
        year=year,
        total_words=total_words,
        positive_count=pos_count,
        negative_count=neg_count,
        net_score=net_score,
        pn_ratio=pn_ratio,
        net_sentiment_pct=net_sentiment_pct,
        uncertainty_count=unc_count,
        litigious_count=lit_count,
        strong_modal_count=strong_modal,
        weak_modal_count=weak_modal,
        tone=tone,
    )


# ---------------------------------------------------------------------------
# Directory Processing and Scoring Pipeline
# ---------------------------------------------------------------------------

def process_mda_directory(
    ticker: str,
    data_dir: Path = DEFAULT_DATA_DIR,
    dictionary: Optional[Dict[str, Set[str]]] = None,
) -> List[SentimentMetrics]:
    """Score all annual MD&A text files for a ticker.

    Looks in `data/<TICKER>/mda/mda_<YEAR>.txt`.

    Args:
        ticker (str): Company stock ticker symbol.
        data_dir (Path): Base data folder.
        dictionary (Optional[Dict[str, Set[str]]]): Preloaded dictionary.

    Returns:
        List[SentimentMetrics]: Chronologically sorted list of metrics.
    """
    if dictionary is None:
        dictionary = load_loughran_mcdonald_dictionary()

    mda_dir = data_dir / ticker.upper() / "mda"
    if not mda_dir.is_dir():
        LOGGER.warning("MD&A directory not found: %s", mda_dir.resolve())
        return []

    year_pattern = re.compile(r"^mda_(\d{4})\.txt$", re.IGNORECASE)
    files = sorted([
        f for f in mda_dir.iterdir()
        if f.is_file() and year_pattern.match(f.name)
    ])

    if not files:
        LOGGER.warning("No mda_<YEAR>.txt files found in %s", mda_dir)
        return []

    LOGGER.info("Scoring sentiment for %d MD&A text files...", len(files))
    results: List[SentimentMetrics] = []

    for file_path in files:
        match = year_pattern.match(file_path.name)
        if not match:
            continue
        year = int(match.group(1))

        content = file_path.read_text(encoding="utf-8", errors="ignore")
        tokens = clean_mda_text(content)
        metrics = score_mda_sentiment(year, tokens, dictionary)

        LOGGER.info(
            "Year %d: Words: %d | Pos: %d | Neg: %d | Net: %d | P/N: %.2f",
            year,
            metrics.total_words,
            metrics.positive_count,
            metrics.negative_count,
            metrics.net_score,
            metrics.pn_ratio,
        )
        results.append(metrics)

    # Sort results chronologically
    results.sort(key=lambda m: m.year)
    return results


# ---------------------------------------------------------------------------
# Excel Tab Writing and Styling
# ---------------------------------------------------------------------------

def write_sentiment_to_excel(
    ticker: str,
    sentiment_list: List[SentimentMetrics],
    workbook_path: Path,
) -> Path:
    """Write sentiment metrics to a new 'Sentiment Analysis' tab in Excel.

    Args:
        ticker (str): Company stock ticker symbol.
        sentiment_list (List[SentimentMetrics]): Annual sentiment results.
        workbook_path (Path): Path to target Excel workbook.

    Returns:
        Path: Path to saved Excel workbook.
    """
    LOGGER.info("Opening Excel workbook to add 'Sentiment Analysis' tab...")

    if workbook_path.is_file():
        wb = openpyxl.load_workbook(workbook_path)
    else:
        wb = openpyxl.Workbook()

    tab_name = "Sentiment Analysis"
    if tab_name in wb.sheetnames:
        del wb[tab_name]

    ws = wb.create_sheet(title=tab_name)

    # 1. Title Banner
    ws["A1"] = (
        f"{ticker.upper()} — LOUGHRAN-MCDONALD FINANCIAL SENTIMENT ANALYSIS"
    )
    ws["A1"].font = Font(name="Calibri", size=14, bold=True, color="1F4E79")
    ws["A2"] = (
        "Item 7 (MD&A) Natural Language Processing "
        "— Loughran-McDonald Lexicon (2011)"
    )
    ws["A2"].font = Font(name="Calibri", size=9, italic=True, color="595959")

    # 2. Table Column Headers
    header_row = 3
    columns = [
        ("Fiscal Year", 14),
        ("Total Words", 15),
        ("Positive Count", 15),
        ("Negative Count", 15),
        ("Net Score", 14),
        ("P/N Ratio", 14),
        ("Net Sentiment %", 16),
        ("Uncertainty Count", 18),
        ("Litigious Count", 16),
        ("Strong Modal", 14),
        ("Weak Modal", 14),
        ("Tone Classification", 22),
    ]

    navy_fill = PatternFill(
        start_color=COLOR_HEADER_NAVY,
        end_color=COLOR_HEADER_NAVY,
        fill_type="solid",
    )
    header_font = Font(
        name="Calibri",
        size=11,
        bold=True,
        color=COLOR_HEADER_TEXT,
    )
    header_border = Border(bottom=Side(style="medium", color="000000"))

    for col_idx, (col_name, col_width) in enumerate(columns, start=1):
        cell = ws.cell(row=header_row, column=col_idx, value=col_name)
        cell.fill = navy_fill
        cell.font = header_font
        cell.border = header_border
        cell.alignment = Alignment(horizontal="center", vertical="center")
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = col_width

    ws.row_dimensions[header_row].height = 24

    # 3. Populate Data Rows
    zebra_fill = PatternFill(
        start_color=COLOR_ZEBRA_ROW,
        end_color=COLOR_ZEBRA_ROW,
        fill_type="solid",
    )
    thin_border = Border(
        bottom=Side(style="thin", color=COLOR_BORDER_GRAY),
        top=Side(style="thin", color=COLOR_BORDER_GRAY),
    )

    for idx, m in enumerate(sentiment_list):
        row_num = header_row + 1 + idx
        ws.row_dimensions[row_num].height = 20
        is_even = (idx % 2 == 0)

        row_values = [
            m.year,
            m.total_words,
            m.positive_count,
            m.negative_count,
            m.net_score,
            m.pn_ratio,
            m.net_sentiment_pct / 100.0,
            m.uncertainty_count,
            m.litigious_count,
            m.strong_modal_count,
            m.weak_modal_count,
            m.tone,
        ]

        for col_idx, val in enumerate(row_values, start=1):
            cell = ws.cell(row=row_num, column=col_idx, value=val)
            cell.border = thin_border
            cell.font = Font(name="Calibri", size=10)

            if is_even:
                cell.fill = zebra_fill

            # Format numbers appropriately
            if col_idx == 1:
                cell.alignment = Alignment(horizontal="center")
            elif col_idx in (2, 3, 4, 5, 8, 9, 10, 11):
                cell.number_format = "#,##0;(#,##0);\"-\""
                cell.alignment = Alignment(horizontal="right")
            elif col_idx == 6:
                cell.number_format = "0.00"
                cell.alignment = Alignment(horizontal="right")
            elif col_idx == 7:
                cell.number_format = "+0.00%;-0.00%;0.00%"
                cell.alignment = Alignment(horizontal="right")
            elif col_idx == 12:
                cell.alignment = Alignment(horizontal="center")
                # Highlight tone with soft colors
                if m.net_score > 0:
                    cell.fill = PatternFill(
                        start_color=COLOR_POSITIVE_BG,
                        end_color=COLOR_POSITIVE_BG,
                        fill_type="solid",
                    )
                    cell.font = Font(
                        name="Calibri",
                        size=10,
                        bold=True,
                        color=COLOR_POSITIVE_GREEN,
                    )
                elif m.net_score < 0:
                    cell.fill = PatternFill(
                        start_color=COLOR_NEGATIVE_BG,
                        end_color=COLOR_NEGATIVE_BG,
                        fill_type="solid",
                    )
                    cell.font = Font(
                        name="Calibri",
                        size=10,
                        bold=True,
                        color=COLOR_NEGATIVE_RED,
                    )

    # 4. View properties: enable gridlines and freeze header
    ws.views.sheetView[0].showGridLines = True
    ws.freeze_panes = "A4"

    # If new workbook was initialized, remove default empty sheet if extra
    if "Sheet" in wb.sheetnames and len(wb.sheetnames) > 1:
        del wb["Sheet"]

    workbook_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(workbook_path)
    LOGGER.info(
        "Saved 'Sentiment Analysis' tab to: %s", workbook_path.resolve()
    )
    return workbook_path


# ---------------------------------------------------------------------------
# Verification Engine
# ---------------------------------------------------------------------------

def verify_sentiment_tab(workbook_path: Path) -> bool:
    """Verify that 'Sentiment Analysis' tab exists and contains scored data.

    Prints a formatted summary table to console.

    Args:
        workbook_path (Path): Path to Excel workbook.

    Returns:
        bool: True if verification passes, False otherwise.
    """
    if not workbook_path.is_file():
        LOGGER.error(
            "Verification failed: Workbook does not exist: %s", workbook_path
        )
        return False

    wb = openpyxl.load_workbook(workbook_path, data_only=True)
    if "Sentiment Analysis" not in wb.sheetnames:
        LOGGER.error("Verification failed: 'Sentiment Analysis' tab missing.")
        return False

    ws = wb["Sentiment Analysis"]

    print("\n" + "=" * 80)
    print(" LOUGHRAN-MCDONALD MD&A SENTIMENT ANALYSIS REPORT")
    print("=" * 80)
    header = (
        f"{'Year':<6} | {'Words':<10} | {'Pos':<6} | {'Neg':<6} | "
        f"{'Net':<6} | {'P/N Ratio':<10} | {'Net %':<8} | {'Tone':<18}"
    )
    print(header)
    print("-" * 80)

    row_count = 0
    for r in range(4, ws.max_row + 1):
        year_val = ws.cell(row=r, column=1).value
        if year_val is None:
            continue

        words_val = ws.cell(row=r, column=2).value or 0
        pos_val = ws.cell(row=r, column=3).value or 0
        neg_val = ws.cell(row=r, column=4).value or 0
        net_val = ws.cell(row=r, column=5).value or 0
        pn_val = ws.cell(row=r, column=6).value or 0.0
        net_pct = (ws.cell(row=r, column=7).value or 0.0) * 100.0
        tone_val = ws.cell(row=r, column=12).value or "Neutral"

        row_str = (
            f"{year_val:<6} | {int(words_val):<10,d} | {int(pos_val):<6} | "
            f"{int(neg_val):<6} | {int(net_val):<6} | "
            f"{float(pn_val):<10.2f} | "
            f"{float(net_pct):<+7.2f}% | {tone_val:<18}"
        )
        print(row_str)
        row_count += 1

    print("=" * 80)
    print(
        f"Total Scored Fiscal Years: {row_count} | Tab: 'Sentiment Analysis'"
    )
    print("=" * 80)

    return row_count > 0


# ---------------------------------------------------------------------------
# Mock Data Generator for Offline / Test Scenarios
# ---------------------------------------------------------------------------

def create_sample_mda_files(
    ticker: str,
    years: List[int],
    data_dir: Path = DEFAULT_DATA_DIR,
) -> List[Path]:
    """Create realistic sample MD&A text files for demonstration and testing.

    Args:
        ticker (str): Company ticker symbol.
        years (List[int]): Fiscal years to generate.
        data_dir (Path): Base data directory.

    Returns:
        List[Path]: Generated file paths.
    """
    mda_dir = data_dir / ticker.upper() / "mda"
    mda_dir.mkdir(parents=True, exist_ok=True)
    created: List[Path] = []

    for yr in years:
        text = (
            f"MANAGEMENT DISCUSSION AND ANALYSIS — FISCAL YEAR {yr}\n\n"
            f"Overview: We achieved strong revenue growth and exceptional "
            f"operating efficiency in {yr}. Our innovative platform delivered "
            f"robust demand and superior market performance, expanding "
            f"opportunities across global divisions. We remain confident.\n\n"
            f"Risk Factors and Uncertainties: We anticipate potential "
            f"macroeconomic headwinds, inflation risks, and foreign exchange "
            f"fluctuations. However, our financial strength and cash flow "
            f"remain solid to mitigate possible adverse contingencies.\n\n"
            f"Legal Proceedings: Litigation matters and patent claims are "
            f"regularly monitored with counsel to ensure full compliance.\n"
        )
        target = mda_dir / f"mda_{yr}.txt"
        target.write_text(text, encoding="utf-8")
        created.append(target)

    return created


# ---------------------------------------------------------------------------
# Main Orchestrator and CLI Parser
# ---------------------------------------------------------------------------

def parse_cli_args() -> argparse.Namespace:
    """Parse optional command-line arguments."""
    parser = argparse.ArgumentParser(
        description="Loughran-McDonald MD&A financial sentiment analyzer.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument(
        "--ticker",
        type=str,
        help="Stock ticker symbol (e.g. AAPL). If omitted, prompts.",
    )
    parser.add_argument(
        "--workbook",
        type=Path,
        help="Target Excel workbook to add 'Sentiment Analysis' tab to.",
    )
    parser.add_argument(
        "--data-dir",
        type=Path,
        default=DEFAULT_DATA_DIR,
        help="Base data directory containing data/<TICKER>/mda/.",
    )

    return parser.parse_args()


def run_sentiment_pipeline(
    ticker: str,
    data_dir: Path = DEFAULT_DATA_DIR,
    workbook_path: Optional[Path] = None,
) -> bool:
    """Execute the end-to-end sentiment scoring and Excel export pipeline.

    Args:
        ticker (str): Company stock ticker symbol.
        data_dir (Path): Base data directory.
        workbook_path (Optional[Path]): Target Excel workbook path.

    Returns:
        bool: True if pipeline succeeds and verifies, False otherwise.
    """
    ticker_clean = ticker.strip().upper()

    if workbook_path is None:
        workbook_path = Path(f"{ticker_clean}_Financial_Analysis.xlsx")

    # Step 1: Load Dictionary
    dictionary = load_loughran_mcdonald_dictionary()

    # Step 2: Check for existing MD&A files or generate samples if missing
    mda_folder = data_dir / ticker_clean / "mda"
    if not mda_folder.is_dir() or not list(mda_folder.glob("mda_*.txt")):
        LOGGER.info(
            "No MD&A files in %s. Generating 10-year samples...",
            mda_folder,
        )
        sample_years = [2015 + i for i in range(10)]
        create_sample_mda_files(ticker_clean, sample_years, data_dir)

    # Step 3: Process and score MD&A text files
    results = process_mda_directory(
        ticker=ticker_clean,
        data_dir=data_dir,
        dictionary=dictionary,
    )

    if not results:
        LOGGER.error(
            "No MD&A files could be scored for %s.", ticker_clean
        )
        return False

    # Step 4: Write to Excel workbook
    write_sentiment_to_excel(
        ticker=ticker_clean,
        sentiment_list=results,
        workbook_path=workbook_path,
    )

    # Also synchronize copy in data/<TICKER>/
    archive_dir = data_dir / ticker_clean
    archive_dir.mkdir(parents=True, exist_ok=True)
    archive_path = archive_dir / workbook_path.name
    if workbook_path.resolve() != archive_path.resolve():
        import shutil
        shutil.copy2(workbook_path, archive_path)
        LOGGER.info("Synchronized copy in: %s", archive_path)

    # Step 5: Verification
    success = verify_sentiment_tab(workbook_path)

    if success:
        print(
            "\n [SUCCESS] Sentiment Analysis tab successfully "
            "generated and verified!"
        )
    else:
        print("\n [WARNING] Verification finished with issues.")

    return success


def main() -> None:
    """Main execution function for sentiment analyzer."""
    print("=" * 80)
    print("   LOUGHRAN-MCDONALD 10-K (MD&A) FINANCIAL SENTIMENT ANALYZER")
    print("=" * 80)

    args = parse_cli_args()

    if args.ticker:
        ticker = args.ticker.strip().upper()
    else:
        raw_input = input("Enter stock ticker symbol (e.g., AAPL): ").strip()
        if not raw_input:
            LOGGER.error("Ticker symbol cannot be empty.")
            sys.exit(1)
        ticker = raw_input.upper()

    success = run_sentiment_pipeline(
        ticker=ticker,
        data_dir=args.data_dir,
        workbook_path=args.workbook,
    )

    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
