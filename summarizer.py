"""Summary builders for Nessus findings."""
from __future__ import annotations

from collections import Counter, defaultdict
from typing import Dict, Iterable, List, Optional, Tuple

SEVERITY_LABELS = {
    4: "Critical",
    3: "High",
    2: "Medium",
    1: "Low",
    0: "Info",
}

FAILED_RESULTS = {"FAILED", "ERROR", "WARNING", "FAIL"}


def severity_label(value: Optional[int]) -> str:
    if value is None:
        return "Unknown"
    return SEVERITY_LABELS.get(value, "Unknown")


def overall_summary(findings: Iterable[Dict]) -> Dict[str, int]:
    findings_list = list(findings)
    hosts = {finding.get("host") for finding in findings_list if finding.get("host")}
    severity_counts = Counter(severity_label(f.get("severity")) for f in findings_list)

    summary = {
        "Total Hosts": len(hosts),
        "Total Findings": len(findings_list),
        "Critical": severity_counts.get("Critical", 0),
        "High": severity_counts.get("High", 0),
        "Medium": severity_counts.get("Medium", 0),
        "Low": severity_counts.get("Low", 0),
        "Info": severity_counts.get("Info", 0),
    }

    return summary


def host_summary(findings: Iterable[Dict]) -> List[Dict[str, int]]:
    host_counts: Dict[str, Counter] = defaultdict(Counter)
    total_by_host: Counter = Counter()

    for finding in findings:
        host = finding.get("host") or "Unknown"
        label = severity_label(finding.get("severity"))
        host_counts[host][label] += 1
        total_by_host[host] += 1

    rows = []
    for host, counts in sorted(host_counts.items()):
        rows.append({
            "Host": host,
            "Critical": counts.get("Critical", 0),
            "High": counts.get("High", 0),
            "Medium": counts.get("Medium", 0),
            "Low": counts.get("Low", 0),
            "Info": counts.get("Info", 0),
            "Total": total_by_host.get(host, 0),
        })

    return rows


def severity_summary(findings: Iterable[Dict]) -> List[Dict[str, int]]:
    counts = Counter(severity_label(f.get("severity")) for f in findings)
    return [
        {"Severity": "Critical", "Count": counts.get("Critical", 0)},
        {"Severity": "High", "Count": counts.get("High", 0)},
        {"Severity": "Medium", "Count": counts.get("Medium", 0)},
        {"Severity": "Low", "Count": counts.get("Low", 0)},
        {"Severity": "Info", "Count": counts.get("Info", 0)},
    ]


def plugin_summary(findings: Iterable[Dict], top: int = 20) -> List[Dict[str, str]]:
    counts: Dict[Tuple[str, str, str], int] = Counter()

    for finding in findings:
        plugin_id = finding.get("plugin_id") or "Unknown"
        plugin_name = finding.get("plugin_name") or "Unknown"
        severity = severity_label(finding.get("severity"))
        counts[(plugin_id, plugin_name, severity)] += 1

    rows = []
    for (plugin_id, plugin_name, severity), count in counts.most_common(top):
        rows.append({
            "Plugin ID": plugin_id,
            "Plugin Name": plugin_name,
            "Severity": severity,
            "Count": count,
        })

    return rows


def compliance_summary(findings: Iterable[Dict]) -> List[Dict[str, int]]:
    summaries: Dict[str, Counter] = defaultdict(Counter)

    for finding in findings:
        if not finding.get("compliance"):
            continue
        benchmark = finding.get("compliance_benchmark_name") or "Unknown"
        result = (finding.get("compliance_result") or "Unknown").upper()
        summaries[benchmark][result] += 1

    rows = []
    for benchmark, counts in sorted(summaries.items()):
        total = sum(counts.values())
        rows.append({
            "Benchmark": benchmark,
            "Passed": counts.get("PASSED", 0),
            "Failed": counts.get("FAILED", 0),
            "Warning": counts.get("WARNING", 0),
            "Error": counts.get("ERROR", 0),
            "Other": total - (
                counts.get("PASSED", 0)
                + counts.get("FAILED", 0)
                + counts.get("WARNING", 0)
                + counts.get("ERROR", 0)
            ),
            "Total": total,
        })

    return rows


def failed_compliance_checks(findings: Iterable[Dict]) -> List[Dict[str, str]]:
    rows = []
    for finding in findings:
        if not finding.get("compliance"):
            continue
        result = (finding.get("compliance_result") or "Unknown").upper()
        if result not in FAILED_RESULTS:
            continue

        rows.append({
            "Host": finding.get("host") or "Unknown",
            "Check Name": finding.get("compliance_check_name") or "Unknown",
            "Result": result,
            "Policy Value": finding.get("policy_value") or "",
            "Actual Value": finding.get("actual_value") or "",
        })

    return rows


def filter_findings(
    findings: Iterable[Dict],
    host: Optional[str] = None,
    severity: Optional[str] = None,
    compliance_only: bool = False,
    vuln_only: bool = False,
) -> List[Dict]:
    filtered: List[Dict] = []
    severity_normalized = severity.lower() if severity else None

    for finding in findings:
        if host and (finding.get("host") or "").lower() != host.lower():
            continue

        if severity_normalized:
            label = severity_label(finding.get("severity")).lower()
            if label != severity_normalized and str(finding.get("severity")) != severity_normalized:
                continue

        if compliance_only and not finding.get("compliance"):
            continue
        if vuln_only and finding.get("compliance"):
            continue

        filtered.append(finding)

    return filtered
