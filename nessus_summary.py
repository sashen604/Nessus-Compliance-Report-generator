"""CLI entry point for Nessus summaries."""
from __future__ import annotations

import argparse
from typing import List

from parser import parse_nessus
from summarizer import (
    compliance_summary,
    failed_compliance_checks,
    filter_findings,
    host_summary,
    overall_summary,
    plugin_summary,
    severity_summary,
)
from output import export_csv, export_json, print_table


DEFAULT_SUMMARIES = ["overall", "host", "compliance", "failed"]


def _summary_overall(findings: List[dict]) -> None:
    summary = overall_summary(findings)
    rows = [{"Metric": key, "Count": value} for key, value in summary.items()]
    print_table("Overall Summary", rows, headers=["Metric", "Count"])


def _summary_host(findings: List[dict]) -> None:
    rows = host_summary(findings)
    print_table("Findings by Host", rows, headers=["Host", "Critical", "High", "Medium", "Low", "Info", "Total"])


def _summary_severity(findings: List[dict]) -> None:
    rows = severity_summary(findings)
    print_table("Severity Summary", rows, headers=["Severity", "Count"])


def _summary_plugin(findings: List[dict]) -> None:
    rows = plugin_summary(findings)
    print_table("Top Plugins / Repeated Findings", rows, headers=["Plugin ID", "Plugin Name", "Severity", "Count"])


def _summary_compliance(findings: List[dict]) -> None:
    rows = compliance_summary(findings)
    print_table("Compliance Results Summary", rows, headers=["Benchmark", "Passed", "Failed", "Warning", "Error", "Other", "Total"])


def _summary_failed(findings: List[dict]) -> None:
    rows = failed_compliance_checks(findings)
    print_table(
        "Failed Compliance Checks",
        rows,
        headers=["Host", "Check Name", "Result", "Policy Value", "Actual Value"],
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Summarize Nessus XML (.nessus/.xml) results.")
    parser.add_argument("scan_path", help="Path to Nessus .nessus/.xml file")
    parser.add_argument(
        "--summary",
        action="append",
        choices=["overall", "host", "severity", "plugin", "compliance", "failed"],
        help="Select summary table(s) to show. Repeat to show multiple.",
    )
    parser.add_argument("--failed-only", action="store_true", help="Only include failed compliance checks")
    parser.add_argument("--host", help="Filter results by host name")
    parser.add_argument("--severity", help="Filter results by severity label or numeric value")
    parser.add_argument("--compliance-only", action="store_true", help="Only include compliance findings")
    parser.add_argument("--vuln-only", action="store_true", help="Only include vulnerability findings")
    parser.add_argument("--export", dest="export_csv", help="Export findings to CSV")
    parser.add_argument("--export-json", dest="export_json", help="Export findings to JSON")
    return parser


def main() -> None:
    args = build_parser().parse_args()

    findings = parse_nessus(args.scan_path)

    findings = filter_findings(
        findings,
        host=args.host,
        severity=args.severity,
        compliance_only=args.compliance_only,
        vuln_only=args.vuln_only,
    )

    if args.failed_only:
        findings = [
            finding for finding in findings
            if finding.get("compliance")
            and (finding.get("compliance_result") or "").upper() in {"FAILED", "ERROR", "WARNING", "FAIL"}
        ]

    if args.export_csv:
        export_csv(args.export_csv, findings)

    if args.export_json:
        export_json(args.export_json, findings)

    summaries = args.summary or DEFAULT_SUMMARIES

    for summary in summaries:
        if summary == "overall":
            _summary_overall(findings)
        elif summary == "host":
            _summary_host(findings)
        elif summary == "severity":
            _summary_severity(findings)
        elif summary == "plugin":
            _summary_plugin(findings)
        elif summary == "compliance":
            _summary_compliance(findings)
        elif summary == "failed":
            _summary_failed(findings)


if __name__ == "__main__":
    main()
