"""Nessus XML parser for vulnerability and compliance findings."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional
import xml.etree.ElementTree as ET


@dataclass
class Finding:
    host: str
    plugin_id: str
    plugin_name: str
    severity: Optional[int]
    risk_factor: Optional[str]
    port: Optional[str]
    svc_name: Optional[str]
    protocol: Optional[str]
    description: Optional[str]
    solution: Optional[str]
    synopsis: Optional[str]
    family: Optional[str]
    compliance: bool
    compliance_check_name: Optional[str]
    compliance_result: Optional[str]
    compliance_info: Optional[str]
    compliance_solution: Optional[str]
    compliance_benchmark_name: Optional[str]
    compliance_audit_file: Optional[str]
    compliance_profile: Optional[str]
    policy_value: Optional[str]
    actual_value: Optional[str]


def _strip_text(value: Optional[str]) -> Optional[str]:
    if value is None:
        return None
    value = value.strip()
    return value if value else None


def _find_text(element: ET.Element, tag: str) -> Optional[str]:
    child = element.find(tag)
    if child is not None and child.text is not None:
        return _strip_text(child.text)

    # Namespace-insensitive lookup (e.g., cm:compliance-*)
    for child in element:
        if child.tag.endswith(tag):
            if child.text is not None:
                return _strip_text(child.text)

    return None


def _parse_severity(value: Optional[str]) -> Optional[int]:
    if value is None:
        return None
    try:
        return int(value)
    except ValueError:
        return None


def parse_nessus(path: str) -> List[Dict[str, Optional[str]]]:
    tree = ET.parse(path)
    root = tree.getroot()

    findings: List[Dict[str, Optional[str]]] = []

    for report_host in root.iter("ReportHost"):
        host = report_host.attrib.get("name") or "Unknown"

        for report_item in report_host.iter("ReportItem"):
            plugin_id = report_item.attrib.get("pluginID")
            plugin_name = report_item.attrib.get("pluginName")
            family = report_item.attrib.get("pluginFamily")
            severity_value = report_item.attrib.get("severity")

            compliance_text = _find_text(report_item, "compliance")
            compliance_result = _find_text(report_item, "compliance-result")

            compliance_flag = False
            if compliance_text:
                compliance_flag = compliance_text.lower() == "true"
            elif compliance_result:
                compliance_flag = True

            finding = {
                "host": host,
                "plugin_id": plugin_id,
                "plugin_name": plugin_name,
                "severity": _parse_severity(severity_value),
                "risk_factor": _find_text(report_item, "risk_factor"),
                "port": report_item.attrib.get("port"),
                "svc_name": report_item.attrib.get("svc_name"),
                "protocol": report_item.attrib.get("protocol"),
                "description": _find_text(report_item, "description"),
                "solution": _find_text(report_item, "solution"),
                "synopsis": _find_text(report_item, "synopsis"),
                "family": family,
                "compliance": compliance_flag,
                "compliance_check_name": _find_text(report_item, "compliance-check-name"),
                "compliance_result": compliance_result,
                "compliance_info": _find_text(report_item, "compliance-info"),
                "compliance_solution": _find_text(report_item, "compliance-solution"),
                "compliance_benchmark_name": _find_text(report_item, "compliance-benchmark-name"),
                "compliance_audit_file": _find_text(report_item, "compliance-audit-file"),
                "compliance_profile": _find_text(report_item, "compliance-profile"),
                "policy_value": _find_text(report_item, "compliance-policy-value"),
                "actual_value": _find_text(report_item, "compliance-actual-value"),
            }

            findings.append(finding)

    return findings
