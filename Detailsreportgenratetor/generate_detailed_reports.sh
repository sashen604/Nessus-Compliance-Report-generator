#!/usr/bin/env bash
set -euo pipefail

# Usage: ./generate_detailed_reports.sh /path/to/nessus/files /path/to/output
# Default input: current directory
# Default output: ./detailed-docs

INPUT_DIR="${1:-$(pwd)}"
OUTPUT_DIR="${2:-${INPUT_DIR}/detailed-docs}"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT_DIR="$(cd "${SCRIPT_DIR}/.." && pwd)"
PYTHON_BIN="${ROOT_DIR}/.venv/bin/python"
SCRIPT="${SCRIPT_DIR}/details_report_docx.py"

if [[ ! -x "${PYTHON_BIN}" ]]; then
  echo "Python venv not found at ${PYTHON_BIN}. Create it first:" >&2
  echo "  python3 -m venv \"${ROOT_DIR}/.venv\"" >&2
  echo "  \"${ROOT_DIR}/.venv/bin/pip\" install -r \"${ROOT_DIR}/requirements.txt\"" >&2
  exit 1
fi

mkdir -p "${OUTPUT_DIR}"

shopt -s nullglob
for nessus_file in "${INPUT_DIR}"/*.nessus; do
  base_name="$(basename "${nessus_file}")"
  ip_token="$(echo "${base_name}" | grep -oE '[0-9]+(_[0-9]+){3}' | head -n1 || true)"
  if [[ -n "${ip_token}" ]]; then
    ip_address="${ip_token//_/.}"
    output_file="${OUTPUT_DIR}/${ip_address}_detailed.docx"
  else
    output_file="${OUTPUT_DIR}/${base_name%.nessus}_detailed.docx"
  fi

  echo "Generating ${output_file}"
  "${PYTHON_BIN}" "${SCRIPT}" "${nessus_file}" --output "${output_file}"
done

echo "Done. Detailed reports saved in ${OUTPUT_DIR}"
