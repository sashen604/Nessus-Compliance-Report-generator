#!/usr/bin/env bash
set -euo pipefail

# Usage: ./generate_compliance_docs.sh /path/to/nessus/files /path/to/output
# Default input: current directory
# Default output: ./output-docs

INPUT_DIR="${1:-$(pwd)}"
OUTPUT_DIR="${2:-${INPUT_DIR}/output-docs}"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PYTHON_BIN="${SCRIPT_DIR}/.venv/bin/python"
SCRIPT="${SCRIPT_DIR}/compliance_table_docx.py"

if [[ ! -x "${PYTHON_BIN}" ]]; then
  echo "Python venv not found at ${PYTHON_BIN}. Create it first:" >&2
  echo "  python3 -m venv \"${SCRIPT_DIR}/.venv\"" >&2
  echo "  \"${SCRIPT_DIR}/.venv/bin/pip\" install -r \"${SCRIPT_DIR}/requirements.txt\"" >&2
  exit 1
fi

mkdir -p "${OUTPUT_DIR}"

shopt -s nullglob
for nessus_file in "${INPUT_DIR}"/*.nessus; do
  base_name="$(basename "${nessus_file}")"
  # Extract IP-like token from filename (e.g., 192_168_6_22) and convert to dots.
  ip_token="$(echo "${base_name}" | grep -oE '[0-9]+(_[0-9]+){3}' | head -n1 || true)"
  if [[ -n "${ip_token}" ]]; then
    ip_address="${ip_token//_/.}"
    output_file="${OUTPUT_DIR}/${ip_address}.docx"
  else
    output_file="${OUTPUT_DIR}/${base_name%.nessus}.docx"
  fi

  echo "Generating ${output_file}"
  "${PYTHON_BIN}" "${SCRIPT}" "${nessus_file}" --output "${output_file}"
done

echo "Done. Documents saved in ${OUTPUT_DIR}"
