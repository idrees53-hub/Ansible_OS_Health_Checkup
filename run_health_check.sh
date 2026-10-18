#!/bin/bash

set -euo pipefail

PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

echo "============================================================"
echo "        STC RHEL OS HEALTH CHECK AUTOMATION"
echo "============================================================"
echo
echo "Project Directory : $PROJECT_DIR"
echo "Started At        : $(date)"
echo

cd "$PROJECT_DIR"

# ------------------------------------------------------------
# Step 1: Run Ansible OS Health Check
# ------------------------------------------------------------

echo "------------------------------------------------------------"
echo "STEP 1: Running Ansible OS Health Check"
echo "------------------------------------------------------------"

ansible-playbook playbooks/os_health_check.yml

echo
echo "Ansible health check completed successfully."
echo

# ------------------------------------------------------------
# Step 2: Generate Individual Server DOCX Reports
# ------------------------------------------------------------

echo "------------------------------------------------------------"
echo "STEP 2: Generating Individual Server Reports"
echo "------------------------------------------------------------"

python3 report_generator/generate_server_reports.py

echo
echo "Individual server reports generated successfully."
echo

# ------------------------------------------------------------
# Step 3: Generate Master DOCX Report
# ------------------------------------------------------------

echo "------------------------------------------------------------"
echo "STEP 3: Generating Master DOCX Report"
echo "------------------------------------------------------------"

python3 report_generator/generate_master_report.py

echo
echo "Master DOCX report generated successfully."
echo

# ------------------------------------------------------------
# Step 4: Generate Master XLSX Report
# ------------------------------------------------------------

echo "------------------------------------------------------------"
echo "STEP 4: Generating Master XLSX Report"
echo "------------------------------------------------------------"

python3 report_generator/generate_master_excel.py

echo
echo "Master XLSX report generated successfully."
echo

# ------------------------------------------------------------
# Final Summary
# ------------------------------------------------------------

echo "============================================================"
echo "             HEALTH CHECK COMPLETED"
echo "============================================================"
echo
echo "Reports are available under:"
echo
echo "  $PROJECT_DIR/reports/2026-Q4/"
echo
echo "Generated reports:"
echo
echo "  - Individual server JSON reports"
echo "  - Individual server DOCX reports"
echo "  - Master DOCX report"
echo "  - Master XLSX report"
echo
echo "Completed At : $(date)"
echo
echo "============================================================"
