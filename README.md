# RHEL OS Health Check Automation

## 1. Project Overview

The **RHEL OS Health Check Automation** project uses Ansible to automate routine operating system health checks on Red Hat Enterprise Linux (RHEL) servers.

The purpose of this project is to reduce manual work, collect important system information, identify potential issues, and generate reports with health assessments and recommendations.

The project is designed for RHEL 8 and RHEL 9 servers.

## 2. Key Features

The automation performs the following checks:

- **OS Information:** Collects operating system and version details.
- **CPU Health:** Checks CPU information, utilization, and load.
- **Memory Health:** Collects memory and swap usage information.
- **Filesystem Health:** Checks filesystem capacity and usage.
- **Kernel Health:** Collects kernel messages and identifies potential warnings.
- **System Logs:** Reviews recent system logs to identify errors and warnings.
- **Service Health:** Checks system services and their status.
- **Network Health:** Collects network-related information for assessment.
- **NTP Health:** Checks time synchronization information.
- **Security Assessment:** Collects selected security-related information.
- **Uptime and Performance:** Checks uptime and available performance statistics.

The collected information is assessed to help identify potential issues and provide recommendations for further investigation.

## 3. Project Structure

```text
stc-os-hc/
├── ansible.cfg
├── apply_stc_updates.py
├── inventory/
│   ├── hosts
│   └── group_vars/
│       └── all.yml
├── playbooks/
│   └── os_health_check.yml
├── report_generator/
│   ├── generate_master_excel.py
│   ├── generate_master_report.py
│   ├── generate_server_report.py
│   └── generate_server_reports.py
├── roles/
│   └── os_health/
│       ├── defaults/
│       ├── tasks/
│       └── templates/
├── run_health_check.sh
├── scripts/
├── reports/
└── README.md
```

## 4. Requirements

Before running the health check, ensure that you have:

- Ansible installed on the control node.
- RHEL 8 or RHEL 9 managed servers.
- SSH connectivity from the control node to the managed servers.
- Appropriate permissions to execute the required commands on managed servers.
- Python installed for the report-generation scripts.
- Required Python libraries for generating Word and Excel reports.

## 5. Configuration

### Inventory

Add the managed servers to:

`inventory/hosts`

Example:

```ini
[rhel_servers]
rhel9-node1
rhel9-node2
rhel9-node3
```

Replace the example hostnames with the actual inventory entries used in your environment.

### Health Check Variables

The file `inventory/group_vars/all.yml` contains shared configuration variables used by the health-check tasks.

Review this file before execution and adjust the applicable thresholds and other settings according to the environment.

## 6. Verify Connectivity

From the project directory, test connectivity to the managed servers:

```bash
ansible all -m ping
```

Ensure that the managed servers are reachable before running the complete health check.

## 7. Run the Health Check

Navigate to the project directory:

```bash
cd ~/stc-os-hc
```

Validate the playbook syntax:

```bash
ansible-playbook -i inventory/hosts \
  playbooks/os_health_check.yml --syntax-check
```

Run the complete health check:

```bash
ansible-playbook -i inventory/hosts \
  playbooks/os_health_check.yml
```

Alternatively, use the project wrapper script if it is configured for your environment:

```bash
./run_health_check.sh
```

## 8. Reports and Output

The automation collects system information and generates structured output for analysis.

The project can generate:

- **JSON files:** Contain collected health-check data for individual servers.
- **Individual Word reports:** Present the health assessment for each server.
- **Master Word report:** Provides a consolidated health-check summary.
- **Master Excel report:** Presents consolidated results in spreadsheet format.

Generated reports are stored under the configured reporting directory, for example:

```text
reports/
└── 2026-Q4/
    ├── data/
    ├── servers/
    ├── <NAME>_OS_Health_Summary.docx
    └── <NAME>_OS_Health_Summary.xlsx
```

The reporting period and output location depend on the project configuration.

## 9. Understanding the Results

The assessment results help identify areas that may require attention.

- **OK:** No issue was identified by the relevant check.
- **WARNING:** A potential issue requires review.
- **CRITICAL:** A serious finding requires investigation and appropriate action.

A warning does not always indicate a production incident. Review the evidence and recommendations before taking corrective action.

## 10. Important Notes

- Review the inventory and configuration before running the automation.
- Ensure SSH connectivity and appropriate permissions.
- Review generated reports before sharing them.
- Do not commit passwords, private keys, tokens, or confidential production data to Git.
- Keep generated reports and collected server data out of the Git repository unless approved for version control.
- The health-check automation is intended for assessment and reporting. Do not assume that it automatically fixes identified issues.
- Review `apply_stc_updates.py` separately before using it, as update operations are different from health checks.

## 11. Future Improvements

Potential improvements include:

- Centralized storage of health-check reports.
- Automated scheduling of periodic health checks.
- Email notifications for critical findings.
- Historical comparison of health-check results.
- Improved reporting and recommendation logic.

## 12. Summary

This project automates routine RHEL operating system health checks using Ansible. It collects system information, assesses potential issues, and generates individual and consolidated reports.

The goal is to make health checks more consistent, reduce manual effort, and help administrators identify issues that need attention.


