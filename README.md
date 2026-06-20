
# 📱 App Adoption Report Generator

Automated daily reporting of app user adoption rates
across organisational units (region → area → branch).

## What It Does
- Reads user account Excel exports
- Filters active users, removes test accounts
- Calculates adoption % per organisational unit
- Outputs timestamped Excel with summary + full user list

## Why It Exists
Previously done manually in Excel — time-consuming and error-prone.
This script ran hundreds of times for daily reporting,
saving significant time per cycle.

## Tech
Python · Pandas
