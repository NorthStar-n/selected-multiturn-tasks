#!/bin/bash
set -euo pipefail
cd "$(dirname "$0")"
mkdir -p /app/output
cp 'Emergency Follow-Up Schedule.xlsx' '/app/output/Emergency Follow-Up Schedule.xlsx'
cp 'ExplanatoryFile(N).docx' '/app/output/ExplanatoryFile(N).docx'
cp 'Integrated Executive Master file.xlsx' '/app/output/Integrated Executive Master file.xlsx'
cp 'MEDICAL CLEARANCE LETTER.docx' '/app/output/MEDICAL CLEARANCE LETTER.docx'
cp 'Solution file (Messy vs Master).xlsx' '/app/output/Solution file (Messy vs Master).xlsx'
cp 'Updated revenue reconciliation and lab audit.xlsx' '/app/output/Updated revenue reconciliation and lab audit.xlsx'
