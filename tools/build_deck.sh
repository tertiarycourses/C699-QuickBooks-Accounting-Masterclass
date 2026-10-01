#!/bin/sh
# Rebuild the C699 deck from the WSQ source: convert_deck.py, then fix_deck.py, then PDF export.
set -e
cd "$(dirname "$0")/.."
OUT="courseware/QuickBooks Accounting Masterclass (C699)-v1.0.pptx"
python3 ~/.claude/skills/wsq-to-non-wsq/scripts/convert_deck.py \
  --src "source-wsq/courseware/Advanced Transactional Accounting with QuickBooks Online-v11.1.pptx" --out "$OUT" \
  --old-title "Advanced Transactional Accounting with QuickBooks Online" --new-title "QuickBooks Accounting Masterclass" \
  --old-code TGS-2023018989 --new-code C699 --tsc "ACC-CRP-4009-1.1" \
  --old-version v11.1 --new-version v1.0 \
  --sub "Version v11.1 · 16 August 2026=>Version v1.0 · 1 October 2026" \
  --sub "TGS-2023018989-Advanced-Transactional-Accounting-with-Quickbooks-Online=>C699-QuickBooks-Accounting-Masterclass" \
  --wsq-slides 1,6,7,11,12,13,132,134,135 --feedback-slide 136
python3 tools/fix_deck.py
