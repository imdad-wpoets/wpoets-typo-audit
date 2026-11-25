#!/usr/bin/env python3
"""
typo_audit_app.py
================================================================================
Zuci Systems — Interactive Website Typo & Character Confusion Audit Studio
Features:
  1. Single or Multi-URL Input Box (plus optional CSV/TXT upload & Content Type tagging)
  2. Fresh Scan vs. Append to Master Report Mode (with automatic deduplication)
  3. Live Editable Character Confusion Matrix Rules Table (205 default rules + custom rules)
  4. Live Editable Brand & Technical Whitelist Manager
  5. Playwright Headless Chromium Renderer (with cache-bypass headers, desktop + mobile DOM,
     SEO meta tags, image alt/title, and attribute extraction)
  6. 1-Click 3-Sheet Excel Workbook Download (Executive Summary, Master URL Report,
     All Issue Findings with live formulas & severity formatting)

Run locally:
  pip install streamlit playwright beautifulsoup4 lxml pandas openpyxl pyspellchecker
  playwright install chromium
  streamlit run typo_audit_app.py
================================================================================
"""

from __future__ import annotations

import asyncio
import io
import json
import re
import subprocess
import sys
import time
import unicodedata
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

import pandas as pd
import streamlit as st
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

try:
    from spellchecker import SpellChecker
except ImportError:
    SpellChecker = None

try:
    from playwright.async_api import async_playwright
except ImportError:
    async_playwright = None

# Windows asyncio subprocess policy fix for Playwright inside Streamlit
if sys.platform.startswith("win"):
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

# ==============================================================================
# 1. PERSISTENT RULE & WHITELIST STORAGE
# ==============================================================================
RULES_STORE_PATH = Path("rules_store.json")
WHITELIST_STORE_PATH = Path("whitelist_store.json")
MASTER_EXCEL_PATH = Path("Master_Zuci_Typo_Audit_Report.xlsx")

DEFAULT_WHITELIST: List[str] = [
    "Zuci", "ZuciSystems", "TruZt", "TeZt", "IGNITEZ", "ZoY", "ZCortex",
    "OPTIMUS", "PRIMAL", "Lyzr", "Tricentis", "Snowflake", "UiPath",
    "Boomi", "iPaaS", "MLOps", "FinTech", "HealthTech", "RegTech",
    "InsurTech", "DevOps", "DevSecOps", "GenAI", "Agentic", "Ellowhale",
    "Intics", "Radiare", "Swadesh", "DPIIT", "NBFC", "RBI", "HIPAA",
    "SOC2", "ISO", "GDPR", "SDLC", "MVP", "PoC", "PoV", "ROI", "SME",
    "LOS", "RPA", "APA", "OCR", "CX", "QA", "QE", "BI", "API", "APIs",
    "SaaS", "PaaS", "IaaS", "AI", "LLM", "LLMs", "NLP", "UI", "UX",
    "SQL", "NoSQL", "ETLs", "ETL", "CI/CD", "24x7", "5x", "10x",
    "B2B", "B2C", "OAuth", "SAML", "SSO", "WebP", "PDF", "PDFs",
    "iOS", "Android", "Rapid7", "Metasploit", "Sezzle", "Dynamics",
    "SAP", "MACH", "AWS", "Azure", "GCP", "Kubernetes", "Docker",
    "IoT", "Web3", "FinOps", "DataOps", "AIOps", "GitHub", "GitLab",
    "Jira", "Selenium", "Playwright", "Appium", "Tosca", "qTest",
    "Cybergenie", "ADAPT", "GOAL"
]

DEFAULT_RULES: List[Dict] = [
  {
    "Rule_ID": "CCM-001",
    "Category": "Alphabetic",
    "Actual Character(s)": "I",
    "Potential Confusion": "l",
    "Example / Notes": "Uppercase I may resemble lowercase l, digit 1, or vertical bar",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-002",
    "Category": "Alphabetic",
    "Actual Character(s)": "I",
    "Potential Confusion": "1",
    "Example / Notes": "Uppercase I may resemble lowercase l, digit 1, or vertical bar",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-003",
    "Category": "Alphabetic",
    "Actual Character(s)": "I",
    "Potential Confusion": "|",
    "Example / Notes": "Uppercase I may resemble lowercase l, digit 1, or vertical bar",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-004",
    "Category": "Alphabetic",
    "Actual Character(s)": "l",
    "Potential Confusion": "I",
    "Example / Notes": "Lowercase L may resemble uppercase I, digit 1, or vertical bar",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-005",
    "Category": "Alphabetic",
    "Actual Character(s)": "l",
    "Potential Confusion": "1",
    "Example / Notes": "Lowercase L may resemble uppercase I, digit 1, or vertical bar",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-006",
    "Category": "Alphabetic",
    "Actual Character(s)": "l",
    "Potential Confusion": "|",
    "Example / Notes": "Lowercase L may resemble uppercase I, digit 1, or vertical bar",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-007",
    "Category": "Alphabetic",
    "Actual Character(s)": "O",
    "Potential Confusion": "0",
    "Example / Notes": "Uppercase O may resemble zero, Q, or D depending on font",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-008",
    "Category": "Alphabetic",
    "Actual Character(s)": "O",
    "Potential Confusion": "Q",
    "Example / Notes": "Uppercase O may resemble zero, Q, or D depending on font",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-009",
    "Category": "Alphabetic",
    "Actual Character(s)": "O",
    "Potential Confusion": "D",
    "Example / Notes": "Uppercase O may resemble zero, Q, or D depending on font",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-010",
    "Category": "Alphabetic",
    "Actual Character(s)": "o",
    "Potential Confusion": "0",
    "Example / Notes": "Lowercase o may resemble zero, uppercase O, or a",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-011",
    "Category": "Alphabetic",
    "Actual Character(s)": "o",
    "Potential Confusion": "O",
    "Example / Notes": "Lowercase o may resemble zero, uppercase O, or a",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-012",
    "Category": "Alphabetic",
    "Actual Character(s)": "o",
    "Potential Confusion": "a",
    "Example / Notes": "Lowercase o may resemble zero, uppercase O, or a",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-013",
    "Category": "Alphabetic",
    "Actual Character(s)": "S",
    "Potential Confusion": "5",
    "Example / Notes": "S may resemble 5 or dollar sign",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-014",
    "Category": "Alphabetic",
    "Actual Character(s)": "S",
    "Potential Confusion": "$",
    "Example / Notes": "S may resemble 5 or dollar sign",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-015",
    "Category": "Alphabetic",
    "Actual Character(s)": "s",
    "Potential Confusion": "5",
    "Example / Notes": "Lowercase s may resemble 5 or uppercase S",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-016",
    "Category": "Alphabetic",
    "Actual Character(s)": "s",
    "Potential Confusion": "S",
    "Example / Notes": "Lowercase s may resemble 5 or uppercase S",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-017",
    "Category": "Alphabetic",
    "Actual Character(s)": "B",
    "Potential Confusion": "8",
    "Example / Notes": "B may resemble 8 or 3",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-018",
    "Category": "Alphabetic",
    "Actual Character(s)": "B",
    "Potential Confusion": "3",
    "Example / Notes": "B may resemble 8 or 3",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-019",
    "Category": "Alphabetic",
    "Actual Character(s)": "Z",
    "Potential Confusion": "2",
    "Example / Notes": "Z may resemble 2 or 7",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-020",
    "Category": "Alphabetic",
    "Actual Character(s)": "Z",
    "Potential Confusion": "7",
    "Example / Notes": "Z may resemble 2 or 7",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-021",
    "Category": "Alphabetic",
    "Actual Character(s)": "G",
    "Potential Confusion": "6",
    "Example / Notes": "G may resemble 6 or C",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-022",
    "Category": "Alphabetic",
    "Actual Character(s)": "G",
    "Potential Confusion": "C",
    "Example / Notes": "G may resemble 6 or C",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-023",
    "Category": "Alphabetic",
    "Actual Character(s)": "C",
    "Potential Confusion": "G",
    "Example / Notes": "C may resemble G, O, or lowercase c",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-024",
    "Category": "Alphabetic",
    "Actual Character(s)": "C",
    "Potential Confusion": "O",
    "Example / Notes": "C may resemble G, O, or lowercase c",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-025",
    "Category": "Alphabetic",
    "Actual Character(s)": "C",
    "Potential Confusion": "c",
    "Example / Notes": "C may resemble G, O, or lowercase c",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-026",
    "Category": "Alphabetic",
    "Actual Character(s)": "D",
    "Potential Confusion": "O",
    "Example / Notes": "D may resemble O or zero",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-027",
    "Category": "Alphabetic",
    "Actual Character(s)": "D",
    "Potential Confusion": "0",
    "Example / Notes": "D may resemble O or zero",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-028",
    "Category": "Alphabetic",
    "Actual Character(s)": "Q",
    "Potential Confusion": "O",
    "Example / Notes": "Q may resemble O or zero",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-029",
    "Category": "Alphabetic",
    "Actual Character(s)": "Q",
    "Potential Confusion": "0",
    "Example / Notes": "Q may resemble O or zero",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-030",
    "Category": "Alphabetic",
    "Actual Character(s)": "U",
    "Potential Confusion": "V",
    "Example / Notes": "U may resemble V or J",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-031",
    "Category": "Alphabetic",
    "Actual Character(s)": "U",
    "Potential Confusion": "J",
    "Example / Notes": "U may resemble V or J",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-032",
    "Category": "Alphabetic",
    "Actual Character(s)": "V",
    "Potential Confusion": "U",
    "Example / Notes": "V may resemble U or Y",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-033",
    "Category": "Alphabetic",
    "Actual Character(s)": "V",
    "Potential Confusion": "Y",
    "Example / Notes": "V may resemble U or Y",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-034",
    "Category": "Alphabetic",
    "Actual Character(s)": "W",
    "Potential Confusion": "VV",
    "Example / Notes": "W may be segmented as VV or resemble M",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-035",
    "Category": "Alphabetic",
    "Actual Character(s)": "W",
    "Potential Confusion": "M",
    "Example / Notes": "W may be segmented as VV or resemble M",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-036",
    "Category": "Alphabetic",
    "Actual Character(s)": "M",
    "Potential Confusion": "N",
    "Example / Notes": "M may resemble N, W, or the sequence rn",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-037",
    "Category": "Alphabetic",
    "Actual Character(s)": "M",
    "Potential Confusion": "W",
    "Example / Notes": "M may resemble N, W, or the sequence rn",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-038",
    "Category": "Alphabetic",
    "Actual Character(s)": "M",
    "Potential Confusion": "rn",
    "Example / Notes": "M may resemble N, W, or the sequence rn",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-039",
    "Category": "Alphabetic",
    "Actual Character(s)": "N",
    "Potential Confusion": "M",
    "Example / Notes": "N may resemble M, H, or the sequence ri",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-040",
    "Category": "Alphabetic",
    "Actual Character(s)": "N",
    "Potential Confusion": "H",
    "Example / Notes": "N may resemble M, H, or the sequence ri",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-041",
    "Category": "Alphabetic",
    "Actual Character(s)": "N",
    "Potential Confusion": "ri",
    "Example / Notes": "N may resemble M, H, or the sequence ri",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-042",
    "Category": "Alphabetic",
    "Actual Character(s)": "H",
    "Potential Confusion": "N",
    "Example / Notes": "H may resemble N or K",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-043",
    "Category": "Alphabetic",
    "Actual Character(s)": "H",
    "Potential Confusion": "K",
    "Example / Notes": "H may resemble N or K",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-044",
    "Category": "Alphabetic",
    "Actual Character(s)": "K",
    "Potential Confusion": "X",
    "Example / Notes": "K may resemble X or R",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-045",
    "Category": "Alphabetic",
    "Actual Character(s)": "K",
    "Potential Confusion": "R",
    "Example / Notes": "K may resemble X or R",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-046",
    "Category": "Alphabetic",
    "Actual Character(s)": "E",
    "Potential Confusion": "F",
    "Example / Notes": "E may resemble F or C",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-047",
    "Category": "Alphabetic",
    "Actual Character(s)": "E",
    "Potential Confusion": "C",
    "Example / Notes": "E may resemble F or C",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-048",
    "Category": "Alphabetic",
    "Actual Character(s)": "F",
    "Potential Confusion": "E",
    "Example / Notes": "F may resemble E or P",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-049",
    "Category": "Alphabetic",
    "Actual Character(s)": "F",
    "Potential Confusion": "P",
    "Example / Notes": "F may resemble E or P",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-050",
    "Category": "Alphabetic",
    "Actual Character(s)": "P",
    "Potential Confusion": "F",
    "Example / Notes": "P may resemble F or R",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-051",
    "Category": "Alphabetic",
    "Actual Character(s)": "P",
    "Potential Confusion": "R",
    "Example / Notes": "P may resemble F or R",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-052",
    "Category": "Alphabetic",
    "Actual Character(s)": "R",
    "Potential Confusion": "P",
    "Example / Notes": "R may resemble P or K",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-053",
    "Category": "Alphabetic",
    "Actual Character(s)": "R",
    "Potential Confusion": "K",
    "Example / Notes": "R may resemble P or K",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-054",
    "Category": "Alphabetic",
    "Actual Character(s)": "A",
    "Potential Confusion": "4",
    "Example / Notes": "A may resemble digit 4 or Greek capital lambda",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-055",
    "Category": "Alphabetic",
    "Actual Character(s)": "A",
    "Potential Confusion": "Λ",
    "Example / Notes": "A may resemble digit 4 or Greek capital lambda",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-056",
    "Category": "Alphabetic",
    "Actual Character(s)": "T",
    "Potential Confusion": "7",
    "Example / Notes": "T may resemble 7 or I",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-057",
    "Category": "Alphabetic",
    "Actual Character(s)": "T",
    "Potential Confusion": "I",
    "Example / Notes": "T may resemble 7 or I",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-058",
    "Category": "Alphabetic",
    "Actual Character(s)": "Y",
    "Potential Confusion": "V",
    "Example / Notes": "Y may resemble V or yen sign",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-059",
    "Category": "Alphabetic",
    "Actual Character(s)": "Y",
    "Potential Confusion": "¥",
    "Example / Notes": "Y may resemble V or yen sign",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-060",
    "Category": "Alphabetic",
    "Actual Character(s)": "J",
    "Potential Confusion": "I",
    "Example / Notes": "J may resemble I or T",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-061",
    "Category": "Alphabetic",
    "Actual Character(s)": "J",
    "Potential Confusion": "T",
    "Example / Notes": "J may resemble I or T",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-062",
    "Category": "Alphabetic",
    "Actual Character(s)": "X",
    "Potential Confusion": "K",
    "Example / Notes": "X may resemble K or multiplication sign",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-063",
    "Category": "Alphabetic",
    "Actual Character(s)": "X",
    "Potential Confusion": "×",
    "Example / Notes": "X may resemble K or multiplication sign",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-064",
    "Category": "Alphabetic",
    "Actual Character(s)": "a",
    "Potential Confusion": "o",
    "Example / Notes": "Lowercase a can be confused in degraded or unusual fonts",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-065",
    "Category": "Alphabetic",
    "Actual Character(s)": "a",
    "Potential Confusion": "e",
    "Example / Notes": "Lowercase a can be confused in degraded or unusual fonts",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-066",
    "Category": "Alphabetic",
    "Actual Character(s)": "a",
    "Potential Confusion": "d",
    "Example / Notes": "Lowercase a can be confused in degraded or unusual fonts",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-067",
    "Category": "Alphabetic",
    "Actual Character(s)": "e",
    "Potential Confusion": "c",
    "Example / Notes": "Lowercase e may resemble c or o",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-068",
    "Category": "Alphabetic",
    "Actual Character(s)": "e",
    "Potential Confusion": "o",
    "Example / Notes": "Lowercase e may resemble c or o",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-069",
    "Category": "Alphabetic",
    "Actual Character(s)": "g",
    "Potential Confusion": "q",
    "Example / Notes": "Lowercase g may resemble q, 9, or y",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-070",
    "Category": "Alphabetic",
    "Actual Character(s)": "g",
    "Potential Confusion": "9",
    "Example / Notes": "Lowercase g may resemble q, 9, or y",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-071",
    "Category": "Alphabetic",
    "Actual Character(s)": "g",
    "Potential Confusion": "y",
    "Example / Notes": "Lowercase g may resemble q, 9, or y",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-072",
    "Category": "Alphabetic",
    "Actual Character(s)": "q",
    "Potential Confusion": "g",
    "Example / Notes": "Lowercase q may resemble g or 9",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-073",
    "Category": "Alphabetic",
    "Actual Character(s)": "q",
    "Potential Confusion": "9",
    "Example / Notes": "Lowercase q may resemble g or 9",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-074",
    "Category": "Alphabetic",
    "Actual Character(s)": "h",
    "Potential Confusion": "b",
    "Example / Notes": "Lowercase h may resemble b, k, or n in degraded text",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-075",
    "Category": "Alphabetic",
    "Actual Character(s)": "h",
    "Potential Confusion": "k",
    "Example / Notes": "Lowercase h may resemble b, k, or n in degraded text",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-076",
    "Category": "Alphabetic",
    "Actual Character(s)": "h",
    "Potential Confusion": "n",
    "Example / Notes": "Lowercase h may resemble b, k, or n in degraded text",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-077",
    "Category": "Alphabetic",
    "Actual Character(s)": "n",
    "Potential Confusion": "r",
    "Example / Notes": "Lowercase n may resemble r or h",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-078",
    "Category": "Alphabetic",
    "Actual Character(s)": "n",
    "Potential Confusion": "h",
    "Example / Notes": "Lowercase n may resemble r or h",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-079",
    "Category": "Alphabetic",
    "Actual Character(s)": "r",
    "Potential Confusion": "n",
    "Example / Notes": "Lowercase r may resemble n or i",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-080",
    "Category": "Alphabetic",
    "Actual Character(s)": "r",
    "Potential Confusion": "i",
    "Example / Notes": "Lowercase r may resemble n or i",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-081",
    "Category": "Alphabetic",
    "Actual Character(s)": "m",
    "Potential Confusion": "rn",
    "Example / Notes": "Lowercase m may be segmented into rn or nn",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-082",
    "Category": "Alphabetic",
    "Actual Character(s)": "m",
    "Potential Confusion": "nn",
    "Example / Notes": "Lowercase m may be segmented into rn or nn",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-083",
    "Category": "Alphabetic",
    "Actual Character(s)": "u",
    "Potential Confusion": "v",
    "Example / Notes": "Lowercase u may resemble v or n",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-084",
    "Category": "Alphabetic",
    "Actual Character(s)": "u",
    "Potential Confusion": "n",
    "Example / Notes": "Lowercase u may resemble v or n",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-085",
    "Category": "Alphabetic",
    "Actual Character(s)": "v",
    "Potential Confusion": "u",
    "Example / Notes": "Lowercase v may resemble u or y",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-086",
    "Category": "Alphabetic",
    "Actual Character(s)": "v",
    "Potential Confusion": "y",
    "Example / Notes": "Lowercase v may resemble u or y",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-087",
    "Category": "Alphabetic",
    "Actual Character(s)": "w",
    "Potential Confusion": "vv",
    "Example / Notes": "Lowercase w may be segmented into vv or resemble m",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-088",
    "Category": "Alphabetic",
    "Actual Character(s)": "w",
    "Potential Confusion": "m",
    "Example / Notes": "Lowercase w may be segmented into vv or resemble m",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-089",
    "Category": "Alphabetic",
    "Actual Character(s)": "y",
    "Potential Confusion": "v",
    "Example / Notes": "Lowercase y may resemble v or g",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-090",
    "Category": "Alphabetic",
    "Actual Character(s)": "y",
    "Potential Confusion": "g",
    "Example / Notes": "Lowercase y may resemble v or g",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-091",
    "Category": "Numeric",
    "Actual Character(s)": "0",
    "Potential Confusion": "O",
    "Example / Notes": "Zero may be confused with letter O/o or Q",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-092",
    "Category": "Numeric",
    "Actual Character(s)": "0",
    "Potential Confusion": "o",
    "Example / Notes": "Zero may be confused with letter O/o or Q",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-093",
    "Category": "Numeric",
    "Actual Character(s)": "0",
    "Potential Confusion": "Q",
    "Example / Notes": "Zero may be confused with letter O/o or Q",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-094",
    "Category": "Numeric",
    "Actual Character(s)": "1",
    "Potential Confusion": "I",
    "Example / Notes": "One may be confused with I, lowercase l, vertical bar, or 7",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-095",
    "Category": "Numeric",
    "Actual Character(s)": "1",
    "Potential Confusion": "l",
    "Example / Notes": "One may be confused with I, lowercase l, vertical bar, or 7",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-096",
    "Category": "Numeric",
    "Actual Character(s)": "1",
    "Potential Confusion": "|",
    "Example / Notes": "One may be confused with I, lowercase l, vertical bar, or 7",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-097",
    "Category": "Numeric",
    "Actual Character(s)": "1",
    "Potential Confusion": "7",
    "Example / Notes": "One may be confused with I, lowercase l, vertical bar, or 7",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-098",
    "Category": "Numeric",
    "Actual Character(s)": "2",
    "Potential Confusion": "Z",
    "Example / Notes": "Two may resemble Z or 7",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-099",
    "Category": "Numeric",
    "Actual Character(s)": "2",
    "Potential Confusion": "7",
    "Example / Notes": "Two may resemble Z or 7",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-100",
    "Category": "Numeric",
    "Actual Character(s)": "3",
    "Potential Confusion": "B",
    "Example / Notes": "Three may resemble B or 8",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-101",
    "Category": "Numeric",
    "Actual Character(s)": "3",
    "Potential Confusion": "8",
    "Example / Notes": "Three may resemble B or 8",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-102",
    "Category": "Numeric",
    "Actual Character(s)": "4",
    "Potential Confusion": "A",
    "Example / Notes": "Four may resemble A or 9",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-103",
    "Category": "Numeric",
    "Actual Character(s)": "4",
    "Potential Confusion": "9",
    "Example / Notes": "Four may resemble A or 9",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-104",
    "Category": "Numeric",
    "Actual Character(s)": "5",
    "Potential Confusion": "S",
    "Example / Notes": "Five may resemble S or 6",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-105",
    "Category": "Numeric",
    "Actual Character(s)": "5",
    "Potential Confusion": "6",
    "Example / Notes": "Five may resemble S or 6",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-106",
    "Category": "Numeric",
    "Actual Character(s)": "6",
    "Potential Confusion": "G",
    "Example / Notes": "Six may resemble G, 5, or 8",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-107",
    "Category": "Numeric",
    "Actual Character(s)": "6",
    "Potential Confusion": "5",
    "Example / Notes": "Six may resemble G, 5, or 8",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-108",
    "Category": "Numeric",
    "Actual Character(s)": "6",
    "Potential Confusion": "8",
    "Example / Notes": "Six may resemble G, 5, or 8",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-109",
    "Category": "Numeric",
    "Actual Character(s)": "7",
    "Potential Confusion": "T",
    "Example / Notes": "Seven may resemble T, Z, or 1",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-110",
    "Category": "Numeric",
    "Actual Character(s)": "7",
    "Potential Confusion": "Z",
    "Example / Notes": "Seven may resemble T, Z, or 1",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-111",
    "Category": "Numeric",
    "Actual Character(s)": "7",
    "Potential Confusion": "1",
    "Example / Notes": "Seven may resemble T, Z, or 1",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-112",
    "Category": "Numeric",
    "Actual Character(s)": "8",
    "Potential Confusion": "B",
    "Example / Notes": "Eight may resemble B, 3, or 6",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-113",
    "Category": "Numeric",
    "Actual Character(s)": "8",
    "Potential Confusion": "3",
    "Example / Notes": "Eight may resemble B, 3, or 6",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-114",
    "Category": "Numeric",
    "Actual Character(s)": "8",
    "Potential Confusion": "6",
    "Example / Notes": "Eight may resemble B, 3, or 6",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-115",
    "Category": "Numeric",
    "Actual Character(s)": "9",
    "Potential Confusion": "g",
    "Example / Notes": "Nine may resemble lowercase g/q or 4",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-116",
    "Category": "Numeric",
    "Actual Character(s)": "9",
    "Potential Confusion": "q",
    "Example / Notes": "Nine may resemble lowercase g/q or 4",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-117",
    "Category": "Numeric",
    "Actual Character(s)": "9",
    "Potential Confusion": "4",
    "Example / Notes": "Nine may resemble lowercase g/q or 4",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-118",
    "Category": "Multi-character",
    "Actual Character(s)": "rn",
    "Potential Confusion": "m",
    "Example / Notes": "Two letters rn may merge into one m",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-119",
    "Category": "Multi-character",
    "Actual Character(s)": "m",
    "Potential Confusion": "rn",
    "Example / Notes": "One letter m may be split into rn or nn",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-120",
    "Category": "Multi-character",
    "Actual Character(s)": "m",
    "Potential Confusion": "nn",
    "Example / Notes": "One letter m may be split into rn or nn",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-121",
    "Category": "Multi-character",
    "Actual Character(s)": "cl",
    "Potential Confusion": "d",
    "Example / Notes": "c followed by l may resemble d",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-122",
    "Category": "Multi-character",
    "Actual Character(s)": "d",
    "Potential Confusion": "cl",
    "Example / Notes": "d may be segmented or recognized as cl",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-123",
    "Category": "Multi-character",
    "Actual Character(s)": "vv",
    "Potential Confusion": "w",
    "Example / Notes": "Two v characters may merge into w",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-124",
    "Category": "Multi-character",
    "Actual Character(s)": "w",
    "Potential Confusion": "vv",
    "Example / Notes": "w may be segmented into vv",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-125",
    "Category": "Multi-character",
    "Actual Character(s)": "ri",
    "Potential Confusion": "n",
    "Example / Notes": "r followed by i may merge into n",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-126",
    "Category": "Multi-character",
    "Actual Character(s)": "n",
    "Potential Confusion": "ri",
    "Example / Notes": "n may be split into ri",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-127",
    "Category": "Multi-character",
    "Actual Character(s)": "nn",
    "Potential Confusion": "m",
    "Example / Notes": "Double n may merge into m",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-128",
    "Category": "Multi-character",
    "Actual Character(s)": "ii",
    "Potential Confusion": "u",
    "Example / Notes": "Double i may merge into u",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-129",
    "Category": "Multi-character",
    "Actual Character(s)": "li",
    "Potential Confusion": "h",
    "Example / Notes": "l followed by i may resemble h",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-130",
    "Category": "Multi-character",
    "Actual Character(s)": "h",
    "Potential Confusion": "li",
    "Example / Notes": "h may be segmented as li",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-131",
    "Category": "Multi-character",
    "Actual Character(s)": "ni",
    "Potential Confusion": "n",
    "Example / Notes": "i may be dropped or merged in a sequence",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-132",
    "Category": "Multi-character",
    "Actual Character(s)": "cj",
    "Potential Confusion": "g",
    "Example / Notes": "c followed by j may resemble g in some fonts",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-133",
    "Category": "Multi-character",
    "Actual Character(s)": "oe",
    "Potential Confusion": "œ",
    "Example / Notes": "Separate letters may be rendered as a ligature",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-134",
    "Category": "Multi-character",
    "Actual Character(s)": "ae",
    "Potential Confusion": "æ",
    "Example / Notes": "Separate letters may be rendered as a ligature",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-135",
    "Category": "Multi-character",
    "Actual Character(s)": "fi",
    "Potential Confusion": "ﬁ",
    "Example / Notes": "f+i may be represented by the fi ligature",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-136",
    "Category": "Multi-character",
    "Actual Character(s)": "fl",
    "Potential Confusion": "ﬂ",
    "Example / Notes": "f+l may be represented by the fl ligature",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-137",
    "Category": "Multi-character",
    "Actual Character(s)": "ff",
    "Potential Confusion": "ﬀ",
    "Example / Notes": "f+f may be represented by the ff ligature",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-138",
    "Category": "Multi-character",
    "Actual Character(s)": "ffi",
    "Potential Confusion": "ﬃ",
    "Example / Notes": "f+f+i may be represented by the ffi ligature",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-139",
    "Category": "Multi-character",
    "Actual Character(s)": "ffl",
    "Potential Confusion": "ﬄ",
    "Example / Notes": "f+f+l may be represented by the ffl ligature",
    "Priority": "Low"
  },
  {
    "Rule_ID": "CCM-140",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": ".",
    "Potential Confusion": ",",
    "Example / Notes": "Period may be mistaken for comma",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-141",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": ",",
    "Potential Confusion": ".",
    "Example / Notes": "Comma may be mistaken for period",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-142",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": ":",
    "Potential Confusion": ";",
    "Example / Notes": "Colon may be mistaken for semicolon",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-143",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": ":",
    "Potential Confusion": ";",
    "Example / Notes": "Colon may be mistaken for semicolon",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-144",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": ";",
    "Potential Confusion": ":",
    "Example / Notes": "Semicolon may be mistaken for colon",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-145",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "'",
    "Potential Confusion": "\\`",
    "Example / Notes": "Apostrophe may resemble grave or acute accent",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-146",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "'",
    "Potential Confusion": "´",
    "Example / Notes": "Apostrophe may resemble grave or acute accent",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-147",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "\"",
    "Potential Confusion": "''",
    "Example / Notes": "Quotation mark may be split into two apostrophes",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-148",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "-",
    "Potential Confusion": "–",
    "Example / Notes": "Hyphen may be confused with en dash or em dash",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-149",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "-",
    "Potential Confusion": "—",
    "Example / Notes": "Hyphen may be confused with en dash or em dash",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-150",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "–",
    "Potential Confusion": "-",
    "Example / Notes": "En dash may be confused with hyphen or em dash",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-151",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "–",
    "Potential Confusion": "—",
    "Example / Notes": "En dash may be confused with hyphen or em dash",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-152",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "—",
    "Potential Confusion": "-",
    "Example / Notes": "Em dash may be confused with hyphen or en dash",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-153",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "—",
    "Potential Confusion": "–",
    "Example / Notes": "Em dash may be confused with hyphen or en dash",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-154",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "/",
    "Potential Confusion": "\\\\",
    "Example / Notes": "Forward slash may be confused with backslash",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-155",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "\\\\",
    "Potential Confusion": "/",
    "Example / Notes": "Backslash may be confused with forward slash",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-156",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "(",
    "Potential Confusion": "\\[",
    "Example / Notes": "Opening parenthesis may resemble bracket or brace",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-157",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "(",
    "Potential Confusion": "{",
    "Example / Notes": "Opening parenthesis may resemble bracket or brace",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-158",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": ")",
    "Potential Confusion": "\\]",
    "Example / Notes": "Closing parenthesis may resemble bracket or brace",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-159",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": ")",
    "Potential Confusion": "}",
    "Example / Notes": "Closing parenthesis may resemble bracket or brace",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-160",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "\\[",
    "Potential Confusion": "(",
    "Example / Notes": "Opening bracket may resemble parenthesis or brace",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-161",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "\\[",
    "Potential Confusion": "{",
    "Example / Notes": "Opening bracket may resemble parenthesis or brace",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-162",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "\\]",
    "Potential Confusion": ")",
    "Example / Notes": "Closing bracket may resemble parenthesis or brace",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-163",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "\\]",
    "Potential Confusion": "}",
    "Example / Notes": "Closing bracket may resemble parenthesis or brace",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-164",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "{",
    "Potential Confusion": "(",
    "Example / Notes": "Opening brace may resemble parenthesis or bracket",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-165",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "{",
    "Potential Confusion": "\\[",
    "Example / Notes": "Opening brace may resemble parenthesis or bracket",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-166",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "}",
    "Potential Confusion": ")",
    "Example / Notes": "Closing brace may resemble parenthesis or bracket",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-167",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "}",
    "Potential Confusion": "\\]",
    "Example / Notes": "Closing brace may resemble parenthesis or bracket",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-168",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "\\<",
    "Potential Confusion": "(",
    "Example / Notes": "Less-than sign may resemble parenthesis or single angle quote",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-169",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "\\<",
    "Potential Confusion": "‹",
    "Example / Notes": "Less-than sign may resemble parenthesis or single angle quote",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-170",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "\\>",
    "Potential Confusion": ")",
    "Example / Notes": "Greater-than sign may resemble parenthesis or single angle quote",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-171",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "\\>",
    "Potential Confusion": "›",
    "Example / Notes": "Greater-than sign may resemble parenthesis or single angle quote",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-172",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "=",
    "Potential Confusion": "≡",
    "Example / Notes": "Equals sign may resemble triple bar or hyphen in poor scans",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-173",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "=",
    "Potential Confusion": "-",
    "Example / Notes": "Equals sign may resemble triple bar or hyphen in poor scans",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-174",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "+",
    "Potential Confusion": "t",
    "Example / Notes": "Plus sign may resemble lowercase t",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-175",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "@",
    "Potential Confusion": "a",
    "Example / Notes": "At sign may resemble a",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-176",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "&",
    "Potential Confusion": "8",
    "Example / Notes": "Ampersand may resemble 8",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-177",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "%",
    "Potential Confusion": "‰",
    "Example / Notes": "Percent may resemble per mille",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-178",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "\\#",
    "Potential Confusion": "H",
    "Example / Notes": "Hash may resemble H",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-179",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "\\*",
    "Potential Confusion": "×",
    "Example / Notes": "Asterisk may resemble multiplication sign",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-180",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "\\!",
    "Potential Confusion": "l",
    "Example / Notes": "Exclamation mark may resemble lowercase l or uppercase I",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-181",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "\\!",
    "Potential Confusion": "I",
    "Example / Notes": "Exclamation mark may resemble lowercase l or uppercase I",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-182",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "?",
    "Potential Confusion": "7",
    "Example / Notes": "Question mark may resemble 7 in degraded scans",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-183",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "°",
    "Potential Confusion": "o",
    "Example / Notes": "Degree symbol may resemble o or zero",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-184",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "°",
    "Potential Confusion": "0",
    "Example / Notes": "Degree symbol may resemble o or zero",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-185",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "|",
    "Potential Confusion": "I",
    "Example / Notes": "Vertical bar may resemble I, l, or 1",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-186",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "|",
    "Potential Confusion": "l",
    "Example / Notes": "Vertical bar may resemble I, l, or 1",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-187",
    "Category": "Special characters & punctuation",
    "Actual Character(s)": "|",
    "Potential Confusion": "1",
    "Example / Notes": "Vertical bar may resemble I, l, or 1",
    "Priority": "Medium"
  },
  {
    "Rule_ID": "CCM-188",
    "Category": "Unicode lookalikes",
    "Actual Character(s)": "a",
    "Potential Confusion": "а (Cyrillic)",
    "Example / Notes": "Unicode code points: U+0061 / U+0430",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-189",
    "Category": "Unicode lookalikes",
    "Actual Character(s)": "e",
    "Potential Confusion": "е (Cyrillic)",
    "Example / Notes": "Unicode code points: U+0065 / U+0435",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-190",
    "Category": "Unicode lookalikes",
    "Actual Character(s)": "o",
    "Potential Confusion": "о (Cyrillic)",
    "Example / Notes": "Unicode code points: U+006F / U+043E",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-191",
    "Category": "Unicode lookalikes",
    "Actual Character(s)": "p",
    "Potential Confusion": "р (Cyrillic)",
    "Example / Notes": "Unicode code points: U+0070 / U+0440",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-192",
    "Category": "Unicode lookalikes",
    "Actual Character(s)": "c",
    "Potential Confusion": "с (Cyrillic)",
    "Example / Notes": "Unicode code points: U+0063 / U+0441",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-193",
    "Category": "Unicode lookalikes",
    "Actual Character(s)": "x",
    "Potential Confusion": "х (Cyrillic)",
    "Example / Notes": "Unicode code points: U+0078 / U+0445",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-194",
    "Category": "Unicode lookalikes",
    "Actual Character(s)": "i",
    "Potential Confusion": "і (Ukrainian)",
    "Example / Notes": "Unicode code points: U+0069 / U+0456",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-195",
    "Category": "Unicode lookalikes",
    "Actual Character(s)": "j",
    "Potential Confusion": "ј (Cyrillic)",
    "Example / Notes": "Unicode code points: U+006A / U+0458",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-196",
    "Category": "Unicode lookalikes",
    "Actual Character(s)": "A",
    "Potential Confusion": "Α (Greek Alpha)",
    "Example / Notes": "Unicode code points: U+0041 / U+0391",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-197",
    "Category": "Unicode lookalikes",
    "Actual Character(s)": "B",
    "Potential Confusion": "Β (Greek Beta)",
    "Example / Notes": "Unicode code points: U+0042 / U+0392",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-198",
    "Category": "Unicode lookalikes",
    "Actual Character(s)": "E",
    "Potential Confusion": "Ε (Greek Epsilon)",
    "Example / Notes": "Unicode code points: U+0045 / U+0395",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-199",
    "Category": "Unicode lookalikes",
    "Actual Character(s)": "K",
    "Potential Confusion": "Κ (Greek Kappa)",
    "Example / Notes": "Unicode code points: U+004B / U+039A",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-200",
    "Category": "Unicode lookalikes",
    "Actual Character(s)": "M",
    "Potential Confusion": "Μ (Greek Mu)",
    "Example / Notes": "Unicode code points: U+004D / U+039C",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-201",
    "Category": "Unicode lookalikes",
    "Actual Character(s)": "N",
    "Potential Confusion": "Ν (Greek Nu)",
    "Example / Notes": "Unicode code points: U+004E / U+039D",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-202",
    "Category": "Unicode lookalikes",
    "Actual Character(s)": "O",
    "Potential Confusion": "Ο (Greek Omicron)",
    "Example / Notes": "Unicode code points: U+004F / U+039F",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-203",
    "Category": "Unicode lookalikes",
    "Actual Character(s)": "P",
    "Potential Confusion": "Ρ (Greek Rho)",
    "Example / Notes": "Unicode code points: U+0050 / U+03A1",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-204",
    "Category": "Unicode lookalikes",
    "Actual Character(s)": "T",
    "Potential Confusion": "Τ (Greek Tau)",
    "Example / Notes": "Unicode code points: U+0054 / U+03A4",
    "Priority": "High"
  },
  {
    "Rule_ID": "CCM-205",
    "Category": "Unicode lookalikes",
    "Actual Character(s)": "X",
    "Potential Confusion": "Χ (Greek Chi)",
    "Example / Notes": "Unicode code points: U+0058 / U+03A7",
    "Priority": "High"
  }
]


def load_rules_df() -> pd.DataFrame:
    if RULES_STORE_PATH.exists():
        df = pd.read_json(RULES_STORE_PATH)
    else:
        df = pd.DataFrame(DEFAULT_RULES)
        if "Enabled" not in df.columns:
            df["Enabled"] = True
        save_rules_df(df)
    if "Enabled" not in df.columns:
        df["Enabled"] = True
    return df


def save_rules_df(df: pd.DataFrame) -> None:
    df.to_json(RULES_STORE_PATH, orient="records", indent=2, force_ascii=False)


def load_whitelist() -> List[str]:
    if WHITELIST_STORE_PATH.exists():
        return json.loads(WHITELIST_STORE_PATH.read_text(encoding="utf-8"))
    save_whitelist(DEFAULT_WHITELIST)
    return list(DEFAULT_WHITELIST)


def save_whitelist(terms: List[str]) -> None:
    cleaned = sorted(set(t.strip() for t in terms if t.strip()), key=str.lower)
    WHITELIST_STORE_PATH.write_text(json.dumps(cleaned, indent=2, ensure_ascii=False), encoding="utf-8")


# ==============================================================================
# 2. CHARACTER CONFUSION & TYPO DETECTION ENGINE (DRIVEN BY EDITABLE RULES)
# ==============================================================================
UNICODE_LOOKALIKE_MAP: Dict[str, Tuple[str, str, str]] = {
    "\u0430": ("a", "U+0061 / U+0430", "Cyrillic Small Letter A replacing Latin 'a'"),
    "\u0435": ("e", "U+0065 / U+0435", "Cyrillic Small Letter Ie replacing Latin 'e'"),
    "\u043E": ("o", "U+006F / U+043E", "Cyrillic Small Letter O replacing Latin 'o'"),
    "\u0440": ("p", "U+0070 / U+0440", "Cyrillic Small Letter Er replacing Latin 'p'"),
    "\u0441": ("c", "U+0063 / U+0441", "Cyrillic Small Letter Es replacing Latin 'c'"),
    "\u0445": ("x", "U+0078 / U+0445", "Cyrillic Small Letter Ha replacing Latin 'x'"),
    "\u0456": ("i", "U+0069 / U+0456", "Cyrillic/Ukrainian I replacing Latin 'i'"),
    "\u0458": ("j", "U+006A / U+0458", "Cyrillic Small Letter Je replacing Latin 'j'"),
    "\u0391": ("A", "U+0041 / U+0391", "Greek Capital Alpha replacing Latin 'A'"),
    "\u0392": ("B", "U+0042 / U+0392", "Greek Capital Beta replacing Latin 'B'"),
    "\u0395": ("E", "U+0045 / U+0395", "Greek Capital Epsilon replacing Latin 'E'"),
    "\u039A": ("K", "U+004B / U+039A", "Greek Capital Kappa replacing Latin 'K'"),
    "\u039C": ("M", "U+004D / U+039C", "Greek Capital Mu replacing Latin 'M'"),
    "\u039D": ("N", "U+004E / U+039D", "Greek Capital Nu replacing Latin 'N'"),
    "\u039F": ("O", "U+004F / U+039F", "Greek Capital Omicron replacing Latin 'O'"),
    "\u03A1": ("P", "U+0050 / U+03A1", "Greek Capital Rho replacing Latin 'P'"),
    "\u03A4": ("T", "U+0054 / U+03A4", "Greek Capital Tau replacing Latin 'T'"),
    "\u03A7": ("X", "U+0058 / U+03A7", "Greek Capital Chi replacing Latin 'X'"),
    "\u039B": ("A", "U+0041 / U+039B", "Greek Capital Lamda replacing Latin 'A'"),
}

UNICODE_LIGATURE_MAP: Dict[str, Tuple[str, str]] = {
    "\uFB01": ("fi", "fi -> \uFB01 (U+FB01)"),
    "\uFB02": ("fl", "fl -> \uFB02 (U+FB02)"),
    "\uFB00": ("ff", "ff -> \uFB00 (U+FB00)"),
    "\uFB03": ("ffi", "ffi -> \uFB03 (U+FB03)"),
    "\uFB04": ("ffl", "ffl -> \uFB04 (U+FB04)"),
    "\u0153": ("oe", "oe -> \u0153 (U+0153)"),
    "\u00E6": ("ae", "ae -> \u00E6 (U+00E6)"),
}


class DynamicConfusionEngine:
    """Runs character confusion detection directly using the user's live editable rules_df and whitelist."""

    def __init__(self, rules_df: pd.DataFrame, whitelist: List[str]):
        self.rules_df = rules_df[rules_df["Enabled"] == True].copy()
        self.whitelist_lower: Set[str] = {w.lower() for w in whitelist}
        self.spell = SpellChecker(language="en") if SpellChecker else None
        if self.spell and self.whitelist_lower:
            self.spell.word_frequency.load_words(list(self.whitelist_lower))

        # Build dynamic lookup sets from active rules table
        self.active_alpha_pairs: List[Tuple[str, str, str, str]] = []
        self.active_multichar_pairs: List[Tuple[str, str, str, str]] = []
        self.active_digit_to_letter: Dict[str, List[Tuple[str, str, str]]] = {}
        self.active_letter_to_digit: Dict[str, List[Tuple[str, str, str]]] = {}
        self.unicode_enabled = False
        self.punct_enabled = False

        for r in self.rules_df.itertuples(index=False):
            cat = str(getattr(r, "Category", ""))
            act = str(getattr(r, "Actual Character(s)", getattr(r, "Actual_Characters", ""))).strip()
            conf = str(getattr(r, "Potential Confusion", getattr(r, "Potential_Confusion", ""))).strip()
            rid = str(getattr(r, "Rule_ID", "CUSTOM"))
            prio = str(getattr(r, "Priority", "Medium"))
            label = f"{act} -> {conf} ({rid})"

            if cat == "Unicode lookalikes":
                self.unicode_enabled = True
            elif cat == "Special characters & punctuation":
                self.punct_enabled = True
            elif cat == "Multi-character":
                if len(act) >= 1 and len(conf) >= 1 and act != conf:
                    # Observed in text is `conf` when intended was `act`, or vice versa
                    self.active_multichar_pairs.append((conf, act, label, prio))
            elif cat == "Alphabetic":
                if len(act) == 1 and len(conf) == 1 and act.isalpha() and conf.isalpha():
                    self.active_alpha_pairs.append((conf, act, label, prio))
            elif cat == "Numeric":
                if act.isdigit() and len(conf) == 1 and conf.isalpha():
                    self.active_letter_to_digit.setdefault(conf, []).append((act, label, prio))
                elif act.isalpha() and len(act) == 1 and conf.isdigit():
                    self.active_digit_to_letter.setdefault(conf, []).append((act, label, prio))

    def is_valid_word(self, word: str) -> bool:
        clean = word.strip("'\".,;:!?()[]{}")
        if not clean:
            return True
        low = clean.lower()
        if low in self.whitelist_lower or len(clean) <= 2:
            return True
        # Flag anomalous mid-word uppercase (e.g. 'Al-Infused' split token 'Al', 'cIoud', 'MuIti')
        if re.match(r"^[a-z]+[A-Z][a-z]+$", clean) or re.match(r"^[A-Z][a-z]+[A-Z][a-z]+$", clean):
            return False
        if clean.isupper() and len(clean) <= 7:
            return True
        if not self.spell:
            return True
        return low in self.spell

    @staticmethod
    def _sentence_for(text: str, span: str) -> str:
        for s in re.split(r"(?<=[.!?])\s+", text.strip()):
            if span in s:
                return s.strip()[:400]
        return text.strip()[:400]

    def inspect_text(
        self,
        url: str,
        content_type: str,
        priority: str,
        text: str,
        html_el: str,
        css_sel: str,
        attr_name: str,
    ) -> List[Dict]:
        if not text or not text.strip():
            return []

        results: List[Dict] = []
        seen: Set[Tuple[str, str, str]] = set()

        # 1. Unicode Lookalikes & PDF Ligatures
        for m in re.finditer(r"\S+", text):
            tok = m.group(0)
            for i, ch in enumerate(tok):
                cp = f"U+{ord(ch):04X}"
                if self.unicode_enabled and ch in UNICODE_LOOKALIKE_MAP:
                    lat, ref_m, desc = UNICODE_LOOKALIKE_MAP[ch]
                    sug = tok[:i] + lat + tok[i + 1 :]
                    k = (tok, cp, css_sel)
                    if k not in seen:
                        seen.add(k)
                        results.append({
                            "url": url, "content_type": content_type, "audit_priority": priority,
                            "orig_text": text[:280], "suspicious_word": tok,
                            "orig_char": f"{ch} ({unicodedata.name(ch, 'UNKNOWN')})",
                            "suspected_char": lat, "suggested_correction": sug,
                            "category": "Unicode lookalikes", "ref_mapping": ref_m,
                            "ref_severity": "High", "actual_severity": "High",
                            "unicode_cp": cp, "html_element": html_el, "css_selector": css_sel,
                            "attr_name": attr_name, "sentence": self._sentence_for(text, tok),
                            "confidence": 0.99, "classification": "Confirmed Typo",
                            "reason": f"{desc} inside '{tok}'.",
                        })
                elif ch in UNICODE_LIGATURE_MAP:
                    asc, ref_m = UNICODE_LIGATURE_MAP[ch]
                    sug = tok.replace(ch, asc)
                    k = (tok, cp, css_sel)
                    if k not in seen:
                        seen.add(k)
                        results.append({
                            "url": url, "content_type": content_type, "audit_priority": priority,
                            "orig_text": text[:280], "suspicious_word": tok,
                            "orig_char": f"{ch} ({cp})", "suspected_char": asc,
                            "suggested_correction": sug, "category": "Multi-character",
                            "ref_mapping": ref_m, "ref_severity": "Medium", "actual_severity": "Medium",
                            "unicode_cp": cp, "html_element": html_el, "css_selector": css_sel,
                            "attr_name": attr_name, "sentence": self._sentence_for(text, tok),
                            "confidence": 0.94, "classification": "Confirmed Typo",
                            "reason": f"PDF Unicode ligature '{ch}' ({cp}) embedded in '{tok}'.",
                        })

        # 2. Explicit 'Al' -> 'AI' OCR substitution check (catches 'Al-Infused', 'all Al agents', 'Al decisions')
        for m in re.finditer(r"\bAl\b", text):
            k = ("Al", "I -> l (CCM-001)", css_sel)
            if k not in seen:
                seen.add(k)
                results.append({
                    "url": url, "content_type": content_type, "audit_priority": priority,
                    "orig_text": text[:280], "suspicious_word": "Al",
                    "orig_char": "l (U+006C)", "suspected_char": "I (U+0049)",
                    "suggested_correction": "AI", "category": "Alphabetic",
                    "ref_mapping": "I -> l (CCM-001)", "ref_severity": "High", "actual_severity": "High",
                    "unicode_cp": "U+006C (expected U+0049)", "html_element": html_el, "css_selector": css_sel,
                    "attr_name": attr_name, "sentence": self._sentence_for(text, "Al"),
                    "confidence": 0.98, "classification": "Confirmed Typo",
                    "reason": "Lowercase 'l' substituted for uppercase 'I' in 'AI' (renders as 'Al').",
                })

        # 3. Token-level Word & Number Checks
        for wm in re.finditer(r"\b[A-Za-z0-9|]+\b", text):
            tok = wm.group(0)
            if tok.lower() in self.whitelist_lower or tok == "Al":
                continue
            if re.match(r"^\d+(st|nd|rd|th|x|X|D|G|K|M|B|ms|s|hr|hrs|px)?$", tok):
                continue

            # 3A: Digit inside alphabetic word (e.g. 5ystem, 1nvoice)
            if re.search(r"[A-Za-z]", tok) and re.search(r"[0-9]", tok) and not tok[0].isdigit():
                for i, ch in enumerate(tok):
                    if ch in self.active_digit_to_letter:
                        for repl, map_lbl, ref_sev in self.active_digit_to_letter[ch]:
                            cand = tok[:i] + repl + tok[i + 1 :]
                            if self.is_valid_word(cand):
                                k = (tok, map_lbl, css_sel)
                                if k not in seen:
                                    seen.add(k)
                                    results.append({
                                        "url": url, "content_type": content_type, "audit_priority": priority,
                                        "orig_text": text[:280], "suspicious_word": tok,
                                        "orig_char": ch, "suspected_char": repl,
                                        "suggested_correction": cand, "category": "Numeric",
                                        "ref_mapping": map_lbl, "ref_severity": ref_sev, "actual_severity": "High",
                                        "unicode_cp": f"U+{ord(ch):04X}", "html_element": html_el, "css_selector": css_sel,
                                        "attr_name": attr_name, "sentence": self._sentence_for(text, tok),
                                        "confidence": 0.95, "classification": "Confirmed Typo",
                                        "reason": f"Digit '{ch}' embedded in '{tok}'; matches valid word '{cand}'.",
                                    })
                                break

            # 3B: Letter inside numeric token starting with digits (e.g. 2O26, 1OO%)
            elif re.match(r"^[0-9]+[A-Za-z][0-9]*$", tok):
                for i, ch in enumerate(tok):
                    if ch in self.active_letter_to_digit:
                        digit, map_lbl, ref_sev = self.active_letter_to_digit[ch][0]
                        cand = tok[:i] + digit + tok[i + 1 :]
                        k = (tok, map_lbl, css_sel)
                        if k not in seen:
                            seen.add(k)
                            results.append({
                                "url": url, "content_type": content_type, "audit_priority": priority,
                                "orig_text": text[:280], "suspicious_word": tok,
                                "orig_char": ch, "suspected_char": digit,
                                "suggested_correction": cand, "category": "Numeric",
                                "ref_mapping": map_lbl, "ref_severity": ref_sev, "actual_severity": "High",
                                "unicode_cp": f"U+{ord(ch):04X}", "html_element": html_el, "css_selector": css_sel,
                                "attr_name": attr_name, "sentence": self._sentence_for(text, tok),
                                "confidence": 0.93, "classification": "Confirmed Typo",
                                "reason": f"Letter '{ch}' inside numeric token '{tok}' (expected '{cand}').",
                            })
                        break

            # 3C: Pure alphabetic word failing dictionary check (skip all-caps acronyms)
            elif tok.isalpha() and not tok.isupper() and not self.is_valid_word(tok):
                # Multi-character check
                matched_mc = False
                for obs, intended, map_lbl, ref_sev in self.active_multichar_pairs:
                    if obs in tok.lower():
                        for m in re.finditer(re.escape(obs), tok, flags=re.IGNORECASE):
                            si, ei = m.span()
                            repl = intended.upper() if tok[si:ei].isupper() else intended
                            cand = tok[:si] + repl + tok[ei:]
                            if cand.lower() != tok.lower() and self.is_valid_word(cand):
                                k = (tok, map_lbl, css_sel)
                                if k not in seen:
                                    seen.add(k)
                                    results.append({
                                        "url": url, "content_type": content_type, "audit_priority": priority,
                                        "orig_text": text[:280], "suspicious_word": tok,
                                        "orig_char": tok[si:ei], "suspected_char": repl,
                                        "suggested_correction": cand, "category": "Multi-character",
                                        "ref_mapping": map_lbl, "ref_severity": ref_sev, "actual_severity": "High",
                                        "unicode_cp": "N/A (Multi-char)", "html_element": html_el, "css_selector": css_sel,
                                        "attr_name": attr_name, "sentence": self._sentence_for(text, tok),
                                        "confidence": 0.91, "classification": "Confirmed Typo",
                                        "reason": f"Unknown word '{tok}' resolves to '{cand}' via {map_lbl}.",
                                    })
                                matched_mc = True
                                break
                    if matched_mc:
                        break

                if not matched_mc:
                    for obs_ch, int_ch, map_lbl, ref_sev in self.active_alpha_pairs:
                        if obs_ch in tok:
                            for i, ch in enumerate(tok):
                                if ch == obs_ch:
                                    cand = tok[:i] + int_ch + tok[i + 1 :]
                                    if cand.lower() == tok.lower():
                                        continue
                                    if self.is_valid_word(cand):
                                        k = (tok, map_lbl, css_sel)
                                        if k not in seen:
                                            seen.add(k)
                                            results.append({
                                                "url": url, "content_type": content_type, "audit_priority": priority,
                                                "orig_text": text[:280], "suspicious_word": tok,
                                                "orig_char": obs_ch, "suspected_char": int_ch,
                                                "suggested_correction": cand, "category": "Alphabetic",
                                                "ref_mapping": map_lbl, "ref_severity": ref_sev, "actual_severity": "Medium",
                                                "unicode_cp": f"U+{ord(obs_ch):04X}", "html_element": html_el, "css_selector": css_sel,
                                                "attr_name": attr_name, "sentence": self._sentence_for(text, tok),
                                                "confidence": 0.85, "classification": "Needs Manual Review",
                                                "reason": f"Unrecognized word '{tok}' matches '{cand}' via {map_lbl}.",
                                            })
                                        break

        # 4. Punctuation & Spacing Anomalies
        if self.punct_enabled:
            punct_checks = [
                (r"\b[a-z]+\s+:", " :", ":", "systems: (no space before colon)", ": -> ; / spacing (CCM-142)", "Low", "Confirmed Typo", "Erroneous space before colon."),
                (r"\b[A-Za-z]+-\s+[a-z]+\b", "- ", "-", "hyphenated compound", "- -> – / stray space after hyphen (CCM-148)", "Low", "Confirmed Typo", "Erroneous space inserted after hyphen."),
                (r"\b\d+,\d+%\b", ",", ".", "decimal percentage", ". -> , (CCM-141)", "Medium", "Needs Manual Review", "Comma used in decimal percentage instead of period."),
                (r"\b[A-Za-z]*\|[A-Za-z]+\b|\b[A-Za-z]+\|[A-Za-z]*\b", "|", "I / l", "I / l replacement", "I / l -> | (CCM-185)", "High", "Confirmed Typo", "Vertical bar '|' embedded inside alphabetic word."),
            ]
            for pat, det, sug_hint, _, map_lbl, sev, cls, rsn in punct_checks:
                for m in re.finditer(pat, text):
                    tok = m.group(0)
                    k = (tok, map_lbl, css_sel)
                    if k not in seen:
                        seen.add(k)
                        results.append({
                            "url": url, "content_type": content_type, "audit_priority": priority,
                            "orig_text": text[:280], "suspicious_word": tok,
                            "orig_char": det, "suspected_char": sug_hint,
                            "suggested_correction": tok.replace(" :", ":").replace("- ", "-"),
                            "category": "Special characters & punctuation",
                            "ref_mapping": map_lbl, "ref_severity": "Medium", "actual_severity": sev,
                            "unicode_cp": f"U+{ord(det.strip()[0]):04X}", "html_element": html_el, "css_selector": css_sel,
                            "attr_name": attr_name, "sentence": self._sentence_for(text, tok),
                            "confidence": 0.88, "classification": cls, "reason": rsn,
                        })

        return results


# ==============================================================================
# 3. PLAYWRIGHT LIVE SCANNER (CACHE-BYPASS + FULL DOM + META + ATTRS)
# ==============================================================================
def extract_dom_items(html_str: str) -> List[Tuple[str, str, str, str]]:
    """Returns list of (html_element, css_selector, attr_name, text_value)."""
    soup = BeautifulSoup(html_str, "lxml")
    items: List[Tuple[str, str, str, str]] = []
    seen_local: Set[Tuple[str, str]] = set()

    # SEO Metadata
    if soup.title and soup.title.string:
        items.append(("title", "head > title", "innerText", soup.title.string.strip()))
    for meta in soup.find_all("meta"):
        m_name = meta.get("name") or meta.get("property") or ""
        m_content = meta.get("content") or ""
        if m_name.lower() in {"description", "og:title", "og:description", "twitter:title", "twitter:description"} and m_content.strip():
            items.append(("meta", f"meta[name='{m_name}']", "content", m_content.strip()))

    # Strip non-content scripts/styles
    for tag in soup.find_all(["script", "style", "noscript", "template"]):
        tag.decompose()

    def _sel(el) -> str:
        parts, curr, d = [], el, 0
        while curr and getattr(curr, "name", None) and curr.name not in {"[document]", "html", "body"} and d < 4:
            ident = curr.name
            if curr.get("id"):
                parts.append(f"{ident}#{curr.get('id')}")
                break
            cls = [c for c in (curr.get("class") or []) if not c.startswith("elementor-element-")][:2]
            if cls:
                ident += "." + ".".join(cls)
            parts.append(ident)
            curr = curr.parent
            d += 1
        return " > ".join(reversed(parts)) or el.name

    for el in soup.find_all(["h1", "h2", "h3", "h4", "h5", "h6", "p", "li", "td", "th", "blockquote", "figcaption", "caption", "label", "button", "a", "span"]):
        if el.find(["h1", "h2", "h3", "h4", "h5", "h6", "p", "li", "td", "th", "div", "section", "article"]):
            continue
        txt = " ".join(el.get_text(separator=" ", strip=True).split())
        if len(txt) >= 2:
            s = _sel(el)
            if (s, txt) not in seen_local:
                seen_local.add((s, txt))
                items.append((el.name, s, "innerText", txt))

    for img in soup.find_all("img"):
        for attr in ("alt", "title"):
            val = " ".join(str(img.get(attr, "")).split())
            if len(val) >= 2:
                s = _sel(img)
                if (s, val) not in seen_local:
                    seen_local.add((s, val))
                    items.append(("img", s, attr, val))

    for el in soup.find_all(attrs={"aria-label": True}):
        val = " ".join(str(el.get("aria-label", "")).split())
        if len(val) >= 2:
            s = _sel(el)
            items.append((el.name, s, "aria-label", val))

    for inp in soup.find_all(["input", "textarea"], placeholder=True):
        val = " ".join(str(inp.get("placeholder", "")).split())
        if len(val) >= 2:
            s = _sel(inp)
            items.append((inp.name, s, "placeholder", val))

    return items


async def crawl_urls_async(url_records: List[Dict], engine: DynamicConfusionEngine, progress_cb=None) -> Tuple[List[Dict], List[Dict]]:
    if async_playwright is None:
        raise RuntimeError("Playwright is not installed. Run: pip install playwright && playwright install chromium")

    all_findings: List[Dict] = []
    url_summaries: List[Dict] = []

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=True)
        for idx, rec in enumerate(url_records, start=1):
            u = rec["url"]
            ctype = rec["content_type"]
            prio = rec["priority"]
            if progress_cb:
                progress_cb(idx, len(url_records), u)

            context = await browser.new_context(
                viewport={"width": 1440, "height": 900},
                ignore_https_errors=True,
                extra_http_headers={"Cache-Control": "no-cache, no-store, must-revalidate", "Pragma": "no-cache"},
            )
            page = await context.new_page()
            try:
                resp = await page.goto(u, timeout=45000, wait_until="networkidle")
                await page.wait_for_timeout(1200)
                status = resp.status if resp else 0
                if status >= 400:
                    url_summaries.append({**rec, "scan_status": "Failed", "error": f"HTTP {status}"})
                else:
                    html_content = await page.content()
                    dom_items = extract_dom_items(html_content)
                    for html_el, css_sel, attr_name, txt in dom_items:
                        hits = engine.inspect_text(u, ctype, prio, txt, html_el, css_sel, attr_name)
                        all_findings.extend(hits)
                    url_summaries.append({**rec, "scan_status": "Completed", "error": ""})
            except Exception as exc:
                url_summaries.append({**rec, "scan_status": "Failed", "error": str(exc)})
            finally:
                await context.close()
        await browser.close()

    return url_summaries, all_findings


# ==============================================================================
# 4. 3-SHEET EXCEL GENERATOR (SUPPORTS FRESH & APPEND MODES)
# ==============================================================================
HEADER_FILL = PatternFill("solid", fgColor="1B365D")
ZEBRA_FILL = PatternFill("solid", fgColor="F7FAFC")
TOTAL_FILL = PatternFill("solid", fgColor="EDF2F7")
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
TITLE_FONT = Font(name="Calibri", size=15, bold=True, color="1B365D")
SECTION_FONT = Font(name="Calibri", size=12, bold=True, color="1B365D")
BODY_FONT = Font(name="Calibri", size=10, color="1A202C")
BOLD_FONT = Font(name="Calibri", size=10, bold=True, color="1A202C")
LINK_FONT = Font(name="Calibri", size=10, underline="single", color="1E40AF")
THIN_GRAY = Side(style="thin", color="CBD5E0")
MED_NAVY = Side(style="medium", color="1B365D")
CELL_BORDER = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY)
TOTAL_BORDER = Border(left=THIN_GRAY, right=THIN_GRAY, top=MED_NAVY, bottom=MED_NAVY)


def _style_tbl(ws, r1: int, r2: int, c1: int, c2: int, freeze: bool = False, total_row: bool = False):
    for c in range(c1, c2 + 1):
        cell = ws.cell(row=r1, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        cell.border = CELL_BORDER
    ws.row_dimensions[r1].height = 28
    for r in range(r1 + 1, r2 + 1):
        is_tot = total_row and (r == r2)
        fill = TOTAL_FILL if is_tot else (ZEBRA_FILL if (r - r1) % 2 == 0 else PatternFill(fill_type=None))
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            if fill.fill_type:
                cell.fill = fill
            cell.font = BOLD_FONT if is_tot else BODY_FONT
            cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            cell.border = TOTAL_BORDER if is_tot else CELL_BORDER
    if freeze:
        ws.freeze_panes = ws.cell(row=r1 + 1, column=c1)
        ws.auto_filter.ref = f"{get_column_letter(c1)}{r1}:{get_column_letter(c2)}{r2}"


def _apply_cf(ws, col_let: str, r1: int, r2: int):
    if r2 < r1:
        return
    rng = f"{col_let}{r1}:{col_let}{r2}"
    for label, bg, fg in [
        ('"Critical"', "FED7D7", "9B2C2C"),
        ('"High"', "FEEBC8", "9C4221"),
        ('"Medium"', "FEFCBF", "975A16"),
        ('"Low"', "E2E8F0", "2D3748"),
    ]:
        ws.conditional_formatting.add(
            rng,
            CellIsRule(operator="equal", formula=[label], fill=PatternFill("solid", fgColor=bg), font=Font(name="Calibri", size=10, bold=True, color=fg)),
        )


def build_excel_report(
    url_records: List[Dict],
    findings: List[Dict],
    rules_df: pd.DataFrame,
    append_mode: bool = False,
    existing_excel_bytes: Optional[bytes] = None,
) -> bytes:
    # If Append Mode: merge previous Master URL Report & All Issue Findings with new scan results
    if append_mode:
        prev_buf = None
        if existing_excel_bytes:
            prev_buf = io.BytesIO(existing_excel_bytes)
        elif MASTER_EXCEL_PATH.exists():
            prev_buf = MASTER_EXCEL_PATH

        if prev_buf is not None:
            try:
                old_urls_df = pd.read_excel(prev_buf, sheet_name="Master URL Report")
                old_urls_df = old_urls_df[old_urls_df["S.No"].astype(str) != "TOTAL"]
                if prev_buf is not MASTER_EXCEL_PATH:
                    prev_buf.seek(0)
                old_find_df = pd.read_excel(prev_buf, sheet_name="All Issue Findings")

                # Merge URLs (replace status if URL was re-scanned, keep older URLs intact)
                new_url_set = {r["url"] for r in url_records}
                merged_urls = []
                for _, r in old_urls_df.iterrows():
                    if r["URL"] not in new_url_set:
                        merged_urls.append({
                            "url": r["URL"],
                            "content_type": r["Content Type"],
                            "priority": r["Audit Priority"],
                            "scan_status": r["Scan Status"],
                            "last_mod": r.get("Last Modified", ""),
                        })
                merged_urls.extend(url_records)
                url_records = merged_urls

                # Merge Findings (drop old findings for re-scanned URLs, keep other URLs, deduplicate)
                kept_old_findings = []
                for _, r in old_find_df.iterrows():
                    if r["URL"] not in new_url_set:
                        kept_old_findings.append({
                            "url": r["URL"],
                            "content_type": r["Content Type"],
                            "priority": r["Audit Priority"],
                            "orig_text": r["Original Text"],
                            "suspicious_word": r["Suspicious Word"],
                            "orig_char": r["Original Character"],
                            "suspected_char": r["Suspected Replacement Character"],
                            "suggested_correction": r["Suggested Correction"],
                            "category": r["Character Confusion Category"],
                            "ref_mapping": r["Reference Mapping"],
                            "ref_severity": r["Reference Severity"],
                            "actual_severity": r["Actual Issue Severity"],
                            "unicode_cp": r["Unicode Code Point"],
                            "html_element": r["HTML Element"],
                            "css_selector": r["CSS Selector"],
                            "attr_name": r["Attribute Name"],
                            "sentence": r["Surrounding Sentence"],
                            "confidence": r["Confidence Score"],
                            "classification": r["Classification"],
                            "reason": r["Reason for Flagging"],
                        })
                findings = kept_old_findings + findings
            except Exception:
                pass

    # Assign clean sequential Issue IDs
    for i, f in enumerate(findings, start=1):
        f["issue_id"] = f"TYP-{i:05d}"

    wb = Workbook()

    # --- Sheet 3: All Issue Findings ---
    ws3 = wb.active
    ws3.title = "All Issue Findings"
    s3_headers = [
        "Issue ID", "URL", "Content Type", "Audit Priority", "Original Text",
        "Suspicious Word", "Original Character", "Suspected Replacement Character",
        "Suggested Correction", "Character Confusion Category", "Reference Mapping",
        "Reference Severity", "Actual Issue Severity", "Unicode Code Point",
        "HTML Element", "CSS Selector", "Attribute Name", "Surrounding Sentence",
        "Confidence Score", "Classification", "Reason for Flagging",
    ]
    for ci, h in enumerate(s3_headers, start=1):
        ws3.cell(row=1, column=ci, value=h)

    for ro, item in enumerate(findings, start=2):
        vals = [
            item["issue_id"], item["url"], item["content_type"], item.get("audit_priority", item.get("priority", "High")),
            item["orig_text"], item["suspicious_word"], item["orig_char"],
            item["suspected_char"], item["suggested_correction"], item["category"],
            item["ref_mapping"], item["ref_severity"], item["actual_severity"],
            item["unicode_cp"], item["html_element"], item["css_selector"],
            item["attr_name"], item["sentence"], item["confidence"],
            item["classification"], item["reason"],
        ]
        for ci, v in enumerate(vals, start=1):
            c = ws3.cell(row=ro, column=ci, value=v)
            if ci == 2:
                c.hyperlink = v
                c.font = LINK_FONT
            elif ci == 19:
                c.number_format = "0.00"

    last_f = max(2, 1 + len(findings))
    _style_tbl(ws3, 1, last_f, 1, len(s3_headers), freeze=True)
    if findings:
        _apply_cf(ws3, "L", 2, last_f)
        _apply_cf(ws3, "M", 2, last_f)
    for ci, w in {1: 13, 2: 44, 3: 18, 4: 14, 5: 42, 6: 22, 7: 16, 8: 20, 9: 26, 10: 26, 11: 26, 12: 16, 13: 18, 14: 20, 15: 15, 16: 34, 17: 14, 18: 42, 19: 15, 20: 20, 21: 46}.items():
        ws3.column_dimensions[get_column_letter(ci)].width = w

    # --- Sheet 2: Master URL Report ---
    ws2 = wb.create_sheet("Master URL Report")
    s2_headers = [
        "S.No", "URL", "Sitemap Source", "Content Type", "Last Modified",
        "Audit Priority", "QE Status", "Scan Status", "Total Findings",
        "Confirmed Typos", "Manual Review", "False Positives",
        "Critical", "High", "Medium", "Low", "Overall Result",
    ]
    for ci, h in enumerate(s2_headers, start=1):
        ws2.cell(row=1, column=ci, value=h)

    f_end = 1 + len(findings)
    for idx, rec in enumerate(url_records, start=2):
        u = rec["url"]
        ws2.cell(row=idx, column=1, value=idx - 1)
        cu = ws2.cell(row=idx, column=2, value=u)
        cu.hyperlink = u
        cu.font = LINK_FONT
        ws2.cell(row=idx, column=3, value=rec.get("sitemap", "Manual / Custom Input"))
        ws2.cell(row=idx, column=4, value=rec["content_type"])
        ws2.cell(row=idx, column=5, value=rec.get("last_mod", datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")))
        ws2.cell(row=idx, column=6, value=rec["priority"])
        ws2.cell(row=idx, column=7, value="Audited")
        ws2.cell(row=idx, column=8, value=rec.get("scan_status", "Completed"))

        if len(findings) > 0:
            ws2.cell(row=idx, column=9, value=f'=COUNTIF(\'All Issue Findings\'!\$B\$2:\$B\${f_end}, B{idx})')
            ws2.cell(row=idx, column=10, value=f'=COUNTIFS(\'All Issue Findings\'!\$B\$2:\$B\${f_end}, B{idx}, \'All Issue Findings\'!\$T\$2:\$T\${f_end}, "Confirmed Typo")')
            ws2.cell(row=idx, column=11, value=f'=COUNTIFS(\'All Issue Findings\'!\$B\$2:\$B\${f_end}, B{idx}, \'All Issue Findings\'!\$T\$2:\$T\${f_end}, "Needs Manual Review")')
            ws2.cell(row=idx, column=12, value=f'=COUNTIFS(\'All Issue Findings\'!\$B\$2:\$B\${f_end}, B{idx}, \'All Issue Findings\'!\$T\$2:\$T\${f_end}, "False Positive")')
            ws2.cell(row=idx, column=13, value=f'=COUNTIFS(\'All Issue Findings\'!\$B\$2:\$B\${f_end}, B{idx}, \'All Issue Findings\'!\$M\$2:\$M\${f_end}, "Critical")')
            ws2.cell(row=idx, column=14, value=f'=COUNTIFS(\'All Issue Findings\'!\$B\$2:\$B\${f_end}, B{idx}, \'All Issue Findings\'!\$M\$2:\$M\${f_end}, "High")')
            ws2.cell(row=idx, column=15, value=f'=COUNTIFS(\'All Issue Findings\'!\$B\$2:\$B\${f_end}, B{idx}, \'All Issue Findings\'!\$M\$2:\$M\${f_end}, "Medium")')
            ws2.cell(row=idx, column=16, value=f'=COUNTIFS(\'All Issue Findings\'!\$B\$2:\$B\${f_end}, B{idx}, \'All Issue Findings\'!\$M\$2:\$M\${f_end}, "Low")')
        else:
            for col_zero in range(9, 17):
                ws2.cell(row=idx, column=col_zero, value=0)
        ws2.cell(row=idx, column=17, value=f'=IF(H{idx}="Failed","Scan Failed",IF(J{idx}>0,"Action Required",IF(K{idx}>0,"Needs Manual Review","Clean (0 Issues)")))')

    tot_r = 2 + len(url_records)
    ws2.cell(row=tot_r, column=1, value="TOTAL")
    ws2.cell(row=tot_r, column=2, value=f"{len(url_records)} Audited URLs")
    for ci in range(9, 17):
        cl = get_column_letter(ci)
        ws2.cell(row=tot_r, column=ci, value=f"=SUM({cl}2:{cl}{tot_r-1})")
    ws2.cell(row=tot_r, column=17, value="Complete")
    _style_tbl(ws2, 1, tot_r, 1, len(s2_headers), freeze=True, total_row=True)
    for ci, w in {1: 8, 2: 50, 3: 24, 4: 18, 5: 22, 6: 14, 7: 14, 8: 16, 9: 15, 10: 16, 11: 15, 12: 15, 13: 11, 14: 11, 15: 11, 16: 11, 17: 24}.items():
        ws2.column_dimensions[get_column_letter(ci)].width = w

    # --- Sheet 1: Executive Summary ---
    ws1 = wb.create_sheet("Executive Summary")
    wb._sheets = [ws1, ws2, ws3]

    ws1["A1"] = "Zuci Systems — Website Typo & Character Confusion Audit Report"
    ws1["A1"].font = TITLE_FONT
    ws1["A2"] = f"Generated: {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC')} | Mode: {'Append / Cumulative' if append_mode else 'Fresh Scan'}"
    ws1["A2"].font = Font(name="Calibri", size=10, italic=True, color="4A5568")

    ws1["A4"] = "Overall Audit KPIs"
    ws1["A4"].font = SECTION_FONT
    ws1["A5"] = "Audit Metric"
    ws1["B5"] = "Value"

    u_end = tot_r - 1
    kpis = [
        ("Total URLs", f"=COUNTA('Master URL Report'!B2:B{u_end})"),
        ("Successfully scanned URLs", f'=COUNTIF(\'Master URL Report\'!H2:H{u_end},"Completed")'),
        ("Partially scanned URLs", f'=COUNTIF(\'Master URL Report\'!H2:H{u_end},"Partially Completed")'),
        ("Failed URLs", f'=COUNTIF(\'Master URL Report\'!H2:H{u_end},"Failed")'),
        ("URLs containing findings", f"=COUNTIF('Master URL Report'!I2:I{u_end},\">0\")"),
        ("URLs without detected issues", f"=COUNTIF('Master URL Report'!I2:I{u_end},0)"),
        ("Total findings", len(findings)),
        ("Confirmed typos", f'=COUNTIF(\'All Issue Findings\'!T2:T{last_f},"Confirmed Typo")' if findings else 0),
        ("Needs Manual Review", f'=COUNTIF(\'All Issue Findings\'!T2:T{last_f},"Needs Manual Review")' if findings else 0),
        ("False Positives", f'=COUNTIF(\'All Issue Findings\'!T2:T{last_f},"False Positive")' if findings else 0),
    ]
    for idx, (lbl, val) in enumerate(kpis, start=6):
        ws1.cell(row=idx, column=1, value=lbl)
        ws1.cell(row=idx, column=2, value=val)
    _style_tbl(ws1, 5, 15, 1, 2)

    # Category Table
    ws1["A17"] = "Findings by Character Confusion Category"
    ws1["A17"].font = SECTION_FONT
    ws1["A18"] = "Character Confusion Category"
    ws1["B18"] = "Active Rules"
    ws1["C18"] = "Total Findings"

    cats = ["Alphabetic", "Numeric", "Multi-character", "Special characters & punctuation", "Unicode lookalikes"]
    active_df = rules_df[rules_df["Enabled"] == True]
    for ro, cname in enumerate(cats, start=19):
        rcnt = int((active_df["Category"] == cname).sum())
        ws1.cell(row=ro, column=1, value=cname)
        ws1.cell(row=ro, column=2, value=rcnt)
        ws1.cell(row=ro, column=3, value=f'=COUNTIF(\'All Issue Findings\'!\$J\$2:\$J\${last_f}, A{ro})' if findings else 0)
    ws1.cell(row=24, column=1, value="Total")
    ws1.cell(row=24, column=2, value="=SUM(B19:B23)")
    ws1.cell(row=24, column=3, value="=SUM(C19:C23)")
    _style_tbl(ws1, 18, 24, 1, 3, total_row=True)

    # Severity Table
    ws1["D4"] = "Findings by Severity"
    ws1["D4"].font = SECTION_FONT
    ws1["D5"] = "Severity"
    ws1["E5"] = "Total Findings"
    ws1["F5"] = "Confirmed Typos"
    for ro, sev in enumerate(["Critical", "High", "Medium", "Low"], start=6):
        ws1.cell(row=ro, column=4, value=sev)
        ws1.cell(row=ro, column=5, value=f'=COUNTIF(\'All Issue Findings\'!\$M\$2:\$M\${last_f}, D{ro})' if findings else 0)
        ws1.cell(row=ro, column=6, value=f'=COUNTIFS(\'All Issue Findings\'!\$M\$2:\$M\${last_f}, D{ro}, \'All Issue Findings\'!\$T\$2:\$T\${last_f}, "Confirmed Typo")' if findings else 0)
    ws1.cell(row=10, column=4, value="Total")
    ws1.cell(row=10, column=5, value="=SUM(E6:E9)")
    ws1.cell(row=10, column=6, value="=SUM(F6:F9)")
    _style_tbl(ws1, 5, 10, 4, 6, total_row=True)

    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "Findings by Character Confusion Category"
    chart.y_axis.title = "Findings"
    chart.add_data(Reference(ws1, min_col=3, min_row=18, max_row=23), titles_from_data=True)
    chart.set_categories(Reference(ws1, min_col=1, min_row=19, max_row=23))
    chart.legend = None
    chart.width = 15
    chart.height = 8.5
    ws1.add_chart(chart, "E12")

    for ci, w in {1: 34, 2: 18, 3: 16, 4: 22, 5: 16, 6: 18}.items():
        ws1.column_dimensions[get_column_letter(ci)].width = w

    out_bytes = io.BytesIO()
    wb.save(out_bytes)
    data = out_bytes.getvalue()
    MASTER_EXCEL_PATH.write_bytes(data)
    return data


# ==============================================================================
# 5. STREAMLIT WEB APPLICATION UI
# ==============================================================================

@st.cache_resource
def ensure_playwright_browser() -> None:
    """Ensures headless Chromium is installed on Streamlit Community Cloud."""
    try:
        subprocess.run(
            [sys.executable, "-m", "playwright", "install", "chromium"],
            check=False,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
    except Exception:
        pass


def main():
    st.set_page_config(
        page_title="Zuci Systems — Typo & Character Confusion Audit Studio",
        page_icon="🔍",
        layout="wide",
    )

    ensure_playwright_browser()
    st.title("🔍 Zuci Systems — Website Typo & Character Confusion Audit Studio")
    st.caption("Live Playwright DOM Scanner · Editable 205-Rule Confusion Matrix · Fresh & Append Excel Reporting")

    rules_df = load_rules_df()
    whitelist_terms = load_whitelist()

    tab_scan, tab_rules, tab_whitelist = st.tabs([
        "🚀 1. Run URL Audit & Download Report",
        f"📋 2. Character Confusion Rules ({int((rules_df['Enabled'] == True).sum())}/{len(rules_df)} Active)",
        f"🛡️ 3. Brand & Tech Whitelist ({len(whitelist_terms)} Terms)",
    ])

    # --------------------------------------------------------------------------
    # TAB 1: URL INPUT, FRESH VS APPEND MODE, LIVE SCAN & 1-CLICK EXCEL DOWNLOAD
    # --------------------------------------------------------------------------
    with tab_scan:
        col_left, col_right = st.columns([2, 1])

        with col_left:
            urls_raw = st.text_area(
                "Enter Single or Multiple URLs (one URL per line)",
                value="https://www.zucisystems.com/\nhttps://www.zucisystems.com/truzt-ai/\nhttps://www.zucisystems.com/services/artificial-intelligence/qe-for-ai/",
                height=160,
                help="Paste any number of URLs, one per line.",
            )
            uploaded_url_file = st.file_uploader(
                "Or upload a URL list (.txt or .csv with a 'URL' column)",
                type=["txt", "csv"],
            )

        with col_right:
            report_mode = st.radio(
                "Report Generation Mode",
                options=["Fresh Report (New Excel)", "Append to Previous Master Report"],
                index=0,
                help="Fresh creates a standalone report for these URLs. Append merges results into your cumulative Master report.",
            )
            append_file = None
            if report_mode.startswith("Append"):
                st.info(
                    f"Local cumulative file: `{MASTER_EXCEL_PATH}` "
                    f"({'Found' if MASTER_EXCEL_PATH.exists() else 'Will be created'})"
                )
                append_file = st.file_uploader(
                    "Optional: Upload an existing .xlsx report to append into",
                    type=["xlsx"],
                )

            default_ctype = st.selectbox("Default Content Type Tag", ["Service Page", "Product Page", "Homepage", "Case Study", "Blog Post", "Page"])
            default_prio = st.selectbox("Default Audit Priority", ["High", "Critical", "Medium", "Low"])

        run_btn = st.button("▶️ Start Live Typo & Character Audit", type="primary", use_container_width=True)

        if run_btn:
            url_list: List[str] = [ln.strip() for ln in urls_raw.splitlines() if ln.strip().startswith("http")]
            if uploaded_url_file is not None:
                if uploaded_url_file.name.endswith(".csv"):
                    udf = pd.read_csv(uploaded_url_file)
                    url_col = next((c for c in udf.columns if c.strip().lower() == "url"), udf.columns[0])
                    url_list.extend([str(u).strip() for u in udf[url_col].dropna() if str(u).strip().startswith("http")])
                else:
                    txt_lines = uploaded_url_file.read().decode("utf-8", errors="ignore").splitlines()
                    url_list.extend([ln.strip() for ln in txt_lines if ln.strip().startswith("http")])

            # Deduplicate preserving order
            seen_u: Set[str] = set()
            clean_urls: List[str] = []
            for u in url_list:
                if u not in seen_u:
                    seen_u.add(u)
                    clean_urls.append(u)

            if not clean_urls:
                st.error("Please enter at least one valid `http://` or `https://` URL.")
            else:
                records = [
                    {
                        "url": u,
                        "content_type": "Homepage" if u.rstrip("/") == "https://www.zucisystems.com" else default_ctype,
                        "priority": default_prio,
                    }
                    for u in clean_urls
                ]
                engine = DynamicConfusionEngine(rules_df=rules_df, whitelist=whitelist_terms)
                prog_bar = st.progress(0.0, text="Launching headless Chromium...")

                def _on_progress(curr: int, total: int, cur_url: str):
                    prog_bar.progress(curr / total, text=f"Scanning ({curr}/{total}): {cur_url}")

                try:
                    t0 = time.perf_counter()
                    url_sums, found_items = asyncio.run(crawl_urls_async(records, engine, progress_cb=_on_progress))
                    elapsed = round(time.perf_counter() - t0, 2)
                    prog_bar.progress(1.0, text=f"Completed {len(clean_urls)} URLs in {elapsed}s!")

                    is_append = report_mode.startswith("Append")
                    prev_bytes = append_file.getvalue() if append_file is not None else None
                    excel_bytes = build_excel_report(
                        url_records=url_sums,
                        findings=found_items,
                        rules_df=rules_df,
                        append_mode=is_append,
                        existing_excel_bytes=prev_bytes,
                    )

                    st.session_state["last_excel_bytes"] = excel_bytes
                    st.session_state["last_findings_df"] = pd.DataFrame(found_items)
                    st.session_state["last_urls_df"] = pd.DataFrame(url_sums)
                    st.session_state["last_report_name"] = (
                        "Zuci_Systems_Master_Typo_Audit_Report.xlsx"
                        if is_append
                        else f"Zuci_Systems_Typo_Audit_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    )
                except Exception as exc:
                    st.error(f"Scan error: {exc}")

        if "last_excel_bytes" in st.session_state:
            st.divider()
            fdf = st.session_state["last_findings_df"]
            udf = st.session_state["last_urls_df"]

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("URLs Scanned (This Run)", len(udf))
            m2.metric("Findings (This Run)", len(fdf))
            m3.metric("Confirmed Typos", int((fdf["classification"] == "Confirmed Typo").sum()) if not fdf.empty else 0)
            m4.metric("Needs Manual Review", int((fdf["classification"] == "Needs Manual Review").sum()) if not fdf.empty else 0)

            st.download_button(
                label="📥 Download Excel Audit Report (.xlsx)",
                data=st.session_state["last_excel_bytes"],
                file_name=st.session_state["last_report_name"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )

            if not fdf.empty:
                st.subheader("Detected Findings Preview")
                st.dataframe(fdf, use_container_width=True)
            else:
                st.success("Zero character-confusion typos detected on the scanned URLs!")

    # --------------------------------------------------------------------------
    # TAB 2: EDITABLE CHARACTER CONFUSION MATRIX RULES TABLE
    # --------------------------------------------------------------------------
    with tab_rules:
        st.subheader("Editable Character Confusion Reference Rules")
        st.write(
            "Add new rows, edit mappings, change priority, or uncheck **Enabled** to disable any rule. "
            "Changes are saved permanently to `rules_store.json` and used on your next scan."
        )

        cat_filter = st.selectbox(
            "Filter by Category",
            ["All Categories", "Alphabetic", "Numeric", "Multi-character", "Special characters & punctuation", "Unicode lookalikes"],
        )

        display_df = rules_df if cat_filter == "All Categories" else rules_df[rules_df["Category"] == cat_filter].copy()

        edited_df = st.data_editor(
            display_df,
            num_rows="dynamic",
            use_container_width=True,
            height=460,
            column_config={
                "Enabled": st.column_config.CheckboxColumn("Enabled", default=True),
                "Rule_ID": st.column_config.TextColumn("Rule ID", required=True),
                "Category": st.column_config.SelectboxColumn(
                    "Category",
                    options=["Alphabetic", "Numeric", "Multi-character", "Special characters & punctuation", "Unicode lookalikes"],
                    required=True,
                ),
                "Actual Character(s)": st.column_config.TextColumn("Actual Character(s)", required=True),
                "Potential Confusion": st.column_config.TextColumn("Potential Confusion", required=True),
                "Example / Notes": st.column_config.TextColumn("Example / Notes"),
                "Priority": st.column_config.SelectboxColumn("Priority", options=["High", "Medium", "Low"], default="Medium"),
            },
        )

        b1, b2 = st.columns([1, 1])
        with b1:
            if st.button("💾 Save Rules Changes", type="primary", use_container_width=True):
                if cat_filter == "All Categories":
                    updated_full = edited_df.reset_index(drop=True)
                else:
                    other_rows = rules_df[rules_df["Category"] != cat_filter]
                    updated_full = pd.concat([other_rows, edited_df], ignore_index=True)
                save_rules_df(updated_full)
                st.success(f"Saved {len(updated_full)} rules to `rules_store.json`!")
                st.rerun()

        with b2:
            if st.button("🔄 Reset to Default 205 Matrix Rules", use_container_width=True):
                default_df = pd.DataFrame(DEFAULT_RULES)
                default_df["Enabled"] = True
                save_rules_df(default_df)
                st.warning("Reset rules to the original 205 Character Confusion Matrix entries.")
                st.rerun()

    # --------------------------------------------------------------------------
    # TAB 3: EDITABLE BRAND & TECHNICAL WHITELIST
    # --------------------------------------------------------------------------
    with tab_whitelist:
        st.subheader("Brand, Product & Technical Terminology Whitelist")
        st.write(
            "Words in this list (case-insensitive) will never be flagged as spelling or character-confusion errors. "
            "Saved in `whitelist_store.json`."
        )
        wl_text = st.text_area(
            "Whitelisted Terms (one per line or comma-separated)",
            value="\n".join(whitelist_terms),
            height=340,
        )
        if st.button("💾 Save Whitelist", type="primary"):
            raw_items = re.split(r"[,\n]+", wl_text)
            save_whitelist(raw_items)
            st.success("Whitelist updated and saved to `whitelist_store.json`!")
            st.rerun()


if __name__ == "__main__":
    main()
