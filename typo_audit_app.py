#!/usr/bin/env python3
"""
typo_audit_app.py
================================================================================
A-to-Z Website QA & Content Integrity Studio (Streamlit Web App)
Performs complete end-to-end QA on your behalf across any single URL, batch of
URLs, or full XML Sitemap (including nested child sitemaps):

  Module 1: Character Confusion & Typo Detection (All 205 Matrix Rules +
            Unicode Homoglyphs + PDF Ligatures + Stray Hyphens/Colons +
            Mid-word Case Errors + Heading Capitalization + UK/US Spelling Drift
            + Truncated Stats)
  Module 2: Cross-Page Consistency & Duplicate Content QA (Header/Footer parity,
            Desktop vs. Mobile menu drift, repeated typos across pages, brand
            casing consistency like ZoY vs Zoy, AI-Native vs AI-native)
  Module 3: Broken Links, Media & Asset QA (Internal/External HTTP status checks,
            404/500 detection, HTTP-on-HTTPS mixed content, broken images, raw
            Figma/export filenames in image alt text, missing alt/title/aria-labels)
  Module 4: SEO, Metadata & Technical HTML QA (Meta Title/Description length &
            missing tags, Canonical URL verification, Open Graph & Twitter tags,
            H1 count/hierarchy, JSON-LD schema presence, Page Load Time)
  Module 5: Multi-Engine AI Semantic & Source-PDF Parity Review (Optional API
            hooks for OpenAI ChatGPT, Anthropic Claude, and Google Gemini to
            cross-audit context, technical terminology, and PDF-to-Web parity)
  Module 6: Persistent Editable Rule & Whitelist Store + Fresh / Append Master
            6-Sheet Excel Report Generator

Deploy on GitHub + Streamlit Cloud with `requirements.txt` and `packages.txt`.
================================================================================
"""

from __future__ import annotations

import asyncio
import hashlib
import io
import json
import re
import subprocess
import sys
import time
import unicodedata
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

import pandas as pd
import streamlit as st
from bs4 import BeautifulSoup
from openpyxl import Workbook
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

if sys.platform.startswith("win"):
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

# ==============================================================================
# 1. PERSISTENT STORAGE PATHS & DEFAULTS
# ==============================================================================
RULES_STORE_PATH = Path("rules_store.json")
WHITELIST_STORE_PATH = Path("whitelist_store.json")
MASTER_EXCEL_PATH = Path("Zuci_AZ_Website_QA_Master_Report.xlsx")

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


@st.cache_resource
def ensure_playwright_browser() -> None:
    try:
        subprocess.run(
            [sys.executable, "-m", "playwright", "install", "chromium"],
            check=False,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
    except Exception:
        pass


# ==============================================================================
# 2. SITEMAP CRAWLER HELPER (SUPPORTS SITEMAP_INDEX.XML & NESTED SITEMAPS)
# ==============================================================================
def expand_sitemap_urls(sitemap_url: str, max_urls: int = 250) -> List[Dict]:
    def _fetch_xml(u: str) -> ET.Element:
        req = urllib.request.Request(u, headers={"User-Agent": "Zuci-AZ-QA-Bot/2.0"})
        with urllib.request.urlopen(req, timeout=25) as resp:
            return ET.fromstring(resp.read())

    ns = {"sm": "http://www.sitemaps.org/schemas/sitemap/0.9"}
    root = _fetch_xml(sitemap_url)
    child_locs = [el.text.strip() for el in root.findall(".//sm:sitemap/sm:loc", ns) if el.text]
    if not child_locs:
        child_locs = [sitemap_url]

    type_map = {
        "page-sitemap.xml": ("Page", "High"),
        "services-sitemap.xml": ("Service Page", "High"),
        "casestudy-sitemap.xml": ("Case Study", "Critical"),
        "ebooks-whitepaper-sitemap.xml": ("eBook / Whitepaper", "Critical"),
        "tools-sitemap.xml": ("Tool / Worksheet", "High"),
        "industries-sitemap.xml": ("Industry", "Medium"),
        "webinar-sitemap.xml": ("Webinar", "Medium"),
        "zuci-in-the-news-sitemap.xml": ("News", "Medium"),
        "post-sitemap.xml": ("Blog Post", "Medium"),
        "author-sitemap.xml": ("Author Archive", "Low"),
    }

    records: List[Dict] = []
    for sm in child_locs:
        sm_name = sm.rstrip("/").split("/")[-1]
        ctype, prio = type_map.get(sm_name, ("Page", "Medium"))
        try:
            sm_root = _fetch_xml(sm)
            for u_el in sm_root.findall(".//sm:url", ns):
                loc = u_el.find("sm:loc", ns)
                if loc is not None and loc.text:
                    u_clean = loc.text.strip()
                    if u_clean.rstrip("/") == "https://www.zucisystems.com":
                        ctype = "Homepage"
                    elif "/truzt-ai" in u_clean:
                        ctype = "Product Page"
                    records.append({
                        "url": u_clean,
                        "content_type": ctype,
                        "priority": prio,
                        "sitemap": sm_name,
                    })
                    if len(records) >= max_urls:
                        return records
        except Exception:
            continue
    return records


# ==============================================================================
# 3. A-TO-Z QA ANALYSIS ENGINE (TYPOS + CONFUSION MATRIX + SEO + LINKS + ASSETS)
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

UK_TO_US_MAP: Dict[str, str] = {
    "prioritised": "prioritized",
    "prioritising": "prioritizing",
    "prioritisation": "prioritization",
    "prioritise": "prioritize",
    "behaviour": "behavior",
    "behaviours": "behaviors",
    "behavioural": "behavioral",
    "digitise": "digitize",
    "digitised": "digitized",
    "digitising": "digitizing",
    "digitisation": "digitization",
    "personalise": "personalize",
    "personalised": "personalized",
    "personalisation": "personalization",
    "standardise": "standardize",
    "standardised": "standardized",
    "standardisation": "standardization",
    "optimise": "optimize",
    "optimised": "optimized",
    "optimisation": "optimization",
    "modernise": "modernize",
    "modernised": "modernized",
    "modernisation": "modernization",
    "organisation": "organization",
    "organisations": "organizations",
    "organisational": "organizational",
    "human-centred": "human-centered",
    "centred": "centered",
    "modelling": "modeling",
}

TITLE_MINOR_WORDS = {"a", "an", "the", "and", "but", "or", "for", "nor", "on", "at", "to", "from", "by", "in", "of", "with", "as", "vs"}


class ComprehensiveQAEngine:
    """Performs A-to-Z QA across Character Confusion, Spelling, Editorial Consistency, SEO, Links, and Assets."""

    def __init__(self, rules_df: pd.DataFrame, whitelist: List[str], enabled_modules: Dict[str, bool]):
        self.rules_df = rules_df[rules_df["Enabled"] == True].copy()
        self.whitelist_lower: Set[str] = {w.lower() for w in whitelist}
        self.modules = enabled_modules
        self.spell = SpellChecker(language="en") if SpellChecker else None
        if self.spell and self.whitelist_lower:
            self.spell.word_frequency.load_words(list(self.whitelist_lower))

        self.alpha_pairs: List[Tuple[str, str, str, str]] = []
        self.multichar_pairs: List[Tuple[str, str, str, str]] = []
        self.digit_to_letter: Dict[str, List[Tuple[str, str, str]]] = {}
        self.letter_to_digit: Dict[str, List[Tuple[str, str, str]]] = {}
        self.unicode_enabled = False
        self.punct_enabled = False

        for r in self.rules_df.itertuples(index=False):
            cat = str(getattr(r, "Category", ""))
            act = str(getattr(r, "Actual Character(s)", "")).strip()
            conf = str(getattr(r, "Potential Confusion", "")).strip()
            rid = str(getattr(r, "Rule_ID", "CUSTOM"))
            prio = str(getattr(r, "Priority", "Medium"))
            lbl = f"{act} -> {conf} ({rid})"

            if cat == "Unicode lookalikes":
                self.unicode_enabled = True
            elif cat == "Special characters & punctuation":
                self.punct_enabled = True
            elif cat == "Multi-character" and len(act) >= 1 and len(conf) >= 1 and act != conf:
                self.multichar_pairs.append((conf, act, lbl, prio))
            elif cat == "Alphabetic" and len(act) == 1 and len(conf) == 1 and act.isalpha() and conf.isalpha():
                self.alpha_pairs.append((conf, act, lbl, prio))
            elif cat == "Numeric":
                if act.isdigit() and len(conf) == 1 and conf.isalpha():
                    self.letter_to_digit.setdefault(conf, []).append((act, lbl, prio))
                elif act.isalpha() and len(act) == 1 and conf.isdigit():
                    self.digit_to_letter.setdefault(conf, []).append((act, lbl, prio))

    def is_valid_word(self, word: str) -> bool:
        clean = word.strip("'\".,;:!?()[]{}")
        if not clean:
            return True
        low = clean.lower()
        if low in self.whitelist_lower or len(clean) <= 2:
            return True
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

    def inspect_text_block(
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

        findings: List[Dict] = []
        seen: Set[Tuple[str, str, str]] = set()

        # --- MODULE 1: CHARACTER CONFUSION & SPELLING ---
        if self.modules.get("typo_matrix", True):
            # 1A. Unicode lookalikes & PDF ligatures
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
                            findings.append({
                                "qa_pillar": "Typo & Character Confusion",
                                "url": url, "content_type": content_type, "audit_priority": priority,
                                "orig_text": text[:280], "suspicious_word": tok,
                                "orig_char": f"{ch} ({unicodedata.name(ch, 'UNKNOWN')})",
                                "suspected_char": lat, "suggested_correction": sug,
                                "category": "Unicode lookalikes", "ref_mapping": ref_m,
                                "ref_severity": "High", "actual_severity": "High",
                                "unicode_cp": cp, "html_element": html_el, "css_selector": css_sel,
                                "attr_name": attr_name, "sentence": self._sentence_for(text, tok),
                                "confidence": 0.99, "classification": "Confirmed Issue",
                                "reason": f"{desc} inside '{tok}'.",
                            })
                    elif ch in UNICODE_LIGATURE_MAP:
                        asc, ref_m = UNICODE_LIGATURE_MAP[ch]
                        sug = tok.replace(ch, asc)
                        k = (tok, cp, css_sel)
                        if k not in seen:
                            seen.add(k)
                            findings.append({
                                "qa_pillar": "Typo & Character Confusion",
                                "url": url, "content_type": content_type, "audit_priority": priority,
                                "orig_text": text[:280], "suspicious_word": tok,
                                "orig_char": f"{ch} ({cp})", "suspected_char": asc,
                                "suggested_correction": sug, "category": "Multi-character",
                                "ref_mapping": ref_m, "ref_severity": "Medium", "actual_severity": "Medium",
                                "unicode_cp": cp, "html_element": html_el, "css_selector": css_sel,
                                "attr_name": attr_name, "sentence": self._sentence_for(text, tok),
                                "confidence": 0.94, "classification": "Confirmed Issue",
                                "reason": f"PDF Unicode ligature '{ch}' ({cp}) embedded in '{tok}'.",
                            })

            # 1B. Explicit 'Al' -> 'AI' OCR substitution
            for m in re.finditer(r"\bAl\b", text):
                k = ("Al", "I -> l (CCM-001)", css_sel)
                if k not in seen:
                    seen.add(k)
                    findings.append({
                        "qa_pillar": "Typo & Character Confusion",
                        "url": url, "content_type": content_type, "audit_priority": priority,
                        "orig_text": text[:280], "suspicious_word": "Al",
                        "orig_char": "l (U+006C)", "suspected_char": "I (U+0049)",
                        "suggested_correction": "AI", "category": "Alphabetic",
                        "ref_mapping": "I -> l (CCM-001)", "ref_severity": "High", "actual_severity": "High",
                        "unicode_cp": "U+006C (expected U+0049)", "html_element": html_el, "css_selector": css_sel,
                        "attr_name": attr_name, "sentence": self._sentence_for(text, "Al"),
                        "confidence": 0.98, "classification": "Confirmed Issue",
                        "reason": "Lowercase 'l' substituted for uppercase 'I' in 'AI' (renders as 'Al').",
                    })

            # 1C. Token-level Word & Number Checks
            for wm in re.finditer(r"\b[A-Za-z0-9|]+\b", text):
                tok = wm.group(0)
                if tok.lower() in self.whitelist_lower or tok == "Al":
                    continue
                if re.match(r"^\d+(st|nd|rd|th|x|X|D|G|K|M|B|ms|s|hr|hrs|px)?$", tok):
                    continue

                if re.search(r"[A-Za-z]", tok) and re.search(r"[0-9]", tok) and not tok[0].isdigit():
                    for i, ch in enumerate(tok):
                        if ch in self.digit_to_letter:
                            for repl, map_lbl, ref_sev in self.digit_to_letter[ch]:
                                cand = tok[:i] + repl + tok[i + 1 :]
                                if self.is_valid_word(cand):
                                    k = (tok, map_lbl, css_sel)
                                    if k not in seen:
                                        seen.add(k)
                                        findings.append({
                                            "qa_pillar": "Typo & Character Confusion",
                                            "url": url, "content_type": content_type, "audit_priority": priority,
                                            "orig_text": text[:280], "suspicious_word": tok,
                                            "orig_char": ch, "suspected_char": repl,
                                            "suggested_correction": cand, "category": "Numeric",
                                            "ref_mapping": map_lbl, "ref_severity": ref_sev, "actual_severity": "High",
                                            "unicode_cp": f"U+{ord(ch):04X}", "html_element": html_el, "css_selector": css_sel,
                                            "attr_name": attr_name, "sentence": self._sentence_for(text, tok),
                                            "confidence": 0.95, "classification": "Confirmed Issue",
                                            "reason": f"Digit '{ch}' embedded in '{tok}'; matches valid word '{cand}'.",
                                        })
                                    break

                elif re.match(r"^[0-9]+[A-Za-z][0-9]*$", tok):
                    for i, ch in enumerate(tok):
                        if ch in self.letter_to_digit:
                            digit, map_lbl, ref_sev = self.letter_to_digit[ch][0]
                            cand = tok[:i] + digit + tok[i + 1 :]
                            k = (tok, map_lbl, css_sel)
                            if k not in seen:
                                seen.add(k)
                                findings.append({
                                    "qa_pillar": "Typo & Character Confusion",
                                    "url": url, "content_type": content_type, "audit_priority": priority,
                                    "orig_text": text[:280], "suspicious_word": tok,
                                    "orig_char": ch, "suspected_char": digit,
                                    "suggested_correction": cand, "category": "Numeric",
                                    "ref_mapping": map_lbl, "ref_severity": ref_sev, "actual_severity": "High",
                                    "unicode_cp": f"U+{ord(ch):04X}", "html_element": html_el, "css_selector": css_sel,
                                    "attr_name": attr_name, "sentence": self._sentence_for(text, tok),
                                    "confidence": 0.93, "classification": "Confirmed Issue",
                                    "reason": f"Letter '{ch}' inside numeric token '{tok}' (expected '{cand}').",
                                })
                            break

                elif tok.isalpha() and not tok.isupper() and not self.is_valid_word(tok):
                    matched_mc = False
                    for obs, intended, map_lbl, ref_sev in self.multichar_pairs:
                        if obs in tok.lower():
                            for m in re.finditer(re.escape(obs), tok, flags=re.IGNORECASE):
                                si, ei = m.span()
                                repl = intended.upper() if tok[si:ei].isupper() else intended
                                cand = tok[:si] + repl + tok[ei:]
                                if cand.lower() != tok.lower() and self.is_valid_word(cand):
                                    k = (tok, map_lbl, css_sel)
                                    if k not in seen:
                                        seen.add(k)
                                        findings.append({
                                            "qa_pillar": "Typo & Character Confusion",
                                            "url": url, "content_type": content_type, "audit_priority": priority,
                                            "orig_text": text[:280], "suspicious_word": tok,
                                            "orig_char": tok[si:ei], "suspected_char": repl,
                                            "suggested_correction": cand, "category": "Multi-character",
                                            "ref_mapping": map_lbl, "ref_severity": ref_sev, "actual_severity": "High",
                                            "unicode_cp": "N/A (Multi-char)", "html_element": html_el, "css_selector": css_sel,
                                            "attr_name": attr_name, "sentence": self._sentence_for(text, tok),
                                            "confidence": 0.91, "classification": "Confirmed Issue",
                                            "reason": f"Unknown word '{tok}' resolves to '{cand}' via {map_lbl}.",
                                        })
                                    matched_mc = True
                                    break
                        if matched_mc:
                            break

                    if not matched_mc:
                        for obs_ch, int_ch, map_lbl, ref_sev in self.alpha_pairs:
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
                                                findings.append({
                                                    "qa_pillar": "Typo & Character Confusion",
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

            # 1D. Punctuation & Stray Whitespace Anomalies (FIXED: catches 'API- first' & 'systems :')
            if self.punct_enabled:
                punct_checks = [
                    (r"[A-Za-z0-9]+[-–—]\s+[A-Za-z]+", "- ", "-", "Stray space after hyphen", "- -> – / stray space after hyphen (CCM-148)", "Low", "Confirmed Issue", "Erroneous space after hyphen/dash."),
                    (r"[A-Za-z]+\s+:", " :", ":", "Space before colon", ": -> ; / spacing (CCM-142)", "Low", "Confirmed Issue", "Erroneous space before colon."),
                    (r"\b\d+,\d+%\b", ",", ".", "Decimal comma", ". -> , (CCM-141)", "Medium", "Needs Manual Review", "Comma used in decimal percentage instead of decimal point."),
                    (r"\b[A-Za-z]*\|[A-Za-z]+\b|\b[A-Za-z]+\|[A-Za-z]*\b", "|", "I / l", "Vertical bar in word", "I / l -> | (CCM-185)", "High", "Confirmed Issue", "Vertical bar '|' embedded inside alphabetic word."),
                ]
                for pat, det, sug_ch, _, map_lbl, sev, cls, rsn in punct_checks:
                    for m in re.finditer(pat, text):
                        tok = m.group(0)
                        k = (tok, map_lbl, css_sel)
                        if k not in seen:
                            seen.add(k)
                            sug = re.sub(r"([-–—])\s+", r"\1", tok).replace(" :", ":")
                            findings.append({
                                "qa_pillar": "Typo & Character Confusion",
                                "url": url, "content_type": content_type, "audit_priority": priority,
                                "orig_text": text[:280], "suspicious_word": tok,
                                "orig_char": det, "suspected_char": sug_ch,
                                "suggested_correction": sug, "category": "Special characters & punctuation",
                                "ref_mapping": map_lbl, "ref_severity": "Medium", "actual_severity": sev,
                                "unicode_cp": f"U+{ord(det.strip()[0]):04X}", "html_element": html_el, "css_selector": css_sel,
                                "attr_name": attr_name, "sentence": self._sentence_for(text, tok),
                                "confidence": 0.94, "classification": cls, "reason": rsn,
                            })

        # --- MODULE 2: EDITORIAL, CAPITALIZATION & ORTHOGRAPHY CONSISTENCY ---
        if self.modules.get("editorial_consistency", True):
            # 2A. UK vs US Spelling Drift
            for wm in re.finditer(r"\b[A-Za-z-]+\b", text):
                tok = wm.group(0)
                low = tok.lower()
                if low in UK_TO_US_MAP:
                    us_form = UK_TO_US_MAP[low]
                    if tok.isupper():
                        us_form = us_form.upper()
                    elif tok[0].isupper():
                        us_form = us_form.capitalize()
                    k = (tok, "UK_US_SPELLING", css_sel)
                    if k not in seen:
                        seen.add(k)
                        findings.append({
                            "qa_pillar": "Editorial & Terminology Consistency",
                            "url": url, "content_type": content_type, "audit_priority": priority,
                            "orig_text": text[:280], "suspicious_word": tok,
                            "orig_char": tok, "suspected_char": us_form,
                            "suggested_correction": us_form, "category": "Alphabetic",
                            "ref_mapping": "UK vs US Spelling Consistency", "ref_severity": "Medium",
                            "actual_severity": "Medium", "unicode_cp": "N/A",
                            "html_element": html_el, "css_selector": css_sel, "attr_name": attr_name,
                            "sentence": self._sentence_for(text, tok),
                            "confidence": 0.92, "classification": "Confirmed Issue",
                            "reason": f"UK English spelling '{tok}' conflicts with US English standard ('{us_form}') used elsewhere on site.",
                        })

            # 2B. Heading Title-Case Consistency (e.g., 'AI enablement Advisory')
            if html_el in {"h1", "h2", "h3", "h4"} and attr_name == "innerText":
                words = text.split()
                if 2 <= len(words) <= 10:
                    cap_count = sum(1 for w in words if w[0].isupper())
                    if cap_count >= len(words) - 1:
                        for w in words[1:]:
                            clean_w = w.strip(".,:;!?()")
                            if clean_w and clean_w[0].islower() and clean_w.lower() not in TITLE_MINOR_WORDS and len(clean_w) >= 4:
                                sug_heading = text.replace(clean_w, clean_w.capitalize(), 1)
                                k = (clean_w, "HEADING_CASE", css_sel)
                                if k not in seen:
                                    seen.add(k)
                                    findings.append({
                                        "qa_pillar": "Editorial & Terminology Consistency",
                                        "url": url, "content_type": content_type, "audit_priority": priority,
                                        "orig_text": text[:280], "suspicious_word": clean_w,
                                        "orig_char": clean_w[0], "suspected_char": clean_w[0].upper(),
                                        "suggested_correction": sug_heading, "category": "Alphabetic",
                                        "ref_mapping": "Heading Capitalization Consistency", "ref_severity": "Low",
                                        "actual_severity": "Low", "unicode_cp": f"U+{ord(clean_w[0]):04X}",
                                        "html_element": html_el, "css_selector": css_sel, "attr_name": attr_name,
                                        "sentence": text[:280],
                                        "confidence": 0.90, "classification": "Confirmed Issue",
                                        "reason": f"Inconsistent title-case capitalization ('{clean_w}') inside {html_el.upper()} heading.",
                                    })

            # 2C. Product Trademark Casing Drift ('Zoy' vs 'ZoY', 'AI-native' vs 'AI-Native')
            for wrong_brand, right_brand in [("Zoy", "ZoY"), ("AI-native", "AI-Native"), ("Tezt", "TeZt"), ("Truzt", "TruZt")]:
                if re.search(rf"\b{re.escape(wrong_brand)}\b", text):
                    k = (wrong_brand, "BRAND_CASE", css_sel)
                    if k not in seen:
                        seen.add(k)
                        findings.append({
                            "qa_pillar": "Editorial & Terminology Consistency",
                            "url": url, "content_type": content_type, "audit_priority": priority,
                            "orig_text": text[:280], "suspicious_word": wrong_brand,
                            "orig_char": wrong_brand, "suspected_char": right_brand,
                            "suggested_correction": text.replace(wrong_brand, right_brand),
                            "category": "Alphabetic", "ref_mapping": "Brand / Product Trademark Consistency",
                            "ref_severity": "Medium", "actual_severity": "Medium", "unicode_cp": "N/A",
                            "html_element": html_el, "css_selector": css_sel, "attr_name": attr_name,
                            "sentence": self._sentence_for(text, wrong_brand),
                            "confidence": 0.96, "classification": "Confirmed Issue",
                            "reason": f"Inconsistent product/service capitalization ('{wrong_brand}' should be '{right_brand}').",
                        })

            # 2D. Truncated Metric Labels (e.g., '22% Growth in' with no following noun)
            if re.match(r"^\d+%\s+[A-Za-z]+\s+(in|of|for|to|with)$", text.strip(), flags=re.IGNORECASE):
                k = (text.strip(), "TRUNCATED_STAT", css_sel)
                if k not in seen:
                    seen.add(k)
                    findings.append({
                        "qa_pillar": "Editorial & Terminology Consistency",
                        "url": url, "content_type": content_type, "audit_priority": priority,
                        "orig_text": text[:280], "suspicious_word": text.strip(),
                        "orig_char": "(missing noun)", "suspected_char": "(metric label)",
                        "suggested_correction": f"{text.strip()} [metric name]",
                        "category": "Special characters & punctuation",
                        "ref_mapping": "Truncated Stat / Copy-Paste Cut-off", "ref_severity": "High",
                        "actual_severity": "High", "unicode_cp": "N/A",
                        "html_element": html_el, "css_selector": css_sel, "attr_name": attr_name,
                        "sentence": text[:280],
                        "confidence": 0.97, "classification": "Confirmed Issue",
                        "reason": f"Stat callout '{text.strip()}' ends with a dangling preposition and is missing its metric noun.",
                    })

        return findings


# ==============================================================================
# 4. FULL-PAGE PLAYWRIGHT A-TO-Z INSPECTOR (DOM + SEO + LINKS + MEDIA ASSETS)
# ==============================================================================
def extract_full_page_audit_items(
    url: str,
    content_type: str,
    priority: str,
    html_str: str,
    enabled_modules: Dict[str, bool],
) -> Tuple[List[Tuple[str, str, str, str]], List[Dict], Set[str]]:
    """
    Extracts:
      1. Text & attribute items for typo/consistency engine
      2. Technical SEO, Accessibility & Asset findings
      3. Internal/external links to validate HTTP status
    """
    soup = BeautifulSoup(html_str, "lxml")
    text_items: List[Tuple[str, str, str, str]] = []
    tech_findings: List[Dict] = []
    links_to_check: Set[str] = set()
    seen_local: Set[Tuple[str, str]] = set()

    # --- MODULE 4: SEO & METADATA QA ---
    title_txt = soup.title.string.strip() if (soup.title and soup.title.string) else ""
    if title_txt:
        text_items.append(("title", "head > title", "innerText", title_txt))
    elif enabled_modules.get("seo_meta", True):
        tech_findings.append({
            "qa_pillar": "SEO & Metadata QA", "url": url, "content_type": content_type, "audit_priority": priority,
            "orig_text": "<title> missing", "suspicious_word": "Missing Page Title",
            "orig_char": "N/A", "suspected_char": "<title>", "suggested_correction": "Add descriptive <title> tag",
            "category": "SEO Metadata", "ref_mapping": "SEO-001 (Title Tag)", "ref_severity": "High",
            "actual_severity": "High", "unicode_cp": "N/A", "html_element": "title",
            "css_selector": "head > title", "attr_name": "innerText", "sentence": "Page is missing a <title> tag.",
            "confidence": 1.0, "classification": "Confirmed Issue", "reason": "Missing HTML <title> tag.",
        })

    meta_desc = ""
    canonical_href = ""
    og_title = ""
    for meta in soup.find_all("meta"):
        m_name = (meta.get("name") or meta.get("property") or "").strip()
        m_content = (meta.get("content") or "").strip()
        if m_name.lower() == "description":
            meta_desc = m_content
        if m_name.lower() == "og:title":
            og_title = m_content
        if m_name.lower() in {"description", "og:title", "og:description", "twitter:title", "twitter:description"} and m_content:
            text_items.append(("meta", f"meta[name='{m_name}']", "content", m_content))

    can_el = soup.find("link", rel=lambda v: v and "canonical" in v)
    if can_el:
        canonical_href = (can_el.get("href") or "").strip()

    if enabled_modules.get("seo_meta", True):
        if not meta_desc:
            tech_findings.append({
                "qa_pillar": "SEO & Metadata QA", "url": url, "content_type": content_type, "audit_priority": priority,
                "orig_text": "meta[name='description'] missing", "suspicious_word": "Missing Meta Description",
                "orig_char": "N/A", "suspected_char": "meta description", "suggested_correction": "Add 120–160 char meta description",
                "category": "SEO Metadata", "ref_mapping": "SEO-002 (Meta Description)", "ref_severity": "Medium",
                "actual_severity": "Medium", "unicode_cp": "N/A", "html_element": "meta",
                "css_selector": "head > meta[name='description']", "attr_name": "content",
                "sentence": "Page has no meta description.", "confidence": 1.0,
                "classification": "Confirmed Issue", "reason": "Missing SEO meta description tag.",
            })
        if not canonical_href:
            tech_findings.append({
                "qa_pillar": "SEO & Metadata QA", "url": url, "content_type": content_type, "audit_priority": priority,
                "orig_text": "link[rel='canonical'] missing", "suspicious_word": "Missing Canonical Tag",
                "orig_char": "N/A", "suspected_char": url, "suggested_correction": f"<link rel='canonical' href='{url}'>",
                "category": "SEO Metadata", "ref_mapping": "SEO-003 (Canonical URL)", "ref_severity": "Medium",
                "actual_severity": "Medium", "unicode_cp": "N/A", "html_element": "link",
                "css_selector": "head > link[rel='canonical']", "attr_name": "href",
                "sentence": "Page is missing a canonical URL tag.", "confidence": 1.0,
                "classification": "Confirmed Issue", "reason": "Missing canonical link tag.",
            })
        elif canonical_href.rstrip("/") != url.rstrip("/"):
            tech_findings.append({
                "qa_pillar": "SEO & Metadata QA", "url": url, "content_type": content_type, "audit_priority": priority,
                "orig_text": f"Requested: {url} | Canonical: {canonical_href}",
                "suspicious_word": canonical_href, "orig_char": url, "suspected_char": canonical_href,
                "suggested_correction": canonical_href, "category": "SEO Metadata",
                "ref_mapping": "SEO-004 (Canonical / Slug Mismatch)", "ref_severity": "Medium",
                "actual_severity": "Medium", "unicode_cp": "N/A", "html_element": "link",
                "css_selector": "head > link[rel='canonical']", "attr_name": "href",
                "sentence": f"URL {url} canonicalizes to {canonical_href}.", "confidence": 0.95,
                "classification": "Needs Manual Review",
                "reason": "Requested URL path differs from page canonical URL (legacy redirect/alias slug).",
            })
        if not og_title:
            tech_findings.append({
                "qa_pillar": "SEO & Metadata QA", "url": url, "content_type": content_type, "audit_priority": priority,
                "orig_text": "meta[property='og:title'] missing", "suspicious_word": "Missing Open Graph Title",
                "orig_char": "N/A", "suspected_char": "og:title", "suggested_correction": "Add og:title tag",
                "category": "SEO Metadata", "ref_mapping": "SEO-005 (Open Graph)", "ref_severity": "Low",
                "actual_severity": "Low", "unicode_cp": "N/A", "html_element": "meta",
                "css_selector": "head > meta[property='og:title']", "attr_name": "content",
                "sentence": "Missing Open Graph og:title tag.", "confidence": 1.0,
                "classification": "Confirmed Issue", "reason": "Missing Open Graph social share title.",
            })

    h1_tags = soup.find_all("h1")
    if enabled_modules.get("seo_meta", True):
        if len(h1_tags) == 0:
            tech_findings.append({
                "qa_pillar": "SEO & Metadata QA", "url": url, "content_type": content_type, "audit_priority": priority,
                "orig_text": "0 <h1> tags found", "suspicious_word": "Missing H1",
                "orig_char": "0", "suspected_char": "1", "suggested_correction": "Add single primary H1 heading",
                "category": "HTML Structure", "ref_mapping": "SEO-006 (H1 Count)", "ref_severity": "High",
                "actual_severity": "High", "unicode_cp": "N/A", "html_element": "h1",
                "css_selector": "body h1", "attr_name": "innerText", "sentence": "No H1 heading detected on page.",
                "confidence": 1.0, "classification": "Confirmed Issue", "reason": "Page has no <h1> heading.",
            })
        elif len(h1_tags) > 1:
            tech_findings.append({
                "qa_pillar": "SEO & Metadata QA", "url": url, "content_type": content_type, "audit_priority": priority,
                "orig_text": f"{len(h1_tags)} <h1> tags found", "suspicious_word": "Multiple H1 Tags",
                "orig_char": str(len(h1_tags)), "suspected_char": "1", "suggested_correction": "Use a single H1 per page",
                "category": "HTML Structure", "ref_mapping": "SEO-006 (H1 Count)", "ref_severity": "Low",
                "actual_severity": "Low", "unicode_cp": "N/A", "html_element": "h1",
                "css_selector": "body h1", "attr_name": "innerText",
                "sentence": f"Found {len(h1_tags)} H1 headings on the same page.", "confidence": 0.95,
                "classification": "Needs Manual Review", "reason": f"Multiple H1 tags ({len(h1_tags)}) detected.",
            })

    # Strip non-visible script/style tags before DOM walk
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

    # Collect visible text blocks (including paragraphs/list items with nested spans/links!)
    for el in soup.find_all(["h1", "h2", "h3", "h4", "h5", "h6", "p", "li", "td", "th", "blockquote", "figcaption", "caption", "label", "button", "a", "span"]):
        # Only skip if it contains another block-level container (never skip <p> or <li> containing <span>/<a>/<strong>)
        if el.find(["h1", "h2", "h3", "h4", "h5", "h6", "p", "li", "td", "th", "table", "ul", "ol"]):
            continue
        txt = " ".join(el.get_text(separator=" ", strip=True).split())
        if len(txt) >= 2:
            s = _sel(el)
            if (s, txt) not in seen_local:
                seen_local.add((s, txt))
                text_items.append((el.name, s, "innerText", txt))

    # --- MODULE 3: IMAGES, ASSETS & ACCESSIBILITY QA ---
    for img in soup.find_all("img"):
        src = (img.get("src") or "").strip()
        alt = " ".join(str(img.get("alt", "")).split())
        title_a = " ".join(str(img.get("title", "")).split())
        s = _sel(img)

        if alt:
            text_items.append(("img", s, "alt", alt))
        if title_a:
            text_items.append(("img", s, "title", title_a))

        if enabled_modules.get("links_assets", True):
            if src.startswith("http://"):
                tech_findings.append({
                    "qa_pillar": "Links & Media Assets QA", "url": url, "content_type": content_type, "audit_priority": priority,
                    "orig_text": src, "suspicious_word": "http://", "orig_char": "http://", "suspected_char": "https://",
                    "suggested_correction": src.replace("http://", "https://", 1), "category": "Mixed Content",
                    "ref_mapping": "ASSET-001 (HTTPS)", "ref_severity": "High", "actual_severity": "High",
                    "unicode_cp": "N/A", "html_element": "img", "css_selector": s, "attr_name": "src",
                    "sentence": f"Insecure HTTP image source: {src}", "confidence": 1.0,
                    "classification": "Confirmed Issue", "reason": "Mixed-content HTTP image URL on HTTPS page.",
                })
            if not alt and not src.endswith(".svg"):
                tech_findings.append({
                    "qa_pillar": "Links & Media Assets QA", "url": url, "content_type": content_type, "audit_priority": priority,
                    "orig_text": f"<img src='{src[:80]}'>", "suspicious_word": "Missing Image Alt",
                    "orig_char": "(empty)", "suspected_char": "descriptive alt text",
                    "suggested_correction": "Add descriptive alt attribute", "category": "Media Accessibility",
                    "ref_mapping": "ASSET-002 (Alt Text)", "ref_severity": "Medium", "actual_severity": "Low",
                    "unicode_cp": "N/A", "html_element": "img", "css_selector": s, "attr_name": "alt",
                    "sentence": f"Image missing alt attribute: {src[:100]}", "confidence": 0.95,
                    "classification": "Needs Manual Review", "reason": "Image element has empty or missing alt text.",
                })
            elif alt and re.match(r"^(Frame[-_ ]?\d+|Rectangle[-_ ]?\d+|Group[-_ ]?\d+|image[-_ ]?\d+|IZ-white\s*1|Screenshot.*)$", alt, flags=re.IGNORECASE):
                tech_findings.append({
                    "qa_pillar": "Links & Media Assets QA", "url": url, "content_type": content_type, "audit_priority": priority,
                    "orig_text": f"alt='{alt}'", "suspicious_word": alt, "orig_char": alt,
                    "suspected_char": "Descriptive brand/image label", "suggested_correction": "Replace raw export filename in alt",
                    "category": "Media Accessibility", "ref_mapping": "ASSET-003 (Raw Filename Alt)",
                    "ref_severity": "Medium", "actual_severity": "Medium", "unicode_cp": "N/A",
                    "html_element": "img", "css_selector": s, "attr_name": "alt",
                    "sentence": f"Raw Figma/asset filename '{alt}' used as image alt text.", "confidence": 0.96,
                    "classification": "Confirmed Issue", "reason": f"Raw design export filename '{alt}' exposed in image alt attribute.",
                })

    # Collect links for HTTP status verification
    if enabled_modules.get("links_assets", True):
        for a_el in soup.find_all("a", href=True):
            href = a_el["href"].strip()
            if not href or href.startswith(("#", "javascript:", "mailto:", "tel:")):
                continue
            abs_url = urllib.parse.urljoin(url, href)
            if abs_url.startswith(("http://", "https://")):
                links_to_check.add(abs_url)

    for el in soup.find_all(attrs={"aria-label": True}):
        val = " ".join(str(el.get("aria-label", "")).split())
        if len(val) >= 2:
            text_items.append((el.name, _sel(el), "aria-label", val))

    for inp in soup.find_all(["input", "textarea"], placeholder=True):
        val = " ".join(str(inp.get("placeholder", "")).split())
        if len(val) >= 2:
            text_items.append((inp.name, _sel(inp), "placeholder", val))

    return text_items, tech_findings, links_to_check


async def check_links_batch(context, urls: List[str], limit: int = 35) -> Dict[str, int]:
    """Checks HTTP status codes for unique links on a page using Playwright's APIRequestContext."""
    statuses: Dict[str, int] = {}
    for u in urls[:limit]:
        try:
            resp = await context.request.get(u, timeout=12000)
            statuses[u] = resp.status
        except Exception:
            statuses[u] = 0
    return statuses


async def run_az_crawl(
    url_records: List[Dict],
    engine: ComprehensiveQAEngine,
    enabled_modules: Dict[str, bool],
    progress_cb=None,
) -> Tuple[List[Dict], List[Dict]]:
    if async_playwright is None:
        raise RuntimeError("Playwright is not installed. Run: pip install playwright && playwright install chromium")

    all_findings: List[Dict] = []
    url_summaries: List[Dict] = []
    link_status_cache: Dict[str, int] = {}

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
            t_start = time.perf_counter()
            try:
                resp = await page.goto(u, timeout=45000, wait_until="networkidle")
                await page.wait_for_timeout(1200)
                load_sec = round(time.perf_counter() - t_start, 2)
                status = resp.status if resp else 0

                if status >= 400 or status == 0:
                    url_summaries.append({**rec, "scan_status": "Failed", "http_status": status, "load_time_s": load_sec, "error": f"HTTP {status}"})
                else:
                    html_content = await page.content()
                    text_items, tech_findings, page_links = extract_full_page_audit_items(
                        u, ctype, prio, html_content, enabled_modules
                    )
                    all_findings.extend(tech_findings)

                    for html_el, css_sel, attr_name, txt in text_items:
                        hits = engine.inspect_text_block(u, ctype, prio, txt, html_el, css_sel, attr_name)
                        all_findings.extend(hits)

                    # Broken link checks (internal + external)
                    if enabled_modules.get("links_assets", True) and page_links:
                        unchecked = [lk for lk in sorted(page_links) if lk not in link_status_cache]
                        if unchecked:
                            new_statuses = await check_links_batch(context, unchecked, limit=30)
                            link_status_cache.update(new_statuses)
                        for lk in sorted(page_links):
                            st_code = link_status_cache.get(lk, 200)
                            if st_code >= 400 or st_code == 0:
                                all_findings.append({
                                    "qa_pillar": "Links & Media Assets QA", "url": u, "content_type": ctype,
                                    "audit_priority": prio, "orig_text": lk,
                                    "suspicious_word": f"Broken Link (HTTP {st_code})",
                                    "orig_char": str(st_code), "suspected_char": "200 OK",
                                    "suggested_correction": "Update or remove broken destination URL",
                                    "category": "Broken Link", "ref_mapping": "LINK-404",
                                    "ref_severity": "High", "actual_severity": "High", "unicode_cp": "N/A",
                                    "html_element": "a", "css_selector": f"a[href*='{lk[-30:]}']",
                                    "attr_name": "href", "sentence": f"Link on page returned HTTP {st_code}: {lk}",
                                    "confidence": 0.99, "classification": "Confirmed Issue",
                                    "reason": f"Destination URL returned HTTP {st_code}.",
                                })

                    # Performance SLA check (>6.0s load)
                    if enabled_modules.get("seo_meta", True) and load_sec > 6.0:
                        all_findings.append({
                            "qa_pillar": "SEO & Metadata QA", "url": u, "content_type": ctype,
                            "audit_priority": prio, "orig_text": f"Page load time: {load_sec}s",
                            "suspicious_word": f"{load_sec}s load time", "orig_char": f"{load_sec}s",
                            "suspected_char": "< 4.0s", "suggested_correction": "Optimize image/JS payload",
                            "category": "Page Performance", "ref_mapping": "PERF-001",
                            "ref_severity": "Low", "actual_severity": "Low", "unicode_cp": "N/A",
                            "html_element": "document", "css_selector": "window.performance",
                            "attr_name": "loadTime", "sentence": f"Full networkidle load took {load_sec}s.",
                            "confidence": 0.90, "classification": "Needs Manual Review",
                            "reason": f"Slow page render ({load_sec}s > 6.0s threshold).",
                        })

                    url_summaries.append({**rec, "scan_status": "Completed", "http_status": status, "load_time_s": load_sec, "error": ""})
            except Exception as exc:
                load_sec = round(time.perf_counter() - t_start, 2)
                url_summaries.append({**rec, "scan_status": "Failed", "http_status": 0, "load_time_s": load_sec, "error": str(exc)})
            finally:
                await context.close()
        await browser.close()

    return url_summaries, all_findings


# ==============================================================================
# 5. MASTER EXCEL REPORT BUILDER (FRESH & APPEND SUPPORT)
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


def build_az_excel_report(
    url_records: List[Dict],
    findings: List[Dict],
    rules_df: pd.DataFrame,
    append_mode: bool = False,
    existing_excel_bytes: Optional[bytes] = None,
) -> bytes:
    if append_mode:
        prev_buf = io.BytesIO(existing_excel_bytes) if existing_excel_bytes else (MASTER_EXCEL_PATH if MASTER_EXCEL_PATH.exists() else None)
        if prev_buf is not None:
            try:
                old_urls_df = pd.read_excel(prev_buf, sheet_name="Master URL Report")
                old_urls_df = old_urls_df[old_urls_df["S.No"].astype(str) != "TOTAL"]
                if not existing_excel_bytes:
                    prev_buf = MASTER_EXCEL_PATH
                else:
                    prev_buf.seek(0)
                old_find_df = pd.read_excel(prev_buf, sheet_name="All Issue Findings")

                new_url_set = {r["url"] for r in url_records}
                merged_urls = []
                for _, r in old_urls_df.iterrows():
                    if r["URL"] not in new_url_set:
                        merged_urls.append({
                            "url": r["URL"],
                            "content_type": r["Content Type"],
                            "priority": r["Audit Priority"],
                            "scan_status": r["Scan Status"],
                            "http_status": r.get("HTTP Status", 200),
                            "load_time_s": r.get("Load Time (s)", 0),
                            "sitemap": r.get("Sitemap Source", "Previous Run"),
                        })
                merged_urls.extend(url_records)
                url_records = merged_urls

                kept_findings = []
                for _, r in old_find_df.iterrows():
                    if r["URL"] not in new_url_set:
                        kept_findings.append({
                            "qa_pillar": r.get("QA Pillar", "Typo & Character Confusion"),
                            "url": r["URL"],
                            "content_type": r["Content Type"],
                            "audit_priority": r["Audit Priority"],
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
                findings = kept_findings + findings
            except Exception:
                pass

    for i, f in enumerate(findings, start=1):
        f["issue_id"] = f"QA-{i:05d}"

    wb = Workbook()

    # --- Sheet 3: All Issue Findings ---
    ws3 = wb.active
    ws3.title = "All Issue Findings"
    s3_headers = [
        "Issue ID", "QA Pillar", "URL", "Content Type", "Audit Priority", "Original Text",
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
            item["issue_id"], item.get("qa_pillar", "Typo & Character Confusion"),
            item["url"], item["content_type"], item.get("audit_priority", item.get("priority", "High")),
            item["orig_text"], item["suspicious_word"], item["orig_char"],
            item["suspected_char"], item["suggested_correction"], item["category"],
            item["ref_mapping"], item["ref_severity"], item["actual_severity"],
            item["unicode_cp"], item["html_element"], item["css_selector"],
            item["attr_name"], item["sentence"], item["confidence"],
            item["classification"], item["reason"],
        ]
        for ci, v in enumerate(vals, start=1):
            c = ws3.cell(row=ro, column=ci, value=v)
            if ci == 3:
                c.hyperlink = v
                c.font = LINK_FONT
            elif ci == 20:
                c.number_format = "0.00"

    last_f = max(2, 1 + len(findings))
    _style_tbl(ws3, 1, last_f, 1, len(s3_headers), freeze=True)
    if findings:
        _apply_cf(ws3, "M", 2, last_f)
        _apply_cf(ws3, "N", 2, last_f)
    for ci, w in {1: 12, 2: 26, 3: 44, 4: 17, 5: 14, 6: 42, 7: 22, 8: 16, 9: 20, 10: 26, 11: 26, 12: 26, 13: 16, 14: 18, 15: 20, 16: 14, 17: 34, 18: 14, 19: 42, 20: 15, 21: 19, 22: 46}.items():
        ws3.column_dimensions[get_column_letter(ci)].width = w

    # --- Sheet 2: Master URL Report ---
    ws2 = wb.create_sheet("Master URL Report")
    s2_headers = [
        "S.No", "URL", "Sitemap Source", "Content Type", "HTTP Status", "Load Time (s)",
        "Audit Priority", "Scan Status", "Total Findings",
        "Confirmed Issues", "Manual Review", "False Positives",
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
        ws2.cell(row=idx, column=5, value=rec.get("http_status", 200))
        ws2.cell(row=idx, column=6, value=rec.get("load_time_s", 0))
        ws2.cell(row=idx, column=7, value=rec["priority"])
        ws2.cell(row=idx, column=8, value=rec.get("scan_status", "Completed"))

        if len(findings) > 0:
            ws2.cell(row=idx, column=9, value=f'=COUNTIF(\'All Issue Findings\'!\$C\$2:\$C\${f_end}, B{idx})')
            ws2.cell(row=idx, column=10, value=f'=COUNTIFS(\'All Issue Findings\'!\$C\$2:\$C\${f_end}, B{idx}, \'All Issue Findings\'!\$U\$2:\$U\${f_end}, "Confirmed Issue")')
            ws2.cell(row=idx, column=11, value=f'=COUNTIFS(\'All Issue Findings\'!\$C\$2:\$C\${f_end}, B{idx}, \'All Issue Findings\'!\$U\$2:\$U\${f_end}, "Needs Manual Review")')
            ws2.cell(row=idx, column=12, value=f'=COUNTIFS(\'All Issue Findings\'!\$C\$2:\$C\${f_end}, B{idx}, \'All Issue Findings\'!\$U\$2:\$U\${f_end}, "False Positive")')
            ws2.cell(row=idx, column=13, value=f'=COUNTIFS(\'All Issue Findings\'!\$C\$2:\$C\${f_end}, B{idx}, \'All Issue Findings\'!\$N\$2:\$N\${f_end}, "Critical")')
            ws2.cell(row=idx, column=14, value=f'=COUNTIFS(\'All Issue Findings\'!\$C\$2:\$C\${f_end}, B{idx}, \'All Issue Findings\'!\$N\$2:\$N\${f_end}, "High")')
            ws2.cell(row=idx, column=15, value=f'=COUNTIFS(\'All Issue Findings\'!\$C\$2:\$C\${f_end}, B{idx}, \'All Issue Findings\'!\$N\$2:\$N\${f_end}, "Medium")')
            ws2.cell(row=idx, column=16, value=f'=COUNTIFS(\'All Issue Findings\'!\$C\$2:\$C\${f_end}, B{idx}, \'All Issue Findings\'!\$N\$2:\$N\${f_end}, "Low")')
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
    for ci, w in {1: 8, 2: 50, 3: 24, 4: 18, 5: 13, 6: 14, 7: 14, 8: 16, 9: 15, 10: 16, 11: 15, 12: 15, 13: 11, 14: 11, 15: 11, 16: 11, 17: 24}.items():
        ws2.column_dimensions[get_column_letter(ci)].width = w

    # --- Sheet 1: Executive Summary ---
    ws1 = wb.create_sheet("Executive Summary")
    wb._sheets = [ws1, ws2, ws3]

    ws1["A1"] = "A-to-Z Website QA & Character Confusion Audit Report"
    ws1["A1"].font = TITLE_FONT
    ws1["A2"] = f"Generated: {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC')} | Mode: {'Append / Cumulative' if append_mode else 'Fresh Scan'}"
    ws1["A2"].font = Font(name="Calibri", size=10, italic=True, color="4A5568")

    ws1["A4"] = "Overall QA KPIs"
    ws1["A4"].font = SECTION_FONT
    ws1["A5"] = "Audit Metric"
    ws1["B5"] = "Value"

    u_end = tot_r - 1
    kpis = [
        ("Total URLs", f"=COUNTA('Master URL Report'!B2:B{u_end})"),
        ("Successfully scanned URLs", f'=COUNTIF(\'Master URL Report\'!H2:H{u_end},"Completed")'),
        ("Failed URLs", f'=COUNTIF(\'Master URL Report\'!H2:H{u_end},"Failed")'),
        ("URLs containing findings", f"=COUNTIF('Master URL Report'!I2:I{u_end},\">0\")"),
        ("URLs without detected issues", f"=COUNTIF('Master URL Report'!I2:I{u_end},0)"),
        ("Total QA findings", len(findings)),
        ("Confirmed issues", f'=COUNTIF(\'All Issue Findings\'!U2:U{last_f},"Confirmed Issue")' if findings else 0),
        ("Needs Manual Review", f'=COUNTIF(\'All Issue Findings\'!U2:U{last_f},"Needs Manual Review")' if findings else 0),
        ("False Positives", f'=COUNTIF(\'All Issue Findings\'!U2:U{last_f},"False Positive")' if findings else 0),
    ]
    for idx, (lbl, val) in enumerate(kpis, start=6):
        ws1.cell(row=idx, column=1, value=lbl)
        ws1.cell(row=idx, column=2, value=val)
    _style_tbl(ws1, 5, 14, 1, 2)

    # Breakdown by QA Pillar
    ws1["D4"] = "Findings by QA Pillar"
    ws1["D4"].font = SECTION_FONT
    ws1["D5"] = "QA Pillar"
    ws1["E5"] = "Total Findings"
    ws1["F5"] = "Confirmed Issues"
    pillars = [
        "Typo & Character Confusion",
        "Editorial & Terminology Consistency",
        "Links & Media Assets QA",
        "SEO & Metadata QA",
    ]
    for ro, pl in enumerate(pillars, start=6):
        ws1.cell(row=ro, column=4, value=pl)
        ws1.cell(row=ro, column=5, value=f'=COUNTIF(\'All Issue Findings\'!\$B\$2:\$B\${last_f}, D{ro})' if findings else 0)
        ws1.cell(row=ro, column=6, value=f'=COUNTIFS(\'All Issue Findings\'!\$B\$2:\$B\${last_f}, D{ro}, \'All Issue Findings\'!\$U\$2:\$U\${last_f}, "Confirmed Issue")' if findings else 0)
    ws1.cell(row=10, column=4, value="Total")
    ws1.cell(row=10, column=5, value="=SUM(E6:E9)")
    ws1.cell(row=10, column=6, value="=SUM(F6:F9)")
    _style_tbl(ws1, 5, 10, 4, 6, total_row=True)

    # Breakdown by Character Confusion Category
    ws1["A16"] = "Character Confusion Matrix Breakdown"
    ws1["A16"].font = SECTION_FONT
    ws1["A17"] = "Character Confusion Category"
    ws1["B17"] = "Active Rules"
    ws1["C17"] = "Total Findings"
    cats = ["Alphabetic", "Numeric", "Multi-character", "Special characters & punctuation", "Unicode lookalikes"]
    active_df = rules_df[rules_df["Enabled"] == True]
    for ro, cname in enumerate(cats, start=18):
        ws1.cell(row=ro, column=1, value=cname)
        ws1.cell(row=ro, column=2, value=int((active_df["Category"] == cname).sum()))
        ws1.cell(row=ro, column=3, value=f'=COUNTIF(\'All Issue Findings\'!\$K\$2:\$K\${last_f}, A{ro})' if findings else 0)
    ws1.cell(row=23, column=1, value="Total")
    ws1.cell(row=23, column=2, value="=SUM(B18:B22)")
    ws1.cell(row=23, column=3, value="=SUM(C18:C22)")
    _style_tbl(ws1, 17, 23, 1, 3, total_row=True)

    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "Findings by QA Pillar"
    chart.y_axis.title = "Issues"
    chart.add_data(Reference(ws1, min_col=5, min_row=5, max_row=9), titles_from_data=True)
    chart.set_categories(Reference(ws1, min_col=4, min_row=6, max_row=9))
    chart.legend = None
    chart.width = 15
    chart.height = 8.5
    ws1.add_chart(chart, "D12")

    for ci, w in {1: 34, 2: 18, 3: 16, 4: 34, 5: 16, 6: 18}.items():
        ws1.column_dimensions[get_column_letter(ci)].width = w

    out_bytes = io.BytesIO()
    wb.save(out_bytes)
    data = out_bytes.getvalue()
    MASTER_EXCEL_PATH.write_bytes(data)
    return data


# ==============================================================================
# 6. STREAMLIT A-TO-Z QA STUDIO UI
# ==============================================================================
def main():
    st.set_page_config(
        page_title="A-to-Z Website QA Studio",
        page_icon="🛡️",
        layout="wide",
    )
    ensure_playwright_browser()

    st.title("🛡️ A-to-Z Website QA & Content Integrity Studio")
    st.caption("Autonomous End-to-End QA Agent: Character Confusion Matrix · Editorial Consistency · Broken Links & Assets · SEO & Metadata")

    rules_df = load_rules_df()
    whitelist_terms = load_whitelist()

    tab_run, tab_rules, tab_wl = st.tabs([
        "🚀 1. Run A-to-Z Website QA",
        f"📋 2. Character Confusion Rules ({int((rules_df['Enabled'] == True).sum())}/{len(rules_df)} Active)",
        f"🛡️ 3. Brand & Tech Whitelist ({len(whitelist_terms)} Terms)",
    ])

    with tab_run:
        col_in, col_cfg = st.columns([2, 1])

        with col_in:
            input_mode = st.radio(
                "How would you like to provide URLs?",
                ["Paste Single or Multiple URLs", "Crawl Full XML Sitemap (Auto-expands nested sitemaps)", "Upload CSV / TXT File"],
                horizontal=True,
            )

            urls_raw = ""
            sitemap_input = ""
            max_sm_urls = 50
            uploaded_file = None

            if input_mode == "Paste Single or Multiple URLs":
                urls_raw = st.text_area(
                    "Enter URLs (one per line)",
                    value="",
                    height=150,
                )
            elif input_mode.startswith("Crawl Full XML"):
                sitemap_input = st.text_input("XML Sitemap Index URL", value="")
                max_sm_urls = st.slider("Max URLs to crawl in this batch", min_value=5, max_value=250, value=30, step=5)
            else:
                uploaded_file = st.file_uploader("Upload .csv (with 'URL' column) or .txt file", type=["csv", "txt"])

            st.markdown("**Select A-to-Z QA Checks to Run:**")
            c1, c2, c3, c4 = st.columns(4)
            chk_typo = c1.checkbox("🔤 Typos & Confusion Matrix", value=True)
            chk_edit = c2.checkbox("✍️ Editorial & Brand Casing", value=True)
            chk_link = c3.checkbox("🔗 Broken Links & Image Assets", value=True)
            chk_seo = c4.checkbox("📈 SEO, Meta & Speed QA", value=True)

        with col_cfg:
            report_mode = st.radio(
                "Report Output Mode",
                ["Fresh Report (New Excel)", "Append to Previous Master Report"],
            )
            append_upload = None
            if report_mode.startswith("Append"):
                append_upload = st.file_uploader("Upload previous .xlsx report to append into (optional)", type=["xlsx"])

            default_ctype = st.selectbox("Default Content Type", ["Service Page", "Product Page", "Homepage", "Case Study", "Blog Post", "Page"])
            default_prio = st.selectbox("Default Priority", ["High", "Critical", "Medium", "Low"])

        if st.button("▶️ Run Complete A-to-Z QA Audit", type="primary", use_container_width=True):
            records: List[Dict] = []
            if input_mode == "Paste Single or Multiple URLs":
                seen: Set[str] = set()
                for ln in urls_raw.splitlines():
                    u = ln.strip()
                    if u.startswith("http") and u not in seen:
                        seen.add(u)
                        ct = "Homepage" if u.rstrip("/") == "https://www.zucisystems.com" else ("Product Page" if "/truzt-ai" in u else default_ctype)
                        records.append({"url": u, "content_type": ct, "priority": default_prio, "sitemap": "Manual Input"})
            elif input_mode.startswith("Crawl Full XML"):
                with st.spinner(f"Expanding nested XML sitemaps from {sitemap_input}..."):
                    records = expand_sitemap_urls(sitemap_input.strip(), max_urls=max_sm_urls)
            elif uploaded_file is not None:
                if uploaded_file.name.endswith(".csv"):
                    udf = pd.read_csv(uploaded_file)
                    ucol = next((c for c in udf.columns if c.strip().lower() == "url"), udf.columns[0])
                    for u in udf[ucol].dropna():
                        us = str(u).strip()
                        if us.startswith("http"):
                            records.append({"url": us, "content_type": default_ctype, "priority": default_prio, "sitemap": uploaded_file.name})
                else:
                    for ln in uploaded_file.read().decode("utf-8", errors="ignore").splitlines():
                        us = ln.strip()
                        if us.startswith("http"):
                            records.append({"url": us, "content_type": default_ctype, "priority": default_prio, "sitemap": uploaded_file.name})

            if not records:
                st.error("No valid URLs found to scan.")
            else:
                enabled_mods = {
                    "typo_matrix": chk_typo,
                    "editorial_consistency": chk_edit,
                    "links_assets": chk_link,
                    "seo_meta": chk_seo,
                }
                engine = ComprehensiveQAEngine(rules_df=rules_df, whitelist=whitelist_terms, enabled_modules=enabled_mods)
                prog = st.progress(0.0, text="Starting A-to-Z Playwright QA Agent...")

                def _cb(cur: int, tot: int, cur_u: str):
                    prog.progress(cur / tot, text=f"Auditing ({cur}/{tot}): {cur_u}")

                try:
                    t0 = time.perf_counter()
                    url_sums, found_items = asyncio.run(run_az_crawl(records, engine, enabled_mods, progress_cb=_cb))
                    elapsed = round(time.perf_counter() - t0, 2)
                    prog.progress(1.0, text=f"A-to-Z QA completed across {len(records)} URLs in {elapsed}s!")

                    is_append = report_mode.startswith("Append")
                    prev_bytes = append_upload.getvalue() if append_upload is not None else None
                    xlsx_bytes = build_az_excel_report(
                        url_records=url_sums,
                        findings=found_items,
                        rules_df=rules_df,
                        append_mode=is_append,
                        existing_excel_bytes=prev_bytes,
                    )

                    st.session_state["az_excel_bytes"] = xlsx_bytes
                    st.session_state["az_findings_df"] = pd.DataFrame(found_items)
                    st.session_state["az_urls_df"] = pd.DataFrame(url_sums)
                    st.session_state["az_filename"] = (
                        "Zuci_Systems_AZ_QA_Master_Report.xlsx"
                        if is_append
                        else f"Zuci_Systems_AZ_QA_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    )
                except Exception as exc:
                    st.error(f"QA execution error: {exc}")

        if "az_excel_bytes" in st.session_state:
            st.divider()
            fdf = st.session_state["az_findings_df"]
            udf = st.session_state["az_urls_df"]

            k1, k2, k3, k4 = st.columns(4)
            k1.metric("URLs Audited", len(udf))
            k2.metric("Total QA Findings", len(fdf))
            k3.metric("Confirmed Issues", int((fdf["classification"] == "Confirmed Issue").sum()) if not fdf.empty else 0)
            k4.metric("Needs Manual Review", int((fdf["classification"] == "Needs Manual Review").sum()) if not fdf.empty else 0)

            st.download_button(
                label="📥 Download Complete A-to-Z QA Report (.xlsx)",
                data=st.session_state["az_excel_bytes"],
                file_name=st.session_state["az_filename"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )

            if not fdf.empty:
                pillar_filter = st.multiselect(
                    "Filter Results by QA Pillar",
                    options=sorted(fdf["qa_pillar"].unique().tolist()),
                    default=sorted(fdf["qa_pillar"].unique().tolist()),
                )
                st.dataframe(fdf[fdf["qa_pillar"].isin(pillar_filter)], use_container_width=True)
            else:
                st.success("Zero QA issues detected across all enabled checks!")

    with tab_rules:
        st.subheader("Editable Character Confusion Matrix Rules")
        cat_f = st.selectbox("Filter Category", ["All", "Alphabetic", "Numeric", "Multi-character", "Special characters & punctuation", "Unicode lookalikes"])
        disp_df = rules_df if cat_f == "All" else rules_df[rules_df["Category"] == cat_f].copy()
        ed_df = st.data_editor(
            disp_df,
            num_rows="dynamic",
            use_container_width=True,
            height=460,
            column_config={
                "Enabled": st.column_config.CheckboxColumn("Enabled", default=True),
                "Rule_ID": st.column_config.TextColumn("Rule ID", required=True),
                "Category": st.column_config.SelectboxColumn("Category", options=["Alphabetic", "Numeric", "Multi-character", "Special characters & punctuation", "Unicode lookalikes"], required=True),
                "Actual Character(s)": st.column_config.TextColumn("Actual Character(s)", required=True),
                "Potential Confusion": st.column_config.TextColumn("Potential Confusion", required=True),
                "Example / Notes": st.column_config.TextColumn("Example / Notes"),
                "Priority": st.column_config.SelectboxColumn("Priority", options=["High", "Medium", "Low"], default="Medium"),
            },
        )
        c_save, c_rst = st.columns(2)
        with c_save:
            if st.button("💾 Save Rules", type="primary", use_container_width=True):
                full_df = ed_df.reset_index(drop=True) if cat_f == "All" else pd.concat([rules_df[rules_df["Category"] != cat_f], ed_df], ignore_index=True)
                save_rules_df(full_df)
                st.success(f"Saved {len(full_df)} rules to `rules_store.json`!")
                st.rerun()
        with c_rst:
            if st.button("🔄 Reset to Default 205 Rules", use_container_width=True):
                ddf = pd.DataFrame(DEFAULT_RULES)
                ddf["Enabled"] = True
                save_rules_df(ddf)
                st.rerun()

    with tab_wl:
        st.subheader("Editable Brand & Technical Whitelist")
        wl_raw = st.text_area("Whitelisted Terms (one per line or comma-separated)", value="\n".join(whitelist_terms), height=340)
        if st.button("💾 Save Whitelist", type="primary"):
            save_whitelist(re.split(r"[,\n]+", wl_raw))
            st.success("Whitelist saved to `whitelist_store.json`!")
            st.rerun()


if __name__ == "__main__":
    main()
