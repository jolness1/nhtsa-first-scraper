#!/usr/bin/env python3
"""Fetch FIRST CrashReport.xlsx for all states using Playwright to obtain cookies.

Usage:
  - create and activate a venv: `python3 -m venv .venv && source .venv/bin/activate`
  - install deps: `pip install -r requirements.txt`
  - install playwright browsers: `playwright install` (after installing package)
  - run: `python fetch-first-dui-data.py`
"""

import json
import time
from pathlib import Path
from urllib.parse import urlencode

from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright

WORKDIR = Path(__file__).parent
STATE_FILE = WORKDIR / "state-list.json"
OUTDIR = WORKDIR / "scraped"
SAS_URL = "https://cdan.dot.gov/SASJobExecution/?sso_guest=true"
QUERY_URL = "https://cdan.dot.gov/query"
PROGRAM = "/Public/OTRA/Apps/FIRST/FIRST"
APPHOST = "cdan.dot.gov"
BASE = "https://cdan.dot.gov"


def ensure_outdir():
	OUTDIR.mkdir(exist_ok=True)


def load_states():
	if not STATE_FILE.exists():
		raise FileNotFoundError(f"State file not found: {STATE_FILE}")
	text = STATE_FILE.read_text(encoding="utf-8")
	print(f"DEBUG: read {len(text)} bytes from {STATE_FILE}")
	preview = text[:200].replace("\n", "\\n")
	print(f"DEBUG PREVIEW: {preview}")
	if not text.strip():
		raise ValueError(f"State file is empty: {STATE_FILE}")
	return json.loads(text)


def post_query_and_download(request_api, page, sid, name):
	sas_query = (
		"&topic_num=26&metric_num=33&metrictype_num=37"
		"&CrashYear=2010,2011,2012,2013,2014,2015,2016,2017,2018,2019,2020,2021,2022,2023,"
		f"&State={sid}&A_PTYPE=1&DRIMPAIR_A=9&TableRows=YEAR&TableCols=MONTH"
		"&ReleaseDate=Version 9.2.1, released Nov 13, 2025&ReportType=1&Criteria=Years: 2010-2023"
	)
	data = {
		"SASQueryString": sas_query,
		"_program": PROGRAM,
		"_apphostname": APPHOST,
	}

	headers = {
		"x-requested-with": "XMLHttpRequest",
		"referer": QUERY_URL,
		"origin": "https://cdan.dot.gov",
		"accept": "*/*",
		"content-type": "application/x-www-form-urlencoded",
	}

	body = urlencode(data)
	resp = request_api.post(SAS_URL, data=body, headers=headers, timeout=60000)
	status = resp.status
	text = resp.text()
	print(f"DEBUG: POST response length={len(text)} for {name} status={status}")
	print(f"DEBUG PREVIEW POST: {text[:400].replace('\n','\\n')}")

	# handle 449 SSO response by following the returned uri and retrying once
	if status == 449:
		try:
			j = resp.json()
			uri = j.get("uri") or j.get("URI") or j.get("uri")
		except Exception:
			uri = None
		if uri:
			auth_url = uri if uri.startswith("http") else BASE + uri
			print(f"Following SSO auth URI: {auth_url}")
			_ = request_api.get(auth_url, headers={"referer": QUERY_URL, "origin": "https://cdan.dot.gov"}, timeout=60000)
			resp = request_api.post(SAS_URL, data=body, headers=headers, timeout=60000)
			status = resp.status
			text = resp.text()
			print(f"DEBUG: POST retry length={len(text)} for {name} status={status}")
			print(f"DEBUG PREVIEW POST RETRY: {text[:400].replace('\n','\\n')}")

	soup = BeautifulSoup(text, "html.parser")

	# if there's an auto-submitting ProgressForm, submit it once via request_api and load the result into the page
	form = soup.find("form", attrs={"name": "ProgressForm"}) or soup.find("form", id="ProgressForm")
	if form:
		action = form.get("action") or SAS_URL
		if action.startswith("/"):
			action = BASE + action
		inputs = {}
		for inp in form.find_all(["input", "textarea"]):
			namei = inp.get("name")
			if not namei:
				continue
			inputs[namei] = inp.get("value", "")
		print(f"Submitting ProgressForm to {action} (once)")
		body2 = urlencode(inputs)
		resp2 = request_api.post(action, data=body2, headers={"referer": QUERY_URL, "origin": "https://cdan.dot.gov", "content-type": "application/x-www-form-urlencoded"}, timeout=60000)
		page_html = resp2.text()
		print(f"DEBUG: ProgressForm submit returned {resp2.status} length={len(page_html)}")
	else:
		page_html = text

	# load HTML into Playwright page so inline scripts run and DOM can be queried
	try:
		page.set_content(page_html, wait_until="load")
	except Exception:
		from urllib.parse import quote

		data_url = "data:text/html," + quote(page_html)
		page.goto(data_url, timeout=60000)

	# wait for the xlsx anchor to appear in the page DOM (no repeated POSTS)
	try:
		handle = page.wait_for_selector('a[download$=".xlsx"], a[href$=".xlsx"], a[href*="/files/files/"]', timeout=180000)
		link = handle.get_attribute("href")
	except Exception:
		# try a JS fallback to find the anchor
		try:
			link = page.evaluate("() => { const a = document.querySelector('a[download$=\".xlsx\"], a[href$=\".xlsx\"], a[href*=\"/files/files/\"]'); return a ? a.getAttribute('href') : null }")
		except Exception:
			link = None

	if not link:
		print(f"Could not find xlsx link for {name} via DOM polling")
		return False


	# normalize link and download
	if link.startswith("/"):
		url = BASE + link
	elif link.startswith("http"):
		url = link
	else:
		url = BASE + "/" + link

	r2 = request_api.get(url, headers={"referer": QUERY_URL, "origin": "https://cdan.dot.gov"}, timeout=120000)
	if r2.status == 200:
		outpath = OUTDIR / f"{name}-dui-data.xlsx"
		bodybytes = r2.body()
		with open(outpath, "wb") as fh:
			fh.write(bodybytes)
		print(f"Saved {outpath}")
		return True
	else:
		print(f"Failed to download file for {name}: {r2.status}")
		return False


def main():
	ensure_outdir()
	states = load_states()
	with sync_playwright() as p:
		browser = p.chromium.launch(headless=True)
		context = browser.new_context()
		page = context.new_page()
		page.goto(QUERY_URL, timeout=60000)
		request_api = context.request
		for s in states:
			sid = s.get("Id")
			name = s.get("StateName", "unknown").replace("/", "-")
			print(f"Processing {sid} - {name}")
			try:
				ok = post_query_and_download(request_api, page, sid, name)
			except Exception as e:
				print(f"Error for {name}: {e}")
			time.sleep(1)
		browser.close()


if __name__ == "__main__":
	main()
