"""
Capture dashboard screenshots for the assignment submission document.

Uses Playwright with the locally installed Chrome browser to avoid downloading
Chromium during the assignment workflow.
"""
from __future__ import annotations

import subprocess
import sys
import time
from pathlib import Path

from playwright.sync_api import sync_playwright

PROJECT_ROOT = Path(__file__).resolve().parent
OUTPUT_DIR = PROJECT_ROOT / "submission_assets"
BASE_URL = "http://127.0.0.1:8010"
PERIOD = "2024Q4"

SHOTS = [
    ("dashboard_hero.png", f"{BASE_URL}/?granularity=Q&period={PERIOD}", 1440, 1200, 0),
    ("dashboard_advanced_rollups.png", f"{BASE_URL}/?granularity=A&period=2024", 1440, 1200, 900),
    ("dashboard_tables.png", f"{BASE_URL}/?granularity=Q&period={PERIOD}", 1440, 1200, 3000),
]


def start_server() -> subprocess.Popen:
  return subprocess.Popen(
      [sys.executable, "manage.py", "runserver", "127.0.0.1:8010", "--noreload"],
      cwd=PROJECT_ROOT,
      stdout=subprocess.DEVNULL,
      stderr=subprocess.DEVNULL,
  )


def wait_for_server(timeout_seconds: int = 30) -> None:
  import urllib.error
  import urllib.request

  deadline = time.time() + timeout_seconds
  while time.time() < deadline:
      try:
          with urllib.request.urlopen(BASE_URL, timeout=2) as response:
              if response.status == 200:
                  return
      except (urllib.error.URLError, TimeoutError):
          time.sleep(0.5)
  raise RuntimeError("Django server did not become ready in time.")


def capture() -> list[Path]:
  OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
  server = start_server()
  saved: list[Path] = []

  try:
      wait_for_server()
      with sync_playwright() as playwright:
          browser = playwright.chromium.launch(channel="chrome", headless=True)
          page = browser.new_page(viewport={"width": 1440, "height": 900})
          for filename, url, width, height, scroll_y in SHOTS:
              page.set_viewport_size({"width": width, "height": height})
              page.goto(url, wait_until="networkidle")
              page.wait_for_timeout(1200)
              if scroll_y:
                  page.evaluate(f"window.scrollTo(0, {scroll_y})")
                  page.wait_for_timeout(600)
              target = OUTPUT_DIR / filename
              page.screenshot(path=str(target), full_page=False)
              saved.append(target)
          browser.close()
  finally:
      server.terminate()
      server.wait(timeout=10)

  return saved


def main() -> None:
  paths = capture()
  for path in paths:
      print(f"Created: {path}")


if __name__ == "__main__":
  main()
