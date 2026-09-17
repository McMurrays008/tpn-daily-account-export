from __future__ import annotations

import os
from datetime import date
from pathlib import Path

from playwright.sync_api import sync_playwright

HERE = Path(__file__).resolve().parent
DOWNLOADS = HERE / "downloads"
DOWNLOADS.mkdir(exist_ok=True)


def _visible(page, selectors):
    for selector in selectors:
        try:
            locator = page.locator(selector)
            if locator.count() and locator.first.is_visible():
                return locator.first
        except Exception:
            pass
    return None


def _click_named(page, text):
    candidates = [
        page.get_by_role("menuitem", name=text, exact=False),
        page.get_by_role("button", name=text, exact=True),
        page.get_by_role("link", name=text, exact=True),
        page.get_by_text(text, exact=True),
    ]
    for group in candidates:
        try:
            for index in range(group.count()):
                item = group.nth(index)
                if item.is_visible():
                    item.click(no_wait_after=True, timeout=20000)
                    return
        except Exception:
            pass
    raise RuntimeError(f"Could not find visible control: {text}")


def download_today_export(target_day: date) -> Path:
    username = os.getenv("TPN_USERNAME")
    password = os.getenv("TPN_PASSWORD")
    if not username or not password:
        raise RuntimeError("TPN_USERNAME and TPN_PASSWORD are not configured")

    date_text = target_day.strftime(os.getenv("DATE_FORMAT", "%d/%m/%Y"))
    login_url = os.getenv("TPN_LOGIN_URL", "https://pilot.tpnconnect.com/")
    headless = os.getenv("HEADLESS", "true").lower() == "true"

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(
            headless=headless,
            args=[
                "--no-sandbox",
                "--disable-dev-shm-usage",
                "--disable-gpu",
                "--renderer-process-limit=1",
            ],
        )
        context = browser.new_context(
            accept_downloads=True,
            locale="en-GB",
            viewport={"width": 1280, "height": 800},
        )
        page = context.new_page()
        page.set_default_timeout(20000)
        page.set_default_navigation_timeout(45000)
        page.route(
            "**/*",
            lambda route: route.abort()
            if route.request.resource_type in {"image", "media", "font"}
            else route.continue_(),
        )

        try:
            print("[tpn] Opening login", flush=True)
            page.goto(login_url, wait_until="domcontentloaded", timeout=45000)
            user_field = _visible(
                page,
                [
                    "input[placeholder*='Username' i]",
                    "input[name='Username']",
                    "input[name='username']",
                    "#Username",
                    "#username",
                ],
            )
            password_field = _visible(
                page,
                [
                    "input[placeholder*='Password' i]",
                    "input[name='Password']",
                    "input[name='password']",
                    "input[type='password']",
                ],
            )
            if not user_field or not password_field:
                raise RuntimeError("Could not identify TPN login fields")

            user_field.fill(username)
            password_field.fill(password)
            print("[tpn] Submitting login", flush=True)

            login = page.get_by_role("button", name="Login", exact=True)
            if login.count():
                login.first.click(no_wait_after=True, timeout=20000)
            else:
                password_field.press("Enter", no_wait_after=True)

            # TPN navigation can keep network activity open. Wait for the login
            # controls to disappear rather than waiting for networkidle.
            page.locator("input[type='password']:visible").wait_for(
                state="hidden", timeout=45000
            )
            page.wait_for_timeout(1200)
            if page.locator("input[placeholder*='Username' i]:visible").count():
                raise RuntimeError("TPN login page remained visible after sign-in")

            print("[tpn] Opening top-level Browse", flush=True)
            _click_named(page, "Browse")
            page.wait_for_timeout(1000)

            for name in ("DateFrom", "DateTo"):
                field = page.locator(
                    f'input[name="{name}"], input[id*="{name}" i]'
                ).first
                field.wait_for(state="visible", timeout=30000)
                field.fill(date_text)
                page.keyboard.press("Tab")

            # Select All consignments when the control exists; otherwise TPN's
            # existing All/default selection is retained.
            for label in ("Consignments", "Consignment Type", "Browse Type"):
                try:
                    box = page.get_by_label(label, exact=False)
                    if box.count() and box.first.is_visible():
                        box.first.select_option(label="All consignments")
                        break
                except Exception:
                    pass

            print(f"[tpn] Loading Browse for {date_text}", flush=True)
            browse_buttons = page.get_by_role("button", name="Browse", exact=True)
            if browse_buttons.count():
                browse_buttons.last.click(no_wait_after=True, timeout=20000)
            else:
                _click_named(page, "Browse")

            page.locator(".k-grid-content tbody tr, table tbody tr").first.wait_for(
                state="attached", timeout=120000
            )

            print("[tpn] Exporting Excel", flush=True)
            export = page.get_by_text("Export To Excel", exact=False).first
            export.wait_for(state="visible", timeout=30000)
            with page.expect_download(timeout=120000) as download_info:
                export.click(force=True, no_wait_after=True, timeout=20000)
            download = download_info.value
            suffix = Path(download.suggested_filename).suffix.lower()
            if suffix not in {".xls", ".xlsx"}:
                suffix = ".xlsx"
            target = DOWNLOADS / f"TPN_All_{target_day.isoformat()}{suffix}"
            download.save_as(target)
            print(f"[tpn] Downloaded {target.name}", flush=True)
            return target
        except Exception:
            try:
                page.screenshot(path=str(HERE / "tpn_failure.png"), timeout=5000)
            except Exception:
                pass
            raise
        finally:
            context.close()
            browser.close()
