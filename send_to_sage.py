# ─── web_login_demo.py ────────────────────────────────────────────────────
from playwright.sync_api import sync_playwright
from pathlib import Path


EMAIL    = "" #Enter your own email and password bellow
PASSWORD = ""

BASE_URL  = "https://techs.sageserviceops.com"
LOGIN_URL = BASE_URL + "/quote"
POST_LOGIN_EXPECTED = BASE_URL + "/quote"

selectors = {
    "email":  'input#txt-email',
    "pw":     'input#txt-password',
    "login":  'button#btn-login',
    "menu":   'button:has-text("Menu")',
    "new_q":  'a[href="quote_edit"]',
}

# ───────── helpers ───────────────────────────────────────────────────────
def flash(page, selector):
    """Outline element in red for 0.4 s so you can see every step."""
    page.eval_on_selector(
        selector,
        "el => { const o = el.style.outline;"
        "el.style.outline = '3px solid red';"
        "setTimeout(() => el.style.outline = o, 400); }",
    )

def click_visible(page, selector, *, timeout=10_000):
    """Click the first visible desktop copy (ignores hidden mob-icon)."""
    flash(page, selector)
    loc = page.locator(f"{selector}:not(.mob-icon)").first
    loc.wait_for(state="visible", timeout=timeout)
    loc.click()

# ───────── main workflow ────────────────────────────────────────────────
def send_to_sage(data: dict, *, headless: bool = False, slow_mo: int = 0) -> bool:

    excel_folder = Path(data["_SOURCE_FILE"]).parent
    excel_folder.mkdir(exist_ok=True)    

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless, slow_mo=slow_mo)

        context = browser.new_context(accept_downloads=True)
        page = browser.new_page()

        # 1) login ---------------------------------------------------------
        page.goto(LOGIN_URL, wait_until="domcontentloaded")
        flash(page, selectors["email"]);  page.fill(selectors["email"], EMAIL)
        flash(page, selectors["pw"]);     page.fill(selectors["pw"], PASSWORD)
        flash(page, selectors["login"]);  page.click(selectors["login"])
        page.wait_for_url(lambda u: u.startswith(POST_LOGIN_EXPECTED), timeout=15_000)



        # 2) Menu → New Quote ---------------------------------------------
        click_visible(page, selectors["menu"])
        with page.expect_navigation():
            click_visible(page, selectors["new_q"])
        

        # 3) new-quote header ---------------------------------------------
        flash(page, 'input#quote_desc')
        page.fill('input#quote_desc', "Mechanical Proposal")
        click_visible(page, 'button.new_site_lnk')
        flash(page, 'textarea#site_info')
        page.fill('textarea#site_info', data["JOB"])
        with page.expect_navigation():
            click_visible(page, 'button#quote_form_submit')
        print("✓ Quote created →", page.url)
        quote_url = page.url
        quote_id  = quote_url.split("quote_id=")[-1]   

        # 4) CONTACT modal -------------------------------------------------
        if not page.is_visible('#dialog_edit_gen_contact'):
            click_visible(page, 'a:has-text("Edit Contact")')
        page.wait_for_selector('#dialog_edit_gen_contact', state="visible")
        page.fill('input#contact_name',  data.get("CONTACT", ""))
        page.fill('input#contact_phone', data.get("PHONE",   ""))
        page.fill('input#contact_email', data.get("EMAIL",   ""))
        click_visible(page, 'button:has-text("Update Contact")')
        page.wait_for_selector('#dialog_edit_gen_contact', state="hidden")

        # 5) NOTES modal ---------------------------------------------------
        click_visible(page, 'a[href="#tab-notes"]')
        click_visible(page, 'a:has-text("Edit Notes")')
        page.wait_for_selector('#dialog_quote_notes_edit', state="visible")
        page.fill('textarea#customer_viewable_comments', data["SCOPE_OF_WORK"])
        click_visible(page, 'button:has-text("Update General Notes")')
        page.wait_for_selector('#dialog_quote_notes_edit', state="hidden")

        # 6) MISC item -----------------------------------------------------
        click_visible(page, 'a[href="#tab-misc"]')
        click_visible(page, 'a.dialog_quote_add_misc_part')

        # wait until modal body is present
        page.wait_for_selector('#quote_misc_description_new', state="visible")

        page.fill('input#quote_misc_description_new', "Labor & Materials")
        page.fill('input#quote_misc_quantity_new', "1")
        page.select_option('select#quote_misc_rate_sheet_markup_new', value="0")
        page.fill('input#quote_misc_unit_price_new', data["TOTAL_PRICE"])

        # --- click the blue  Save and New  button -------------------------------
        save_and_new_btn = page.locator('button.update_quote_misc.no-close')
        save_and_new_btn.wait_for(state="visible")   # ensure we get the visible one
        save_and_new_btn.click(force=True)

        # press Esc to close the modal
        page.keyboard.press("Escape")

        # wait until the modal closes before moving on
        page.wait_for_selector('#quote_misc_description_new', state="hidden")

        # 7) REVIEW / PRINT -----------------------------------------------
        click_visible(page, 'a[href="#tab-review"]')
        with page.expect_navigation():
            click_visible(page, 'a:has-text("Print")')
        page.click('label[for="cmn-toggle-quantity"]')
        page.click('label[for="cmn-toggle-uprice"]')
        with page.expect_navigation():
            click_visible(page, 'a#save_quote')

        # 8) ATTACHMENTS download -----------------------------------------
        click_visible(page, 'a[href="#tab-content"]')

        # wait until the tab body is the active pane (table rendered)
        page.wait_for_selector('#tab-content.active', timeout=15_000)

        # locate the first *visible* dropdown toggle inside that pane
        dropdown = page.locator('#tab-content button.dropdown-toggle').first
        dropdown.wait_for(state="visible", timeout=15_000)
        dropdown.scroll_into_view_if_needed()
        dropdown.click()

        with page.expect_download() as dl2:
            page.locator('#tab-content a:has-text("Download")').click()

        file2 = dl2.value
        file2_path = excel_folder / file2.suggested_filename     # build full path
        file2.save_as(file2_path)                                # save in same folder as Excel
        print("Attachment saved →", file2_path)

        page.wait_for_timeout(3_000)
        browser.close()
    return {"pdf": file2_path, "quote_id": quote_id}


