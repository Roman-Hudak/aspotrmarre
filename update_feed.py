import asyncio
import os
import sys
import glob as globmod
from playwright.async_api import async_playwright
import openpyxl
import xml.etree.ElementTree as ET
from xml.dom import minidom
from datetime import datetime, date

# Config
LOGIN_URL = "https://predajca.festool.sk"
NETTO_CENNIK_URL = "https://predajca.festool.sk/objedn%C3%A1vky/netto-cenn%C3%ADk-pre-predajcov-"
USERNAME = os.environ.get("FESTOOL_USERNAME", "")
PASSWORD = os.environ.get("FESTOOL_PASSWORD", "")
OUTPUT_FILE = "feed.xml"
DEBUG_LOCAL = os.environ.get("DEBUG_LOCAL", "0") == "1"


async def screenshot(page, name):
    path = f"debug_{name}.png"
    await page.screenshot(path=path, full_page=True)
    print(f"   Screenshot: {path}")


async def dismiss_cookies(page):
    """Odklikne cookie consent banner ak sa zobrazi."""
    try:
        await page.wait_for_timeout(2000)
        for txt in ["Prijať všetky", "Prijat všetky", "Accept All", "Accept all"]:
            cookie_btn = await page.query_selector(f'button:has-text("{txt}")')
            if cookie_btn and await cookie_btn.is_visible():
                await cookie_btn.click()
                print("   Cookie consent odkliknuty!")
                await page.wait_for_timeout(1000)
                return
    except Exception as e:
        print(f"   Cookie handling: {e}")


async def download_excel():
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=not DEBUG_LOCAL,
            slow_mo=500 if DEBUG_LOCAL else 0
        )
        context = await browser.new_context(
            accept_downloads=True,
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        page = await context.new_page()

        # KROK 1: Otvorime dealer portal - presmeruje na login
        print("1. Otvarame dealer portal (presmeruje na login)...")
        await page.goto(LOGIN_URL, wait_until="networkidle", timeout=60000)
        await page.wait_for_timeout(3000)
        await screenshot(page, "01_redirected_login")
        print(f"   URL: {page.url}")

        await dismiss_cookies(page)
        await screenshot(page, "01b_after_cookies")

        # KROK 2: Login ak treba
        if "login.festool.com" in page.url:
            print("2. Zadavame prihlasovacie udaje...")
            await page.wait_for_timeout(2000)

            # Email
            email_filled = False
            for sel in ['input[name="Input.Email"]', 'input[name="Email"]', 'input[type="email"]', '#Email', '#Input_Email']:
                try:
                    el = await page.query_selector(sel)
                    if el and await el.is_visible():
                        await el.fill(USERNAME)
                        print(f"   Email vyplneny cez: {sel}")
                        email_filled = True
                        break
                except:
                    continue

            # Heslo
            password_filled = False
            for sel in ['input[name="Input.Password"]', 'input[name="Password"]', 'input[type="password"]']:
                try:
                    el = await page.query_selector(sel)
                    if el and await el.is_visible():
                        await el.fill(PASSWORD)
                        print(f"   Heslo vyplnene cez: {sel}")
                        password_filled = True
                        break
                except:
                    continue

            if not email_filled or not password_filled:
                print(f"   CHYBA: email_filled={email_filled}, password_filled={password_filled}")
                await screenshot(page, "02_error")
                await browser.close()
                return None

            await screenshot(page, "02_filled_form")

            # KROK 3: Kliknutie na prihlasenie
            print("3. Prihlasujeme sa...")
            for sel in ['button[type="submit"]', 'input[type="submit"]', '.btn-primary', 'button.btn']:
                try:
                    el = await page.query_selector(sel)
                    if el and await el.is_visible():
                        await el.click()
                        print(f"   Kliknute na: {sel}")
                        break
                except:
                    continue

            await page.wait_for_timeout(5000)
            await page.wait_for_load_state("networkidle", timeout=60000)
            await screenshot(page, "03_after_login")
            print(f"   URL po prihlaseni: {page.url}")

            if "login.festool.com" in page.url:
                print("   CHYBA: Prihlasenie zlyhalo!")
                await browser.close()
                return None
        else:
            print("   Uz sme prihlaseni, pokracujeme...")

        # KROK 4: Navigacia na Netto cennik
        print("4. Navigujeme na Netto cennik pre predajcov...")
        await page.goto(NETTO_CENNIK_URL, wait_until="networkidle", timeout=60000)
        await page.wait_for_timeout(3000)
        await dismiss_cookies(page)
        await screenshot(page, "04_netto_cennik")
        print(f"   URL: {page.url}")

        # KROK 5: Kliknutie na "Excel export všetko"
        print("5. Hladame tlacidlo 'Excel export všetko'...")
        download_path = None

        for sel in [
            'button:has-text("Excel export všetko")',
            'a:has-text("Excel export všetko")',
            'button:has-text("Excel export v")',
            'a:has-text("Excel export v")',
        ]:
            try:
                el = await page.query_selector(sel)
                if el and await el.is_visible():
                    print(f"   Nasiel som tlacidlo: {sel}")
                    async with page.expect_download(timeout=120000) as download_info:
                        await el.click()
                    download = await download_info.value
                    download_path = f"downloads/{download.suggested_filename}"
                    os.makedirs("downloads", exist_ok=True)
                    await download.save_as(download_path)
                    print(f"   Subor stiahnuty: {download_path}")
                    break
            except Exception as e:
                print(f"   Selektor {sel} - chyba: {e}")
                continue

        # Fallback - hladame vsetky buttony/linky s textom "export"
        if not download_path:
            print("   Fallback: prehladavame vsetky elementy...")
            elements = await page.query_selector_all('button, a, [role="button"]')
            for el in elements:
                try:
                    text = (await el.inner_text()).strip().lower()
                    if 'export' in text and 'všetko' in text:
                        print(f"   Nasiel som element s textom: '{text}'")
                        async with page.expect_download(timeout=120000) as download_info:
                            await el.click()
                        download = await download_info.value
                        download_path = f"downloads/{download.suggested_filename}"
                        os.makedirs("downloads", exist_ok=True)
                        await download.save_as(download_path)
                        print(f"   Subor stiahnuty: {download_path}")
                        break
                except:
                    continue

        if not download_path:
            await screenshot(page, "05_no_download")
            print("   CHYBA: Nepodarilo sa najst/stiahnut subor!")
            # Debug - vypiseme vsetky viditelne elementy
            elements = await page.query_selector_all('button, a')
            for el in elements:
                try:
                    text = (await el.inner_text()).strip()
                    if text:
                        print(f"   ELEMENT: '{text}'")
                except:
                    pass
            await browser.close()
            return None

        await browser.close()
        return download_path


def generate_feed(excel_path):
    wb = openpyxl.load_workbook(excel_path)
    ws = wb[wb.sheetnames[0]]
    headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]

    today = date.today()
    print(f"Dnesny datum: {today}")

    root = ET.Element("CHANNEL")
    root.set("xmlns", "http://www.mergado.com/ns/1.10")

    ET.SubElement(root, "LINK").text = "https://github.com/Roman-Hudak/aspotrmarre"
    ET.SubElement(root, "GENERATOR").text = "custom_feed_generator_1.0"

    count = 0
    in_stock = 0
    out_stock = 0

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        data = dict(zip(headers, row))
        if not data.get('Obj. číslo') or not data.get('Opis'):
            continue

        item = ET.SubElement(root, "ITEM")
        ET.SubElement(item, "ITEM_ID").text = str(data.get('Obj. číslo', ''))
        ET.SubElement(item, "NAME_EXACT").text = str(data.get('Opis', ''))

        if data.get('Cena EUR') is not None:
            ET.SubElement(item, "PRICE_VAT").text = str(data['Cena EUR'])
        if data.get('Netto NC EUR') is not None:
            ET.SubElement(item, "COST").text = str(data['Netto NC EUR'])

        ET.SubElement(item, "CURRENCY").text = "EUR"

        if data.get('EAN'):
            ET.SubElement(item, "EAN").text = str(data['EAN'])
        if data.get('Typ'):
            ET.SubElement(item, "CATEGORY").text = str(data['Typ'])
        if data.get('Hierarchia produktov'):
            ET.SubElement(item, "CATEGORYTEXT").text = str(data['Hierarchia produktov'])

        for label, key in [('Výška', 'Výška'), ('Šírka', 'Šírka'), ('Dĺžka', 'Dĺžka'), ('Hmotnosť', 'Hmotnosť')]:
            if data.get(key):
                param = ET.SubElement(item, "PARAM")
                ET.SubElement(param, "n").text = label
                ET.SubElement(param, "VALUE").text = str(data[key])

        if data.get('CoO'):
            ET.SubElement(item, "COUNTRY_OF_ORIGIN").text = str(data['CoO'])

        # AVAILABILITY: in stock iba ak datum dodania <= dnes
        delivery_date = data.get('Dátum dodania')
        if delivery_date and hasattr(delivery_date, 'date'):
            delivery = delivery_date.date()
        elif delivery_date and hasattr(delivery_date, 'strftime'):
            delivery = delivery_date
        else:
            delivery = None

        if delivery and delivery <= today:
            ET.SubElement(item, "AVAILABILITY").text = "in stock"
            in_stock += 1
        else:
            ET.SubElement(item, "AVAILABILITY").text = "out of stock"
            out_stock += 1

        ET.SubElement(item, "CONDITION").text = "new"

        if data.get('Partnerská zľava') is not None:
            param = ET.SubElement(item, "PARAM")
            ET.SubElement(param, "n").text = "Partnerská zľava"
            ET.SubElement(param, "VALUE").text = str(data['Partnerská zľava'])

        if delivery:
            ET.SubElement(item, "DELIVERY_DATE").text = delivery.strftime('%Y-%m-%d')

        if data.get('Toolpoints'):
            param = ET.SubElement(item, "PARAM")
            ET.SubElement(param, "n").text = "Toolpoints"
            ET.SubElement(param, "VALUE").text = str(data['Toolpoints'])

        count += 1

    xml_str = minidom.parseString(ET.tostring(root, encoding='unicode')).toprettyxml(indent="  ")
    lines = xml_str.split(chr(10))
    xml_str = '<?xml version="1.0" encoding="utf-8"?>\n' + chr(10).join(lines[1:])

    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        f.write(xml_str)

    print(f"Feed vygenerovany: {count} produktov")
    print(f"  IN STOCK:     {in_stock}")
    print(f"  OUT OF STOCK: {out_stock}")
    return count


async def main():
    os.makedirs("downloads", exist_ok=True)

    print("=" * 50)
    print("FESTOOL XML FEED GENERATOR")
    if DEBUG_LOCAL:
        print(">>> LOKALNY DEBUG MOD <<<")
    print("=" * 50)

    # Ak je zadany argument s cestou k excelu, pouzijeme ho priamo
    if len(sys.argv) > 1:
        excel_path = sys.argv[1]
        print(f"Pouzivam zadany Excel: {excel_path}")
    else:
        excel_path = await download_excel()

        # Ak download zlyhal a sme v debug mode, skusime najst lokalny xlsx
        if not excel_path and DEBUG_LOCAL:
            local_files = globmod.glob("*.xlsx") + globmod.glob("downloads/*.xlsx")
            if local_files:
                excel_path = local_files[0]
                print(f"Download zlyhal, pouzivam lokalny subor: {excel_path}")

    if excel_path:
        count = generate_feed(excel_path)
        if count > 0:
            print(f"\nUSPECH! Feed s {count} produktmi bol vygenerovany.")
        else:
            print("\nCHYBA: Ziadne produkty neboli najdene v subore.")
            exit(1)
    else:
        print("\nCHYBA: Nepodarilo sa stiahnut Excel subor.")
        exit(1)

if __name__ == "__main__":
    asyncio.run(main())
