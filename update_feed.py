import asyncio
import os
from playwright.async_api import async_playwright
import openpyxl
import xml.etree.ElementTree as ET
from xml.dom import minidom
from datetime import datetime, date

# Config
LOGIN_URL = "https://login.festool.com/Account/Login"
DOWNLOAD_URL = "https://predajca.festool.sk/b5152248-3b3d-426f-b973-89b78b2022ca"
USERNAME = os.environ.get("FESTOOL_USERNAME", "")
PASSWORD = os.environ.get("FESTOOL_PASSWORD", "")
OUTPUT_FILE = "feed.xml"


async def screenshot(page, name):
    path = f"debug_{name}.png"
    await page.screenshot(path=path, full_page=True)
    print(f"   Screenshot: {path}")


async def download_excel():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            accept_downloads=True,
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        page = await context.new_page()

        # KROK 1: Otvor login stranku
        print("1. Otvarame prihlasovaci formular...")
        await page.goto(LOGIN_URL, wait_until="networkidle", timeout=60000)
        await screenshot(page, "01_login_page")
        print(f"   URL: {page.url}")

        # KROK 2: Vyplnime udaje
        print("2. Zadavame prihlasovacie udaje...")
        await page.wait_for_timeout(2000)

        for sel in ['input[name="Input.Email"]', 'input[name="Email"]', 'input[type="email"]', '#Email']:
            el = await page.query_selector(sel)
            if el and await el.is_visible():
                await el.fill(USERNAME)
                print(f"   Email vyplneny cez: {sel}")
                break

        for sel in ['input[name="Input.Password"]', 'input[name="Password"]', 'input[type="password"]', '#Password']:
            el = await page.query_selector(sel)
            if el and await el.is_visible():
                await el.fill(PASSWORD)
                print(f"   Heslo vyplnene cez: {sel}")
                break

        await screenshot(page, "02_filled_form")

        # KROK 3: Klikni na prihlasenie
        print("3. Prihlasujeme sa...")
        for sel in ['button[type="submit"]', 'input[type="submit"]', '.btn-primary']:
            el = await page.query_selector(sel)
            if el and await el.is_visible():
                await el.click()
                print(f"   Kliknute na: {sel}")
                break

        await page.wait_for_timeout(5000)
        await page.wait_for_load_state("networkidle", timeout=60000)
        await screenshot(page, "03_after_login")
        print(f"   URL po prihlaseni: {page.url}")

        # KROK 4: Prejdi na stranku s cennikom
        print("4. Prechadzame na stranku s cennikom...")
        await page.goto(DOWNLOAD_URL, wait_until="networkidle", timeout=60000)
        await page.wait_for_timeout(3000)
        await screenshot(page, "04_download_page")
        print(f"   URL: {page.url}")

        # Ak nas presmerovalo na login, prihlasenie zlyhalo
        if "login.festool.com" in page.url:
            print("   CHYBA: Prihlasenie zlyhalo - stale sme na login stranke!")
            await browser.close()
            return None

        # KROK 5: Hladame subor na stiahnutie
        print("5. Hladame subor na stiahnutie...")
        download_path = None

        # Metoda 1: Link s .xlsx alebo download
        links = await page.query_selector_all('a[href*=".xlsx"], a[href*="download"], a[href*="export"]')
        if links:
            for link in links:
                try:
                    async with page.expect_download(timeout=30000) as download_info:
                        await link.click()
                    download = await download_info.value
                    download_path = f"downloads/{download.suggested_filename}"
                    os.makedirs("downloads", exist_ok=True)
                    await download.save_as(download_path)
                    print(f"   Subor stiahnuty: {download_path}")
                    break
                except Exception as e:
                    print(f"   Link nefungoval: {e}")
                    continue

        # Metoda 2: Tlacidla s textom download/stiahnuť
        if not download_path:
            elements = await page.query_selector_all('button, a, [role="button"]')
            for el in elements:
                try:
                    text = (await el.inner_text()).lower()
                    if any(w in text for w in ['download', 'stiahnuť', 'stiahnut', 'export', 'xlsx', 'cenník', 'cennik', 'excel']):
                        async with page.expect_download(timeout=30000) as download_info:
                            await el.click()
                        download = await download_info.value
                        download_path = f"downloads/{download.suggested_filename}"
                        os.makedirs("downloads", exist_ok=True)
                        await download.save_as(download_path)
                        print(f"   Subor stiahnuty cez button: {download_path}")
                        break
                except:
                    continue

        # Metoda 3: Mozno stranka priamo ponuka download
        if not download_path:
            print("   Skusame priamy download...")
            try:
                async with page.expect_download(timeout=15000) as download_info:
                    pass
                download = await download_info.value
                download_path = f"downloads/{download.suggested_filename}"
                os.makedirs("downloads", exist_ok=True)
                await download.save_as(download_path)
                print(f"   Automaticky download: {download_path}")
            except:
                await screenshot(page, "05_no_download")
                print("   CHYBA: Nepodarilo sa stiahnut subor!")
                page_content = await page.content()
                with open("debug_page.html", "w") as f:
                    f.write(page_content)
                print("   HTML stranky ulozene do debug_page.html")
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
    print("=" * 50)

    excel_path = await download_excel()

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
