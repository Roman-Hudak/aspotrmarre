import asyncio
import os
from playwright.async_api import async_playwright
import openpyxl
import xml.etree.ElementTree as ET
from xml.dom import minidom
import glob

# Config
LOGIN_URL = "https://login.festool.com/Account/Login"
DOWNLOAD_URL = "https://predajca.festool.sk/b5152248-3b3d-426f-b973-89b78b2022ca"
USERNAME = os.environ.get("FESTOOL_USERNAME", "")
PASSWORD = os.environ.get("FESTOOL_PASSWORD", "")
OUTPUT_FILE = "feed.xml"

async def download_excel():
    """Prihlasi sa na Festool portal a stiahne Excel cennik."""
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(accept_downloads=True)
        page = await context.new_page()

        print("1. Otvarame prihlasovaci formular...")
        await page.goto(DOWNLOAD_URL, wait_until="networkidle", timeout=60000)

        # Pockame na formular
        await page.wait_for_selector('input[type="email"], input[name="Email"], input[name="Username"], input[name="Input.Email"], #Email, #Username', timeout=30000)

        print("2. Zadavame prihlasovacie udaje...")
        # Skusime rozne varianty input fieldov
        email_selectors = ['input[name="Input.Email"]', 'input[name="Email"]', 'input[type="email"]', '#Email', '#Username']
        password_selectors = ['input[name="Input.Password"]', 'input[name="Password"]', 'input[type="password"]', '#Password']

        for sel in email_selectors:
            try:
                el = await page.query_selector(sel)
                if el:
                    await el.fill(USERNAME)
                    print(f"   Email vyplneny cez: {sel}")
                    break
            except:
                continue

        for sel in password_selectors:
            try:
                el = await page.query_selector(sel)
                if el:
                    await el.fill(PASSWORD)
                    print(f"   Heslo vyplnene cez: {sel}")
                    break
            except:
                continue

        # Klikni na prihlasenie
        submit_selectors = ['button[type="submit"]', 'input[type="submit"]', '.btn-primary']
        for sel in submit_selectors:
            try:
                el = await page.query_selector(sel)
                if el:
                    await el.click()
                    print(f"   Kliknute na: {sel}")
                    break
            except:
                continue

        print("3. Cakame na presmerovanie po prihlaseni...")
        await page.wait_for_load_state("networkidle", timeout=60000)

        # Ak sme boli presmerovani na login, skusime znova ist na download URL
        if "login.festool.com" in page.url:
            print("   Este sme na login stranke, cakame...")
            await page.wait_for_url("**/predajca.festool.sk/**", timeout=60000)

        # Ak nie sme na spravnej stranke, prejdeme tam
        if DOWNLOAD_URL not in page.url:
            print(f"4. Prechadzame na stranku s cennikom: {DOWNLOAD_URL}")
            await page.goto(DOWNLOAD_URL, wait_until="networkidle", timeout=60000)

        print("5. Hladame tlacidlo na stiahnutie...")
        await page.wait_for_timeout(3000)

        # Skusime najst download link/button
        download_path = None

        # Metoda 1: Hladame link s .xlsx
        xlsx_links = await page.query_selector_all('a[href*=".xlsx"], a[href*="download"], a[href*="export"]')
        if xlsx_links:
            async with page.expect_download(timeout=60000) as download_info:
                await xlsx_links[0].click()
            download = await download_info.value
            download_path = f"downloads/{download.suggested_filename}"
            await download.save_as(download_path)
            print(f"   Subor stiahnuty: {download_path}")
        else:
            # Metoda 2: Skusime vsetky tlacidla s textom download/stiahnout/export
            buttons = await page.query_selector_all('button, a.btn, .download, [class*="download"], [class*="export"]')
            for btn in buttons:
                text = await btn.inner_text()
                if any(word in text.lower() for word in ['download', 'stiahnuť', 'stiahnut', 'export', 'xlsx', 'cenník', 'cennik']):
                    async with page.expect_download(timeout=60000) as download_info:
                        await btn.click()
                    download = await download_info.value
                    download_path = f"downloads/{download.suggested_filename}"
                    await download.save_as(download_path)
                    print(f"   Subor stiahnuty cez button: {download_path}")
                    break

        if not download_path:
            # Metoda 3: Mozno sa subor stiahne automaticky pri navsteve stranky
            print("   Skusame priamy download zo stranky...")
            try:
                async with page.expect_download(timeout=30000) as download_info:
                    await page.reload()
                download = await download_info.value
                download_path = f"downloads/{download.suggested_filename}"
                await download.save_as(download_path)
                print(f"   Subor stiahnuty automaticky: {download_path}")
            except:
                # Ulozime screenshot pre debugging
                await page.screenshot(path="debug_screenshot.png")
                print("   CHYBA: Nepodarilo sa stiahnut subor!")
                print(f"   Aktualna URL: {page.url}")
                print("   Screenshot ulozeny ako debug_screenshot.png")
                await browser.close()
                return None

        await browser.close()
        return download_path


def generate_feed(excel_path):
    """Vygeneruje Mergado XML feed z Excel suboru."""
    wb = openpyxl.load_workbook(excel_path)
    ws = wb[wb.sheetnames[0]]
    headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]

    root = ET.Element("CHANNEL")
    root.set("xmlns", "http://www.mergado.com/ns/1.10")

    ET.SubElement(root, "LINK").text = "https://github.com/Roman-Hudak/aspotrmarre"
    ET.SubElement(root, "GENERATOR").text = "custom_feed_generator_1.0"

    count = 0
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        data = dict(zip(headers, row))
        if not data.get('Obj. \u010d\u00edslo') or not data.get('Opis'):
            continue

        item = ET.SubElement(root, "ITEM")
        ET.SubElement(item, "ITEM_ID").text = str(data.get('Obj. \u010d\u00edslo', ''))
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

        for label, key in [('V\u00fd\u0161ka', 'V\u00fd\u0161ka'), ('\u0160\u00edrka', '\u0160\u00edrka'), ('D\u013a\u017eka', 'D\u013a\u017eka'), ('Hmotnos\u0165', 'Hmotnos\u0165')]:
            if data.get(key):
                param = ET.SubElement(item, "PARAM")
                ET.SubElement(param, "n").text = label
                ET.SubElement(param, "VALUE").text = str(data[key])

        if data.get('CoO'):
            ET.SubElement(item, "COUNTRY_OF_ORIGIN").text = str(data['CoO'])
        if data.get('Stav'):
            if data['Stav'] == 'Aktu\u00e1lne':
                ET.SubElement(item, "AVAILABILITY").text = "in stock"
            else:
                ET.SubElement(item, "AVAILABILITY").text = str(data['Stav'])

        ET.SubElement(item, "CONDITION").text = "new"

        if data.get('Partnersk\u00e1 z\u013eava') is not None:
            param = ET.SubElement(item, "PARAM")
            ET.SubElement(param, "n").text = "Partnersk\u00e1 z\u013eava"
            ET.SubElement(param, "VALUE").text = str(data['Partnersk\u00e1 z\u013eava'])

        if data.get('D\u00e1tum dodania'):
            val = data['D\u00e1tum dodania']
            if hasattr(val, 'strftime'):
                ET.SubElement(item, "DELIVERY_DATE").text = val.strftime('%Y-%m-%d')
            else:
                ET.SubElement(item, "DELIVERY_DATE").text = str(val)

        if data.get('Toolpoints'):
            param = ET.SubElement(item, "PARAM")
            ET.SubElement(param, "n").text = "Toolpoints"
            ET.SubElement(param, "VALUE").text = str(data['Toolpoints'])

        count += 1

    xml_str = minidom.parseString(ET.tostring(root, encoding='unicode')).toprettyxml(indent="  ")
    lines = xml_str.split('\n')
    xml_str = '<?xml version="1.0" encoding="utf-8"?>\n' + '\n'.join(lines[1:])

    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        f.write(xml_str)

    print(f"Feed vygenerovany: {count} produktov -> {OUTPUT_FILE}")
    return count


async def main():
    os.makedirs("downloads", exist_ok=True)

    print("=" * 50)
    print("FESTOOL XML FEED GENERATOR")
    print("=" * 50)

    # Stiahni Excel
    excel_path = await download_excel()

    if excel_path:
        # Vygeneruj feed
        count = generate_feed(excel_path)
        if count > 0:
            print(f"\nUSPECH! Feed s {count} produktmi bol vygenerovany.")
        else:
            print("\nCHYBA: Ziadne produkty neboli najdene v subore.")
    else:
        print("\nCHYBA: Nepodarilo sa stiahnut Excel subor.")
        exit(1)

if __name__ == "__main__":
    asyncio.run(main())
