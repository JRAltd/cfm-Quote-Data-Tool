import zipfile
import xml.etree.ElementTree as ET
import re
import json
import base64
import os

xlsx_path = r"Quote Data For Release 092126.xlsx"
logo_path = r"cfm_logo.png"
output_html = r"cfm_quote_tool.html"

def load_logo_b64():
    if os.path.exists(logo_path):
        with open(logo_path, "rb") as f:
            return base64.b64encode(f.read()).decode("ascii")
    return ""

def is_number(s):
    if not s:
        return False
    s_clean = s.replace("$", "").replace(",", "").strip()
    try:
        float(s_clean)
        return True
    except:
        return False

def clean_num(val):
    if not val:
        return ""
    val = val.replace("$", "").replace(",", "").strip()
    try:
        f = float(val)
        if f <= 0:
            return ""
        return f"{f:.2f}" if f != int(f) else f"{int(f)}"
    except:
        return ""

def sanitize_text(text):
    if not text:
        return ""
    t = re.sub(r'\bJulie(?:\s+O\'?Flaherty|\s+O)?\b', 'CFM Commercial Support', text, flags=re.IGNORECASE)
    t = re.sub(r'\bJulie\s+Notes\b', 'Technical Notes', t, flags=re.IGNORECASE)
    t = re.sub(r'\bJulie\b', 'CFM Team', t, flags=re.IGNORECASE)
    t = re.sub(r'\bUpdates?\s+on\s+this\s+Version\s+Release[^\n.]*', '', t, flags=re.IGNORECASE)
    t = re.sub(r'\bPrice\s+change\s+updates?[^\n.]*', '', t, flags=re.IGNORECASE)
    t = re.sub(r'\bUpdated\s+Commercial\s+IPA\s+Guidance\b', 'Commercial IPA Guidance', t, flags=re.IGNORECASE)
    t = re.sub(r'\bUpdated\s+RTU\s+stock\s+page[^\n.]*', '', t, flags=re.IGNORECASE)
    t = re.sub(r'\bUpdated\b', 'Effective', t, flags=re.IGNORECASE)
    t = re.sub(r'\bUpdates?\b', 'Revisions', t, flags=re.IGNORECASE)
    return t.strip()

def categorize_accessory(desc, model):
    d = (desc + " " + model).lower()
    if any(k in d for k in ['curb', 'curb adapter', 'transition']):
        return 'Roof Curbs & Adapters'
    elif any(k in d for k in ['economizer', 'damper', 'fresh air', 'relief', 'barometric', 'hood']):
        return 'Economizers & Dampers'
    elif any(k in d for k in ['hail guard', 'coil guard', 'burglar bar', 'wire guard']):
        return 'Hail & Coil Guards'
    elif any(k in d for k in ['smoke detector', 'co2', 'sensor', 'enthalpy', 'temperature sensor', 'humidity']):
        return 'Smoke Detectors & Sensors'
    elif any(k in d for k in ['bacnet', 'smart equipment', 'secom', 'control', 'thermostat', 'vfd', 'remote']):
        return 'Controls & Comm Cards'
    elif any(k in d for k in ['disconnect', 'fuse', 'transformer', 'phase monitor', 'low ambient', 'electric heat', 'heater kit', 'wiring', 'cable']):
        return 'Electrical & Power'
    elif any(k in d for k in ['propane', 'lp conversion', 'gas conversion', 'flue', 'exhaust kit']):
        return 'Gas & Flue Kits'
    elif any(k in d for k in ['pump', 'drain', 'valve', 'vent', 'filter', 'expansion tank']):
        return 'Pumps & Piping'
    else:
        return 'General Accessories'

def parse_workbook():
    print("Parsing Quote Data workbook and extracting Equipment & Paired Accessories...")
    with zipfile.ZipFile(xlsx_path, "r") as z:
        shared_strings = []
        if "xl/sharedStrings.xml" in z.namelist():
            tree = ET.parse(z.open("xl/sharedStrings.xml"))
            for si in tree.getroot():
                text_parts = [elem.text or "" for elem in si.iter() if elem.tag.endswith("}t")]
                shared_strings.append("".join(text_parts))

        wb_tree = ET.parse(z.open("xl/workbook.xml"))
        sheets = []
        for elem in wb_tree.getroot().iter():
            if elem.tag.endswith("}sheet"):
                sheets.append((
                    elem.attrib.get("name"),
                    elem.attrib.get("sheetId"),
                    elem.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                ))

        rels_tree = ET.parse(z.open("xl/_rels/workbook.xml.rels"))
        rel_map = {elem.attrib["Id"]: elem.attrib["Target"] for elem in rels_tree.getroot()}

        catalog_items = []
        matrix_data = {}
        reference_sheets = {}
        sheet_stats = []

        ref_sheet_names = {
            "Stocking Location Craneyards",
            "York Lead Time Tracker",
            "Replacing a York Millennium RTU",
            "York Norman Mod Shop",
            "Warranty Plans",
            "Verbiage Quote Entry",
            "Evergy RebatesComm KS 2026",
            "Evergy Rebates Comm MO 2026",
            "Source 1 Compressor Cross Ref",
            "Cross Ref Others to York Model#",
            "Freight Estimation",
            "Freight Damage",
            "Item Classes",
            "Vendor Codes",
            "BABA Build America Act"
        }

        matrix_sheet_names = {
            "Matrix on stock ",
            "Matrix on factory order models"
        }

        excluded_sheets = {"Updates"}

        # Series Catalog Definition
        series_catalog = {
            'sunline-3-6': {
                'title': 'York Sunline (3 - 6 Ton)',
                'badge': 'RTU 3-6T',
                'desc': 'Commercial packaged rooftop units (Gas/Electric, Electric/Electric, Heat Pump)',
                'accessories': []
            },
            'sunpro-3-12': {
                'title': 'York Sun Pro (3 - 12.5 Ton)',
                'badge': 'RTU 3-12.5T',
                'desc': 'High-efficiency commercial packaged units with Smart Equipment controls',
                'accessories': []
            },
            'suncore-sm': {
                'title': 'York SunCore / Prestige (Small Cab 3 - 6 Ton)',
                'badge': 'RTU Small Cab',
                'desc': 'Standard efficiency commercial RTU small footprint cabinets',
                'accessories': []
            },
            'suncore-lg': {
                'title': 'York SunCore / Prestige (Large Cab 7.5 - 12.5 Ton)',
                'badge': 'RTU Large Cab',
                'desc': 'Standard efficiency commercial RTU large footprint cabinets',
                'accessories': []
            },
            'sunchoice-15-27': {
                'title': 'York SunChoice (15 - 27.5 Ton)',
                'badge': 'RTU 15-27.5T',
                'desc': 'Commercial packaged units AV15, AV18, AV20, AV25, AV28',
                'accessories': []
            },
            'sunselect-27-50': {
                'title': 'York Sun Select (27.5 - 50 Ton)',
                'badge': 'RTU 27.5-50T',
                'desc': 'Large tonnage commercial rooftop packaged systems',
                'accessories': []
            },
            'sunpremier-25-150': {
                'title': 'York Sun Premier (25 - 150 Ton)',
                'badge': 'RTU Applied',
                'desc': 'Custom commercial and applied packaged rooftop systems',
                'accessories': []
            },
            'lx-pkg-2-5': {
                'title': 'York LX Series Packaged Units (2 - 5 Ton)',
                'badge': 'Package 2-5T',
                'desc': 'Residential & light commercial packaged units (PCE, PHE, PCG, PHG)',
                'accessories': []
            },
            'csplit-2-50': {
                'title': 'Commercial Split Systems (2.5 - 50 Ton)',
                'badge': 'Commercial Splits',
                'desc': 'Commercial condensing units and matched air handlers',
                'accessories': []
            },
            'res-splits-furnace': {
                'title': 'Residential Furnaces & Splits',
                'badge': 'Furnaces & Splits',
                'desc': 'Gas furnaces, modular air handlers, and residential split accessories',
                'accessories': []
            },
            'reznor-heaters': {
                'title': 'Reznor Unit Heaters & Duct Furnaces',
                'badge': 'Reznor Heaters',
                'desc': 'UDXC, UDZ, and HX series gas-fired unit heaters and duct furnaces',
                'accessories': []
            },
            'amana-ptac': {
                'title': 'Amana PTAC Systems',
                'badge': 'Amana PTAC',
                'desc': 'Packaged Terminal Air Conditioners and heat pumps',
                'accessories': []
            },
            'boilers': {
                'title': 'Commercial & Residential Boilers',
                'badge': 'Boilers & Hydronics',
                'desc': 'Hydronic boilers, expansion tanks, and piping specialties',
                'accessories': []
            },
            'lg-ductless': {
                'title': 'LG Ductless Systems',
                'badge': 'LG Ductless',
                'desc': 'Single and multi-zone mini split heat pumps and accessories',
                'accessories': []
            }
        }

        # Accessory section regex mapping
        section_to_series = [
            (r'3\s*-\s*6\s*ton\s*sunline\s*package\s*accessories', 'sunline-3-6'),
            (r'sunline.*field\s*installed\s*options', 'sunline-3-6'),
            (r'sun\s*pro\s*3\s*-\s*12\.5\s*ton.*accessories', 'sunpro-3-12'),
            (r'sun\s*pro.*field\s*installed\s*accessories', 'sunpro-3-12'),
            (r'sun\s*core.*small\s*cabinet\s*package\s*accessories', 'suncore-sm'),
            (r'prestige\s*small\s*cabinet\s*package\s*accessories', 'suncore-sm'),
            (r'sun\s*core.*large\s*cabinet\s*package\s*accessories', 'suncore-lg'),
            (r'prestige\s*large\s*cabinet\s*package\s*accessories', 'suncore-lg'),
            (r'av15\s*and\s*av20\s*sun\s*choice\s*stocked\s*accessories', 'sunchoice-15-27'),
            (r'sun\s*choice.*field\s*installed\s*options', 'sunchoice-15-27'),
            (r'sun\s*select.*field\s*installed|sun\s*select.*accessories', 'sunselect-27-50'),
            (r'sun\s*premier.*accessories|sun\s*premier.*options', 'sunpremier-25-150'),
            (r'lx\s*series\s*cabinet\s*package\s*accessories', 'lx-pkg-2-5'),
            (r'lx\s*series.*field\s*installed\s*accessories', 'lx-pkg-2-5'),
            (r'7\.5\s*-\s*25\s*ton\s*condensing\s*unit\s*field\s*installed\s*accessories', 'csplit-2-50'),
            (r'commercial\s*splits.*field\s*installed', 'csplit-2-50'),
            (r'furnace\s*field\s*installed\s*accessories', 'res-splits-furnace'),
            (r'basic\s*ptac\s*accessories', 'amana-ptac'),
            (r'accessories\s*for\s*all\s*boilers', 'boilers'),
            (r'lg\s*stocked\s*accessories|ductless\s*accessories', 'lg-ductless'),
            (r'cfm\s*stock\s*accessories\s*for\s*the\s*hx\s*duct\s*furnace|field\s*installed\s*options\s*for\s*gravity\s*vented\s*duct\s*heaters|reznor.*accessories\s*&\s*options', 'reznor-heaters')
        ]

        all_accessory_models = set()

        # PASS 1: Extract all accessories for each series
        for name, sid, rid in sheets:
            if name in excluded_sheets or name in matrix_sheet_names or name in ref_sheet_names:
                continue
            target = rel_map[rid]
            sheet_path = "xl/" + target if not target.startswith("xl/") else target
            if sheet_path not in z.namelist():
                continue

            sheet_tree = ET.parse(z.open(sheet_path))
            current_s_key = None

            for row in sheet_tree.getroot().iter():
                if not row.tag.endswith("}row"):
                    continue
                cells = {}
                for c in row:
                    if not c.tag.endswith("}c"):
                        continue
                    r_attr = c.attrib.get("r", "")
                    col_letter = "".join([ch for ch in r_attr if ch.isalpha()])
                    t_attr = c.attrib.get("t")
                    val = ""
                    for child in c:
                        if child.tag.endswith("}v") and child.text:
                            if t_attr == "s":
                                try:
                                    s_idx = int(child.text)
                                    val = shared_strings[s_idx] if s_idx < len(shared_strings) else ""
                                except:
                                    pass
                            else:
                                val = child.text
                    if val.strip():
                        cells[col_letter] = sanitize_text(val.strip())

                line = " ".join(cells.values())
                found_s = False
                for pat, skey in section_to_series:
                    if re.search(pat, line, re.I):
                        current_s_key = skey
                        found_s = True
                        break

                if not found_s:
                    if '3-6 ton sunline' in name.lower() and 'field installed' in line.lower():
                        current_s_key = 'sunline-3-6'
                    elif '3-12.5t sun pro' in name.lower() and 'field installed' in line.lower():
                        current_s_key = 'sunpro-3-12'
                    elif 'sunchoice' in name.lower() and 'field installed' in line.lower():
                        current_s_key = 'sunchoice-15-27'
                    elif 'sun select' in name.lower() and ('economizer' in line.lower() or 'hail guard' in line.lower()):
                        current_s_key = 'sunselect-27-50'
                    elif 'sun premier' in name.lower() and ('economizer' in line.lower() or 'curb' in line.lower()):
                        current_s_key = 'sunpremier-25-150'
                    elif '2.5-50 splits' in name.lower() and 'field installed' in line.lower():
                        current_s_key = 'csplit-2-50'

                if current_s_key and len(cells) >= 2:
                    m = cells.get("B") or cells.get("A")
                    d = cells.get("D") or cells.get("C") or cells.get("E")
                    if m and len(m) >= 3 and not any(k in m.lower() for k in ['availab', 'model', 'note', 'price', 'total', 'target sell', 'standard sell', 'below is']):
                        if d and not any(k in d.lower() for k in ['target sell', 'standard sell', 'for non-stock']):
                            price = 0.0
                            for col_k in ['E', 'G', 'H', 'L']:
                                try:
                                    v_clean = re.sub(r'[^\d.]', '', cells.get(col_k, ''))
                                    if v_clean and float(v_clean) > 0:
                                        price = float(v_clean)
                                        break
                                except:
                                    pass

                            avail = cells.get('A', '')
                            stock = '*Norm Stk' if '*Norm Stk' in avail or 'In Stock' in avail else ('KC' if 'KC' in avail else 'Call')
                            cat = categorize_accessory(d, m)

                            # Existing models in this series
                            existing = {x['m'] for x in series_catalog[current_s_key]['accessories']}
                            if m not in existing:
                                series_catalog[current_s_key]['accessories'].append({
                                    'm': m.strip(),
                                    'd': d.strip(),
                                    'c': cat,
                                    'price': price,
                                    'stock': stock,
                                    's': name
                                })
                                all_accessory_models.add(m.strip().upper())

        # Function to map equipment item to series key
        def get_series_key(item_dict):
            m = (item_dict.get('m') or '').upper()
            d = (item_dict.get('d') or '').upper()
            s = (item_dict.get('s') or '').lower()
            f = (item_dict.get('f') or '').lower()
            c = (item_dict.get('c') or '').lower()

            # RTU Series
            if 'sunline' in d.lower() or 'sunline' in s or 'sunline' in f or any(m.startswith(p) for p in ['KE036', 'KE048', 'KE060', 'ZF036', 'ZF048', 'ZF060', 'ZR036', 'ZR048', 'ZR060', 'XP036', 'XP048', 'XP060']):
                return 'sunline-3-6'
            if 'sun pro' in d.lower() or 'sun pro' in s or 'sun pro' in f or any(m.startswith(p) for p in ['KJ', 'KB', 'KT', 'WP078', 'WP090', 'WP102', 'WP120', 'WP150', 'ZT078', 'ZT090', 'ZT102', 'ZT120', 'ZT150']):
                return 'sunpro-3-12'
            if 'sun core' in d.lower() or any(m.startswith(p) for p in ['KQG', 'KXGA']):
                return 'suncore-sm'
            if 'suncore' in s or any(m.startswith(p) for p in ['ZX08', 'ZX09', 'ZX10', 'ZX12', 'ZX14', 'ZY07', 'ZY08', 'ZY09', 'ZY10', 'ZY12', 'XY07']):
                return 'suncore-lg'
            if 'sun choice' in d.lower() or 'sunchoice' in s or 'sun choice' in f or any(m.startswith(p) for p in ['AV15', 'AV18', 'AV20', 'AV25', 'KV15', 'KV18', 'KV20', 'KV25', 'KV28']):
                return 'sunchoice-15-27'
            if 'sun select' in d.lower() or 'sun select' in s or 'sun select' in f:
                return 'sunselect-27-50'
            if 'sun premier' in d.lower() or 'sun premier' in s or 'sun premier' in f:
                return 'sunpremier-25-150'
            if 'lx' in s or 'lx series' in f or any(m.startswith(p) for p in ['PCE', 'PHE', 'PCG', 'PHG', 'PH3E', 'PC3E', 'PZ3E']):
                return 'lx-pkg-2-5'
            
            # Other Equipment
            if 'ptac' in s or 'amana' in s:
                return 'amana-ptac'
            if 'boiler' in s:
                return 'boilers'
            if 'lg ductless' in s or 'dfs ductless' in s or 'lg' in f:
                return 'lg-ductless'
            if 'reznor' in s or 'reznor' in c:
                return 'reznor-heaters'
            if '2.5-50 splits' in s or 'csplit' in s:
                return 'csplit-2-50'
            if 'res split' in s or 'furnace' in s:
                return 'res-splits-furnace'
            return None

        # PASS 2: Parse Catalog Sheets, Matrix Sheets, and Reference Sheets
        for name, sid, rid in sheets:
            if name in excluded_sheets:
                print(f"Skipping excluded sheet: {name}")
                continue

            target = rel_map[rid]
            sheet_path = "xl/" + target if not target.startswith("xl/") else target
            if sheet_path not in z.namelist():
                continue

            sheet_tree = ET.parse(z.open(sheet_path))

            # Matrix Sheets
            if name in matrix_sheet_names:
                matrix_rows = []
                for row in sheet_tree.getroot().iter():
                    if not row.tag.endswith("}row"):
                        continue
                    cells = {}
                    for c in row:
                        if not c.tag.endswith("}c"):
                            continue
                        r_attr = c.attrib.get("r", "")
                        col_letter = "".join([ch for ch in r_attr if ch.isalpha()])
                        t_attr = c.attrib.get("t")
                        val = ""
                        for child in c:
                            if child.tag.endswith("}v") and child.text:
                                if t_attr == "s":
                                    try:
                                        s_idx = int(child.text)
                                        val = shared_strings[s_idx] if s_idx < len(shared_strings) else ""
                                    except:
                                        pass
                                else:
                                    val = child.text
                        if val.strip():
                            cells[col_letter] = sanitize_text(val.strip())
                    if cells:
                        matrix_rows.append(cells)
                matrix_data[name.strip()] = matrix_rows
                continue

            # Reference Sheets
            if name in ref_sheet_names:
                ref_rows = []
                for row in sheet_tree.getroot().iter():
                    if not row.tag.endswith("}row"):
                        continue
                    cells = {}
                    for c in row:
                        if not c.tag.endswith("}c"):
                            continue
                        r_attr = c.attrib.get("r", "")
                        col_letter = "".join([ch for ch in r_attr if ch.isalpha()])
                        t_attr = c.attrib.get("t")
                        val = ""
                        for child in c:
                            if child.tag.endswith("}v") and child.text:
                                if t_attr == "s":
                                    try:
                                        s_idx = int(child.text)
                                        val = shared_strings[s_idx] if s_idx < len(shared_strings) else ""
                                    except:
                                        pass
                                else:
                                    val = child.text
                        if val.strip():
                            cells[col_letter] = sanitize_text(val.strip())
                    if cells:
                        ref_rows.append(cells)
                reference_sheets[name] = ref_rows
                continue

            # Catalog Sheets
            sheet_items = []
            current_category = name
            raw_rows = []

            for row in sheet_tree.getroot().iter():
                if not row.tag.endswith("}row"):
                    continue
                r_num = row.attrib.get("r", "")
                cells = {}
                for c in row:
                    if not c.tag.endswith("}c"):
                        continue
                    r_attr = c.attrib.get("r", "")
                    col_letter = "".join([ch for ch in r_attr if ch.isalpha()])
                    t_attr = c.attrib.get("t")
                    val = ""
                    for child in c:
                        if child.tag.endswith("}v") and child.text:
                            if t_attr == "s":
                                try:
                                    s_idx = int(child.text)
                                    val = shared_strings[s_idx] if s_idx < len(shared_strings) else ""
                                except:
                                    pass
                            else:
                                val = child.text
                    if val.strip():
                        cells[col_letter] = sanitize_text(val.strip())
                if cells:
                    raw_rows.append((r_num, cells))

            for r_idx, cells in raw_rows:
                all_text = " ".join(cells.values()).strip()
                all_lower = all_text.lower()

                if ("model" in all_lower and ("mlp" in all_lower or "desc" in all_lower or "price" in all_lower or "cost" in all_lower)):
                    continue

                if len(cells) <= 2 and not any(is_number(v) for v in cells.values()):
                    first_val = list(cells.values())[0]
                    if len(first_val) > 2 and not first_val.startswith("•"):
                        current_category = first_val
                    continue

                model = cells.get("B", "").strip()
                desc = cells.get("D", "").strip() or cells.get("C", "").strip()
                avail = cells.get("A", "").strip()
                mlp = cells.get("E", "").strip()
                cost_x = cells.get("F", "").strip()
                sell_x = cells.get("G", "").strip()
                sell = cells.get("H", "").strip()
                family = cells.get("I", "").strip()

                if model.strip() in ["***", "**", "*", "-", "—"]:
                    model = ""

                if not model and avail and len(avail) < 35 and re.search(r'^[A-Z0-9-]{4,}$', avail):
                    model = avail
                    avail = ""

                if len(model) > 45 and not desc:
                    desc = model
                    model = ""

                if not model and cells.get("C") and len(cells.get("C")) < 30 and re.search(r'^[A-Z0-9-]{4,}$', cells.get("C").strip()):
                    model = cells.get("C").strip()

                mlp_clean = clean_num(mlp)
                sell_clean = clean_num(sell)
                cost_x_clean = clean_num(cost_x)
                sell_x_clean = clean_num(sell_x)
                has_price = bool(mlp_clean or sell_clean)

                invalid_models = [
                    "MODEL", "AVAILABILITY", "DESCRIPTION", "DAMAGE AND OBSOLETE",
                    "NOTE", "FOR SUBMITTALS", "PER PAUL HARMS", "YORK RESIDENTIAL",
                    "CHECK YORK INVENTORY", "MISTY DAWN STOVER", "ORDER CORRESPONDENT"
                ]
                if any(model.upper().startswith(im) for im in invalid_models):
                    continue

                if "@" in model or "phone:" in model.lower():
                    continue

                if not has_price and (not model or len(model.split()) > 3 or len(model) > 35):
                    title = desc or model or avail
                    if len(title) > 3 and not title.startswith("•"):
                        current_category = title
                    continue

                if not model and not has_price:
                    continue

                generic_models = ["SUNLINE", "SUN PRO", "SUN CORE", "SUN CHOICE", "COMMERCIAL SPLITS", "COMMERCIAL RTU", "CONTACTS", ""]
                if model.upper() in generic_models or not model:
                    code_match = re.search(r'\b([A-Z]{2,4}\d{2,3}[A-Z0-9_-]*)\b', desc)
                    part_match = re.search(r'\b([A-Z0-9]{3,}-[A-Z0-9-]{2,})\b', desc)
                    if code_match:
                        extracted = code_match.group(1)
                        model = f"{extracted} ({model})" if model and model.upper() in generic_models else extracted
                    elif part_match:
                        model = part_match.group(1)

                refrigerant = ""
                full_spec = f"{model} {desc} {current_category}".upper()
                if "454B" in full_spec or "R454B" in full_spec or "R-454B" in full_spec:
                    refrigerant = "R-454B"
                elif "410A" in full_spec or "R410A" in full_spec or "R-410A" in full_spec:
                    refrigerant = "R-410A"
                elif "R-32" in full_spec or "R32" in full_spec:
                    refrigerant = "R-32"
                elif "R-22" in full_spec or "R22" in full_spec:
                    refrigerant = "R-22"

                phase = ""
                if "1 PHASE" in full_spec or "1PH" in full_spec or "SINGLE PHASE" in full_spec or "-1 " in full_spec:
                    phase = "1-Phase"
                elif "3 PHASE" in full_spec or "3PH" in full_spec or "THREE PHASE" in full_spec:
                    phase = "3-Phase"

                if not sell_clean and mlp_clean and sell_x_clean:
                    try:
                        calculated_sell = float(mlp_clean) * float(sell_x_clean)
                        sell_clean = f"{calculated_sell:.2f}"
                    except:
                        pass

                temp_item = {"m": model, "d": desc, "s": name, "f": family, "c": current_category}
                series_k = get_series_key(temp_item)
                
                # Check if item itself is an accessory
                is_accessory_item = False
                if model.upper() in all_accessory_models:
                    is_accessory_item = True
                elif any(k in desc.lower() for k in ['curb', 'economizer', 'hail guard', 'smoke detector', 'burglar bar', 'twinning kit', 'wall slv', 'flange']):
                    is_accessory_item = True

                if model or desc:
                    sheet_items.append({
                        "id": len(catalog_items) + len(sheet_items) + 1,
                        "s": name,
                        "c": current_category,
                        "m": model,
                        "d": desc,
                        "a": avail,
                        "mlp": mlp_clean,
                        "cx": cost_x_clean,
                        "sx": sell_x_clean,
                        "sp": sell_clean,
                        "f": family,
                        "ref": refrigerant,
                        "ph": phase,
                        "sk": series_k or "",
                        "is_acc": is_accessory_item
                    })

            catalog_items.extend(sheet_items)
            sheet_stats.append({"name": name, "type": "catalog", "count": len(sheet_items)})

        return catalog_items, matrix_data, reference_sheets, sheet_stats, series_catalog

def generate_html():
    catalog_items, matrix_data, ref_sheets, sheet_stats, series_catalog = parse_workbook()
    logo_b64 = load_logo_b64()

    print("Structuring Pricing Matrix data...")
    formatted_matrices = {}
    for mat_name, rows in matrix_data.items():
        tiers = ["8%", "9%", "10%", "11%", "12%", "13%", "14%", "15%", "16%", "17%", "18%", "19%", "20%", "21%"]
        col_keys = [
            ("C", "D"), ("E", "F"), ("G", "H"), ("I", "J"), ("K", "L"),
            ("M", "N"), ("O", "P"), ("Q", "R"), ("S", "T"), ("U", "V"),
            ("W", "X"), ("Y", "Z"), ("AA", "AB"), ("AC", "AD")
        ]
        
        body_rows = []
        for r in rows:
            cat = r.get("A", "").strip()
            desc = r.get("B", "").strip()
            if not cat or not desc or "tier" in cat.lower() or "tier" in desc.lower():
                continue

            multipliers = []
            for buy_col, sell_col in col_keys:
                buy_val = clean_num(r.get(buy_col, ""))
                sell_val = clean_num(r.get(sell_col, ""))
                multipliers.append({
                    "buy": buy_val or "—",
                    "sell": sell_val or "—"
                })

            body_rows.append({
                "category": cat,
                "description": desc,
                "tiers": multipliers
            })

        formatted_matrices[mat_name] = {
            "tiers": tiers,
            "rows": body_rows
        }

    sheet_groups = {
        "Commercial Rooftops & KC Stock": [
            "KC rtu and split 454B", "3-6 ton Sunline new", "3-12.5T Sun Pro new",
            "SunCore New", "SunChoice new", "Sun Select Current", "Sun Premier 25-150"
        ],
        "Residential Split Systems": [
            "Res Split Sys Current", "Res Split Sys with ASG Sell", "Value Tier Residential", "2-5 T LX Current"
        ],
        "Commercial Splits (2.5-50T)": [
            "2.5-50 Splits Current"
        ],
        "Ductless & VRF Systems": [
            "LG Ductless", "York Hitachi DFS Ductless", "Hitachi York VRF"
        ],
        "Applied Systems & Chillers": [
            "Quantech Chillers", "Enviro-Tec", "Water Source Heat Pumps", "Boilers"
        ],
        "Specialties, Heaters & Accessories": [
            "Reznor Duct Furnaces", "Reznor Unit Heaters", "Superior Radiant Tube Heaters",
            "Soler & Palau Fans", "Specialties", "Thermostats", "Zip A Duct", "Guardian",
            "Amana PTAC", "Bard Wall Mount", "Mobile Home", "EWC"
        ]
    }

    db_json = json.dumps({
        "items": catalog_items,
        "matrices": formatted_matrices,
        "references": ref_sheets,
        "sheet_groups": sheet_groups,
        "series_catalog": series_catalog
    }, separators=(',', ':'))

    print(f"Clean Database JSON size: {len(db_json):,} bytes ({len(catalog_items):,} items). Assembling HTML application...")

    html_template = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>CFM Distributors - HVACR Quote Data & Equipment Inventory Database</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=DM+Sans:ital,opsz,wght@0,9..40,300;0,9..40,400;0,9..40,500;0,9..40,600;0,9..40,700;1,9..40,400&family=Playfair+Display:wght@600;700&family=JetBrains+Mono:wght@400;500;600&display=swap" rel="stylesheet">
  <style>
    :root {{
      --primary-dark: #0d1f2d;
      --primary-navy: #132738;
      --cfm-blue: #1e82c8;
      --cfm-blue-hover: #156fab;
      --cfm-blue-light: #e8f3fb;
      --cfm-teal: #1ec8a0;
      --cfm-teal-light: #e6faf5;
      --cfm-gold: #f59e0b;
      --cfm-red: #ef4444;
      --bg-body: #f4f8fb;
      --bg-card: #ffffff;
      --bg-muted: #f8fafc;
      --border-subtle: #e2e8f0;
      --border-strong: #cbd5e1;
      --text-main: #0f172a;
      --text-muted: #64748b;
      --text-light: #94a3b8;
      --shadow-sm: 0 1px 3px rgba(0,0,0,0.05), 0 1px 2px rgba(0,0,0,0.03);
      --shadow-md: 0 4px 6px -1px rgba(0,0,0,0.07), 0 2px 4px -2px rgba(0,0,0,0.05);
      --shadow-lg: 0 10px 15px -3px rgba(0,0,0,0.08), 0 4px 6px -4px rgba(0,0,0,0.04);
      --radius-sm: 6px;
      --radius-md: 10px;
      --radius-lg: 16px;
      --radius-full: 9999px;
    }}

    * {{ box-sizing: border-box; margin: 0; padding: 0; }}
    body {{
      font-family: 'DM Sans', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      background-color: var(--bg-body);
      color: var(--text-main);
      min-height: 100vh;
      display: flex;
      flex-direction: column;
      overflow-x: hidden;
    }}

    /* HEADER */
    .header-band {{
      background: linear-gradient(135deg, #091722 0%, #0d1f2d 60%, #122b40 100%);
      color: #ffffff;
      padding: 0.85rem 2rem;
      border-bottom: 3px solid var(--cfm-blue);
      position: sticky;
      top: 0;
      z-index: 1000;
      box-shadow: 0 4px 12px rgba(0,0,0,0.2);
    }}
    .header-inner {{
      max-width: 1700px;
      margin: 0 auto;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 1.5rem;
    }}
    .header-brand {{
      display: flex;
      align-items: center;
      gap: 1.25rem;
      text-decoration: none;
      color: inherit;
    }}
    .header-logo {{
      height: 44px;
      width: 44px;
      object-fit: contain;
      border-radius: 6px;
      transition: transform 0.2s ease;
      box-shadow: 0 2px 6px rgba(0,0,0,0.25);
    }}
    .header-logo:hover {{
      transform: scale(1.04);
    }}
    .brand-titles h1 {{
      font-size: 1.15rem;
      font-weight: 700;
      letter-spacing: -0.01em;
      color: #ffffff;
      display: flex;
      align-items: center;
      gap: 0.6rem;
    }}
    .brand-titles p {{
      font-size: 0.76rem;
      color: #94a3b8;
      letter-spacing: 0.02em;
    }}
    .version-tag {{
      background: rgba(30, 130, 200, 0.25);
      border: 1px solid rgba(30, 130, 200, 0.5);
      color: #7dd3fc;
      font-size: 0.66rem;
      font-weight: 600;
      padding: 2px 8px;
      border-radius: var(--radius-full);
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }}

    /* HEADER SEARCH */
    .header-search {{
      flex: 1;
      max-width: 580px;
      position: relative;
    }}
    .header-search input {{
      width: 100%;
      background: rgba(255, 255, 255, 0.1);
      border: 1px solid rgba(255, 255, 255, 0.2);
      border-radius: var(--radius-full);
      padding: 0.65rem 1.25rem 0.65rem 2.8rem;
      color: #ffffff;
      font-size: 0.92rem;
      font-family: inherit;
      outline: none;
      transition: all 0.2s ease;
    }}
    .header-search input:focus {{
      background: rgba(255, 255, 255, 0.18);
      border-color: var(--cfm-blue);
      box-shadow: 0 0 0 3px rgba(30, 130, 200, 0.35);
    }}
    .header-search input::placeholder {{
      color: #94a3b8;
    }}
    .search-icon {{
      position: absolute;
      left: 1rem;
      top: 50%;
      transform: translateY(-50%);
      color: #94a3b8;
      pointer-events: none;
      display: flex;
      align-items: center;
    }}
    .search-shortcut {{
      position: absolute;
      right: 0.85rem;
      top: 50%;
      transform: translateY(-50%);
      background: rgba(255, 255, 255, 0.12);
      color: #cbd5e1;
      font-size: 0.68rem;
      padding: 2px 6px;
      border-radius: 4px;
      font-family: 'JetBrains Mono', monospace;
    }}

    /* HEADER ACTIONS */
    .header-actions {{
      display: flex;
      align-items: center;
      gap: 0.85rem;
    }}
    .btn {{
      display: inline-flex;
      align-items: center;
      gap: 0.5rem;
      padding: 0.6rem 1.1rem;
      border-radius: var(--radius-md);
      font-size: 0.86rem;
      font-weight: 500;
      font-family: inherit;
      cursor: pointer;
      border: 1px solid transparent;
      transition: all 0.18s ease;
      text-decoration: none;
    }}
    .btn-outline-light {{
      background: rgba(255, 255, 255, 0.08);
      color: #ffffff;
      border-color: rgba(255, 255, 255, 0.2);
    }}
    .btn-outline-light:hover {{
      background: rgba(255, 255, 255, 0.16);
      border-color: rgba(255, 255, 255, 0.4);
    }}
    .btn-primary {{
      background: var(--cfm-blue);
      color: #ffffff;
    }}
    .btn-primary:hover {{
      background: var(--cfm-blue-hover);
    }}
    .btn-quote {{
      background: #059669;
      color: #ffffff;
      position: relative;
    }}
    .btn-quote:hover {{
      background: #047857;
    }}
    .cart-badge {{
      background: #ffffff;
      color: #059669;
      font-weight: 700;
      font-size: 0.72rem;
      padding: 2px 7px;
      border-radius: var(--radius-full);
      margin-left: 4px;
    }}

    /* NAVIGATION TABS */
    .subnav-band {{
      background: #ffffff;
      border-bottom: 1px solid var(--border-subtle);
      box-shadow: var(--shadow-sm);
      position: sticky;
      top: 68px;
      z-index: 900;
    }}
    .nav-tabs-container {{
      max-width: 1700px;
      margin: 0 auto;
      display: flex;
      overflow-x: auto;
      white-space: nowrap;
      padding: 0 1.5rem;
      scrollbar-width: none;
    }}
    .nav-tabs-container::-webkit-scrollbar {{
      display: none;
    }}
    .nav-tab {{
      padding: 0.85rem 1.25rem;
      font-size: 0.88rem;
      font-weight: 500;
      color: var(--text-muted);
      cursor: pointer;
      border-bottom: 3px solid transparent;
      transition: all 0.15s ease;
      display: flex;
      align-items: center;
      gap: 0.5rem;
      user-select: none;
    }}
    .nav-tab:hover {{
      color: var(--cfm-blue);
    }}
    .nav-tab.active {{
      color: var(--cfm-blue);
      border-bottom-color: var(--cfm-blue);
      font-weight: 600;
    }}
    .nav-count {{
      background: #f1f5f9;
      color: #64748b;
      font-size: 0.72rem;
      padding: 2px 6px;
      border-radius: var(--radius-full);
      font-weight: 600;
    }}
    .nav-tab.active .nav-count {{
      background: var(--cfm-blue-light);
      color: var(--cfm-blue);
    }}

    /* MAIN CONTAINER */
    .main-content {{
      max-width: 1700px;
      width: 100%;
      margin: 1.5rem auto;
      padding: 0 1.5rem 3rem;
      flex: 1;
    }}

    /* VIEWS */
    .view-section {{
      display: none;
    }}
    .view-section.active {{
      display: block;
    }}

    /* CONTROL CARDS & FILTERS */
    .control-card {{
      background: var(--bg-card);
      border-radius: var(--radius-md);
      border: 1px solid var(--border-subtle);
      padding: 1.1rem 1.5rem;
      margin-bottom: 1.25rem;
      box-shadow: var(--shadow-sm);
    }}
    .filters-row {{
      display: flex;
      align-items: center;
      justify-content: space-between;
      flex-wrap: wrap;
      gap: 1.25rem;
    }}
    .filter-group {{
      display: flex;
      align-items: center;
      gap: 0.85rem;
      flex-wrap: wrap;
    }}
    .filter-label {{
      font-size: 0.82rem;
      font-weight: 600;
      color: var(--text-muted);
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }}
    .filter-select {{
      padding: 0.5rem 2rem 0.5rem 0.9rem;
      border-radius: var(--radius-sm);
      border: 1px solid var(--border-strong);
      background-color: #ffffff;
      color: var(--text-main);
      font-size: 0.85rem;
      outline: none;
      cursor: pointer;
      appearance: none;
      background-image: url("data:image/svg+xml;charset=UTF-8,%3csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%2364748b' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'%3e%3cpolyline points='6 9 12 15 18 9'%3e%3c/polyline%3e%3c/svg%3e");
      background-repeat: no-repeat;
      background-position: right 0.6rem center;
      background-size: 14px;
    }}
    .filter-pills {{
      display: flex;
      gap: 0.35rem;
      background: #f1f5f9;
      padding: 3px;
      border-radius: var(--radius-sm);
    }}
    .pill-btn {{
      padding: 0.35rem 0.85rem;
      font-size: 0.8rem;
      font-weight: 500;
      border-radius: 4px;
      border: none;
      background: transparent;
      color: var(--text-muted);
      cursor: pointer;
      transition: all 0.15s ease;
    }}
    .pill-btn:hover {{
      color: var(--text-main);
    }}
    .pill-btn.active {{
      background: #ffffff;
      color: var(--cfm-blue);
      font-weight: 600;
      box-shadow: 0 1px 2px rgba(0,0,0,0.06);
    }}
    .result-summary {{
      font-size: 0.85rem;
      color: var(--text-muted);
    }}

    /* DATA TABLES */
    .table-container {{
      background: var(--bg-card);
      border-radius: var(--radius-md);
      border: 1px solid var(--border-subtle);
      box-shadow: var(--shadow-sm);
      overflow-x: auto;
    }}
    .data-table {{
      width: 100%;
      border-collapse: collapse;
      text-align: left;
      font-size: 0.86rem;
    }}
    .data-table th {{
      background: #0d1f2d;
      color: #ffffff;
      font-weight: 600;
      font-size: 0.78rem;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      padding: 0.85rem 1rem;
      border-bottom: 1px solid #1e293b;
      cursor: pointer;
      user-select: none;
      white-space: nowrap;
    }}
    .data-table th:hover {{
      background: #152c3f;
    }}
    .data-table td {{
      padding: 0.85rem 1rem;
      border-bottom: 1px solid var(--border-subtle);
      color: var(--text-main);
      vertical-align: middle;
    }}
    .data-table tbody tr:hover {{
      background-color: #f8fafc;
    }}

    /* BADGES & PILLS */
    .avail-badge {{
      display: inline-block;
      font-size: 0.72rem;
      font-weight: 600;
      padding: 2px 8px;
      border-radius: 4px;
      text-align: center;
      white-space: nowrap;
    }}
    .avail-stock {{
      background: #ecfdf5;
      color: #047857;
      border: 1px solid #a7f3d0;
    }}
    .avail-call {{
      background: #fffbeb;
      color: #b45309;
      border: 1px solid #fde68a;
    }}
    .avail-other {{
      background: #f1f5f9;
      color: #475569;
      border: 1px solid #cbd5e1;
    }}
    .badge-a2l {{
      background: #eff6ff;
      color: #1d4ed8;
      border: 1px solid #bfdbfe;
      font-size: 0.68rem;
      font-weight: 700;
      padding: 2px 6px;
      border-radius: 4px;
      margin-left: 6px;
      display: inline-block;
    }}
    .badge-r410a {{
      background: #f8fafc;
      color: #64748b;
      border: 1px solid #cbd5e1;
      font-size: 0.68rem;
      font-weight: 600;
      padding: 2px 6px;
      border-radius: 4px;
      margin-left: 6px;
      display: inline-block;
    }}
    .model-cell {{
      font-family: 'JetBrains Mono', monospace;
      font-weight: 600;
      color: #0f172a;
      cursor: pointer;
      display: flex;
      align-items: center;
      gap: 6px;
    }}
    .model-cell:hover {{
      color: var(--cfm-blue);
      text-decoration: underline;
    }}
    .price-mlp {{
      font-family: 'JetBrains Mono', monospace;
      color: var(--text-muted);
    }}
    .price-sell {{
      font-family: 'JetBrains Mono', monospace;
      font-weight: 700;
      color: #047857;
      font-size: 0.94rem;
    }}
    .multiplier-tag {{
      font-family: 'JetBrains Mono', monospace;
      font-size: 0.78rem;
      color: #475569;
      background: #f1f5f9;
      padding: 2px 6px;
      border-radius: 4px;
    }}

    /* ACTION BUTTONS */
    .btn-add {{
      background: #f0fdf4;
      color: #166534;
      border: 1px solid #bbf7d0;
      padding: 0.4rem 0.8rem;
      border-radius: var(--radius-sm);
      font-size: 0.8rem;
      font-weight: 600;
      cursor: pointer;
      transition: all 0.15s ease;
      white-space: nowrap;
    }}
    .btn-add:hover {{
      background: #16a34a;
      color: #ffffff;
      border-color: #16a34a;
    }}
    .btn-acc {{
      display: inline-flex;
      align-items: center;
      gap: 4px;
      background: #eef6fc;
      color: #1e82c8;
      border: 1px solid #bae0f7;
      border-radius: var(--radius-sm);
      padding: 0.4rem 0.75rem;
      font-size: 0.78rem;
      font-weight: 600;
      cursor: pointer;
      transition: all 0.15s ease;
      margin-left: 6px;
      white-space: nowrap;
    }}
    .btn-acc:hover {{
      background: #1e82c8;
      color: #ffffff;
      border-color: #1e82c8;
      transform: translateY(-1px);
      box-shadow: 0 2px 4px rgba(30, 130, 200, 0.25);
    }}

    /* PAGINATION */
    .pagination-bar {{
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 1rem 1.5rem;
      background: #ffffff;
      border-top: 1px solid var(--border-subtle);
    }}
    .page-buttons {{
      display: flex;
      gap: 0.4rem;
    }}
    .page-btn {{
      padding: 0.4rem 0.8rem;
      font-size: 0.82rem;
      border-radius: var(--radius-sm);
      border: 1px solid var(--border-strong);
      background: #ffffff;
      color: var(--text-main);
      cursor: pointer;
      font-weight: 500;
    }}
    .page-btn:hover:not(:disabled) {{
      border-color: var(--cfm-blue);
      color: var(--cfm-blue);
    }}
    .page-btn.active {{
      background: var(--cfm-blue);
      color: #ffffff;
      border-color: var(--cfm-blue);
    }}
    .page-btn:disabled {{
      opacity: 0.5;
      cursor: not-allowed;
    }}

    /* SLIDE-OUT QUOTE DRAWER */
    .drawer-overlay {{
      position: fixed;
      top: 0;
      left: 0;
      width: 100vw;
      height: 100vh;
      background: rgba(13, 31, 45, 0.4);
      backdrop-filter: blur(2px);
      z-index: 1500;
      display: none;
    }}
    .drawer-overlay.open {{
      display: block;
    }}
    .quote-drawer {{
      position: fixed;
      top: 0;
      right: -520px;
      width: 500px;
      max-width: 90vw;
      height: 100vh;
      background: #ffffff;
      box-shadow: -4px 0 24px rgba(0,0,0,0.15);
      z-index: 1600;
      transition: right 0.25s cubic-bezier(0.16, 1, 0.3, 1);
      display: flex;
      flex-direction: column;
    }}
    .quote-drawer.open {{
      right: 0;
    }}
    .drawer-header {{
      padding: 1.25rem 1.5rem;
      background: var(--primary-dark);
      color: #ffffff;
      display: flex;
      align-items: center;
      justify-content: space-between;
      border-bottom: 2px solid var(--cfm-blue);
    }}
    .drawer-header h3 {{
      font-size: 1.1rem;
      font-weight: 700;
      display: flex;
      align-items: center;
      gap: 0.5rem;
    }}
    .drawer-close-btn {{
      background: transparent;
      border: none;
      color: #ffffff;
      font-size: 1.5rem;
      cursor: pointer;
      line-height: 1;
    }}
    .drawer-body {{
      flex: 1;
      overflow-y: auto;
      padding: 1.25rem 1.5rem;
    }}
    .drawer-item {{
      background: #f8fafc;
      border: 1px solid var(--border-subtle);
      border-radius: var(--radius-sm);
      padding: 0.85rem;
      margin-bottom: 0.75rem;
    }}
    .drawer-item-title {{
      font-family: 'JetBrains Mono', monospace;
      font-weight: 700;
      font-size: 0.88rem;
      color: var(--primary-dark);
    }}
    .drawer-item-desc {{
      font-size: 0.78rem;
      color: var(--text-muted);
      margin: 3px 0 6px;
    }}
    .drawer-item-calc {{
      display: flex;
      align-items: center;
      justify-content: space-between;
      margin-top: 6px;
    }}
    .qty-controls {{
      display: flex;
      align-items: center;
      gap: 0.4rem;
    }}
    .qty-btn {{
      width: 26px;
      height: 26px;
      border-radius: 4px;
      border: 1px solid var(--border-strong);
      background: #ffffff;
      cursor: pointer;
      font-weight: 700;
      display: flex;
      align-items: center;
      justify-content: center;
    }}
    .qty-val {{
      font-size: 0.88rem;
      font-weight: 600;
      min-width: 24px;
      text-align: center;
    }}
    .btn-del-item {{
      color: #ef4444;
      background: transparent;
      border: none;
      cursor: pointer;
      font-size: 0.9rem;
      padding: 2px 6px;
    }}
    .drawer-footer {{
      padding: 1.25rem 1.5rem;
      background: #ffffff;
      border-top: 1px solid var(--border-subtle);
      box-shadow: 0 -4px 12px rgba(0,0,0,0.03);
    }}
    .summary-row {{
      display: flex;
      justify-content: space-between;
      font-size: 0.88rem;
      margin-bottom: 0.5rem;
      color: var(--text-muted);
    }}
    .summary-row.total {{
      font-size: 1.15rem;
      font-weight: 700;
      color: var(--text-main);
      padding-top: 0.5rem;
      border-top: 1px dashed var(--border-strong);
      margin-bottom: 1rem;
    }}

    /* QUOTE RECS (SMART ACCESSORIES) */
    .quote-recs-box {{
      margin-top: 1.25rem;
      background: #f0f9ff;
      border: 1px solid #bae6fd;
      border-radius: var(--radius-md);
      padding: 1rem;
    }}
    .quote-recs-title {{
      font-size: 0.88rem;
      font-weight: 700;
      color: #0369a1;
      display: flex;
      align-items: center;
      gap: 0.4rem;
      margin-bottom: 0.75rem;
    }}
    .quote-recs-list {{
      display: flex;
      flex-direction: column;
      gap: 0.5rem;
    }}
    .quote-rec-item {{
      background: #ffffff;
      border: 1px solid #e0f2fe;
      border-radius: var(--radius-sm);
      padding: 0.6rem 0.85rem;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 0.5rem;
    }}
    .quote-rec-item-title {{
      font-family: 'JetBrains Mono', monospace;
      font-size: 0.8rem;
      font-weight: 700;
      color: #0f172a;
    }}
    .quote-rec-item-desc {{
      font-size: 0.74rem;
      color: #64748b;
    }}
    .quote-rec-item-price {{
      font-family: 'JetBrains Mono', monospace;
      font-size: 0.82rem;
      font-weight: 700;
      color: #047857;
      white-space: nowrap;
    }}

    /* MODAL (ACCESSORY POPOVER) */
    .modal-overlay {{
      position: fixed;
      top: 0;
      left: 0;
      width: 100vw;
      height: 100vh;
      background: rgba(13, 31, 45, 0.65);
      backdrop-filter: blur(4px);
      z-index: 2500;
      display: none;
      align-items: center;
      justify-content: center;
      padding: 1.5rem;
    }}
    .modal-overlay.open {{
      display: flex;
    }}
    .modal-box {{
      background: #ffffff;
      border-radius: var(--radius-lg);
      width: 100%;
      max-width: 1100px;
      max-height: 90vh;
      display: flex;
      flex-direction: column;
      box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.25);
      overflow: hidden;
      border: 1px solid var(--border-subtle);
    }}
    .modal-header {{
      padding: 1.1rem 1.75rem;
      background: linear-gradient(135deg, #091722 0%, #0d1f2d 100%);
      color: #ffffff;
      display: flex;
      align-items: center;
      justify-content: space-between;
      border-bottom: 2px solid var(--cfm-blue);
    }}
    .modal-header h3 {{
      font-size: 1.15rem;
      font-weight: 700;
      display: flex;
      align-items: center;
      gap: 0.6rem;
    }}
    .modal-header p {{
      font-size: 0.78rem;
      color: #94a3b8;
      margin-top: 2px;
    }}
    .modal-close-btn {{
      background: rgba(255, 255, 255, 0.12);
      border: none;
      color: #ffffff;
      font-size: 1.3rem;
      width: 32px;
      height: 32px;
      border-radius: 50%;
      display: flex;
      align-items: center;
      justify-content: center;
      cursor: pointer;
      transition: background 0.15s ease;
    }}
    .modal-close-btn:hover {{
      background: rgba(239, 68, 68, 0.8);
    }}
    .modal-toolbar {{
      padding: 0.9rem 1.75rem;
      background: #f8fafc;
      border-bottom: 1px solid var(--border-subtle);
      display: flex;
      align-items: center;
      justify-content: space-between;
      flex-wrap: wrap;
      gap: 1rem;
    }}
    .modal-search {{
      flex: 1;
      max-width: 350px;
      position: relative;
    }}
    .modal-search input {{
      width: 100%;
      padding: 0.45rem 1rem 0.45rem 2.2rem;
      font-size: 0.85rem;
      border-radius: var(--radius-full);
      border: 1px solid var(--border-strong);
      outline: none;
    }}
    .modal-search input:focus {{
      border-color: var(--cfm-blue);
      box-shadow: 0 0 0 2px rgba(30, 130, 200, 0.2);
    }}
    .modal-search svg {{
      position: absolute;
      left: 0.8rem;
      top: 50%;
      transform: translateY(-50%);
      color: #94a3b8;
    }}
    .modal-body {{
      flex: 1;
      overflow-y: auto;
      padding: 0;
    }}
    .modal-footer {{
      padding: 0.9rem 1.75rem;
      background: #f8fafc;
      border-top: 1px solid var(--border-subtle);
      display: flex;
      align-items: center;
      justify-content: space-between;
      font-size: 0.85rem;
      color: var(--text-muted);
    }}

    /* EQUIPMENT & ACCESSORY GUIDE VIEW */
    .series-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(290px, 1fr));
      gap: 1.1rem;
      margin-bottom: 1.75rem;
    }}
    .series-card {{
      background: #ffffff;
      border: 1px solid var(--border-subtle);
      border-radius: var(--radius-md);
      padding: 1.25rem;
      cursor: pointer;
      transition: all 0.2s ease;
      display: flex;
      flex-direction: column;
      justify-content: space-between;
      box-shadow: var(--shadow-sm);
    }}
    .series-card:hover {{
      border-color: var(--cfm-blue);
      transform: translateY(-2px);
      box-shadow: var(--shadow-md);
    }}
    .series-card.active {{
      border-color: var(--cfm-blue);
      background: #f0f7fc;
      box-shadow: 0 0 0 2px var(--cfm-blue);
    }}
    .series-card-top {{
      display: flex;
      align-items: flex-start;
      justify-content: space-between;
      margin-bottom: 0.5rem;
    }}
    .series-card-title {{
      font-size: 1rem;
      font-weight: 700;
      color: var(--primary-dark);
    }}
    .series-card-desc {{
      font-size: 0.78rem;
      color: var(--text-muted);
      line-height: 1.4;
      margin-bottom: 0.85rem;
    }}
    .series-card-badge {{
      background: var(--cfm-blue-light);
      color: var(--cfm-blue);
      font-size: 0.72rem;
      font-weight: 700;
      padding: 3px 8px;
      border-radius: var(--radius-full);
      white-space: nowrap;
    }}

    /* TOAST NOTIFICATION */
    .toast {{
      position: fixed;
      bottom: 2rem;
      right: 2rem;
      background: #0d1f2d;
      color: #ffffff;
      padding: 0.85rem 1.5rem;
      border-radius: var(--radius-md);
      box-shadow: var(--shadow-lg);
      font-size: 0.88rem;
      display: flex;
      align-items: center;
      gap: 0.75rem;
      z-index: 9999;
      transform: translateY(150%);
      opacity: 0;
      transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1);
      border-left: 4px solid var(--cfm-teal);
    }}
    .toast.show {{
      transform: translateY(0);
      opacity: 1;
    }}

    
    /* DARK THEME SYSTEM */
    body.dark-theme {{
      --primary-dark: #07131d;
      --primary-navy: #0a1b28;
      --cfm-blue-hover: #299ae8;
      --cfm-blue-light: #13334d;
      --cfm-teal-light: #0d382e;
      --bg-body: #0a131b;
      --bg-card: #0f202e;
      --bg-muted: #142838;
      --border-subtle: #1c384e;
      --border-strong: #2a506e;
      --text-main: #f1f5f9;
      --text-muted: #94a3b8;
      --text-light: #64748b;
    }}
    body.dark-theme .subnav-band {{
      background: #0f202e;
      border-bottom-color: #1c384e;
    }}
    body.dark-theme .nav-tab {{
      color: #94a3b8;
    }}
    body.dark-theme .nav-tab:hover {{
      color: #38bdf8;
    }}
    body.dark-theme .nav-tab.active {{
      color: #38bdf8;
      border-bottom-color: #38bdf8;
    }}
    body.dark-theme .nav-count {{
      background: #183348;
      color: #94a3b8;
    }}
    body.dark-theme .nav-tab.active .nav-count {{
      background: #143e5c;
      color: #38bdf8;
    }}
    body.dark-theme .control-card {{
      background: #0f202e;
      border-color: #1c384e;
    }}
    body.dark-theme .table-container {{
      background: #0f202e;
      border-color: #1c384e;
    }}
    body.dark-theme .data-table th {{
      background: #091621;
      border-bottom-color: #1c384e;
    }}
    body.dark-theme .data-table td {{
      border-bottom-color: #1c384e;
      color: #f1f5f9;
    }}
    body.dark-theme .data-table tbody tr:hover {{
      background-color: #162c3e;
    }}
    body.dark-theme .model-cell {{
      color: #f1f5f9;
    }}
    body.dark-theme .filter-select {{
      background-color: #142838;
      color: #f1f5f9;
      border-color: #2a506e;
      background-image: url("data:image/svg+xml;charset=UTF-8,%3csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%2394a3b8' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'%3e%3cpolyline points='6 9 12 15 18 9'%3e%3c/polyline%3e%3c/svg%3e");
    }}
    body.dark-theme .filter-pills {{
      background: #142838;
    }}
    body.dark-theme .pill-btn {{
      color: #94a3b8;
    }}
    body.dark-theme .pill-btn:hover {{
      color: #f1f5f9;
    }}
    body.dark-theme .pill-btn.active {{
      background: #1e82c8;
      color: #ffffff;
    }}
    body.dark-theme .multiplier-tag {{
      background: #183348;
      color: #cbd5e1;
    }}
    body.dark-theme .pagination-bar {{
      background: #0f202e;
      border-top-color: #1c384e;
    }}
    body.dark-theme .page-btn {{
      background: #142838;
      border-color: #2a506e;
      color: #f1f5f9;
    }}
    body.dark-theme .page-btn.active {{
      background: #1e82c8;
      border-color: #1e82c8;
      color: #ffffff;
    }}
    body.dark-theme .series-card {{
      background: #0f202e;
      border-color: #1c384e;
    }}
    body.dark-theme .series-card:hover {{
      border-color: #1e82c8;
      background: #132a3d;
    }}
    body.dark-theme .series-card.active {{
      background: #14354e;
      border-color: #38bdf8;
    }}
    body.dark-theme .series-card-title {{
      color: #f1f5f9;
    }}
    body.dark-theme .modal-box {{
      background: #0f202e;
      border-color: #2a506e;
    }}
    body.dark-theme .modal-toolbar,
    body.dark-theme .modal-footer {{
      background: #142838;
      border-color: #1c384e;
    }}
    body.dark-theme .modal-search input {{
      background: #0f202e;
      color: #f1f5f9;
      border-color: #2a506e;
    }}
    body.dark-theme .quote-drawer {{
      background: #0f202e;
    }}
    body.dark-theme .drawer-body {{
      background: #0f202e;
    }}
    body.dark-theme .drawer-item {{
      background: #142838;
      border-color: #1c384e;
    }}
    body.dark-theme .drawer-item-title {{
      color: #f1f5f9;
    }}
    body.dark-theme .qty-btn {{
      background: #0f202e;
      border-color: #2a506e;
      color: #f1f5f9;
    }}
    body.dark-theme .drawer-footer {{
      background: #0f202e;
      border-top-color: #1c384e;
    }}
    body.dark-theme .summary-row.total {{
      color: #f1f5f9;
      border-top-color: #2a506e;
    }}
    body.dark-theme .quote-recs-box {{
      background: #0f2c42;
      border-color: #1c5277;
    }}
    body.dark-theme .quote-recs-title {{
      color: #38bdf8;
    }}
    body.dark-theme .quote-rec-item {{
      background: #0f202e;
      border-color: #1c384e;
    }}
    body.dark-theme .quote-rec-item-title {{
      color: #f1f5f9;
    }}
    body.dark-theme .proposal-doc {{
      background: #0f202e !important;
      border-color: #1c384e !important;
      color: #f1f5f9 !important;
    }}
    body.dark-theme .proposal-doc input {{
      background: #142838 !important;
      color: #f1f5f9 !important;
      border-color: #2a506e !important;
    }}
    body.dark-theme .proposal-doc div[style*="color: #0d1f2d"],
    body.dark-theme .proposal-doc div[style*="color: #0f172a"],
    body.dark-theme .proposal-doc span[style*="color: #0f172a"] {{
      color: #f1f5f9 !important;
    }}
    body.dark-theme .proposal-doc div[style*="color: #64748b"],
    body.dark-theme .proposal-doc span[style*="color: #64748b"],
    body.dark-theme .proposal-doc p {{
      color: #94a3b8 !important;
    }}
    body.dark-theme .proposal-doc th {{
      background: #142838 !important;
      color: #cbd5e1 !important;
      border-color: #243c52 !important;
    }}
    body.dark-theme .proposal-doc td {{
      border-color: #1c384e !important;
      color: #f1f5f9 !important;
    }}
    body.dark-theme #matrix-table td:nth-child(1),
    body.dark-theme #matrix-table td:nth-child(2) {{
      background: #0f202e !important;
    }}
    body.dark-theme .calc-box-cost {{
      background: #142838 !important;
      border-color: #2a506e !important;
      color: #94a3b8 !important;
    }}
    body.dark-theme .calc-box-cost div:last-child {{
      color: #f1f5f9 !important;
    }}
    body.dark-theme .calc-box-sell {{
      background: #0d382e !important;
      border-color: #15803d !important;
      color: #4ade80 !important;
    }}
    body.dark-theme .calc-box-sell div:last-child {{
      color: #4ade80 !important;
    }}
    body.dark-theme .calc-box-margin {{
      background: #13334d !important;
      border-color: #0284c7 !important;
      color: #38bdf8 !important;
    }}
    body.dark-theme .calc-box-margin div:last-child {{
      color: #38bdf8 !important;
    }}
    body.dark-theme .proposal-info-box {{
      background: #142838 !important;
      border: 1px solid #2a506e !important;
      color: #f1f5f9 !important;
    }}
    body.dark-theme .proposal-terms-box {{
      background: #142838 !important;
      border: 1px solid #2a506e !important;
      color: #cbd5e1 !important;
    }}
    body.dark-theme #matrix-table th:first-child,
    body.dark-theme #matrix-table th:nth-child(2) {{
      background: #091621 !important;
    }}


    /* PRINT PROPOSAL STYLES */
    @media print {{
      body {{
        background: #ffffff !important;
        color: #000000 !important;
      }}
      .proposal-doc, .proposal-doc div, .proposal-doc span, .proposal-doc p, .proposal-doc td, .proposal-doc th {{
        background: #ffffff !important;
        color: #000000 !important;
      }}
      .proposal-info-box, .proposal-terms-box {{
        background: #f8fafc !important;
        border: 1px solid #cbd5e1 !important;
        color: #000000 !important;
      }}
      .header-band, .subnav-band, .control-card, .btn, .nav-tabs-container,
      .pagination-bar, .modal-overlay, .drawer-overlay, .quote-drawer, .no-print {{
        display: none !important;
      }}
      .main-content {{
        margin: 0 !important;
        padding: 0 !important;
        max-width: 100% !important;
      }}
      #view-quote {{
        display: block !important;
      }}
      .proposal-doc {{
        border: none !important;
        box-shadow: none !important;
        padding: 0 !important;
      }}
    }}
  </style>
</head>
<body>

  <!-- HEADER -->
  <header class="header-band">
    <div class="header-inner">
      <a href="#" class="header-brand" onclick="switchTab('tab-all')">
        {f'<img src="data:image/png;base64,{logo_b64}" alt="CFM Distributors" class="header-logo">' if logo_b64 else '<div style="font-size: 1.5rem; font-weight: 800; color: #1e82c8;">CFM</div>'}
        <div class="brand-titles">
          <h1>CFM Distributors, Inc. <span class="version-tag">REL 092126</span></h1>
          <p>HVACR Equipment Inventory, Pricing Matrix & Proposal System</p>
        </div>
      </a>

      <!-- Instant Search -->
      <div class="header-search">
        <span class="search-icon">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
        </span>
        <input type="text" id="global-search" placeholder="Search 7,000+ products by model, description, or series..." oninput="handleGlobalSearch(this.value)">
        <span class="search-shortcut">/</span>
      </div>

      <!-- Actions -->
      <div class="header-actions">
        <button class="btn btn-outline-light" id="theme-toggle-btn" onclick="toggleTheme()" title="Toggle Dark / Light Theme" style="padding: 0.55rem 0.9rem;">
          <span id="theme-icon">🌙</span>
          <span id="theme-text">Dark Mode</span>
        </button>

        <label class="btn btn-outline-light" style="cursor: pointer;">
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
          Import .xlsx
          <input type="file" id="xlsx-uploader" accept=".xlsx" style="display: none;" onchange="handleFileImport(this)">
        </label>
        <button class="btn btn-quote" onclick="toggleDrawer()">
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="9" cy="21" r="1"/><circle cx="20" cy="21" r="1"/><path d="M1 1h4l2.68 13.39a2 2 0 0 0 2 1.61h9.72a2 2 0 0 0 2-1.61L23 6H6"/></svg>
          Active Quote <span class="cart-badge" id="cart-badge-count">0</span>
        </button>
      </div>
    </div>
  </header>

  <!-- SUB-NAV / TABS -->
  <nav class="subnav-band">
    <div class="nav-tabs-container" id="nav-tabs">
      <div class="nav-tab active" data-tab="tab-all" onclick="switchTab('tab-all')">
        All Products <span class="nav-count" id="count-all">{len(catalog_items):,}</span>
      </div>
      <div class="nav-tab" data-tab="tab-rtu" onclick="switchTab('tab-rtu')">
        Commercial RTUs & Stock
      </div>
      <div class="nav-tab" data-tab="tab-accessories" onclick="switchTab('tab-accessories')">
        ⚙️ Equipment & Accessory Guide <span class="nav-count" style="background:#e0f2fe;color:#0369a1;">14 Series</span>
      </div>
      <div class="nav-tab" data-tab="tab-res" onclick="switchTab('tab-res')">
        Residential Split Systems
      </div>
      <div class="nav-tab" data-tab="tab-split" onclick="switchTab('tab-split')">
        Commercial Splits
      </div>
      <div class="nav-tab" data-tab="tab-ductless" onclick="switchTab('tab-ductless')">
        Ductless & VRF
      </div>
      <div class="nav-tab" data-tab="tab-applied" onclick="switchTab('tab-applied')">
        Applied & Chillers
      </div>
      <div class="nav-tab" data-tab="tab-specialties" onclick="switchTab('tab-specialties')">
        Specialties, Parts & Heaters
      </div>
      <div class="nav-tab" data-tab="tab-matrix" onclick="switchTab('tab-matrix')">
        IPA Pricing Matrices & Calculator
      </div>
      <div class="nav-tab" data-tab="tab-ref" onclick="switchTab('tab-ref')">
        Reference Manuals & Policies
      </div>
      <div class="nav-tab" data-tab="tab-quote" onclick="switchTab('tab-quote')">
        Quote Proposal Builder
      </div>
    </div>
  </nav>

  <!-- MAIN VIEW CONTAINER -->
  <main class="main-content">

    <!-- CATALOG VIEW -->
    <section id="view-catalog" class="view-section active">
      <div class="control-card">
        <div class="filters-row">
          <div class="filter-group">
            <span class="filter-label">Worksheet Tab:</span>
            <select id="sheet-filter" class="filter-select" onchange="applyFilters()">
              <option value="">All Catalog Sheets</option>
            </select>

            <span class="filter-label">Refrigerant:</span>
            <div class="filter-pills" id="ref-pills">
              <button class="pill-btn active" onclick="setFilter('ref', '', this)">All</button>
              <button class="pill-btn" onclick="setFilter('ref', 'R-454B', this)">R-454B (A2L)</button>
              <button class="pill-btn" onclick="setFilter('ref', 'R-410A', this)">R-410A</button>
            </div>

            <span class="filter-label">Stock Status:</span>
            <div class="filter-pills" id="avail-pills">
              <button class="pill-btn active" onclick="setFilter('avail', '', this)">All</button>
              <button class="pill-btn" onclick="setFilter('avail', 'stock', this)">In Stock</button>
              <button class="pill-btn" onclick="setFilter('avail', 'call', this)">Call / Order</button>
            </div>
          </div>

          <div class="result-summary">
            Showing <b id="filtered-count">0</b> items (<span id="active-tab-name">All Products</span>)
          </div>
        </div>
      </div>

      <!-- CATALOG TABLE -->
      <div class="table-container">
        <table class="data-table">
          <thead>
            <tr>
              <th style="width: 100px;">Status</th>
              <th onclick="sortTable('m')" style="width: 240px;">Model Number ↕</th>
              <th onclick="sortTable('d')">Description & Specifications ↕</th>
              <th onclick="sortTable('s')" style="width: 190px;">Sheet / Category ↕</th>
              <th onclick="sortTable('mlp')" style="width: 120px; text-align: right;">List (MLP) ↕</th>
              <th style="width: 110px; text-align: center;">Mult (C/S)</th>
              <th onclick="sortTable('sp')" style="width: 130px; text-align: right;">Sell Price ↕</th>
              <th style="width: 160px; text-align: center;">Action</th>
            </tr>
          </thead>
          <tbody id="catalog-table-body">
            <!-- Dynamic rows rendered here -->
          </tbody>
        </table>

        <!-- PAGINATION BAR -->
        <div class="pagination-bar">
          <div style="font-size: 0.85rem; color: var(--text-muted);">
            Items per page:
            <select id="items-per-page" class="filter-select" onchange="changePageSize(this.value)">
              <option value="25">25</option>
              <option value="50" selected>50</option>
              <option value="100">100</option>
              <option value="250">250</option>
            </select>
          </div>
          <div class="page-buttons" id="pagination-controls">
            <!-- Dynamic pagination buttons -->
          </div>
        </div>
      </div>
    </section>

    <!-- EQUIPMENT & ACCESSORY GUIDE VIEW -->
    <section id="view-accessories" class="view-section">
      <div class="control-card">
        <h2 style="font-size: 1.25rem; font-weight: 700; margin-bottom: 0.25rem; color: var(--primary-dark);">Equipment & Compatible Accessory Directory</h2>
        <p style="font-size: 0.85rem; color: var(--text-muted); margin-bottom: 1.25rem;">Select any RTU series or equipment family below to review paired roof curbs, economizers, hail guards, electrical kits, sensors, and direct-fit accessories.</p>
        
        <div class="series-grid" id="series-selector-grid">
          <!-- Rendered dynamically -->
        </div>

        <div id="series-accessory-display" style="display: none; border-top: 1px solid var(--border-subtle); padding-top: 1.5rem;">
          <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 1rem; flex-wrap: wrap; gap: 1rem;">
            <div>
              <h3 id="active-guide-title" style="font-size: 1.15rem; font-weight: 700; color: var(--primary-dark);">Series Accessories</h3>
              <p id="active-guide-desc" style="font-size: 0.82rem; color: var(--text-muted);"></p>
            </div>
            <div style="display: flex; gap: 0.5rem; align-items: center;">
              <input type="text" id="guide-acc-search" placeholder="Search series accessories..." class="filter-select" style="width: 240px;" oninput="filterGuideAccessories()">
              <button class="btn btn-outline-light" style="color: var(--cfm-blue); border-color: var(--cfm-blue);" onclick="exportGuideCSV()">Export CSV</button>
            </div>
          </div>

          <div class="filter-pills" id="guide-category-pills" style="margin-bottom: 1rem;">
            <!-- Category filter pills rendered here -->
          </div>

          <div class="table-container">
            <table class="data-table">
              <thead>
                <tr>
                  <th style="width: 160px;">Category</th>
                  <th style="width: 220px;">Accessory Model</th>
                  <th>Description & Compatibility Details</th>
                  <th style="width: 110px; text-align: center;">Stock Status</th>
                  <th style="width: 120px; text-align: right;">CFM Sell Price</th>
                  <th style="width: 120px; text-align: center;">Action</th>
                </tr>
              </thead>
              <tbody id="guide-acc-table-body">
                <!-- Rendered dynamically -->
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </section>

    <!-- MATRIX VIEW -->
    <section id="view-matrix" class="view-section">
      <div class="control-card">
        <div class="filters-row">
          <div class="filter-group">
            <span class="filter-label">Pricing Matrix:</span>
            <div class="filter-pills">
              <button class="pill-btn active" id="btn-mat-stock" onclick="switchMatrix('Matrix on stock')">In-Stock Schedule (IPA 753978)</button>
              <button class="pill-btn" id="btn-mat-fo" onclick="switchMatrix('Matrix on factory order models')">Factory Order Schedule</button>
            </div>

            <span class="filter-label" style="margin-left: 1rem;">Highlight Target Tier:</span>
            <select id="matrix-tier-select" class="filter-select" onchange="highlightMatrixTier(this.value)">
              <option value="">No Highlight</option>
              <option value="8%">Tier 8%</option>
              <option value="10%">Tier 10%</option>
              <option value="12%">Tier 12%</option>
              <option value="15%">Tier 15%</option>
              <option value="18%">Tier 18%</option>
              <option value="20%">Tier 20%</option>
              <option value="21%">Tier 21%</option>
            </select>
          </div>
        </div>
      </div>

      <!-- LIVE MARGIN CALCULATOR CARD -->
      <div class="control-card" style="border-left: 4px solid var(--cfm-blue);">
        <h3 style="font-size: 0.95rem; font-weight: 700; margin-bottom: 0.85rem; color: var(--primary-dark); display: flex; align-items: center; gap: 0.5rem;">
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="4" y="2" width="16" height="20" rx="2"/><line x1="8" y1="6" x2="16" y2="6"/><line x1="16" y1="14" x2="16" y2="18"/><path d="M16 10h.01"/><path d="M12 10h.01"/><path d="M8 10h.01"/><path d="M12 14h.01"/><path d="M8 14h.01"/><path d="M12 18h.01"/><path d="M8 18h.01"/></svg>
          Interactive IPA Multiplier & Gross Margin Calculator
        </h3>
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 1rem; align-items: end;">
          <div>
            <label style="font-size: 0.78rem; font-weight: 600; color: var(--text-muted);">Equipment Category:</label>
            <select id="calc-category" class="filter-select" style="width: 100%; margin-top: 4px;" onchange="recalcPricing()">
              <!-- Dynamic matrix categories -->
            </select>
          </div>
          <div>
            <label style="font-size: 0.78rem; font-weight: 600; color: var(--text-muted);">Contractor Discount Tier:</label>
            <select id="calc-tier" class="filter-select" style="width: 100%; margin-top: 4px;" onchange="recalcPricing()">
              <option value="8%">8%</option>
              <option value="9%">9%</option>
              <option value="10%">10%</option>
              <option value="11%">11%</option>
              <option value="12%">12%</option>
              <option value="13%">13%</option>
              <option value="14%">14%</option>
              <option value="15%">15%</option>
              <option value="16%">16%</option>
              <option value="17%">17%</option>
              <option value="18%">18%</option>
              <option value="19%">19%</option>
              <option value="20%">20%</option>
              <option value="21%">21%</option>
            </select>
          </div>
          <div>
            <label style="font-size: 0.78rem; font-weight: 600; color: var(--text-muted);">List Price (MLP) $:</label>
            <input type="number" id="calc-mlp" value="10000" step="50" class="filter-select" style="width: 100%; margin-top: 4px;" oninput="recalcPricing()">
          </div>
          <div class="calc-box-cost" style="padding: 0.6rem 0.8rem; border-radius: var(--radius-sm); border: 1px solid var(--border-subtle);">
            <div style="font-size: 0.72rem; color: var(--text-muted);">Buy Multiplier (Cost):</div>
            <div id="calc-res-buy" style="font-family: 'JetBrains Mono', monospace; font-weight: 700; font-size: 1rem; color: #475569;">—</div>
          </div>
          <div class="calc-box-sell" style="padding: 0.6rem 0.8rem; border-radius: var(--radius-sm); border: 1px solid #bbf7d0;">
            <div style="font-size: 0.72rem; color: #166534;">Target Sell Price:</div>
            <div id="calc-res-sell" style="font-family: 'JetBrains Mono', monospace; font-weight: 700; font-size: 1.15rem; color: #15803d;">$0.00</div>
          </div>
          <div class="calc-box-margin" style="padding: 0.6rem 0.8rem; border-radius: var(--radius-sm); border: 1px solid #bfdbfe;">
            <div style="font-size: 0.72rem; color: #1d4ed8;">Estimated Gross Margin:</div>
            <div id="calc-res-margin" style="font-family: 'JetBrains Mono', monospace; font-weight: 700; font-size: 1rem; color: #1e40af;">—</div>
          </div>
        </div>
      </div>

      <!-- MATRIX DATA TABLE -->
      <div class="table-container">
        <table class="data-table" id="matrix-table">
          <!-- Rendered via JS -->
        </table>
      </div>
    </section>

    <!-- REFERENCES VIEW -->
    <section id="view-ref" class="view-section">
      <div class="control-card">
        <div class="filters-row">
          <div class="filter-group">
            <span class="filter-label">Select Document / Manual:</span>
            <select id="ref-select" class="filter-select" style="min-width: 320px;" onchange="loadReferenceDoc(this.value)">
              <!-- Dynamic reference sheet options -->
            </select>
          </div>
          <div style="display: flex; gap: 0.5rem;">
            <button class="btn btn-outline-light" style="color: var(--cfm-blue); border-color: var(--cfm-blue);" onclick="printRefDoc()">Print Manual</button>
          </div>
        </div>
      </div>

      <div class="control-card" id="ref-doc-container">
        <h2 id="ref-doc-title" style="font-size: 1.25rem; font-weight: 700; margin-bottom: 1rem; color: var(--primary-dark);"></h2>
        <div id="ref-doc-content">
          <!-- Dynamic tabular / document content -->
        </div>
      </div>
    </section>

    <!-- QUOTE PROPOSAL BUILDER VIEW -->
    <section id="view-quote" class="view-section">
      <div class="control-card no-print">
        <div style="display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 1rem;">
          <div>
            <h2 style="font-size: 1.3rem; font-weight: 700; color: var(--primary-dark);">Contractor Equipment Proposal Builder</h2>
            <p style="font-size: 0.82rem; color: var(--text-muted);">Generate customer-facing quotes with freight calculations, paired accessories, and signature authorization lines.</p>
          </div>
          <div style="display: flex; gap: 0.75rem;">
            <button class="btn btn-outline-light" style="color: #ef4444; border-color: #fca5a5;" onclick="clearQuote()">Clear Quote</button>
            <button class="btn btn-outline-light" style="color: var(--cfm-blue); border-color: var(--cfm-blue);" onclick="exportQuoteCSV()">Export CSV</button>
            <button class="btn btn-primary" onclick="window.print()">
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="6 9 6 2 18 2 18 9"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect x="6" y="14" width="12" height="8"/></svg>
              Print / Save PDF
            </button>
          </div>
        </div>
      </div>

      <!-- PROPOSAL DOCUMENT -->
      <div class="proposal-doc control-card" style="padding: 2.5rem;">
        <div style="display: flex; justify-content: space-between; align-items: flex-start; border-bottom: 2px solid var(--primary-dark); padding-bottom: 1.5rem; margin-bottom: 1.75rem;">
          <div style="display: flex; align-items: center; gap: 1.25rem;">
            {f'<img src="data:image/png;base64,{logo_b64}" alt="CFM Distributors" style="height: 52px; width: 52px; border-radius: 8px; object-fit: contain; box-shadow: 0 2px 6px rgba(0,0,0,0.12);">' if logo_b64 else '<div style="font-size: 1.75rem; font-weight: 800; color: #1e82c8;">CFM</div>'}
            <div>
              <div style="font-size: 1.35rem; font-weight: 800; color: #0d1f2d;">CFM Distributors, Inc.</div>
              <div style="font-size: 0.82rem; color: #64748b;">3513 E 14th St, Kansas City, MO 64127 | (816) 483-3111 | cfmdistributors.com</div>
            </div>
          </div>
          <div style="text-align: right;">
            <div style="font-size: 1.4rem; font-weight: 800; color: var(--cfm-blue); text-transform: uppercase;">Equipment Quotation</div>
            <div style="font-size: 0.82rem; color: #64748b; margin-top: 4px;">Quote #: <b id="quote-number">CFM-2026-0921</b></div>
            <div style="font-size: 0.82rem; color: #64748b;">Date: <span id="quote-date"></span></div>
          </div>
        </div>

        <!-- CONTRACTOR & JOB DETAILS -->
        <div class="proposal-info-box" style="display: grid; grid-template-columns: 1fr 1fr; gap: 2rem; margin-bottom: 1.75rem; padding: 1.25rem; border-radius: var(--radius-sm);">
          <div>
            <div style="font-size: 0.76rem; font-weight: 700; color: #64748b; text-transform: uppercase; margin-bottom: 6px;">Prepared For (Contractor):</div>
            <input type="text" id="quote-cust-name" value="ABC Mechanical Services" style="width: 100%; font-weight: 600; padding: 4px 8px; margin-bottom: 4px; border: 1px solid var(--border-subtle); border-radius: 4px;">
            <input type="text" id="quote-cust-contact" value="Attn: Purchasing & Estimating" style="width: 100%; font-size: 0.82rem; padding: 4px 8px; border: 1px solid var(--border-subtle); border-radius: 4px;">
          </div>
          <div>
            <div style="font-size: 0.76rem; font-weight: 700; color: #64748b; text-transform: uppercase; margin-bottom: 6px;">Project / Job Information:</div>
            <input type="text" id="quote-job-name" value="Commercial HVAC Replacement" style="width: 100%; font-weight: 600; padding: 4px 8px; margin-bottom: 4px; border: 1px solid var(--border-subtle); border-radius: 4px;">
            <input type="text" id="quote-sales-rep" value="CFM Sales Rep: Commercial Counter / Technical Support" style="width: 100%; font-size: 0.82rem; padding: 4px 8px; border: 1px solid var(--border-subtle); border-radius: 4px;">
          </div>
        </div>

        <!-- LINE ITEMS TABLE -->
        <table class="data-table" style="margin-bottom: 1.5rem;">
          <thead>
            <tr>
              <th style="width: 40px;">#</th>
              <th style="width: 200px;">Model Number</th>
              <th>Description & Equipment Details</th>
              <th style="width: 120px; text-align: right;">Unit Price</th>
              <th style="width: 70px; text-align: center;">Qty</th>
              <th style="width: 130px; text-align: right;">Extended Price</th>
              <th class="no-print" style="width: 50px; text-align: center;">Del</th>
            </tr>
          </thead>
          <tbody id="quote-table-body">
            <!-- Rendered dynamically -->
          </tbody>
        </table>

        <!-- SMART ACCESSORIES RECOMMENDATIONS INSIDE PROPOSAL -->
        <div id="quote-smart-recs" class="quote-recs-box no-print" style="display: none;">
          <div class="quote-recs-title">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/></svg>
            Recommended Job Accessories for Quoted Units:
          </div>
          <div class="quote-recs-list" id="quote-smart-recs-list">
            <!-- Dynamic recommended accessories chips -->
          </div>
        </div>

        <!-- TOTALS & FREIGHT -->
        <div style="display: flex; justify-content: flex-end; margin-top: 1.5rem;">
          <div style="width: 320px;">
            <div style="display: flex; justify-content: space-between; padding: 6px 0; font-size: 0.9rem; color: #64748b;">
              <span>Equipment Subtotal:</span>
              <span id="doc-subtotal" style="font-family: 'JetBrains Mono', monospace; font-weight: 600; color: #0f172a;">$0.00</span>
            </div>
            <div style="display: flex; justify-content: space-between; padding: 6px 0; font-size: 0.9rem; color: #64748b;">
              <span>Freight (CFM Policy Tier):</span>
              <span id="doc-freight" style="font-family: 'JetBrains Mono', monospace; font-weight: 600; color: #0f172a;">$0.00</span>
            </div>
            <div style="display: flex; justify-content: space-between; padding: 10px 0; font-size: 1.25rem; font-weight: 800; color: #0d1f2d; border-top: 2px solid #0d1f2d; margin-top: 6px;">
              <span>Total Investment:</span>
              <span id="doc-grand-total" style="font-family: 'JetBrains Mono', monospace; color: #059669;">$0.00</span>
            </div>
          </div>
        </div>

        <!-- TERMS & SIGNATURE -->
        <div style="margin-top: 2.5rem; padding-top: 1.5rem; border-top: 1px solid var(--border-subtle); font-size: 0.74rem; color: #64748b; line-height: 1.5;">
          <p><b>Quotation Terms & Conditions:</b> Prices valid for 30 days from quote date. Equipment subject to prior sale and manufacturer lead times. Standard freight policy applies based on total order value (Orders over $12,000 qualify for free standard freight). All curb adapters and custom mod shop equipment require customer dimension sign-off prior to fabrication.</p>
          <div style="margin-top: 2rem; display: flex; justify-content: space-between;">
            <div style="width: 260px; border-top: 1px solid #94a3b8; padding-top: 6px;">Authorized CFM Representative</div>
            <div style="width: 260px; border-top: 1px solid #94a3b8; padding-top: 6px;">Accepted By (Customer Signature & Date)</div>
          </div>
        </div>
      </div>
    </section>

  </main>

  <!-- ACCESSORY MODAL -->
  <div id="acc-modal-overlay" class="modal-overlay" onclick="closeAccessoryModal(event)">
    <div class="modal-box" onclick="event.stopPropagation()">
      <div class="modal-header">
        <div>
          <h3 id="modal-title">⚙️ Compatible Accessories</h3>
          <p id="modal-subtitle">Direct fit accessories, roof curbs, economizers, hail guards & sensors</p>
        </div>
        <button class="modal-close-btn" onclick="closeAccessoryModal()">&times;</button>
      </div>

      <div class="modal-toolbar">
        <div class="filter-pills" id="modal-category-pills">
          <!-- Rendered dynamically -->
        </div>
        <div class="modal-search">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
          <input type="text" id="modal-search-input" placeholder="Search accessories in this series..." oninput="filterModalAccessories()">
        </div>
      </div>

      <div class="modal-body">
        <table class="data-table">
          <thead>
            <tr>
              <th style="width: 140px;">Category</th>
              <th style="width: 220px;">Model Number</th>
              <th>Description & Notes</th>
              <th style="width: 100px; text-align: center;">Stock</th>
              <th style="width: 120px; text-align: right;">Sell Price</th>
              <th style="width: 120px; text-align: center;">Action</th>
            </tr>
          </thead>
          <tbody id="modal-acc-table-body">
            <!-- Rendered dynamically -->
          </tbody>
        </table>
      </div>

      <div class="modal-footer">
        <span id="modal-items-count">0 accessories available</span>
        <button class="btn btn-outline-light" style="color: var(--text-main); border-color: var(--border-strong);" onclick="closeAccessoryModal()">Done</button>
      </div>
    </div>
  </div>

  <!-- SLIDE-OUT QUOTE DRAWER -->
  <div id="drawer-overlay" class="drawer-overlay" onclick="toggleDrawer()"></div>
  <aside id="quote-drawer" class="quote-drawer">
    <div class="drawer-header">
      <h3>
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="9" cy="21" r="1"/><circle cx="20" cy="21" r="1"/><path d="M1 1h4l2.68 13.39a2 2 0 0 0 2 1.61h9.72a2 2 0 0 0 2-1.61L23 6H6"/></svg>
        Active Quote (<span id="drawer-item-count">0</span>)
      </h3>
      <button class="drawer-close-btn" onclick="toggleDrawer()">&times;</button>
    </div>

    <div class="drawer-body">
      <div id="drawer-items-list">
        <!-- Dynamic quote line items -->
      </div>

      <!-- Quick Suggested Accessories inside drawer -->
      <div id="drawer-smart-recs" class="quote-recs-box" style="display: none;">
        <div class="quote-recs-title">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/></svg>
          Suggested Series Accessories:
        </div>
        <div class="quote-recs-list" id="drawer-smart-recs-list">
          <!-- Suggested accessories rendered here -->
        </div>
      </div>
    </div>

    <div class="drawer-footer">
      <div class="summary-row">
        <span>Subtotal:</span>
        <span id="drawer-subtotal" style="font-family: 'JetBrains Mono', monospace; font-weight: 600;">$0.00</span>
      </div>
      <div class="summary-row">
        <span>Freight:</span>
        <span id="drawer-freight" style="font-family: 'JetBrains Mono', monospace;">$0.00</span>
      </div>
      <div class="summary-row total">
        <span>Total:</span>
        <span id="drawer-total" style="font-family: 'JetBrains Mono', monospace; color: #047857;">$0.00</span>
      </div>
      <div style="display: flex; gap: 0.75rem;">
        <button class="btn btn-outline-light" style="flex: 1; justify-content: center; color: var(--text-main); border-color: var(--border-strong);" onclick="toggleDrawer(); switchTab('tab-quote');">Open Full Proposal</button>
        <button class="btn btn-primary" style="flex: 1; justify-content: center;" onclick="toggleDrawer(); switchTab('tab-quote'); window.print();">Print</button>
      </div>
    </div>
  </aside>

  <!-- TOAST NOTIFICATION -->
  <div id="toast" class="toast">
    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#1ec8a0" stroke-width="2"><polyline points="20 6 9 17 4 12"/></svg>
    <span id="toast-msg">Item added to quote</span>
  </div>

  <!-- EMBEDDED APPLICATION DATA & SCRIPT -->
  <script>
    const CFM_DB = {db_json};

    // APPLICATION STATE
    const state = {{
      currentTab: 'tab-all',
      searchQuery: '',
      filters: {{
        sheet: '',
        ref: '',
        avail: ''
      }},
      page: 1,
      pageSize: 50,
      sortCol: 'm',
      sortAsc: true,
      quote: JSON.parse(localStorage.getItem('cfm_quote') || '[]'),
      activeMatrix: 'Matrix on stock',
      activeGuideSeries: 'sunline-3-6',
      modalSeriesKey: '',
      modalCategory: '',
      guideCategory: ''
    }};

    
    // THEME CONTROLLER
    function toggleTheme() {{
      const isDark = document.body.classList.toggle('dark-theme');
      localStorage.setItem('cfm_theme', isDark ? 'dark' : 'light');
      syncThemeControls(isDark);
      showToast(isDark ? 'Dark theme enabled' : 'Light theme enabled');
    }}

    function syncThemeControls(isDark) {{
      const icon = document.getElementById('theme-icon');
      const text = document.getElementById('theme-text');
      if (icon) icon.textContent = isDark ? '☀️' : '🌙';
      if (text) text.textContent = isDark ? 'Light Mode' : 'Dark Mode';
    }}

    function initTheme() {{
      const saved = localStorage.getItem('cfm_theme');
      const prefersDark = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
      const isDark = saved === 'dark' || (!saved && prefersDark);
      if (isDark) {{
        document.body.classList.add('dark-theme');
      }} else {{
        document.body.classList.remove('dark-theme');
      }}
      syncThemeControls(isDark);
    }}

    // INITIALIZATION
    document.addEventListener('DOMContentLoaded', () => {{
      initTheme();
      document.getElementById('quote-date').textContent = new Date().toLocaleDateString('en-US', {{ year: 'numeric', month: 'long', day: 'numeric' }});
      
      populateSheetFilter();
      populateReferenceSelect();
      populateCalculator();
      initAccessoryGuide();
      applyFilters();
      refreshQuoteBadges();
      renderMatrix();

      // Shortcut key '/' to search
      document.addEventListener('keydown', (e) => {{
        if (e.key === '/' && document.activeElement.tagName !== 'INPUT' && document.activeElement.tagName !== 'SELECT') {{
          e.preventDefault();
          document.getElementById('global-search').focus();
        }}
      }});
    }});

    // TOAST
    function showToast(msg) {{
      const toast = document.getElementById('toast');
      document.getElementById('toast-msg').textContent = msg;
      toast.classList.add('show');
      setTimeout(() => toast.classList.remove('show'), 2400);
    }}

    // CLIPBOARD COPY
    function copyText(text, successMsg = 'Copied to clipboard!') {{
      if (!text) return;
      navigator.clipboard.writeText(text).then(() => {{
        showToast(successMsg);
      }}).catch(() => {{
        const el = document.createElement('textarea');
        el.value = text;
        document.body.appendChild(el);
        el.select();
        document.execCommand('copy');
        document.body.removeChild(el);
        showToast(successMsg);
      }});
    }}

    // MONEY FORMATTER
    function formatMoney(num) {{
      if (!num || isNaN(num) || parseFloat(num) <= 0) return '—';
      return '$' + parseFloat(num).toLocaleString('en-US', {{ minimumFractionDigits: 2, maximumFractionDigits: 2 }});
    }}

    // TAB SWITCHING
    function switchTab(tabId) {{
      state.currentTab = tabId;
      document.querySelectorAll('.nav-tab').forEach(t => t.classList.toggle('active', t.dataset.tab === tabId));

      document.querySelectorAll('.view-section').forEach(s => s.classList.remove('active'));

      if (tabId === 'tab-matrix') {{
        document.getElementById('view-matrix').classList.add('active');
        renderMatrix();
      }} else if (tabId === 'tab-ref') {{
        document.getElementById('view-ref').classList.add('active');
        const refSelect = document.getElementById('ref-select');
        if (refSelect.value) loadReferenceDoc(refSelect.value);
      }} else if (tabId === 'tab-quote') {{
        document.getElementById('view-quote').classList.add('active');
        renderQuoteTab();
      }} else if (tabId === 'tab-accessories') {{
        document.getElementById('view-accessories').classList.add('active');
        renderAccessoryGuide(state.activeGuideSeries);
      }} else {{
        document.getElementById('view-catalog').classList.add('active');
        const tabNames = {{
          'tab-all': 'All Products',
          'tab-rtu': 'Commercial Rooftops & Stock',
          'tab-res': 'Residential Split Systems',
          'tab-split': 'Commercial Splits (2.5-50T)',
          'tab-ductless': 'Ductless & VRF Systems',
          'tab-applied': 'Applied Systems & Chillers',
          'tab-specialties': 'Specialties, Parts & Heaters'
        }};
        document.getElementById('active-tab-name').textContent = tabNames[tabId] || 'Equipment Catalog';
        populateSheetFilter();
        state.page = 1;
        applyFilters();
      }}
      window.scrollTo({{ top: 0, behavior: 'smooth' }});
    }}

    // POPULATE WORKSHEET SELECT
    function populateSheetFilter() {{
      const select = document.getElementById('sheet-filter');
      select.innerHTML = '<option value="">All Catalog Sheets</option>';

      let allowedSheets = [];
      if (state.currentTab === 'tab-rtu') {{
        allowedSheets = CFM_DB.sheet_groups["Commercial Rooftops & KC Stock"] || [];
      }} else if (state.currentTab === 'tab-res') {{
        allowedSheets = CFM_DB.sheet_groups["Residential Split Systems"] || [];
      }} else if (state.currentTab === 'tab-split') {{
        allowedSheets = CFM_DB.sheet_groups["Commercial Splits (2.5-50T)"] || [];
      }} else if (state.currentTab === 'tab-ductless') {{
        allowedSheets = CFM_DB.sheet_groups["Ductless & VRF Systems"] || [];
      }} else if (state.currentTab === 'tab-applied') {{
        allowedSheets = CFM_DB.sheet_groups["Applied Systems & Chillers"] || [];
      }} else if (state.currentTab === 'tab-specialties') {{
        allowedSheets = CFM_DB.sheet_groups["Specialties, Heaters & Accessories"] || [];
      }} else {{
        const allSet = new Set(CFM_DB.items.map(i => i.s));
        allowedSheets = Array.from(allSet).sort();
      }}

      allowedSheets.forEach(sName => {{
        const opt = document.createElement('option');
        opt.value = sName;
        opt.textContent = sName;
        select.appendChild(opt);
      }});
      select.value = state.filters.sheet;
    }}

    // GLOBAL SEARCH
    function handleGlobalSearch(query) {{
      state.searchQuery = query.trim().toLowerCase();
      state.page = 1;
      if (state.currentTab === 'tab-matrix' || state.currentTab === 'tab-ref' || state.currentTab === 'tab-quote' || state.currentTab === 'tab-accessories') {{
        switchTab('tab-all');
      }}
      applyFilters();
    }}

    // SET FILTERS
    function setFilter(type, val, btn) {{
      state.filters[type] = val;
      state.page = 1;
      if (btn) {{
        const parent = btn.parentElement;
        parent.querySelectorAll('.pill-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
      }}
      applyFilters();
    }}

    // APPLY FILTERS & RENDER TABLE
    function applyFilters() {{
      const sheetFilterVal = document.getElementById('sheet-filter').value;
      state.filters.sheet = sheetFilterVal;

      let allowedSheets = null;
      if (state.currentTab === 'tab-rtu') {{
        allowedSheets = new Set(CFM_DB.sheet_groups["Commercial Rooftops & KC Stock"] || []);
      }} else if (state.currentTab === 'tab-res') {{
        allowedSheets = new Set(CFM_DB.sheet_groups["Residential Split Systems"] || []);
      }} else if (state.currentTab === 'tab-split') {{
        allowedSheets = new Set(CFM_DB.sheet_groups["Commercial Splits (2.5-50T)"] || []);
      }} else if (state.currentTab === 'tab-ductless') {{
        allowedSheets = new Set(CFM_DB.sheet_groups["Ductless & VRF Systems"] || []);
      }} else if (state.currentTab === 'tab-applied') {{
        allowedSheets = new Set(CFM_DB.sheet_groups["Applied Systems & Chillers"] || []);
      }} else if (state.currentTab === 'tab-specialties') {{
        allowedSheets = new Set(CFM_DB.sheet_groups["Specialties, Heaters & Accessories"] || []);
      }}

      const q = state.searchQuery;
      const refFilter = state.filters.ref;
      const availFilter = state.filters.avail;

      const filtered = CFM_DB.items.filter(item => {{
        if (allowedSheets && !allowedSheets.has(item.s)) return false;
        if (sheetFilterVal && item.s !== sheetFilterVal) return false;
        if (refFilter && item.ref !== refFilter) return false;

        if (availFilter === 'stock') {{
          const a = item.a || '';
          if (!a.includes('Norm') && a !== '1' && !a.includes('KC')) return false;
        }} else if (availFilter === 'call') {{
          const a = item.a || '';
          if (!a.includes('Call')) return false;
        }}

        if (q) {{
          const searchStr = `${{item.m}} ${{item.d}} ${{item.s}} ${{item.c}} ${{item.f}}`.toLowerCase();
          if (!searchStr.includes(q)) return false;
        }}

        return true;
      }});

      // Sorting
      filtered.sort((a, b) => {{
        let valA = a[state.sortCol] || '';
        let valB = b[state.sortCol] || '';
        if (state.sortCol === 'mlp' || state.sortCol === 'sp') {{
          valA = parseFloat(valA) || 0;
          valB = parseFloat(valB) || 0;
        }} else {{
          valA = valA.toString().toLowerCase();
          valB = valB.toString().toLowerCase();
        }}
        if (valA < valB) return state.sortAsc ? -1 : 1;
        if (valA > valB) return state.sortAsc ? 1 : -1;
        return 0;
      }});

      document.getElementById('filtered-count').textContent = filtered.length.toLocaleString();
      renderCatalogTable(filtered);
    }}

    function sortTable(col) {{
      if (state.sortCol === col) {{
        state.sortAsc = !state.sortAsc;
      }} else {{
        state.sortCol = col;
        state.sortAsc = true;
      }}
      applyFilters();
    }}

    function changePageSize(size) {{
      state.pageSize = parseInt(size);
      state.page = 1;
      applyFilters();
    }}

    // RENDER CATALOG TABLE
    function renderCatalogTable(items) {{
      const tbody = document.getElementById('catalog-table-body');
      tbody.innerHTML = '';

      const start = (state.page - 1) * state.pageSize;
      const paginated = items.slice(start, start + state.pageSize);

      if (paginated.length === 0) {{
        tbody.innerHTML = `<tr><td colspan="8" style="text-align: center; color: var(--text-light); padding: 3rem;">No equipment matches your search or filter criteria.</td></tr>`;
        renderPagination(0);
        return;
      }}

      paginated.forEach(item => {{
        const tr = document.createElement('tr');

        let availClass = 'avail-other';
        let availText = item.a || 'Quote';
        if (availText.includes('Norm') || availText === '1' || availText.includes('KC')) {{
          availClass = 'avail-stock';
          if (availText === '*Norm Stk' || availText === '1') availText = 'In Stock';
        }} else if (availText.includes('Call')) {{
          availClass = 'avail-call';
        }}

        let refBadge = '';
        if (item.ref === 'R-454B') refBadge = '<span class="badge-a2l">R-454B (A2L)</span>';
        else if (item.ref === 'R-410A') refBadge = '<span class="badge-r410a">R-410A</span>';

        let sellDisplay = formatMoney(item.sp);
        if (sellDisplay === '—' && item.mlp && item.sx && parseFloat(item.mlp) > 0 && parseFloat(item.sx) > 0) {{
          sellDisplay = formatMoney(parseFloat(item.mlp) * parseFloat(item.sx));
        }}

        // Accessory button for equipment that has paired accessories
        let accBtn = '';
        if (item.sk && CFM_DB.series_catalog[item.sk] && !item.is_acc) {{
          const accCount = CFM_DB.series_catalog[item.sk].accessories.length;
          accBtn = `<button class="btn-acc" onclick="openAccessoryModal('${{item.sk}}', '${{escapeQuotes(item.m)}}')">⚙️ Accessories (${{accCount}})</button>`;
        }}

        tr.innerHTML = `
          <td><span class="avail-badge ${{availClass}}">${{availText}}</span></td>
          <td>
            <div class="model-cell" onclick="copyText('${{item.m}}', 'Model copied!')">
              ${{item.m || '—'}}
              <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>
            </div>
            ${{refBadge}}
          </td>
          <td>
            <div style="font-weight: 500;">${{item.d || '—'}}</div>
            ${{item.f ? `<div style="font-size: 0.72rem; color: var(--text-light);">Family: ${{item.f}}</div>` : ''}}
          </td>
          <td>
            <div style="font-weight: 500; font-size: 0.8rem;">${{item.s}}</div>
            <div style="font-size: 0.72rem; color: var(--text-light);">${{item.c || ''}}</div>
          </td>
          <td style="text-align: right;" class="price-mlp">${{formatMoney(item.mlp)}}</td>
          <td style="text-align: center;" class="multiplier-tag">${{item.cx || item.sx ? `${{item.cx || '—'}} / ${{item.sx || '—'}}` : '—'}}</td>
          <td style="text-align: right;" class="price-sell">${{sellDisplay !== '—' ? sellDisplay : 'Call'}}</td>
          <td style="text-align: center; white-space: nowrap;">
            <button class="btn-add" onclick="addToQuote(${{item.id}})">+ Quote</button>
            ${{accBtn}}
          </td>
        `;
        tbody.appendChild(tr);
      }});

      renderPagination(items.length);
    }}

    function escapeQuotes(str) {{
      if (!str) return '';
      return str.replace(/'/g, "\\'");
    }}

    // PAGINATION
    function renderPagination(total) {{
      const container = document.getElementById('pagination-controls');
      container.innerHTML = '';
      const totalPages = Math.ceil(total / state.pageSize);
      if (totalPages <= 1) return;

      const prevBtn = document.createElement('button');
      prevBtn.className = 'page-btn';
      prevBtn.textContent = '« Prev';
      prevBtn.disabled = state.page === 1;
      prevBtn.onclick = () => {{ state.page--; applyFilters(); }};
      container.appendChild(prevBtn);

      const maxButtons = 5;
      let startPage = Math.max(1, state.page - Math.floor(maxButtons / 2));
      let endPage = Math.min(totalPages, startPage + maxButtons - 1);
      if (endPage - startPage < maxButtons - 1) {{
        startPage = Math.max(1, endPage - maxButtons + 1);
      }}

      for (let i = startPage; i <= endPage; i++) {{
        const btn = document.createElement('button');
        btn.className = `page-btn ${{state.page === i ? 'active' : ''}}`;
        btn.textContent = i;
        btn.onclick = () => {{ state.page = i; applyFilters(); }};
        container.appendChild(btn);
      }}

      const nextBtn = document.createElement('button');
      nextBtn.className = 'page-btn';
      nextBtn.textContent = 'Next »';
      nextBtn.disabled = state.page === totalPages;
      nextBtn.onclick = () => {{ state.page++; applyFilters(); }};
      container.appendChild(nextBtn);
    }}

    // ACCESSORY MODAL CONTROLLER
    function openAccessoryModal(seriesKey, modelCode) {{
      state.modalSeriesKey = seriesKey;
      state.modalCategory = '';
      const series = CFM_DB.series_catalog[seriesKey];
      if (!series) return;

      document.getElementById('modal-title').textContent = `⚙️ Compatible Accessories: ${{series.title}}`;
      document.getElementById('modal-subtitle').textContent = `Paired accessories for ${{modelCode || 'this series'}} (${{series.accessories.length}} accessories)`;
      document.getElementById('modal-search-input').value = '';

      // Populate category pills
      const cats = ['All'];
      const catCounts = {{ 'All': series.accessories.length }};
      series.accessories.forEach(a => {{
        catCounts[a.c] = (catCounts[a.c] || 0) + 1;
        if (!cats.includes(a.c)) cats.push(a.c);
      }});

      const pillContainer = document.getElementById('modal-category-pills');
      pillContainer.innerHTML = '';
      cats.forEach(c => {{
        const btn = document.createElement('button');
        btn.className = `pill-btn ${{c === 'All' ? 'active' : ''}}`;
        btn.textContent = `${{c}} (${{catCounts[c]}})`;
        btn.onclick = () => {{
          pillContainer.querySelectorAll('.pill-btn').forEach(b => b.classList.remove('active'));
          btn.classList.add('active');
          state.modalCategory = c === 'All' ? '' : c;
          filterModalAccessories();
        }};
        pillContainer.appendChild(btn);
      }});

      filterModalAccessories();
      document.getElementById('acc-modal-overlay').classList.add('open');
    }}

    function closeAccessoryModal(e) {{
      if (e && e.target !== e.currentTarget && !e.target.classList.contains('modal-close-btn') && !e.target.closest('.modal-close-btn') && e.target.tagName !== 'BUTTON') {{
        return;
      }}
      document.getElementById('acc-modal-overlay').classList.remove('open');
    }}

    function filterModalAccessories() {{
      const series = CFM_DB.series_catalog[state.modalSeriesKey];
      if (!series) return;

      const q = document.getElementById('modal-search-input').value.trim().toLowerCase();
      const cat = state.modalCategory;

      const list = series.accessories.filter(a => {{
        if (cat && a.c !== cat) return false;
        if (q) {{
          const str = `${{a.m}} ${{a.d}} ${{a.c}}`.toLowerCase();
          if (!str.includes(q)) return false;
        }}
        return true;
      }});

      document.getElementById('modal-items-count').textContent = `${{list.length}} accessories shown`;
      const tbody = document.getElementById('modal-acc-table-body');
      tbody.innerHTML = '';

      if (list.length === 0) {{
        tbody.innerHTML = `<tr><td colspan="6" style="text-align: center; color: var(--text-light); padding: 2.5rem;">No accessories match this filter.</td></tr>`;
        return;
      }}

      list.forEach(a => {{
        const tr = document.createElement('tr');
        let availClass = 'avail-other';
        let availText = a.stock || 'Call';
        if (availText.includes('Norm') || availText === 'In Stock') {{
          availClass = 'avail-stock';
          availText = 'In Stock';
        }}

        tr.innerHTML = `
          <td><span class="nav-count" style="font-size: 0.72rem;">${{a.c}}</span></td>
          <td>
            <div class="model-cell" onclick="copyText('${{a.m}}', 'Accessory model copied!')">
              ${{a.m}}
              <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>
            </div>
          </td>
          <td>
            <div style="font-weight: 500;">${{a.d}}</div>
            <div style="font-size: 0.7rem; color: var(--text-light);">Source: ${{a.s}}</div>
          </td>
          <td style="text-align: center;"><span class="avail-badge ${{availClass}}">${{availText}}</span></td>
          <td style="text-align: right;" class="price-sell">${{formatMoney(a.price)}}</td>
          <td style="text-align: center;">
            <button class="btn-add" onclick="addAccessoryToQuote('${{escapeQuotes(a.m)}}', '${{escapeQuotes(a.d)}}', ${{a.price}})">+ Quote</button>
          </td>
        `;
        tbody.appendChild(tr);
      }});
    }}

    // EQUIPMENT & ACCESSORY GUIDE VIEW CONTROLLER
    function initAccessoryGuide() {{
      const grid = document.getElementById('series-selector-grid');
      grid.innerHTML = '';

      Object.entries(CFM_DB.series_catalog).forEach(([key, s]) => {{
        const card = document.createElement('div');
        card.className = `series-card ${{key === state.activeGuideSeries ? 'active' : ''}}`;
        card.dataset.series = key;
        card.onclick = () => renderAccessoryGuide(key);

        card.innerHTML = `
          <div>
            <div class="series-card-top">
              <span class="series-card-badge">${{s.badge}}</span>
              <span class="nav-count">${{s.accessories.length}} Items</span>
            </div>
            <div class="series-card-title">${{s.title}}</div>
            <div class="series-card-desc">${{s.desc}}</div>
          </div>
          <div style="font-size: 0.78rem; font-weight: 600; color: var(--cfm-blue); display: flex; align-items: center; gap: 4px;">
            Explore Accessories →
          </div>
        `;
        grid.appendChild(card);
      }});
    }}

    function renderAccessoryGuide(seriesKey) {{
      state.activeGuideSeries = seriesKey;
      state.guideCategory = '';
      const series = CFM_DB.series_catalog[seriesKey];
      if (!series) return;

      document.querySelectorAll('.series-card').forEach(c => {{
        c.classList.toggle('active', c.dataset.series === seriesKey);
      }});

      const display = document.getElementById('series-accessory-display');
      display.style.display = 'block';

      document.getElementById('active-guide-title').textContent = `${{series.title}} — Accessory Catalog`;
      document.getElementById('active-guide-desc').textContent = `${{series.desc}} (${{series.accessories.length}} accessories listed)`;
      document.getElementById('guide-acc-search').value = '';

      const cats = ['All'];
      const catCounts = {{ 'All': series.accessories.length }};
      series.accessories.forEach(a => {{
        catCounts[a.c] = (catCounts[a.c] || 0) + 1;
        if (!cats.includes(a.c)) cats.push(a.c);
      }});

      const pillContainer = document.getElementById('guide-category-pills');
      pillContainer.innerHTML = '';
      cats.forEach(c => {{
        const btn = document.createElement('button');
        btn.className = `pill-btn ${{c === 'All' ? 'active' : ''}}`;
        btn.textContent = `${{c}} (${{catCounts[c]}})`;
        btn.onclick = () => {{
          pillContainer.querySelectorAll('.pill-btn').forEach(b => b.classList.remove('active'));
          btn.classList.add('active');
          state.guideCategory = c === 'All' ? '' : c;
          filterGuideAccessories();
        }};
        pillContainer.appendChild(btn);
      }});

      filterGuideAccessories();
      display.scrollIntoView({{ behavior: 'smooth', block: 'nearest' }});
    }}

    function filterGuideAccessories() {{
      const series = CFM_DB.series_catalog[state.activeGuideSeries];
      if (!series) return;

      const q = document.getElementById('guide-acc-search').value.trim().toLowerCase();
      const cat = state.guideCategory;

      const list = series.accessories.filter(a => {{
        if (cat && a.c !== cat) return false;
        if (q) {{
          const str = `${{a.m}} ${{a.d}} ${{a.c}}`.toLowerCase();
          if (!str.includes(q)) return false;
        }}
        return true;
      }});

      const tbody = document.getElementById('guide-acc-table-body');
      tbody.innerHTML = '';

      if (list.length === 0) {{
        tbody.innerHTML = `<tr><td colspan="6" style="text-align: center; color: var(--text-light); padding: 2.5rem;">No accessories match this filter.</td></tr>`;
        return;
      }}

      list.forEach(a => {{
        const tr = document.createElement('tr');
        let availClass = 'avail-other';
        let availText = a.stock || 'Call';
        if (availText.includes('Norm') || availText === 'In Stock') {{
          availClass = 'avail-stock';
          availText = 'In Stock';
        }}

        tr.innerHTML = `
          <td><span class="nav-count" style="font-size: 0.72rem;">${{a.c}}</span></td>
          <td>
            <div class="model-cell" onclick="copyText('${{a.m}}', 'Accessory model copied!')">
              ${{a.m}}
              <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>
            </div>
          </td>
          <td>
            <div style="font-weight: 500;">${{a.d}}</div>
            <div style="font-size: 0.72rem; color: var(--text-light);">Catalog Sheet: ${{a.s}}</div>
          </td>
          <td style="text-align: center;"><span class="avail-badge ${{availClass}}">${{availText}}</span></td>
          <td style="text-align: right;" class="price-sell">${{formatMoney(a.price)}}</td>
          <td style="text-align: center;">
            <button class="btn-add" onclick="addAccessoryToQuote('${{escapeQuotes(a.m)}}', '${{escapeQuotes(a.d)}}', ${{a.price}})">+ Add to Quote</button>
          </td>
        `;
        tbody.appendChild(tr);
      }});
    }}

    function exportGuideCSV() {{
      const series = CFM_DB.series_catalog[state.activeGuideSeries];
      if (!series) return;
      let csv = "Category,Model,Description,Stock Status,Sell Price,Source Sheet\\r\\n";
      series.accessories.forEach(a => {{
        const safeDesc = `"${{(a.d || '').replace(/"/g, '""')}}"`;
        csv += `"${{a.c}}","${{a.m}}",${{safeDesc}},"${{a.stock}}","${{a.price}}","${{a.s}}"\\r\\n`;
      }});
      downloadFile(csv, `${{series.title.replace(/[^a-zA-Z0-9]/g, '_')}}_Accessories.csv`, 'text/csv');
    }}

    // QUOTE WORKBENCH & CART
    function addToQuote(itemId) {{
      const item = CFM_DB.items.find(i => i.id === itemId);
      if (!item) return;

      const existing = state.quote.find(q => q.id === itemId);
      if (existing) {{
        existing.qty += 1;
      }} else {{
        let price = parseFloat(item.sp) || 0;
        if (!price && item.mlp && item.sx) {{
          price = parseFloat(item.mlp) * parseFloat(item.sx);
        }}
        state.quote.push({{
          id: item.id,
          m: item.m || 'Custom Item',
          d: item.d || 'Equipment Item',
          price: price,
          qty: 1,
          mlp: parseFloat(item.mlp) || 0,
          sk: item.sk || ''
        }});
      }}

      saveQuote();
      showToast(`Added ${{item.m || 'Item'}} to Quote`);
    }}

    function addAccessoryToQuote(model, desc, price) {{
      const existing = state.quote.find(q => q.m === model);
      if (existing) {{
        existing.qty += 1;
      }} else {{
        state.quote.push({{
          id: 'acc_' + Date.now() + Math.random().toString(36).substr(2, 4),
          m: model,
          d: desc,
          price: parseFloat(price) || 0,
          qty: 1,
          mlp: parseFloat(price) || 0,
          sk: ''
        }});
      }}
      saveQuote();
      showToast(`Added ${{model}} to Quote`);
    }}

    function saveQuote() {{
      localStorage.setItem('cfm_quote', JSON.stringify(state.quote));
      refreshQuoteBadges();
      renderDrawer();
      if (state.currentTab === 'tab-quote') renderQuoteTab();
    }}

    function refreshQuoteBadges() {{
      const totalQty = state.quote.reduce((acc, i) => acc + i.qty, 0);
      document.getElementById('cart-badge-count').textContent = totalQty;
      document.getElementById('drawer-item-count').textContent = totalQty;
    }}

    function toggleDrawer() {{
      const drawer = document.getElementById('quote-drawer');
      const overlay = document.getElementById('drawer-overlay');
      drawer.classList.toggle('open');
      overlay.classList.toggle('open');
      if (drawer.classList.contains('open')) renderDrawer();
    }}

    function renderDrawer() {{
      const list = document.getElementById('drawer-items-list');
      list.innerHTML = '';

      if (state.quote.length === 0) {{
        list.innerHTML = `<div style="text-align: center; color: var(--text-light); padding: 3rem;">No equipment in active quote.</div>`;
        document.getElementById('drawer-subtotal').textContent = '$0.00';
        document.getElementById('drawer-freight').textContent = '$0.00';
        document.getElementById('drawer-total').textContent = '$0.00';
        document.getElementById('drawer-smart-recs').style.display = 'none';
        return;
      }}

      let subtotal = 0;
      const quotedSeriesKeys = new Set();

      state.quote.forEach((item, idx) => {{
        const itemTotal = item.price * item.qty;
        subtotal += itemTotal;
        if (item.sk) quotedSeriesKeys.add(item.sk);

        const div = document.createElement('div');
        div.className = 'drawer-item';
        div.innerHTML = `
          <div style="display: flex; justify-content: space-between;">
            <div class="drawer-item-title">${{item.m}}</div>
            <button class="btn-del-item" onclick="removeQuoteItem(${{idx}})">&times;</button>
          </div>
          <div class="drawer-item-desc">${{item.d}}</div>
          <div class="drawer-item-calc">
            <div class="qty-controls">
              <button class="qty-btn" onclick="changeItemQty(${{idx}}, -1)">-</button>
              <span class="qty-val">${{item.qty}}</span>
              <button class="qty-btn" onclick="changeItemQty(${{idx}}, 1)">+</button>
            </div>
            <div style="font-weight: 700; font-family: 'JetBrains Mono', monospace; color: #047857;">
              $${{itemTotal.toLocaleString('en-US', {{minimumFractionDigits: 2, maximumFractionDigits: 2}})}}
            </div>
          </div>
        `;
        list.appendChild(div);
      }});

      let freight = 0;
      if (subtotal > 0 && subtotal < 1000) freight = 150;
      else if (subtotal >= 1000 && subtotal < 5000) freight = 275;
      else if (subtotal >= 5000 && subtotal < 12000) freight = 550;
      else freight = 0;

      const grand = subtotal + freight;
      document.getElementById('drawer-subtotal').textContent = '$' + subtotal.toLocaleString('en-US', {{minimumFractionDigits: 2, maximumFractionDigits: 2}});
      document.getElementById('drawer-freight').textContent = freight === 0 ? 'FREE' : '$' + freight.toFixed(2);
      document.getElementById('drawer-total').textContent = '$' + grand.toLocaleString('en-US', {{minimumFractionDigits: 2, maximumFractionDigits: 2}});

      // Render Drawer Recommendations
      renderSmartRecommendations(quotedSeriesKeys, 'drawer-smart-recs', 'drawer-smart-recs-list');
    }}

    function renderSmartRecommendations(seriesKeys, containerId, listId) {{
      const container = document.getElementById(containerId);
      const list = document.getElementById(listId);
      list.innerHTML = '';

      if (seriesKeys.size === 0) {{
        container.style.display = 'none';
        return;
      }}

      let recItems = [];
      seriesKeys.forEach(sKey => {{
        const s = CFM_DB.series_catalog[sKey];
        if (!s) return;
        // Top accessories: Curbs, Economizers, Hail Guards, Smoke Detectors
        const priorityCats = ['Roof Curbs & Adapters', 'Economizers & Dampers', 'Hail & Coil Guards', 'Smoke Detectors & Sensors'];
        priorityCats.forEach(cat => {{
          const match = s.accessories.find(a => a.c === cat && !state.quote.some(q => q.m === a.m));
          if (match && recItems.length < 5) {{
            recItems.push({{ acc: match, seriesTitle: s.title, seriesKey: sKey }});
          }}
        }});
      }});

      if (recItems.length === 0) {{
        container.style.display = 'none';
        return;
      }}

      container.style.display = 'block';
      recItems.forEach(r => {{
        const itemDiv = document.createElement('div');
        itemDiv.className = 'quote-rec-item';
        itemDiv.innerHTML = `
          <div style="flex: 1; min-width: 0;">
            <div class="quote-rec-item-title">${{r.acc.m}} <span style="font-size: 0.68rem; color: #0284c7; font-weight: normal;">(${{r.acc.c}})</span></div>
            <div class="quote-rec-item-desc">${{r.acc.d}}</div>
          </div>
          <div style="display: flex; align-items: center; gap: 0.5rem;">
            <div class="quote-rec-item-price">${{formatMoney(r.acc.price)}}</div>
            <button class="btn-add" onclick="addAccessoryToQuote('${{escapeQuotes(r.acc.m)}}', '${{escapeQuotes(r.acc.d)}}', ${{r.acc.price}})">+ Add</button>
          </div>
        `;
        list.appendChild(itemDiv);
      }});
    }}

    function changeItemQty(idx, delta) {{
      state.quote[idx].qty += delta;
      if (state.quote[idx].qty <= 0) {{
        state.quote.splice(idx, 1);
      }}
      saveQuote();
    }}

    function removeQuoteItem(idx) {{
      state.quote.splice(idx, 1);
      saveQuote();
    }}

    function clearQuote() {{
      if (confirm('Clear all items from active quote?')) {{
        state.quote = [];
        saveQuote();
      }}
    }}

    // FULL QUOTE TAB PROPOSAL VIEW
    function renderQuoteTab() {{
      const tbody = document.getElementById('quote-table-body');
      tbody.innerHTML = '';

      if (state.quote.length === 0) {{
        tbody.innerHTML = `<tr><td colspan="7" style="text-align: center; color: var(--text-light); padding: 2.5rem;">No items in quote. Click "+ Add to Quote" from the equipment catalog.</td></tr>`;
        document.getElementById('doc-subtotal').textContent = '$0.00';
        document.getElementById('doc-freight').textContent = '$0.00';
        document.getElementById('doc-grand-total').textContent = '$0.00';
        document.getElementById('quote-smart-recs').style.display = 'none';
        return;
      }}

      let subtotal = 0;
      const quotedSeriesKeys = new Set();

      state.quote.forEach((item, idx) => {{
        const itemTotal = item.price * item.qty;
        subtotal += itemTotal;
        if (item.sk) quotedSeriesKeys.add(item.sk);

        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td>${{idx + 1}}</td>
          <td style="font-family: 'JetBrains Mono', monospace; font-weight: 700;">${{item.m}}</td>
          <td>${{item.d}}</td>
          <td style="text-align: right; font-family: 'JetBrains Mono', monospace;">$${{item.price.toLocaleString('en-US', {{minimumFractionDigits: 2, maximumFractionDigits: 2}})}}</td>
          <td style="text-align: center; font-weight: 600;">${{item.qty}}</td>
          <td style="text-align: right; font-family: 'JetBrains Mono', monospace; font-weight: 700; color: #047857;">$${{itemTotal.toLocaleString('en-US', {{minimumFractionDigits: 2, maximumFractionDigits: 2}})}}</td>
          <td class="no-print" style="text-align: center;"><button class="btn-del-item" onclick="removeQuoteItem(${{idx}})">&times;</button></td>
        `;
        tbody.appendChild(tr);
      }});

      let freight = 0;
      if (subtotal > 0 && subtotal < 1000) freight = 150;
      else if (subtotal >= 1000 && subtotal < 5000) freight = 275;
      else if (subtotal >= 5000 && subtotal < 12000) freight = 550;
      else freight = 0;

      const grand = subtotal + freight;
      document.getElementById('doc-subtotal').textContent = '$' + subtotal.toLocaleString('en-US', {{minimumFractionDigits: 2, maximumFractionDigits: 2}});
      document.getElementById('doc-freight').textContent = freight === 0 ? 'FREE (Orders over $12,000)' : '$' + freight.toFixed(2);
      document.getElementById('doc-grand-total').textContent = '$' + grand.toLocaleString('en-US', {{minimumFractionDigits: 2, maximumFractionDigits: 2}});

      // Render Smart Recommendations on proposal page
      renderSmartRecommendations(quotedSeriesKeys, 'quote-smart-recs', 'quote-smart-recs-list');
    }}

    function exportQuoteCSV() {{
      if (state.quote.length === 0) {{
        alert('Active quote is empty.');
        return;
      }}
      let csv = "Item#,Model,Description,Unit Price,Quantity,Extended Price\\r\\n";
      state.quote.forEach((item, idx) => {{
        const safeDesc = `"${{(item.d || '').replace(/"/g, '""')}}"`;
        csv += `${{idx + 1}},"${{item.m}}",${{safeDesc}},"${{item.price.toFixed(2)}}",${{item.qty}},"${{(item.price * item.qty).toFixed(2)}}"\\r\\n`;
      }});
      downloadFile(csv, 'CFM_Equipment_Quote.csv', 'text/csv');
    }}

    // PRICING MATRIX VIEW CONTROLLER
    function switchMatrix(matName) {{
      state.activeMatrix = matName;
      document.getElementById('btn-mat-stock').classList.toggle('active', matName === 'Matrix on stock');
      document.getElementById('btn-mat-fo').classList.toggle('active', matName === 'Matrix on factory order models');
      renderMatrix();
    }}

    function highlightMatrixTier(tier) {{
      renderMatrix(tier);
    }}

    function renderMatrix(highlightTier = '') {{
      const mat = CFM_DB.matrices[state.activeMatrix];
      const table = document.getElementById('matrix-table');
      if (!mat) return;

      let html = '<thead><tr>';
      html += '<th style="width: 180px; position: sticky; left: 0; z-index: 20;">Equipment Category</th>';
      html += '<th style="width: 220px; position: sticky; left: 180px; z-index: 20;">Product Family / Description</th>';
      mat.tiers.forEach(t => {{
        const isHl = t === highlightTier ? 'background: #0284c7; color: #ffffff;' : '';
        html += `<th colspan="2" style="text-align: center; border-left: 1px solid rgba(255,255,255,0.2); ${{isHl}}">Tier ${{t}}<div style="font-size: 0.68rem; font-weight: normal; opacity: 0.85;">Buy / Sell</div></th>`;
      }});
      html += '</tr></thead><tbody>';

      mat.rows.forEach(r => {{
        html += '<tr>';
        html += `<td style="font-weight: 700; position: sticky; left: 0; background: #ffffff; z-index: 10;">${{r.category}}</td>`;
        html += `<td style="font-size: 0.8rem; position: sticky; left: 180px; background: #ffffff; z-index: 10;">${{r.description}}</td>`;
        r.tiers.forEach((tierMults, tIdx) => {{
          const tierName = mat.tiers[tIdx];
          const isHl = tierName === highlightTier ? 'background-color: #f0f9ff; font-weight: 700;' : '';
          html += `<td style="text-align: right; border-left: 1px solid var(--border-subtle); font-family: 'JetBrains Mono', monospace; font-size: 0.78rem; color: #64748b; ${{isHl}}">${{tierMults.buy}}</td>`;
          html += `<td style="text-align: right; font-family: 'JetBrains Mono', monospace; font-size: 0.82rem; font-weight: 600; color: #047857; ${{isHl}}">${{tierMults.sell}}</td>`;
        }});
        html += '</tr>';
      }});

      html += '</tbody>';
      table.innerHTML = html;
    }}

    // MARGIN CALCULATOR
    function populateCalculator() {{
      const select = document.getElementById('calc-category');
      select.innerHTML = '';
      const mat = CFM_DB.matrices['Matrix on stock'] || CFM_DB.matrices['Matrix on factory order models'];
      if (!mat) return;
      mat.rows.forEach((r, idx) => {{
        const opt = document.createElement('option');
        opt.value = idx;
        opt.textContent = `${{r.category}} - ${{r.description}}`;
        select.appendChild(opt);
      }});
      recalcPricing();
    }}

    function recalcPricing() {{
      const catIdx = parseInt(document.getElementById('calc-category').value) || 0;
      const tierVal = document.getElementById('calc-tier').value;
      const mlp = parseFloat(document.getElementById('calc-mlp').value) || 0;

      const mat = CFM_DB.matrices[state.activeMatrix];
      if (!mat || !mat.rows[catIdx]) return;

      const row = mat.rows[catIdx];
      const tierIdx = mat.tiers.indexOf(tierVal);
      if (tierIdx === -1) return;

      const mults = row.tiers[tierIdx];
      const buyX = parseFloat(mults.buy) || 0;
      const sellX = parseFloat(mults.sell) || 0;

      document.getElementById('calc-res-buy').textContent = buyX > 0 ? buyX.toFixed(4) : '—';
      if (mlp > 0 && sellX > 0) {{
        const targetSell = mlp * sellX;
        const buyCost = mlp * buyX;
        const marginDollars = targetSell - buyCost;
        const marginPct = targetSell > 0 ? (marginDollars / targetSell) * 100 : 0;

        document.getElementById('calc-res-sell').textContent = '$' + targetSell.toLocaleString('en-US', {{ minimumFractionDigits: 2, maximumFractionDigits: 2 }});
        document.getElementById('calc-res-margin').textContent = `${{marginPct.toFixed(1)}}% ($${{marginDollars.toLocaleString('en-US', {{ minimumFractionDigits: 2, maximumFractionDigits: 2 }})}})`;
      }} else {{
        document.getElementById('calc-res-sell').textContent = '—';
        document.getElementById('calc-res-margin').textContent = '—';
      }}
    }}

    // REFERENCE MANUALS VIEW CONTROLLER
    function populateReferenceSelect() {{
      const select = document.getElementById('ref-select');
      select.innerHTML = '';
      Object.keys(CFM_DB.references).forEach(name => {{
        const opt = document.createElement('option');
        opt.value = name;
        opt.textContent = name;
        select.appendChild(opt);
      }});
    }}

    function loadReferenceDoc(name) {{
      const rows = CFM_DB.references[name];
      const title = document.getElementById('ref-doc-title');
      const content = document.getElementById('ref-doc-content');
      title.textContent = name;
      content.innerHTML = '';

      if (!rows || rows.length === 0) {{
        content.innerHTML = `<div style="color: var(--text-light); padding: 2rem;">No data found in reference sheet.</div>`;
        return;
      }}

      // Check if data is tabular or free text notes
      const hasManyCols = rows.some(r => Object.keys(r).length >= 4);
      if (hasManyCols) {{
        const allCols = Array.from(new Set(rows.flatMap(r => Object.keys(r)))).sort();
        let tableHtml = '<div class="table-container"><table class="data-table"><thead><tr>';
        allCols.forEach(c => {{
          tableHtml += `<th>Column ${{c}}</th>`;
        }});
        tableHtml += '</tr></thead><tbody>';
        rows.forEach(r => {{
          tableHtml += '<tr>';
          allCols.forEach(c => {{
            tableHtml += `<td>${{r[c] || ''}}</td>`;
          }});
          tableHtml += '</tr>';
        }});
        tableHtml += '</tbody></table></div>';
        content.innerHTML = tableHtml;
      }} else {{
        let textHtml = '<div class="proposal-terms-box" style="padding: 1.5rem; border-radius: var(--radius-sm); line-height: 1.7; font-size: 0.9rem;">';
        rows.forEach(r => {{
          const rowText = Object.values(r).join(' | ');
          if (rowText.startsWith('•') || rowText.startsWith('-')) {{
            textHtml += `<div style="padding-left: 1rem; margin-bottom: 4px;">${{rowText}}</div>`;
          }} else if (rowText.length > 50 && rowText.includes(':')) {{
            textHtml += `<div style="margin-bottom: 8px;"><b>${{rowText.split(':')[0]}}:</b> ${{rowText.substring(rowText.indexOf(':') + 1)}}</div>`;
          }} else {{
            textHtml += `<div style="margin-bottom: 6px;">${{rowText}}</div>`;
          }}
        }});
        textHtml += '</div>';
        content.innerHTML = textHtml;
      }}
    }}

    function printRefDoc() {{
      window.print();
    }}

    // UTILS
    function downloadFile(content, fileName, mimeType) {{
      const a = document.createElement('a');
      const blob = new Blob([content], {{ type: mimeType }});
      a.href = URL.createObjectURL(blob);
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
    }}

    function handleFileImport(input) {{
      if (!input.files || !input.files[0]) return;
      alert(`Selected ${{input.files[0].name}}. To process new quarterly workbooks, run "python generate_app.py" to compile latest equipment inventories and matrices into this tool.`);
    }}
  </script>
</body>
</html>
"""

    print("Writing cfm_quote_tool.html...")
    with open(output_html, "w", encoding="utf-8") as f:
        f.write(html_template)
    print(f"Successfully generated clean cfm_quote_tool.html ({os.path.getsize(output_html):,} bytes)!")

if __name__ == "__main__":
    generate_html()
