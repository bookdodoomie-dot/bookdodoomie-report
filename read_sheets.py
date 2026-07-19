# read_sheets.py
import zipfile
import xml.etree.ElementTree as ET
import os
import re

def parse_xlsx(file_path):
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return

    try:
        with zipfile.ZipFile(file_path, 'r') as z:
            # 1. Parse sharedStrings.xml
            shared_strings = []
            if 'xl/sharedStrings.xml' in z.namelist():
                ss_data = z.read('xl/sharedStrings.xml')
                root = ET.fromstring(ss_data)
                ns = {'ns': root.tag.split('}')[0].strip('{')} if '}' in root.tag else {}
                xpath = './/ns:t' if ns else './/t'
                for t in root.findall(xpath, ns):
                    shared_strings.append(t.text or '')
            
            # 2. Parse workbook.xml to get sheet names and their relationship IDs
            sheets = []
            if 'xl/workbook.xml' in z.namelist():
                wb_data = z.read('xl/workbook.xml')
                root = ET.fromstring(wb_data)
                ns = {'ns': root.tag.split('}')[0].strip('{')} if '}' in root.tag else {}
                xpath = './/ns:sheet' if ns else './/sheet'
                for s in root.findall(xpath, ns):
                    name = s.attrib.get('name')
                    sheet_id = s.attrib.get('sheetId')
                    # Find relationship ID r:id
                    r_id = None
                    for key, val in s.attrib.items():
                        if 'id' in key.lower():
                            r_id = val
                            break
                    sheets.append({'name': name, 'id': r_id, 'sheetId': sheet_id})

            # 3. Parse xl/_rels/workbook.xml.rels to map rId to target sheet xml file
            rels = {}
            if 'xl/_rels/workbook.xml.rels' in z.namelist():
                rels_data = z.read('xl/_rels/workbook.xml.rels')
                root = ET.fromstring(rels_data)
                ns = {'ns': root.tag.split('}')[0].strip('{')} if '}' in root.tag else {}
                xpath = './/ns:Relationship' if ns else './/Relationship'
                for r in root.findall(xpath, ns):
                    r_id = r.attrib.get('Id')
                    target = r.attrib.get('Target')
                    rels[r_id] = target

            print(f"Total Sheets Found: {len(sheets)}")
            for sheet in sheets:
                r_id = sheet['id']
                target = rels.get(r_id, f"worksheets/sheet{sheet['sheetId']}.xml")
                
                # Sheet XML path in zip usually: xl/worksheets/sheet1.xml
                sheet_xml_path = f"xl/{target}"
                if sheet_xml_path not in z.namelist():
                    # Fallback
                    sheet_xml_path = f"xl/worksheets/sheet{sheet['sheetId']}.xml"

                if sheet_xml_path in z.namelist():
                    print("-" * 50)
                    print(f"Sheet Name: {sheet['name']} ({sheet_xml_path})")
                    sheet_data = z.read(sheet_xml_path)
                    root = ET.fromstring(sheet_data)
                    ns = {'ns': root.tag.split('}')[0].strip('{')} if '}' in root.tag else {}
                    
                    # We want to inspect D18, F18, H18, J18 and also show general cells in row 18 and surrounding rows like 4, 5 for student name, etc.
                    # Build cell dictionary: cell_ref -> value
                    cell_dict = {}
                    
                    c_xpath = './/ns:c' if ns else './/c'
                    v_xpath = 'ns:v' if ns else 'v'
                    
                    for c in root.findall(c_xpath, ns):
                        r_attr = c.attrib.get('r')
                        t_attr = c.attrib.get('t')
                        v_elem = c.find(v_xpath, ns)
                        val = ""
                        if v_elem is not None:
                            val_str = v_elem.text or ""
                            if t_attr == 's':
                                try:
                                    idx = int(val_str)
                                    val = shared_strings[idx]
                                except (ValueError, IndexError):
                                    val = val_str
                            else:
                                val = val_str
                        cell_dict[r_attr] = val
                    
                    # Print scores
                    print(f"  어휘력 (D18): {cell_dict.get('D18', 'Empty')}")
                    print(f"  분석력 (F18): {cell_dict.get('F18', 'Empty')}")
                    print(f"  추론력 (H18): {cell_dict.get('H18', 'Empty')}")
                    print(f"  비판력 (J18): {cell_dict.get('J18', 'Empty')}")
                    
                    # Print row 18 elements
                    print("  All Row 18 cells:")
                    r18_pattern = re.compile(r'^([A-Z]+)18$')
                    for ref, val in sorted(cell_dict.items()):
                        m = r18_pattern.match(ref)
                        if m:
                            print(f"    {ref}: {val}")
                            
                    # Print typical metadata cells in row 4, 5, etc.
                    print("  Metadata check (Row 3-5):")
                    for row in ['3', '4', '5', '6']:
                        row_cells = []
                        for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G']:
                            ref = f"{col}{row}"
                            if ref in cell_dict and cell_dict[ref]:
                                row_cells.append(f"{ref}:{cell_dict[ref]}")
                        if row_cells:
                            print(f"    Row {row}: {', '.join(row_cells)}")
                else:
                    print(f"Sheet XML path not found in zip: {sheet_xml_path}")
    except Exception as e:
        print(f"Error parsing: {e}")

if __name__ == '__main__':
    excel_path = "꿈동산 엑셀 데이터.xlsx"
    parse_xlsx(excel_path)
