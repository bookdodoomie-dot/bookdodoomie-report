# parse_excel.py
import zipfile
import xml.etree.ElementTree as ET
import os

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
                # The namespace usually is http://schemas.openxmlformats.org/spreadsheetml/2006/main
                ns = {'ns': root.tag.split('}')[0].strip('{')} if '}' in root.tag else {}
                
                # Find all <t> elements
                xpath = './/ns:t' if ns else './/t'
                for t in root.findall(xpath, ns):
                    shared_strings.append(t.text or '')
            
            # 2. Parse sheet1.xml
            if 'xl/worksheets/sheet1.xml' in z.namelist():
                sheet_data = z.read('xl/worksheets/sheet1.xml')
                root = ET.fromstring(sheet_data)
                ns = {'ns': root.tag.split('}')[0].strip('{')} if '}' in root.tag else {}
                
                rows_xpath = './/ns:row' if ns else './/row'
                c_xpath = 'ns:c' if ns else 'c'
                v_xpath = 'ns:v' if ns else 'v'
                
                rows = root.findall(rows_xpath, ns)
                for r in rows:
                    row_data = []
                    cells = r.findall(c_xpath, ns)
                    # Cells might not be in order, but for simplicity let's read them
                    for c in cells:
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
                        row_data.append(val)
                    print("\t".join(row_data))
            else:
                print("sheet1.xml not found in zip.")
    except Exception as e:
        print(f"Error parsing xlsx: {e}")

if __name__ == '__main__':
    excel_path = "꿈동산 엑셀 데이터.xlsx"
    parse_xlsx(excel_path)
