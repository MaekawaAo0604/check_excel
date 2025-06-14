import zipfile
import xml.etree.ElementTree as ET

file_path = '/home/ao0604/check_excel/宮下綾介2023夏期現状報告書 -前川.xlsx'

def parse_cell_ref(cell_ref):
    col = ''
    row = ''
    for char in cell_ref:
        if char.isalpha():
            col += char
        else:
            row += char
    
    col_num = 0
    for char in col:
        col_num = col_num * 26 + (ord(char) - ord('A') + 1)
    
    return (int(row) if row else 0, col_num)

def get_cell_value(cell, shared_strings):
    v_elem = cell.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
    if v_elem is None:
        return ''
    
    cell_type = cell.get('t', 'n')
    
    if cell_type == 's':
        idx = int(v_elem.text)
        if idx < len(shared_strings):
            return shared_strings[idx]
    elif cell_type == 'n' or cell_type is None:
        return v_elem.text
    else:
        return v_elem.text
    
    return ''

with zipfile.ZipFile(file_path, 'r') as zip_file:
    # Load shared strings
    shared_strings = []
    if 'xl/sharedStrings.xml' in zip_file.namelist():
        with zip_file.open('xl/sharedStrings.xml') as f:
            tree = ET.parse(f)
            root = tree.getroot()
            for si in root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si'):
                t = si.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
                if t is not None:
                    shared_strings.append(t.text)
    
    # Create a complete map of all cells
    with zip_file.open('xl/worksheets/sheet1.xml') as f:
        tree = ET.parse(f)
        root = tree.getroot()
        
        cell_map = {}
        
        rows = root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row')
        for row in rows:
            row_num = row.get('r', '')
            cells = row.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c')
            for cell in cells:
                cell_ref = cell.get('r', '')
                value = get_cell_value(cell, shared_strings)
                if value and cell_ref:
                    cell_map[cell_ref] = value
        
        print('=== COMPLETE EXCEL FILE STRUCTURE ===')
        print(f'Sheet name: 2022夏期')
        print(f'Total rows: {len(rows)}')
        print(f'Total cells with data: {len(cell_map)}')
        print('\n=== ALL CELL VALUES ===')
        
        sorted_cells = sorted(cell_map.items(), key=lambda x: parse_cell_ref(x[0]))
        
        current_row = None
        for cell_ref, value in sorted_cells:
            row_num, col_num = parse_cell_ref(cell_ref)
            if current_row is None or current_row != row_num:
                if current_row is not None:
                    print()
                current_row = row_num
                print(f'Row {row_num}:')
            print(f'  {cell_ref}: {value}')