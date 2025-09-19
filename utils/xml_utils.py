def extract_attendance_from_xml_final(file_path: str) -> pd.DataFrame:
    tree = ET.parse(file_path); root = tree.getroot(); data = []; ns = {'ss': 'urn:schemas-microsoft-com:office:spreadsheet'}
    header_row_index = -1; rows = root.findall(".//ss:Row", ns)

    for i, row in enumerate(rows):
        if 'Nome' in [cell.text for cell in row.findall(".//ss:Data", ns)]:
            header_row_index = i; break
    if header_row_index == -1: return pd.DataFrame()
    header_cells = [cell.text for cell in rows[header_row_index].findall(".//ss:Data", ns)]
    try:
        nome_idx = header_cells.index('Nome'); horario_idx = header_cells.index('Horário')
    except ValueError: return pd.DataFrame()
    for row in rows[header_row_index + 1:]:
        cells = [cell.text for cell in row.findall(".//ss:Data", ns)]
        if len(cells) > nome_idx and len(cells) > horario_idx:
            nome = cells[nome_idx]; horario = cells[horario_idx]
            if nome and horario: data.append({'Name': nome.strip(), 'Datetime': horario})
    return pd.DataFrame(data)
