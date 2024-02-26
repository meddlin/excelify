def get_row(row: dict[str, str], column_names: list[str]):
    data = []
    
    for col in column_names:
        data.append(row[col])
    
    return data

def get_row_filtered(row: dict[str, str], column_names: list[str], filter_columns: list[str]):
    """"""
    
    data = []
    for col in column_names:
        if col in filter_columns:
            data.append(row[col])
    
    return data

def format_filter_cols(filter_cs_list: str) -> list[str]:
    """Separate filter columns.
        Remove leading and trailiing whitespace from filter columns.
        Return as a list.
    """
    cols = filter_cs_list.split(',')
    cleaned = []
    for col in cols:
        cleaned.append(col.strip().replace("'", ""))
    
    return cleaned
