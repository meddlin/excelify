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

def parse_optional_bool_flag(option) -> bool:
    """
        Force a nullable field to be a boolean.
        Helpful with CLI flags.

        Default to true, where None (or not set) equals True.
    """
    if option is None:
        return True
    if option is False:
        return False
    
    return True
