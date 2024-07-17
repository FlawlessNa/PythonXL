from openpyxl.utils import get_column_letter, column_index_from_string


def get_range_ref_for_shape(
    num_rows: int,
    num_cols: int,
    first_row: int = 1,
    first_col: str = 'A',
) -> str:
    """
    Get the range reference for a shape in Excel.
    :param num_rows: Number of rows in the shape.
    :param num_cols: Number of columns in the shape.
    :param first_row: The first row of the shape.
    :param first_col: The first column of the shape.
    :return: The range reference for the shape.
    """
    first_col_number = column_index_from_string(first_col)
    last_col = get_column_letter(first_col_number + num_cols)
    return f'{first_col}{first_row}:{last_col}{first_row+num_rows}'
