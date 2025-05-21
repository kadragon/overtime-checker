"""
Provides utility functions for working with Excel files, focusing on styling and formatting.
"""
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill
from openpyxl.utils import get_column_letter

def apply_default_report_styles(ws, num_header_rows=1, num_footer_rows=1):
    """
    Applies predefined styles (fonts, borders, alignment, fill) and adjusts column widths
    to a given openpyxl worksheet.

    Args:
        ws (openpyxl.worksheet.worksheet.Worksheet): The worksheet to style.
        num_header_rows (int): Number of header rows to apply bold font and fill.
        num_footer_rows (int): Number of footer rows to apply bold font and fill.
    """
    font_format = Font(size=11, name='맑은 고딕')
    font_format_bold = Font(size=11, name='맑은 고딕', bold=True)
    border_format = Side(border_style="thin")
    align_format_center = Alignment(horizontal="center", vertical="center")
    align_format_left = Alignment(horizontal="left", vertical="center") # Default for non-header/footer
    fill_style = PatternFill(start_color="00C0C0C0", end_color="00C0C0C0", patternType="solid")

    max_row = ws.max_row
    
    for col_idx, column_cells in enumerate(ws.columns):
        for row_idx, cell in enumerate(column_cells):
            cell.font = font_format
            cell.border = Border(top=border_format, bottom=border_format,
                                 left=border_format, right=border_format)
            cell.alignment = align_format_left # Default alignment

            # Header rows
            if row_idx < num_header_rows:
                cell.font = font_format_bold
                cell.fill = fill_style
                cell.alignment = align_format_center
            # Footer rows (e.g., totals)
            elif row_idx >= (max_row - num_footer_rows):
                cell.font = font_format_bold
                cell.fill = fill_style
                # Center align B and C, or first two columns in general for totals
                if col_idx < 2 : # Assuming first two columns (like 'B', 'C') in totals are centered
                    cell.alignment = align_format_center
            # Data rows
            else:
                # Center align specific columns if needed, e.g. count columns
                # This part might need more specific logic based on column content or index
                # For now, let's assume the second column (index 1, typically 'C') is centered if it's not header/footer
                if col_idx == 1: # Example: Center align the '인원' (count) column
                     cell.alignment = align_format_center
                
                # Apply number formatting for numeric columns (e.g., 'D', 'E')
                # This requires knowing which columns are numeric.
                # For a generic function, this might be passed in or determined dynamically.
                # Let's assume columns from the 3rd one (index 2) onwards could be numeric.
                if col_idx >= 2: # Columns D, E, etc.
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '#,##0'
    
    for column_cells in ws.columns:
        # For column width, consider only data cells, not empty ones if possible
        # and use a reasonable multiplier. Max length of cell.value might be too simple.
        # openpyxl's automatic width calculation is often not perfect.
        # A common approach is to find max length and add some padding or use a fixed width.
        try:
            # Filter out None values before calculating max length
            # Also, ensure values are converted to string
            relevant_values = [str(cell.value) for cell in column_cells if cell.value is not None]
            if not relevant_values:
                new_column_length = 10 # Default width for empty or all-None columns
            else:
                new_column_length = max(len(value) for value in relevant_values)
            
            new_column_letter = get_column_letter(column_cells[0].column)

            # Apply a multiplier and a minimum width
            adjusted_width = (new_column_length * 1.2) + 2
            if adjusted_width < 10: # Minimum width
                adjusted_width = 10
            
            ws.column_dimensions[new_column_letter].width = adjusted_width
        except Exception as e:
            # print(f"Could not set width for column {get_column_letter(column_cells[0].column)}: {e}")
            # Fallback or skip if error occurs
            pass
