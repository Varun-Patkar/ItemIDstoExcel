import os
import glob
import re
import xlsxwriter


def parse_item_line(line):
    # Match pattern: name, then [itemid], then properties string.
    m = re.match(r"^(.*?)\s*\[(.*?)\]\s+(.*)$", line)
    if m:
        name = m.group(1).strip()
        item_id = m.group(2).strip()
        props_str = m.group(3).strip()
        # Parse properties: they are key=value pairs separated by whitespace.
        props = {}
        for token in props_str.split():
            if "=" in token:
                key, value = token.split("=", 1)
                props[key.strip()] = value.strip()
        return name, item_id, props
    return None


def process_file(filepath):
    items = []
    current_section = None
    with open(filepath, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("//"):
                continue
            if line.startswith("="):
                # Remove all "=" characters and see if there's text.
                cleaned = re.sub("=", "", line).strip()
                if cleaned:
                    current_section = cleaned
                continue
            if "[" not in line or "]" not in line:
                continue
            parsed = parse_item_line(line)
            if parsed:
                name, item_id, props = parsed
                items.append((current_section, name, item_id, props))
    return items


def create_sheet_for_file(workbook, filename, items):
    # Group items by section (preserving the order of appearance)
    section_groups = {}
    section_order = []
    for sec, name, item_id, props in items:
        sec_key = sec if sec else "Misc"
        if sec_key not in section_groups:
            section_groups[sec_key] = []
            section_order.append(sec_key)
        section_groups[sec_key].append((name, item_id, props))

    # Use filename without extension as worksheet name.
    sheet_title = os.path.splitext(os.path.basename(filename))[0]
    # Ensure valid and unique worksheet name.
    worksheet = workbook.add_worksheet(sheet_title)

    # Global header (merged across first 10 columns)
    global_header = f"Data from {os.path.basename(filename)}"
    worksheet.merge_range(
        0,
        0,
        0,
        9,
        global_header,
        workbook.add_format({"bold": True, "align": "center", "valign": "vcenter"}),
    )

    # Define section header format (orange background, white text)
    section_format = workbook.add_format(
        {
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "bg_color": "orange",
            "font_color": "white",
        }
    )

    # Start writing tables from row 2 (0-indexed rows; row0 global header)
    current_row = 2
    for section in section_order:
        group = section_groups[section]
        # Determine columns (each section can have different property keys).
        prop_keys = set()
        for name, item_id, props in group:
            prop_keys.update(props.keys())
        prop_keys = sorted(prop_keys)
        headers = ["Name", "ItemId"] + prop_keys
        num_cols = len(headers)
        last_col = num_cols - 1

        # Write section header (merged across the table columns).
        worksheet.merge_range(
            current_row, 0, current_row, last_col, f"Section: {section}", section_format
        )
        current_row += 1

        # Build table data rows from group.
        table_data = []
        for name, item_id, props in group:
            row = [name, item_id] + [props.get(key, "") for key in prop_keys]
            table_data.append(row)

        # Prepare table column headers for XLSXWriter's add_table.
        table_columns = [{"header": h} for h in headers]

        # Determine the table range: if there are N rows, table covers header + data.
        start_row = current_row
        end_row = current_row + len(
            table_data
        )  # add_table automatically creates header row
        # Add the table using XLSXWriter's add_table option.
        worksheet.add_table(
            start_row,
            0,
            end_row,
            last_col,
            {
                "columns": table_columns,
                "data": table_data,
                "style": "Table Style Medium 9",
            },
        )

        # Advance current_row past the table and add a blank row.
        current_row = end_row + 2


def main():
    # Locate all item.*.txt files in the current directory.
    file_pattern = os.path.join(os.getcwd(), "item*.txt")
    file_list = glob.glob(file_pattern)
    if not file_list:
        print("No item text files found.")
        return

    output_filename = "KCD2Items.xlsx"
    workbook = xlsxwriter.Workbook(output_filename)

    for filepath in file_list:
        items = process_file(filepath)
        if items:
            create_sheet_for_file(workbook, filepath, items)
        else:
            print(f"No valid item lines found in {filepath}.")

    workbook.close()
    print(f"Excel file saved as {output_filename}")


if __name__ == "__main__":
    main()
