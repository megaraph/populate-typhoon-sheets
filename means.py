from openpyxl import load_workbook
from pathlib import Path
from indices import VEG_INDICES, veg_index_dict

WB_NAME = "Typhoon Marilyn.xlsx"
WB_PATH = Path.cwd().joinpath(WB_NAME)

WB = load_workbook(filename=WB_NAME, data_only=True)


def indices_results(location, index_dict):
    results = []
    for veg_index in VEG_INDICES:
        values = [row.value for row in location[index_dict[veg_index]]]
        _, *nums = values

        # Perform computation on nums
        result = sum(nums) / len(nums)
        results.append(result)

    return results


def clear_sheet(sheet_name, headers):
    sheet = WB[sheet_name]
    WB.remove(sheet)

    WB.create_sheet(sheet_name)
    new_sheet = WB[sheet_name]
    new_sheet.append(headers)

    WB.save(WB_PATH)

    return new_sheet


before_sheets = [sheet for sheet in WB.sheetnames if sheet.startswith("B - ")]
after_sheets = [sheet for sheet in WB.sheetnames if sheet.startswith("A - ")]


sample_sheet = WB[before_sheets[0]]
sample_sheet_titles = list(sample_sheet.rows)[0]  # gets first row in the sample sheet

# Get titles that are only named "Location" or any of the vegetation indices
titles = [
    title.value
    for title in sample_sheet_titles
    if title.value in ["Location", *VEG_INDICES]
]

corr_before_sheet = clear_sheet(sheet_name="Correlation Before", headers=titles)

# TODO: REMOVE LIMIT
for sheet in before_sheets[:2]:
    loc = WB[sheet]
    loc_name = loc["A2"].value

    index_dict = veg_index_dict()
    results = indices_results(loc, index_dict)
    row = [loc_name, *results]

    corr_before_sheet.append(row)

WB.save(WB_PATH)
