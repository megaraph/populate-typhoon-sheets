from openpyxl.utils import get_column_letter

VEG_INDICES = [
    "VARI",
    "EXG",
    "EXR",
    "EXB",
    "EXGR",
    "GRVI",
    "MGRVI",
    "GLI",
    "RGBVI",
    "IKAW",
    "GBDI",
    "CIVE",
    "GRRI",
    "NGBDI",
    "VDVI",
    "VEG",
    "COM",
    "TGI",
    "gvTGI",
    "gvTeGI",
    "vODGIabc",
    "vODGIfa",
    "vODGIga",
]


def veg_index_dict():
    # Dictionary to hold the indices and their corresponding column letters
    index_column_dict = {}

    # Start from column "I", which is the 9th column in Excel (A=1, B=2, ..., H=8, I=9)
    start_column_number = 9

    # Populate the dictionary
    for i, veg_index in enumerate(VEG_INDICES):
        column_letter = get_column_letter(start_column_number + i)
        index_column_dict[veg_index] = column_letter

    return index_column_dict
