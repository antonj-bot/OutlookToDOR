"""Application and Office constants used across the importer."""

OL_EMBEDDED_ITEM = 5

DESTINATION_WORKBOOK_NAMES = {
    "Ecrash": "Ecrash_Jan 2026_DOR_V2_Anton",
    "CGF": "CGF_Jan 2026_DOR_V2_Anton",
}
DESTINATION_SHEETS = {
    "Ecrash": "Ecrash Raw",
    "CGF": "CGF Raw",
}
WORK_TYPE_MATCH_VALUES = {
    "Ecrash": (
        "ecrash",
        "e crash",
    ),
    "CGF": (
        "cgf",
    ),
}
ALLOWED_EXCEL_EXTENSIONS = {
    ".xlsx",
    ".xlsm",
    ".xlsb",
    ".xls",
}

ALLOWED_OUTLOOK_ITEM_EXTENSIONS = {
    ".msg",
    ".eml",
}

OUTLOOK_MESSAGE_LIMIT = 500


REQUIRED_ATTACHMENT_KEYWORD = (
    "individual"
)

OL_FOLDER_INBOX = 6
XL_UP = -4162
XL_TO_LEFT = -4159
XL_DOWN = -4121
XL_NONE = -4142
IMPORTED_DATA_HIGHLIGHT_COLOR = 65535
DOR_PASS_1_COLOR = 16247773
DOR_PASS_2_COLOR = 14348258
DOR_SOURCE_COLORS = {
    "completed report pass 1": DOR_PASS_1_COLOR,
    "completed report pass 2": DOR_PASS_2_COLOR,
    "total elapsed time": 8696052,
    "average report/hr": 14067636,
    "completed reports": 14539150,
    "late %": 13421812,
    "ontime%": 11065270,
    "auto vs manual manual": 10086143,
    "manual reports average report/hr": 15323865,
    "auto vs manual auto extraction": 14934224,
    "auto extraction average report/hr": 13493756,
}
DOR_COLUMN_SOURCE_COLORS = {
    3: DOR_SOURCE_COLORS["completed report pass 1"],
    4: DOR_SOURCE_COLORS["completed report pass 2"],
    5: DOR_SOURCE_COLORS["total elapsed time"],
    6: DOR_SOURCE_COLORS["average report/hr"],
    7: DOR_SOURCE_COLORS["completed reports"],
    8: DOR_SOURCE_COLORS["late %"],
    9: DOR_SOURCE_COLORS["ontime%"],
    10: DOR_SOURCE_COLORS["auto vs manual manual"],
    12: DOR_SOURCE_COLORS["manual reports average report/hr"],
    13: DOR_SOURCE_COLORS["auto vs manual auto extraction"],
    15: DOR_SOURCE_COLORS["auto extraction average report/hr"],
}
DOR_COMPARISON_SECTION_TITLE = "ECRASH - MANUAL VS AUTO (ALL STATES)"
DOR_COMPARISON_EXCLUDING_TX_SECTION_TITLE = (
    "ECRASH - MANUAL VS AUTO (EXCLUDING TX)"
)
DOR_HEADER_ROWS = (2, 3, 4)

