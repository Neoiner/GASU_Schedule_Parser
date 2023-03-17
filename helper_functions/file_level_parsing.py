from openpyxl import load_workbook
from consts.faculties_abbreviations import faculty_names
from .sheet_level_parsing import parse_worksheet



async def parse_file(folderpath, filename, dryrun=True):
    filepath = folderpath + "/" + filename
    workbook = load_workbook(filepath)

    academic_degree = get_academic_degree(filename)
    faculty = get_faculty_name(filename)
    year = 1

    for worksheet in workbook.worksheets:
        await parse_worksheet(worksheet=worksheet, academic_degree=academic_degree, faculty=faculty,
                              year=year, dryrun=dryrun)
        year += 1





def get_academic_degree(filename: str):
    academic_degree = ''
    filename = filename.replace(".xlsx", "")
    filename = filename.replace(".xls", "")

    BACHELOR_SUFFIX = '_b'
    MASTER_SUFFIX = '_m'

    if filename.endswith(BACHELOR_SUFFIX):
        academic_degree = "БАКАЛАВРИАТ"
    elif filename.endswith(MASTER_SUFFIX):
        academic_degree = "МАГИСТРАТУРА"

    return academic_degree



def get_faculty_name(filename: str):
    faculty_abbreviation = filename.split("_")[1]
    faculty_name = faculty_names.get(faculty_abbreviation)
    return faculty_name



















