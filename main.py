import logging
import sys
import warnings
from io import BytesIO
from itertools import islice
from typing import Dict

import requests
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from pandas import DataFrame, set_option

warnings.filterwarnings("ignore", category=DeprecationWarning)

set_option("display.max_rows", None)
set_option("display.max_columns", None)
set_option("display.width", None)


def init_logging() -> logging.Logger:
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.DEBUG)
    handler = logging.StreamHandler(sys.stdout)
    formatter = logging.Formatter(
            "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    )
    handler.setFormatter(formatter)
    handler.setLevel(logging.INFO)
    logger.addHandler(handler)

    return logger


log = init_logging()

URL = (
        "https://assets.publishing.service.gov.uk/"
        "government/uploads/system/uploads/attachment_data/file/912364/"
        "2020.08.27_Average_road_fuel_sales_and_stock_levels_at_sampled_filling_stations.xlsx"
)

http = requests.Session()


def get_excel_file(url=URL) -> BytesIO:
    """
        Download excel file from url location.
    Args:
        url: Url location of Excel file to be downloaded

    Returns:
        An Excel file in bytes buffer

    """

    global http

    try:
        log.info(f"Get -> {url}")
        response = http.get(url)
        response.raise_for_status()
    except requests.RequestException as e:
        logging.error(str(e) + "\n Error while trying to download excel file")
        http.close()
        sys.exit(-1)
    log.info("File ready")

    return BytesIO(response.content)


def extract_data_sheet(excel_sheet: Worksheet) -> DataFrame:
    """
        Handling function for extracting `Data` sheet from excel.
    Args:
        excel_sheet: An openpyxl worksheet object.

    Returns:
        A panda's data frame with clean sheet data.

    """
    data = excel_sheet.values

    # magic offset from Excel file to extract columns
    cols = list(islice(data, 6, 7))[0]
    data = list(data)

    return DataFrame(data, columns=cols)


def extract_typical_levels_sheet(excel_sheet: Worksheet) -> DataFrame:
    """
        Handling function for extracting `Typical levels` sheet from excel.
    Args:
        excel_sheet: An openpyxl worksheet object.

    Returns:
        A panda's data frame with clean sheet data.

    """
    data = excel_sheet.values

    # magic offset from Excel file to extract columns
    cols = list(islice(data, 8, 9))[0]
    data = list(data)

    return DataFrame(data, columns=cols).dropna()


def extract_stock_data_sheet(excel_sheet: Worksheet) -> DataFrame:
    """
        Handling function for extracting `Stock level` sheet from excel.
    Args:
        excel_sheet: An openpyxl worksheet object.

    Returns:
        A panda's data frame with clean sheet data.

    """
    data = excel_sheet.values

    # magic offset from Excel file to extract columns
    cols = list(islice(data, 6, 7))[0]
    data = list(data)

    return DataFrame(data, columns=cols)


def extract_main_table_sheet(excel_sheet: Worksheet) -> DataFrame:
    """
        Handling function for extracting `Main table` sheet from excel.
    Args:
        excel_sheet: An openpyxl worksheet object.

    Returns:
        A pandas data frame with clean sheet data.

    """
    data = excel_sheet.values

    # magic offset from Excel file to extract columns
    cols = list(islice(data, 7, 8))[0]
    data = list(data)

    df = DataFrame(data, columns=cols)
    return df[df.columns[:-2]]


def extract_data_from_excel(file_data: BytesIO) -> Dict[str, DataFrame]:
    """
        Extract user specified sheets from input excel file.

    Args:
        file_data: excel file in bytes to extract sheets from.

    Returns:
        A python dictionary with keys the sheet name used in excel file
        and as values parsed cleaned up panda's data frames.

    """
    sheets_to_be_extracted = {
            "Main table": extract_main_table_sheet,
            "Typical levels": extract_typical_levels_sheet,
            "Data": extract_data_sheet,
            "Stock data": extract_stock_data_sheet,
    }

    # set `data_only` to evaluate cell functions to values
    wb = load_workbook(file_data, data_only=True)

    extracted_data = {
            "Main table": None,
            "Typical levels": None,
            "Data": None,
            "Stock data": None,
    }
    for sheet, handler_func in sheets_to_be_extracted.items():
        log.info("Export values for `{0}` sheet".format(sheet))
        ws = wb.get_sheet_by_name(sheet)
        extract = handler_func(ws)

        if not extract.empty:
            extracted_data.update({sheet: extract})
            continue

        log.warning(f"Sheet `{sheet}` extracted empty data frame")

    return extracted_data


def write_exported_data_to_file(
        data_map: Dict[str, DataFrame], out_dir="exported_data"
):
    import os

    if not os.path.exists(out_dir):
        log.info("Creating output dir")
        os.mkdir(out_dir)

    log.info("Write data to csv files into `{0}` dir ".format(out_dir))
    for file, df in data_map.items():
        df.to_csv(os.path.join(out_dir, file + ".csv"), index=False)


def main():
    """
    Main entry point for downloader
    """
    global http
    file = get_excel_file()
    all_data = extract_data_from_excel(file)
    write_exported_data_to_file(all_data)
    http.close()
    log.info("Done!")


# Press the green button in the gutter to run the script.
if __name__ == "__main__":
    main()
