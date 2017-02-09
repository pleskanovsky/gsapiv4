from pprint import pprint
import httplib2
import apiclient.discovery
import oauth2client.clientsecrets as clientsecrets
from oauth2client.client import OAuth2WebServerFlow
import googleapiclient.errors
from oauth2client.service_account import ServiceAccountCredentials
import re


def coords_to_range(row, col):
    if (not isinstance(row, int)) or (not isinstance(col, int)):
        raise TypeError
    if row < 1 or col < 1:
        raise ValueError
    res = ""
    if col > 26:
        res += chr(col // 26 + 64)
    res += chr(col % 26 + 64)
    res += str(row)
    return res


class Auth:

    def __init__(self, path_to_secret, scopes):
        self.path_to_secret = path_to_secret
        self.scopes = scopes
        self.flow = None

    def init_flow(self):
        client_type, client_info = clientsecrets.loadfile(self.path_to_secret)
        self.client_id = client_info["client_id"]
        self.client_secret = client_info["client_secret"]
        self.auth_uri = client_info["auth_uri"]
        self.token_uri = client_info["token_uri"]
        self.redirect_uri = client_info["redirect_uris"][0]
        self.flow = OAuth2WebServerFlow(self.client_id, self.client_secret, self.scopes, redirect_uri=self.redirect_uri)

    def get_auth_url(self):
        if self.flow is None:
            self.init_flow()
        return self.flow.step1_get_authorize_url()

    def auth(self, auth_code):
        if self.flow is None:
            self.init_flow()
        self.credentials = self.flow.step2_exchange(auth_code)



class Client:
    def __init__(self, credentials):
        self.service = apiclient.discovery.build('sheets', 'v4', credentials=credentials)
        self.spreadsheet_collection = self.service.spreadsheets()


class SourceRange:
    sheet_id = None
    start_row_index = None
    end_row_index = None
    start_column_index = None
    end_column_index = None

    def __init__(self, start_row_index, start_column_index, end_row_index=None, end_column_index=None):
        if end_row_index is None:
            end_row_index = start_row_index
        if end_column_index is None:
            end_column_index = start_column_index
        self.start_row_index = start_row_index - 1
        self.start_column_index = start_column_index - 1
        self.end_column_index = end_column_index
        self.end_row_index = end_row_index

    @property
    def json(self):
        return {
            "sheetId": self.sheet_id,
            "startRowIndex": self.start_row_index,
            "endRowIndex": self.end_row_index,
            "startColumnIndex": self.start_column_index,
            "endColumnIndex": self.end_column_index
        }


class OverlayPosition:
    sheet_id = None
    row_index = None
    column_index = None
    offset_x_pixels = None
    offset_y_pixels = None

    def __init__(self):
        self.row_index = 0
        self.column_index = 0
        self.offset_y_pixels = 0
        self.offset_x_pixels = 0
        pass

    @property
    def json(self):
        return {
                    "overlayPosition": {
                        "anchorCell": {
                            "sheetId": self.sheet_id,
                            "rowIndex": self.row_index,
                            "columnIndex": self.column_index
                        },
                        "offsetXPixels": self.offset_x_pixels,
                        "offsetYPixels": self.offset_y_pixels
                    }
                }


class Spreadsheet:
    def __init__(self, client: Client, spreadsheet_id):
        if client is not None:
            self.spreadsheet_collection = client.spreadsheet_collection
            self.spreadsheet_id = spreadsheet_id
            self.requests = None
            self.value_ranges = None
            self.current_sheet_id = None
            self.current_sheet_title = None
            self.sheets = None
            self.refresh()
            re_letter_pattern = "([A-Z])+"
            re_digit_pattern = "\d+"
            self.r_exp_letter = re.compile(re_letter_pattern)
            self.r_exp_digit = re.compile(re_digit_pattern)
        else:
            self.requests = []
            self.value_ranges = []

    # spreadsheets.batchUpdate and spreadsheets.values.batchUpdate
    def refresh(self):
        spreadsheet = self.spreadsheet_collection.get(spreadsheetId=self.spreadsheet_id).execute()
        self.requests = []
        self.value_ranges = []
        self.spreadsheet_id = spreadsheet['spreadsheetId']
        self.current_sheet_id = spreadsheet['sheets'][0]['properties']['sheetId']
        self.current_sheet_title = spreadsheet['sheets'][0]['properties']['title']
        self.sheets = {}
        for sheet in spreadsheet['sheets']:
            self.sheets[sheet['properties']['title']] = sheet['properties']['sheetId']

    def execute_queue(self, value_input_option="USER_ENTERED"):
        upd1res = {'replies': []}
        upd2res = {'responses': []}
        try:
            if len(self.requests) > 0:
                upd1res = self.spreadsheet_collection.batchUpdate(spreadsheetId=self.spreadsheet_id,
                                                                  body={"requests": self.requests}).execute()
            if len(self.value_ranges) > 0:
                upd2res = self.spreadsheet_collection.values().batchUpdate(spreadsheetId=self.spreadsheet_id,
                                                                           body={"valueInputOption": value_input_option,
                                                                                 "data": self.value_ranges}).execute()
        finally:
            self.requests = []
            self.value_ranges = []
        return upd1res['replies'], upd2res['responses']

    def prepare_add_sheet(self, sheet_title, rows=100, cols=100):
        self.requests.append({"addSheet": {"properties": {"title": sheet_title,
                                                          'gridProperties': {'rowCount': rows, 'columnCount': cols}}}})

    def prepare_delete_sheet(self, sheet_id):
        self.requests.append(
            {
                "deleteSheet": {
                    "sheetId": sheet_id
                }
            })

    # Adds new sheet to current spreadsheet, sets as current sheet and returns it's id
    def add_sheet(self, sheet_title, rows=100, cols=100):
        self.prepare_add_sheet(sheet_title, rows, cols)
        added_sheet = self.execute_queue()[0][0]['addSheet']['properties']
        self.current_sheet_id = added_sheet['sheetId']
        self.current_sheet_title = added_sheet['title']
        return self.current_sheet_id

    def set_sheet_by_title(self, title):
        if title in self.sheets:
            self.current_sheet_id = self.sheets[title]
            self.current_sheet_title = title
            return True
        else:
            return False

    def cell_to_indexes(self, cell):
        column_index_str = self.r_exp_letter.match(cell).group()
        i = len(column_index_str)
        column_index = 0
        for letter in column_index_str:
            column_index += (26 ** (i - 1)) * (ord(letter) - ord('A') + 1)
            i -= 1
        column_index -= 1
        if self.r_exp_digit.search(cell):
            row_index_str = self.r_exp_digit.search(cell).group()
            row_index = int(row_index_str)
        else:
            row_index = None
        return column_index, row_index


    # Converts string range to GridRange of current sheet; examples:
    # "A3:B4"->{sheetId: id of current sheet, startRowIndex: 2, endRowIndex: 4, startColumnIndex: 0, endColumnIndex: 2}
    # "A5:B" ->{sheetId: id of current sheet, startRowIndex: 4, startColumnIndex: 0, endColumnIndex: 2}
    def to_grid_range(self, cells_range):
        if isinstance(cells_range, str):
            start_cell, end_cell = cells_range.split(":")[0:2]
            cells_range = {}
            cells_range["startColumnIndex"], cells_range["startRowIndex"] = self.cell_to_indexes(start_cell)
            cells_range["endColumnIndex"], cells_range["endRowIndex"] = self.cell_to_indexes(end_cell)
            if cells_range["startRowIndex"] is None:
                del cells_range["startRowIndex"]
            if cells_range["endRowIndex"] is None:
                del cells_range["endRowIndex"]

        cells_range["sheetId"] = self.current_sheet_id
        return cells_range

    def prepare_set_dimension_pixel_size(self, dimension, start_index, end_index, pixel_size):
        self.requests.append({"updateDimensionProperties": {
            "range": {"sheetId": self.current_sheet_id,
                      "dimension": dimension,
                      "startIndex": start_index,
                      "endIndex": end_index},
            "properties": {"pixelSize": pixel_size},
            "fields": "pixelSize"}})

    def prepare_set_columns_width(self, start_col, end_col, width):
        self.prepare_set_dimension_pixel_size("COLUMNS", start_col, end_col + 1, width)

    def prepare_set_column_width(self, col, width):
        self.prepare_set_columns_width(col, col, width)

    def prepare_set_rows_height(self, start_row, end_row, height):
        self.prepare_set_dimension_pixel_size("ROWS", start_row, end_row + 1, height)

    def prepare_set_row_height(self, row, height):
        self.prepare_set_rows_height(row, row, height)

    def prepare_set_values(self, cells_range, values, major_dimension="ROWS"):
        """
        Prepares a range of cells to be updated with given values
        :param cells_range: range of cells to be updated
        :param values: list of lists of values
        :param major_dimension: role of inner lists in values (ex. with param value=ROWS method will write inner lists
        to sheet as rows)
        :return:
        """
        self.value_ranges.append({"range": self.current_sheet_title + "!" + cells_range,
                                  "majorDimension": major_dimension, "values": values})

    def prepare_set_value(self, cell_index, value):
        cells_range = cell_index + ":" + cell_index
        values = [[value]]
        self.prepare_set_values(cells_range, values)

    def prepare_merge_cells(self, cells_range, merge_type ="MERGE_ALL"):
        self.requests.append({"mergeCells": {"range": self.to_grid_range(cells_range), "mergeType": merge_type}})

    # formatJSON should be dict with userEnteredFormat to be applied to each cell
    def prepare_set_cells_format(self, cells_range, format_json, fields="userEnteredFormat"):
        self.requests.append({"repeatCell": {"range": self.to_grid_range(cells_range),
                                             "cell": {"userEnteredFormat": format_json}, "fields": fields}})

    # formatsJSON should be list of lists of dicts with userEnteredFormat for each cell in each row
    def prepare_set_cells_formats(self, cells_range, formats_json, fields="userEnteredFormat"):
        rows_value = [{"values": [{"userEnteredFormat": cellFormat} for cellFormat in rowFormats]} for rowFormats in formats_json]
        self.requests.append({"updateCells": {"range": self.to_grid_range(cells_range),
                                              "rows": rows_value,
                                              "fields": fields}})

    def prepare_set_frozen(self, rows, columns):
        self.requests.append(
            {
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": self.current_sheet_id,
                        "gridProperties": {
                            "frozenRowCount": rows,
                        }
                    },
                    "fields": "gridProperties.frozenRowCount"
                }
            }
        )
        self.requests.append(
            {
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": self.current_sheet_id,
                        "gridProperties": {
                            "frozenColumnCount": columns
                        }
                    },
                    "fields": "gridProperties.frozenColumnCount"
                }
            }
        )

    def prepare_add_pie_chart(self, title, domain: SourceRange, series: SourceRange, position: OverlayPosition):
        # series_list = [{
        #     "series": {
        #         "sourceRange": {
        #             "sources": [
        #                 i.json
        #             ]
        #         }
        #     }
        # } for i in series_ranges]

        self.requests.append(
            {
                "addChart": {
                    "chart": {
                        "spec": {
                            "title": title,
                            "pieChart": {
                                "legendPosition": "RIGHT_LEGEND",
                                "threeDimensional": False,
                                "domain": {
                                    "sourceRange": {
                                        "sources": [
                                            domain.json
                                        ]
                                    }
                                },
                                "series": {
                                    "sourceRange": {
                                        "sources": [
                                            series.json
                                        ]
                                    }
                                }
                            }
                        },
                        "position": position.json
                    }
                }
            }
        )




