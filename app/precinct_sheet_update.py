import pickle
import pygsheets
import time

from datetime import date
import datetime


class PrecinctManager:
    def __init__(self, wks, column_title_mappings):
        self.wks = wks
        self.column_title_mappings = column_title_mappings
        self.SLEEP_TIME = 1
        self.APPEND = False

    def prep_column_headings(self):
        for title in self.column_title_mappings.keys():
            self.wks.update_value('{}1'.format(self.column_title_mappings[title]), title)

    def update_row(self, region, index):
        region_dict = region.flatten()
        new_dict = dict([(key, region_dict[key]) for key in region_dict.keys() if key in set(self.column_title_mappings.keys())])
        new_dict["Precinct"] = region.get_precinct()

        cells = []
        time.sleep(self.SLEEP_TIME)
        for title in new_dict.keys():
            cell = pygsheets.Cell("{letter}{index}".format(letter=self.column_title_mappings[title], index=index), new_dict[title], worksheet=self.wks)
            cells.append(cell)
        self.wks.update_values(cell_list=cells)

    def update_rows(self, regions):
        id_cells = self.wks.get_col(2, returnas="cell", include_tailing_empty=False)[1:]
        end_index = len(id_cells) + 1
        cell_mappings = dict([(cell.value.strip(), cell.label) for cell in id_cells])
        print(cell_mappings)
        for region in regions:
            try:
                print(region.get_precinct())
                index = self.conv_index(cell_mappings[str(region.get_precinct())])
                self.update_row(region, index)
                print(index)
            # row doesn't exist yet
            except Exception as e:
                print('error-start')
                print(e)
                print(region.NAME)
                print("error-end")
                if self.APPEND:
                    end_index+=1
                    self.update_row(region, end_index)
                    # region.display()

    def conv_index(self, address):
        index = int("".join([a for a in str(address) if not a.isalpha()]))
        return index

    def validate_date(self, date_text):
        if date_text[-1:][0] == "}":
            return True
        return False

    def update_heading(self):
        column_headings = self.wks.get_row(1, returnas="cell", include_tailing_empty=False)
        cell_mappings = dict([(cell.value.strip(), cell.label) for cell in column_headings])
        today = date.today()
        today_date = today.strftime("%m/%d/%y")
        cells = []

        set_column_titles = set(title + "1" for title in self.column_title_mappings.values())
        for val, cell in cell_mappings.items():
            if cell in set_column_titles:
                split_val = val.split(" ")
                possible_date = split_val[-1:][0]
                if self.validate_date(possible_date):
                    new_heading = " ".join(split_val[:-1]) + " {" + today_date + "}"
                    real_cell = pygsheets.Cell(cell, new_heading, worksheet=self.wks)
                    cells.append(real_cell)

        print(cells)
        self.wks.update_values(cell_list=cells)

def run_precincts_VBM():
    # load region objects
    with open("VBM_precincts.pkl", "rb") as file:
        regions = pickle.load(file)


    gc = pygsheets.authorize(service_account_file="client_secret.json")

    # Open spreadsheet and then worksheet
    sh = gc.open_by_url('https://docs.google.com/spreadsheets/d/156ta7rVPOJMsLZpBd3t5TIS6bbV-7fJcJ8ceWGgIXK0/edit#gid=864082663')
    wks = sh.worksheet_by_title("Data By Precinct")

    mrman = PrecinctManager(wks, column_title_mappings=dict([("Name", "A"), ("Precinct", "B"), ("People", "K"), ("Doors", "O")]))
    mrman.APPEND = True
    mrman.update_rows(regions)
    mrman.update_heading()


def run_precincts_normal():
    # load region objects
    with open("normal_precincts.pkl", "rb") as file:
        regions = pickle.load(file)


    gc = pygsheets.authorize(service_account_file="client_secret.json")

    # Open spreadsheet and then worksheet
    sh = gc.open_by_url('https://docs.google.com/spreadsheets/d/156ta7rVPOJMsLZpBd3t5TIS6bbV-7fJcJ8ceWGgIXK0/edit#gid=864082663')
    wks = sh.worksheet_by_title("Data By Precinct")

    mrman = PrecinctManager(wks, column_title_mappings=dict([("Name", "A"), ("Precinct", "B"), ("People", "H"), ("Doors", "N")]))
    mrman.APPEND = True
    mrman.update_rows(regions)
    mrman.update_heading()