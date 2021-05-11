import pygsheets
import numpy as np
import pickle
from search_terms import *
import time

class MapRegionManager:
    def __init__(self, wks, column_title_mappings=dict([("ID", "A"), ("Name", "B"), ("People", "C"), ("Doors", "D"), ("Mailboxes", "E"), ("Home Phone Ct", "F"), ("Pref Phone Ct", "G")])):
        self.wks = wks
        self.column_title_mappings = column_title_mappings
        self.SLEEP_TIME = 0.5

    def prep_column_headings(self):
        for title in self.column_title_mappings.keys():
            self.wks.update_value('{}1'.format(self.column_title_mappings[title]), title)

    def update_row(self, region, index):
        region_dict = region.flatten()
        cells = []
        time.sleep(self.SLEEP_TIME)
        for title in region_dict.keys():
            cell = pygsheets.Cell("{letter}{index}".format(letter=self.column_title_mappings[title], index=index), region_dict[title], worksheet=self.wks)
            cells.append(cell)
        self.wks.update_values(cell_list=cells)


    def update_rows(self, regions):
        id_cells = self.wks.get_col(1, returnas="cell", include_tailing_empty=False)[1:]
        end_index = len(id_cells) + 1
        print(id_cells)
        cell_mappings = dict([(cell.value, cell.label) for cell in id_cells])
        for region in regions:
            id = region.ID
            try:
                index = self.conv_index(cell_mappings[id])
                self.update_row(region, index)
            # row doesn't exist yet
            except KeyError:
                # this might be slow
                end_index+=1
                index = end_index
                wks.append_table(values = list(region.flatten().values()), start="A1", end="G{}".format(end_index))
            print("index:{}".format(index))
            # self.update_row(region, index)

    def conv_index(self, address):
        index = int("".join([a for a in str(address) if not a.isalpha()]))
        return index






if __name__ == "__main__":
    # load region objects
    with open("regions.pkl", "rb") as file:
        regions = pickle.load(file)


    gc = pygsheets.authorize(service_account_file="client_secret.json")

    # Open spreadsheet and then worksheet
    sh = gc.open_by_url('https://docs.google.com/spreadsheets/d/10eaL1xkBsXd2I1gxprXJK-L_vP7cj0Zx9a7etL4ZBkY/edit#gid=0')
    wks = sh.sheet1

    mrman = MapRegionManager(wks)
    mrman.prep_column_headings()
    mrman.update_rows(regions)


