from map_region_manager import *
import pickle

class TurfManager:
    def __init__(self, wks, column_title_mappings=dict([("Name", "A"), ("Precinct", "B"), ("Turf", "C"), ("Doors", "Q"), ("People", "M")])):
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
        new_dict["Turf"] = region.get_turf()
        new_dict["Precinct"] = region.get_precinct()
        cells = []
        time.sleep(self.SLEEP_TIME)
        for title in new_dict.keys():
            cell = pygsheets.Cell("{letter}{index}".format(letter=self.column_title_mappings[title], index=index), new_dict[title], worksheet=self.wks)
            cells.append(cell)
        self.wks.update_values(cell_list=cells)

    def update_rows(self, regions):
        name_cells = self.wks.get_col(1, returnas="cell", include_tailing_empty=False)[1:]
        end_index = len(name_cells) + 1
        cell_mappings = dict([(cell.value.strip(), cell.label) for cell in name_cells])
        for region in regions:
            try:
                index = self.conv_index(cell_mappings[region.NAME.strip()])
                self.update_row(region, index)
            # row doesn't exist yet
            except:
                if self.APPEND:
                    end_index+=1
                    self.update_row(region, end_index)
                    # region.display()

    def conv_index(self, address):
        index = int("".join([a for a in str(address) if not a.isalpha()]))
        return index

def run_turfs():
    # load region objects
    with open("turfs.pkl", "rb") as file:
        regions = pickle.load(file)


    gc = pygsheets.authorize(service_account_file="client_secret.json")

    # Open spreadsheet and then worksheet
    sh = gc.open_by_url('https://docs.google.com/spreadsheets/d/156ta7rVPOJMsLZpBd3t5TIS6bbV-7fJcJ8ceWGgIXK0/edit#gid=0')
    wks = sh.worksheet_by_title("Data By Turf")

    mrman = TurfManager(wks)
    mrman.APPEND = True
    # mrman.prep_column_headings()
    mrman.update_rows(regions)

# if __name__ == "__main__":


#     # load region objects
#     with open("turfs.pkl", "rb") as file:
#         regions = pickle.load(file)


#     gc = pygsheets.authorize(service_account_file="client_secret.json")

#     # Open spreadsheet and then worksheet
#     sh = gc.open_by_url('https://docs.google.com/spreadsheets/d/156ta7rVPOJMsLZpBd3t5TIS6bbV-7fJcJ8ceWGgIXK0/edit#gid=0')
#     wks = sh.worksheet_by_title("Data By Turf")

#     mrman = TurfManager(wks)
#     mrman.APPEND = True
#     mrman.prep_column_headings()
#     mrman.update_rows(regions)