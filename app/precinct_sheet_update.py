from map_region_manager import *
import pickle


class PrecinctManager:
    def __init__(self, wks, column_title_mappings=dict([("Name", "A"), ("Precinct", "B"), ("People", "J"), ("Doors", "O")])):
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

def run_precincts_VBM():
    # load region objects
    with open("VBM_precincts.pkl", "rb") as file:
        regions = pickle.load(file)


    gc = pygsheets.authorize(service_account_file="client_secret.json")

    # Open spreadsheet and then worksheet
    sh = gc.open_by_url('https://docs.google.com/spreadsheets/d/156ta7rVPOJMsLZpBd3t5TIS6bbV-7fJcJ8ceWGgIXK0/edit#gid=864082663')
    wks = sh.worksheet_by_title("Data By Precinct")

    mrman = PrecinctManager(wks, column_title_mappings=dict([("Name", "A"), ("Precinct", "B"), ("People", "J"), ("Doors", "O")]))
    mrman.APPEND = True
    mrman.update_rows(regions)


def run_precincts_normal():
    # load region objects
    with open("normal_precincts.pkl", "rb") as file:
        regions = pickle.load(file)


    gc = pygsheets.authorize(service_account_file="client_secret.json")

    # Open spreadsheet and then worksheet
    sh = gc.open_by_url('https://docs.google.com/spreadsheets/d/156ta7rVPOJMsLZpBd3t5TIS6bbV-7fJcJ8ceWGgIXK0/edit#gid=864082663')
    wks = sh.worksheet_by_title("Data By Precinct")

    mrman = PrecinctManager(wks, column_title_mappings=dict([("Name", "A"), ("Precinct", "B"), ("People", "G"), ("Doors", "M")]))
    mrman.APPEND = True
    mrman.update_rows(regions)