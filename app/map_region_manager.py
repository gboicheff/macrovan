import pygsheets
from search_terms import *
import time

class MapRegionManager:
    def __init__(self, wks, column_title_mappings=dict([("Name", "A"), ("Doors", "M")])):
        self.wks = wks
        self.column_title_mappings = column_title_mappings
        self.SLEEP_TIME = 1
        self.APPEND = False

    def prep_column_headings(self):
        for title in self.column_title_mappings.keys():
            self.wks.update_value('{}1'.format(self.column_title_mappings[title]), title)

    def update_row(self, region, index):
        region_dict = region.flatten()
        region_dict["Name"] = region.get_precinct() # fix this hack
        cells = []
        time.sleep(self.SLEEP_TIME)
        for title in region_dict.keys():
            cell = pygsheets.Cell("{letter}{index}".format(letter=self.column_title_mappings[title], index=index), region_dict[title], worksheet=self.wks)
            cells.append(cell)
        self.wks.update_values(cell_list=cells)

    def update_rows(self, regions):
        id_cells = self.wks.get_col(1, returnas="cell", include_tailing_empty=False)[1:]
        end_index = len(id_cells) + 1
        cell_mappings = dict([(cell.value, cell.label) for cell in id_cells])
        for region in regions[1:]:
            try:
                num = region.get_precinct()
                index = self.conv_index(cell_mappings[str(num)])
                self.update_row(region, index)
            # row doesn't exist yet
            except:
                if self.APPEND:
                    end_index+=1
                    self.update_row(region, end_index)
                    region.display()

    def conv_index(self, address):
        index = int("".join([a for a in str(address) if not a.isalpha()]))
        return index






