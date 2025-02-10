# This object generates the Excel table final result based on the given arguments

import xlsxwriter

class Generate_Excel:   
    def setup_sheets(self, path_database):
        # setup workbook
        self.workbook = xlsxwriter.Workbook(path_database)

        # set up color themes
        self.format_black = self.workbook.add_format({'border': 1, 'bold': True, 'bg_color': 'black', 'font_color':'white'})
        self.format_white = self.workbook.add_format({'border': 1})
        self.format_grey = self.workbook.add_format({'border': 1, 'bg_color': '#c4c4c4'})
        self.format_green = self.workbook.add_format({'border': 1, 'bg_color': '#76de68'})
        self.format_red = self.workbook.add_format({'border': 1, 'bg_color': '#db5858'})

        # setup sheets
        # Programs
        self.sheet_Programs = self.workbook.add_worksheet("Programs")
        self.sheet_Programs.freeze_panes(1,0)
        self.sheet_Programs.autofilter("A1:E1")
        self.sheet_Programs.set_column('A:A', 30)
        self.sheet_Programs.set_column('B:B', 20)
        self.sheet_Programs.set_column('C:C', 15)
        self.sheet_Programs.set_column('D:D', 20)
        self.sheet_Programs.set_column('E:E', 15)
        self.sheet_Programs.write_string("A1", "Product Part Number", self.format_black)
        self.sheet_Programs.write_string("B1", "MPT5000L", self.format_black)
        self.sheet_Programs.write_string("C1", "Validation", self.format_black)
        self.sheet_Programs.write_string("D1", "MPT5000", self.format_black)
        self.sheet_Programs.write_string("E1", "Validation", self.format_black)

        # Test Cables
        self.sheet_TestCables = self.workbook.add_worksheet("Test Cables")
        self.sheet_TestCables.freeze_panes(1,0)
        self.sheet_TestCables.autofilter("A1:F1")
        self.sheet_TestCables.set_column('A:A', 15)
        self.sheet_TestCables.set_column('B:B', 10)
        self.sheet_TestCables.set_column('C:C', 40)
        self.sheet_TestCables.set_column('D:D', 40)
        self.sheet_TestCables.set_column('E:E', 10)
        self.sheet_TestCables.set_column('F:F', 10)
        self.sheet_TestCables.write_string("A1", "Test Cable ID", self.format_black)	
        self.sheet_TestCables.write_string("B1", "Branch", self.format_black)
        self.sheet_TestCables.write_string("C1", "Plug Part Number", self.format_black)
        self.sheet_TestCables.write_string("D1", "Flex Plug Part Number", self.format_black)
        self.sheet_TestCables.write_string("E1", "Size", self.format_black)
        self.sheet_TestCables.write_string("F1", "Type", self.format_black)

        # Empty Spaces
        self.sheet_EmptySpaces = self.workbook.add_worksheet("Empty Spaces")
        self.sheet_EmptySpaces.freeze_panes(1,0)
        self.sheet_EmptySpaces.autofilter("A1:B1")
        self.sheet_EmptySpaces.set_column('A:A', 15)
        self.sheet_EmptySpaces.set_column('B:B', 25)
        self.sheet_EmptySpaces.write_string("A1", "Test Cable ID", self.format_black)	
        self.sheet_EmptySpaces.write_string("B1", "Number of empty spaces", self.format_black)

        # Jigs
        self.sheet_Jigs = self.workbook.add_worksheet("Jigs")
        self.sheet_Jigs.freeze_panes(1,0)
        self.sheet_Jigs.autofilter("A1:B1")
        self.sheet_Jigs.set_column('A:A', 15)
        self.sheet_Jigs.set_column('B:B', 15)
        self.sheet_Jigs.write_string("A1", "Jig ID", self.format_black)
        self.sheet_Jigs.write_string("B1", "Test Cables", self.format_black)

        # Products
        self.sheet_Products = self.workbook.add_worksheet("Products")
        self.sheet_Products.freeze_panes(1,0)
        self.sheet_Products.autofilter("A1:D1")
        self.sheet_Products.set_column('A:A', 30)
        self.sheet_Products.set_column('B:B', 20)
        self.sheet_Products.set_column('C:C', 40)
        self.sheet_Products.set_column('D:D', 30)
        self.sheet_Products.write_string("A1", "Product Part Number", self.format_black)
        self.sheet_Products.write_string("B1", "Plug Name", self.format_black)
        self.sheet_Products.write_string("C1", "Plug Part Number", self.format_black)
        self.sheet_Products.write_string("D1", "Connect to Test Cable", self.format_black)

        # Test logs
        self.sheet_TestLogs = self.workbook.add_worksheet("Test Logs")
        self.sheet_TestLogs.freeze_panes(1,0)
        self.sheet_TestLogs.autofilter("A1:C1")
        self.sheet_TestLogs.set_column('A:A', 30)
        self.sheet_TestLogs.set_column('B:B', 30)
        self.sheet_TestLogs.set_column('C:C', 30)
        self.sheet_TestLogs.write_string("A1", "Product Part Number", self.format_black)
        self.sheet_TestLogs.write_string("B1", "Serial Number", self.format_black)
        self.sheet_TestLogs.write_string("C1", "Date", self.format_black)

    def write_data(self, data_programs, data_testingcables, data_jigs, data_products, data_testlogs):
        # Programs
        i = 1
        for line in data_programs[1:]:
            i += 1
            
            # A Product Part Number
            self.sheet_Programs.write_string("A"+str(i), line[0], self.format_white)
            
            # B MPT5000L
            if line[1] == "In Progress":
                self.sheet_Programs.write_string("B"+str(i), line[1], self.format_red)
            else:
                self.sheet_Programs.write_string("B"+str(i), line[1], self.format_green)
            
            # C MPT5000L Validation
            if line[2] == "N":
                self.sheet_Programs.write_string("C"+str(i), line[2], self.format_red)
            else:
                self.sheet_Programs.write_string("C"+str(i), line[2], self.format_green)

            # D MPT5000L
            if line[3] == "In Progress":
                self.sheet_Programs.write_string("D"+str(i), line[3], self.format_red)
            else:
                self.sheet_Programs.write_string("D"+str(i), line[3], self.format_green)
            
            # E MPT5000L Validation
            if line[4] == "N":
                self.sheet_Programs.write_string("E"+str(i), line[4], self.format_red)
            else:
                self.sheet_Programs.write_string("E"+str(i), line[4], self.format_green)

        # Test Cables
        i = 1
        colorkey = ""
        color = self.format_grey
        for line in data_testingcables[0][1:]:
            i += 1

            # determine color
            if colorkey != line[0]:
                colorkey = line[0]
                if color == self.format_grey:
                    color = self.format_white
                else:
                    color = self.format_grey

            # add cells
            self.sheet_TestCables.write_string("A"+str(i), line[0], color) # A Test Cable ID
            self.sheet_TestCables.write_string("B"+str(i), line[1], color) # B Branch
            self.sheet_TestCables.write_string("C"+str(i), line[2], color) # C Plug Part Number
            self.sheet_TestCables.write_string("D"+str(i), line[3], color) # D Flex Plug Part Number
            self.sheet_TestCables.write_string("E"+str(i), str(line[4]), color) # E Size
            self.sheet_TestCables.write_string("F"+str(i), line[5], color) # F Type
        
        # Empty Spaces
        i = 1
        for line in data_testingcables[1][1:]:
            i += 1

            # add cells
            self.sheet_EmptySpaces.write_string("A"+str(i), line[0], self.format_white) # A Test Cable ID
            self.sheet_EmptySpaces.write_string("B"+str(i), str(line[1]), self.format_white) # B Number of empty spaces

        # Jigs
        i = 1
        colorkey = ""
        color = self.format_grey
        for line in data_jigs[1:]:
            i += 1

            # determine color
            if colorkey != line[0]:
                colorkey = line[0]
                if color == self.format_grey:
                    color = self.format_white
                else:
                    color = self.format_grey

            # add cells
            self.sheet_Jigs.write_string("A"+str(i), line[0], color) # A Jig Id
            self.sheet_Jigs.write_string("B"+str(i), line[1], color) # B Test Cables

        # Products
        i = 1
        colorkey = ""
        color = self.format_grey
        for line in data_products[1:]:
            i += 1

            # determine color
            if colorkey != line[0]:
                colorkey = line[0]
                if color == self.format_grey:
                    color = self.format_white
                else:
                    color = self.format_grey

            # add cells
            self.sheet_Products.write_string("A"+str(i), line[0], color) # A Product Part Number
            self.sheet_Products.write_string("B"+str(i), line[1], color) # B Plug Name
            self.sheet_Products.write_string("C"+str(i), line[2], color) # C Plug Part Number 
            self.sheet_Products.write_string("D"+str(i), line[3], color) # D Connect to Test Cable

        # Test logs
        i = 1
        colorkey = ""
        color = self.format_grey
        for line in data_testlogs[1:]:
            i += 1

            # determine color
            if colorkey != line[0]:
                colorkey = line[0]
                if color == self.format_grey:
                    color = self.format_white
                else:
                    color = self.format_grey

            # add cells
            self.sheet_TestLogs.write_string("A"+str(i), line[0], color) # A Product Part Number
            self.sheet_TestLogs.write_string("B"+str(i), line[1], color) # B Serial Number
            self.sheet_TestLogs.write_string("C"+str(i), line[2], color) # C Date

        self.workbook.close()
