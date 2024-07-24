# Download SmartConsole.py from: https://github.com/VladFeldfix/Smart-Console/blob/main/SmartConsole.py
from SmartConsole import *
import xlsxwriter

class main:
    # constructor
    def __init__(self):
        # load smart console
        self.sc = SmartConsole("Database Updater", "1.6")

        # get settings
        self.path_database = self.sc.get_setting("Database Location")
        self.path_programs = self.sc.get_setting("Programs Location")
        self.path_braids = self.sc.get_setting("Braids Location")
        self.path_tms_inv = self.sc.get_setting("TMS Inventory")
        self.path_tms_db = self.sc.get_setting("TMS Database")
        self.path_test_logs = self.sc.get_setting("Test logs")

        # test all paths
        self.sc.test_path(self.path_programs)
        self.sc.test_path(self.path_braids)
        self.sc.test_path(self.path_tms_inv)
        self.sc.test_path(self.path_tms_db)
        self.sc.test_path(self.path_test_logs)

        # display main menu
        self.run()
    
    def run(self):
        self.sc.print("Processing...")
        # SPREADSHEET DESCRIPTION
        # Programs - this spreadsheet shows wheather the program is ready or not, and if not, then what is it missing
        # Braids - information about exsiting braids maps
        # TMS Inventory - The number of boxes left in the TMS inventory
        # Missing connectors - What meeting connectors are missing for what products

        # GATHER DATA
        self.sc.print("\nGATHERING DATA:")
        Programs = []
        Braids = []
        TMSInventory = []
        MissingConnectors = []

        # GATHER DATA on Programs
        self.sc.print("Gathering data on programs...")
        PartNumber = ""
        Status = ""
        test = self.path_programs.replace("\\", "/")
        test = test.split("/")
        test = len(test)
        for root, dirs, files in os.walk(self.path_programs):
            root = root.replace("\\", "/")
            PartNumber = root.split("/")
            Status = "In Progress"
            if len(PartNumber) == test+1:
                PartNumber = PartNumber[-1]
                if not "_" in PartNumber:
                    csv_file = False
                    txt_file = False
                    html_file = False
                    mpt_file = False
                    Missing_plugs = False
                    for file in files:
                        if file == PartNumber+".csv":
                            csv_file = True
                        if file == PartNumber+".txt":
                            txt_file = True
                        if file == PartNumber+".html":
                            html_file = True
                        if file == PartNumber+".mpt_product":
                            mpt_file = True
                        if file == "testcables_to_product.csv":
                            if os.path.isfile(root+"/testcables_to_product.csv"):
                                data = self.sc.load_csv(root+"/testcables_to_product.csv")
                                for line in data[1:]:
                                    if len(line) == 3:
                                        # GATHER DATA on Missing connectors
                                        MissingConnectors.append((PartNumber, line[1], line[2], line[0]))
                                        if line[0] == "":
                                            Missing_plugs = True
                    if csv_file and txt_file and html_file and not Missing_plugs:
                        Status = "Complete"
                    if mpt_file and html_file and not Missing_plugs:
                        Status = "Complete"
                    Programs.append((PartNumber, Status))

        # GATHER DATA on Braids
        self.sc.print("Gathering data on braids...")
        TestCable = ""
        Braid = ""
        CustomerPN = ""
        FlexPN = ""
        Size = 0
        Type = ""
        EmptySpaces = {}
        for root, dirs, files in os.walk(self.path_braids):
            for file in files:
                info = {}
                if ".csv" in file:
                    data = self.sc.load_csv(root+"/"+file)
                    TestCable = file.replace(".csv", "")
                    LegalTestCable = True
                    try:
                        tmp = TestCable[3:]
                        tmp = int(tmp)
                    except:
                        LegalTestCable = False
                    PreviousPin = ""
                    if LegalTestCable:
                        PreviousBraid = ""
                        for line in data[1:]:
                            # 0           , 1   , 2  , 3          , 4               , 5                  , 6
                            # GLOBAL_POINT, PLUG, PIN, PLUG_NUMBER, PART_NUMBER     , RAFAEL_PART_NUMBER , PIN_TYPE 
                            # 1           , 1   , A  , 1          , D38999/20WH53SN , 112065102          , Socket
                            GlobalPoint = line[0]
                            CurrentBraid = line[1]
                            Pin = line[2]
                            Braid = line[3]
                            try:
                                if Braid != "":
                                    Braid = int(Braid)
                                if CurrentBraid != "":
                                    CurrentBraid = int(CurrentBraid)
                            except:
                                self.sc.fatal_error("Check map for test cable #"+str(TestCable))
                            CustomerPN = line[4]
                            FlexPN = line[5]
                            Type = line[6]
                            if Braid != "":
                                info[Braid] = [CustomerPN, FlexPN, Type, 0]
                            if CurrentBraid != "":
                                if CurrentBraid in info:
                                    if len(info[CurrentBraid]) == 4:
                                        if Pin.upper() != "BODY":
                                            if Pin != PreviousPin:
                                                info[CurrentBraid][3] += 1
                                                PreviousPin = Pin
                                            else:
                                                if CurrentBraid != PreviousBraid:
                                                    info[CurrentBraid][3] += 1
                                                    PreviousPin = Pin
                            PreviousBraid = CurrentBraid
                            if not TestCable in EmptySpaces:
                                EmptySpaces[TestCable] = 0
                            if GlobalPoint != "" and CurrentBraid == "":
                                EmptySpaces[TestCable] += 1
                        for Braid, data in info.items():
                            # "TEST CABLE", "BRAID", "CUSTOMER PART NUMBER", "FLEX PART NUMBER", "SIZE", "TYPE"
                            CustomerPN = data[0]
                            FlexPN = data[1]
                            Type = data[2]
                            Size = data[3]
                            Braids.append((TestCable, Braid, CustomerPN, FlexPN, Size, Type))
        Braids = sorted(Braids)
        EmptySpacesSortedList = []
        for key, val in EmptySpaces.items():
            EmptySpacesSortedList.append((key, val))
        EmptySpacesSortedList = sorted(EmptySpacesSortedList)

        # READ JIGS DATA
        self.sc.print("Gathering data on jigs...")
        path = self.path_braids+"/JIGS.csv"
        self.sc.test_path(path)
        Jigs = self.sc.load_csv(path)
        
        # GATHER DATA on TMS Inventory
        self.sc.print("Gathering data on labeling inventory...")
        inv_part_numbers = self.sc.load_csv(self.path_tms_db)[1:] # Flex Part Number	Customer Part Number	Description
        inv_boxid = self.sc.load_csv(self.path_tms_inv)[1:]
        qty = {}
        for x in inv_boxid:
            pn = x[1]
            if not pn in qty:
                qty[pn] = 1
            else:
                qty[pn] += 1
        TMSInventory = []
        for x in inv_part_numbers:
            pn = x[0]
            cpn = x[1]
            desc = x[2]
            if pn in qty:
                pnqty = qty[pn]
            else:
                pnqty = 0
            TMSInventory.append((pn, cpn, desc, pnqty))

        # GATHER DATA ON TEST LOGS
        self.sc.print("Gathering data on test logs...")
        TestLogs = []
        for root, directories, files in os.walk(self.path_test_logs):
            for file in files:
                file = file.split("_")
                part_number = ""
                if len(file) == 3:
                    TestLogs.append((file[0],file[1],file[2].replace(".html", "").replace(".pdf", "")))
        
        
        # CREATE EXCEL TABLES
        self.sc.print("\nGENERATING DATA:")
        self.sc.print("Generating Excel table...")

        # setup workbook
        workbook = xlsxwriter.Workbook(self.path_database)

        # set up color themes
        format_black = workbook.add_format({'border': 1, 'bold': True, 'bg_color': 'black', 'font_color':'white'})
        format_white = workbook.add_format({'border': 1})
        format_grey = workbook.add_format({'border': 1, 'bg_color': '#c4c4c4'})
        format_green = workbook.add_format({'border': 1, 'bg_color': '#76de68'})
        format_red = workbook.add_format({'border': 1, 'bg_color': '#db5858'})
        format_yellow = workbook.add_format({'border': 1, 'bg_color': '#f5f258'})
        
        # setup sheets
        # Programs
        sheet_Programs = workbook.add_worksheet("Programs")
        sheet_Programs.freeze_panes(1,0)
        sheet_Programs.autofilter("A1:B1")
        sheet_Programs.set_column('A:A', 30)
        sheet_Programs.set_column('B:B', 20)
        sheet_Programs.write_string("A1", "PART NUMBER", format_black)
        sheet_Programs.write_string("B1", "MPT PROGRAM", format_black)
        
        # Braids
        sheet_Braids = workbook.add_worksheet("Braids")
        sheet_Braids.freeze_panes(1,0)
        sheet_Braids.autofilter("A1:F1")
        sheet_Braids.set_column('A:A', 15)
        sheet_Braids.set_column('B:B', 10)
        sheet_Braids.set_column('C:C', 40)
        sheet_Braids.set_column('D:D', 40)
        sheet_Braids.set_column('E:E', 10)
        sheet_Braids.set_column('F:F', 10)
        sheet_Braids.write_string("A1", "TEST CABLE", format_black)	
        sheet_Braids.write_string("B1", "BRAID", format_black)
        sheet_Braids.write_string("C1", "CUSTOMER PART NUMBER", format_black)
        sheet_Braids.write_string("D1", "FLEX PART NUMBER", format_black)
        sheet_Braids.write_string("E1", "SIZE", format_black)
        sheet_Braids.write_string("F1", "TYPE", format_black)

        # Empty Spaces
        sheet_BraidsEmptySpaces = workbook.add_worksheet("Empty Spaces")
        sheet_BraidsEmptySpaces.freeze_panes(1,0)
        sheet_BraidsEmptySpaces.autofilter("A1:B1")
        sheet_BraidsEmptySpaces.set_column('A:A', 15)
        sheet_BraidsEmptySpaces.set_column('B:B', 15)
        sheet_BraidsEmptySpaces.write_string("A1", "TEST CABLE", format_black)	
        sheet_BraidsEmptySpaces.write_string("B1", "EMPTY SPACES", format_black)

        # Jigs
        sheet_Jigs = workbook.add_worksheet("Jigs")
        sheet_Jigs.freeze_panes(1,0)
        sheet_Jigs.autofilter("A1:B1")
        sheet_Jigs.set_column('A:A', 15)
        sheet_Jigs.set_column('B:B', 15)
        sheet_Jigs.write_string("A1", "JIG", format_black)
        sheet_Jigs.write_string("B1", "TEST CABLE", format_black)

        # TMSInventory
        sheet_TMSInventory = workbook.add_worksheet("Labeling-Inventory")
        sheet_TMSInventory.freeze_panes(1,0)
        sheet_TMSInventory.autofilter("A1:D1")
        sheet_TMSInventory.set_column('A:A', 20)
        sheet_TMSInventory.set_column('B:B', 20)
        sheet_TMSInventory.set_column('C:C', 20)
        sheet_TMSInventory.set_column('D:D', 10)
        sheet_TMSInventory.write_string("A1", "PART NUMBER", format_black)
        sheet_TMSInventory.write_string("B1", "CUSTOMER PN", format_black)
        sheet_TMSInventory.write_string("C1", "DESCRIPTION", format_black)
        sheet_TMSInventory.write_string("D1", "QTY", format_black)

        # MissingConnectors
        sheet_MissingConnectors = workbook.add_worksheet("Connectors")
        sheet_MissingConnectors.freeze_panes(1,0)
        sheet_MissingConnectors.autofilter("A1:D1")
        sheet_MissingConnectors.set_column('A:A', 30)
        sheet_MissingConnectors.set_column('B:B', 20)
        sheet_MissingConnectors.set_column('C:C', 40)
        sheet_MissingConnectors.set_column('D:D', 30)
        sheet_MissingConnectors.write_string("A1", "ASSEMBLY PART NUMBER", format_black)
        sheet_MissingConnectors.write_string("B1", "PLUG NAME", format_black)
        sheet_MissingConnectors.write_string("C1", "ASSEMBLY PLUG PART NUMBER", format_black)
        sheet_MissingConnectors.write_string("D1", "CONNECT TO TEST CABLE", format_black)

        # Test logs
        sheet_TestLogs = workbook.add_worksheet("Test Logs")
        sheet_TestLogs.freeze_panes(1,0)
        sheet_TestLogs.autofilter("A1:C1")
        sheet_TestLogs.set_column('A:A', 30)
        sheet_TestLogs.set_column('B:B', 30)
        sheet_TestLogs.set_column('C:C', 30)
        sheet_TestLogs.write_string("A1", "PART NUMBER", format_black)
        sheet_TestLogs.write_string("B1", "SERIAL NUMBER", format_black)
        sheet_TestLogs.write_string("C1", "DATE", format_black)

        # insert data
        # Programs
        i = 1
        for line in Programs:
            i += 1
            sheet_Programs.write_string("A"+str(i), line[0], format_white)
            if line[1] == "In Progress":
                sheet_Programs.write_string("B"+str(i), line[1], format_red)
            else:
                sheet_Programs.write_string("B"+str(i), line[1], format_green)

        # Braids
        i = 1
        CurrentBraid = ""
        Color = format_white
        for line in Braids:
            i += 1
            if line[0] != CurrentBraid:
                CurrentBraid = line[0]
                if Color == format_white:
                    Color = format_grey
                else:
                    Color = format_white
            sheet_Braids.write_string("A"+str(i), str(line[0]), Color)
            sheet_Braids.write_string("B"+str(i), str(line[1]), Color)
            sheet_Braids.write_string("C"+str(i), line[2], Color)
            sheet_Braids.write_string("D"+str(i), line[3], Color)
            sheet_Braids.write_string("E"+str(i), str(line[4]), Color)
            sheet_Braids.write_string("F"+str(i), line[5], Color)
        
        # Empty Spaces
        i = 1
        for line in EmptySpacesSortedList:
            i += 1
            sheet_BraidsEmptySpaces.write_string("A"+str(i), str(line[0]), format_white)
            sheet_BraidsEmptySpaces.write_string("B"+str(i), str(line[1]), format_white)

        # Jigs
        i = 1
        CurrentJig = ""
        for line in Jigs:
            i += 1
            if line[0] != CurrentJig:
                CurrentJig = line[0]
                if Color == format_white:
                    Color = format_grey
                else:
                    Color = format_white
            sheet_Jigs.write_string("A"+str(i), line[0], Color)
            sheet_Jigs.write_string("B"+str(i), line[1], Color)

        # TMSInventory
        i = 1
        for line in TMSInventory:
            i += 1
            sheet_TMSInventory.write_string("A"+str(i), line[0], format_white)
            sheet_TMSInventory.write_string("B"+str(i), line[1], format_white)
            sheet_TMSInventory.write_string("C"+str(i), line[2], format_white)
            if line[3] == 0:
                sheet_TMSInventory.write_string("D"+str(i), str(line[3]), format_red)
            if line[3] == 1:
                sheet_TMSInventory.write_string("D"+str(i), str(line[3]), format_yellow)
            if line[3] > 1:
                sheet_TMSInventory.write_string("D"+str(i), str(line[3]), format_green)

        # MissingConnectors
        i = 1
        CurrentProgram = ""
        Color = format_grey
        for line in MissingConnectors:
            i += 1
            if CurrentProgram != line[0]:
                CurrentProgram = line[0]
                if Color == format_grey:
                    Color = format_white
                else:
                    Color = format_grey
            if line[3] == "":
                Color2 = format_red
                line3 = "Missing"
            else:
                Color2 = Color
                line3 = line[3]
            sheet_MissingConnectors.write_string("A"+str(i), line[0], Color)
            sheet_MissingConnectors.write_string("B"+str(i), line[1], Color)
            sheet_MissingConnectors.write_string("C"+str(i), line[2], Color)
            sheet_MissingConnectors.write_string("D"+str(i), line3, Color2)

        # Test Logs
        i = 1
        Color = format_grey
        CurrentProgram = ""
        for line in TestLogs:
            if CurrentProgram != line[2]:
                CurrentProgram = line[2]
                if Color == format_grey:
                    Color = format_white
                else:
                    Color = format_grey
            i += 1
            sheet_TestLogs.write_string("A"+str(i), line[2], Color)
            sheet_TestLogs.write_string("B"+str(i), line[0], Color)
            sheet_TestLogs.write_string("C"+str(i), line[1], Color)
        
        # TRANSFER BRAIDS FROM CSV TO LUA
        self.sc.print("Generating LUA files for braids...")
        # this function saves a copy of all connector maps in csv map mode as a lua map
        for root, dirs, files in os.walk(self.path_braids):
            for file in files:
                if ".csv" in file:
                    try:
                        name = file.replace(".csv","")
                        tmp = name[3:]
                        tmp = int(tmp)
                        error = False
                    except:
                        error = True

                    if not error:
                        path = self.path_braids+"/LUA"
                        newluafilename = path+"/"+file.replace(".csv", ".lua")
                        if not os.path.isdir(path):
                            os.makedirs(path)

                        # generate lua version
                        if not os.path.isfile(newluafilename):
                            # read csv file
                            csv = open(self.path_braids+"/"+file,"r")
                            lines = csv.readlines()
                            csv.close()
                            plugs = {} # plug_index: plug_pins
                            last_pin_name = ""
                            last_plug = ""
                            for line in lines[1:]:
                                # 0           , 1   , 2  , 3          , 4               , 5                  , 6
                                # GLOBAL_POINT, PLUG, PIN, PLUG_NUMBER, PART_NUMBER     , RAFAEL_PART_NUMBER , PIN_TYPE 
                                # 1           , 1   , A  , 1          , D38999/20WH53SN , 112065102          , Socket
                                line = line.replace("\n","")
                                line = line.split(",")
                                GLOBAL_POINT = line[0]
                                PLUG = line[1]
                                PIN = line[2]
                                Kelvin = PIN == last_pin_name and PLUG == last_plug
                                last_pin_name = PIN
                                last_plug = PLUG
                                if PLUG != "":
                                    if not PLUG in plugs:
                                        plugs[PLUG] = []
                                    plugs[PLUG].append((GLOBAL_POINT, PIN, Kelvin))

                            # prepare lua code
                            lua = []
                            lua.append("{\n")
                            lua.append("  version = 1,\n")
                            lua.append("  name = [["+name+"]],\n")
                            lua.append("  connector_list=\n")
                            lua.append("  {\n")
                            for plug_index, plug_pins in plugs.items():
                                lua.append("    {\n")
                                lua.append("      name = [["+name+"_"+plug_index+"]],\n")
                                lua.append("      pins = \n")
                                lua.append("      {\n")
                                for pin in plug_pins:
                                    global_point = pin[0]
                                    local_name = pin[1]
                                    Kelvin = pin[2]
                                    if not Kelvin:
                                        lua.append("        {"+global_point+",name=[["+local_name+"]]},\n")    
                                    else:
                                        lastline = lua.pop()
                                        lua.append(lastline.replace("]]},\n", "]],kelvin="+global_point+"},\n"))
                                lua.append("      }\n")
                                lua.append("    },\n")
                            lua.append("  },\n")
                            lua.append("}\n")
                            
                            # save to lua file
                            luafile = open(newluafilename,"w")
                            for line in lua:
                                luafile.write(line)
                            luafile.close()

        # TRANSFER BRAIDS FROM LUA TO CSV
        # this function saves a copy of all connector maps in lua map mode as a csv map
        
        # END
        workbook.close()
        self.sc.print("")
        self.sc.good("Done!")
        self.sc.exit()

main()