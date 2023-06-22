from SmartConsole import *
import xlsxwriter

class main:
    # constructor
    def __init__(self):
        # load smart console
        self.sc = SmartConsole("Database Updater", "1.0")

        # get settings
        self.path_programs = self.sc.get_setting("Programs Location")
        self.path_braids = self.sc.get_setting("Braids Location")
        self.path_tms_inv = self.sc.get_setting("TMS Inventory")
        self.path_tms_db = self.sc.get_setting("TMS Database")

        # test all paths
        self.sc.test_path(self.path_programs)
        self.sc.test_path(self.path_braids)
        self.sc.test_path(self.path_tms_inv)
        self.sc.test_path(self.path_tms_db)

        # display main menu
        self.run()
    
    def run(self):
        # SPREADSHEET DESCRIPTION
        # Programs - this spreadsheet shows wheather the program is ready or not, and if not, then what is it missing
        # Braids - information about exsiting braids maps
        # TMS Inventory - The number of boxes left in the TMS inventory
        # Missing connectors - What meeting connectors are missing for what products

        # GATHER DATA
        Programs = [("PART NUMBER", "MPT PROGRAM")]
        Braids = [("TEST CABLE", "BRAID", "CUSTOMER PART NUMBER", "FLEX PART NUMBER", "SIZE", "TYPE")]
        TMSInventory = [("TMS PART NUMBER", "SIZE", "COLOR", "QTY")]
        MissingConnectors = []

        # GATHER DATA on Programs
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
                    Missing_plugs = False
                    for file in files:
                        if file == PartNumber+".csv":
                            csv_file = True
                        if file == PartNumber+".txt":
                            txt_file = True
                        if file == PartNumber+".html":
                            html_file = True
                        if file == "testcables_to_product.csv":
                            if os.path.isfile(root+"/testcables_to_product.csv"):
                                data = self.sc.load_csv(root+"/testcables_to_product.csv")
                                for line in data[1:]:
                                    if len(line) == 3:
                                        MissingConnectors.append((PartNumber, line[1], line[2]))
                                        Missing_plugs = True
                    if csv_file and txt_file and html_file and not Missing_plugs:
                        Status = "Complete"
                    Programs.append((PartNumber, Status))

        # GATHER DATA on Braids
        TestCable = ""
        Braid = ""
        CustomerPN = ""
        FlexPN = ""
        Size = 0
        Type = ""
        for root, dirs, files in os.walk(self.path_braids):
            for file in files:
                info = {}
                if ".csv" in file:
                    data = self.sc.load_csv(root+"/"+file)
                    TestCable = file.replace(".csv", "")
                    LegalTestCable = True
                    try:
                        TestCable = int(TestCable)
                    except:
                        LegalTestCable = False
                    PreviousPin = ""
                    if LegalTestCable:
                        for line in data[1:]:
                            GlobalPoint = line[0]
                            CurrentBraid = line[1]
                            Pin = line[2]
                            Braid = line[3]
                            CustomerPN = line[4]
                            FlexPN = line[5]
                            Type = line[6]
                            if Braid != "":
                                info[Braid] = [CustomerPN, FlexPN, Type, 0]
                            if CurrentBraid != "":
                                if CurrentBraid in info:
                                    if len(info[CurrentBraid]) == 4:
                                        if Pin != PreviousPin and Pin.upper() != "BODY":
                                            info[CurrentBraid][3] += 1
                                            PreviousPin = Pin
                        for Braid, data in info.items():
                            # "TEST CABLE", "BRAID", "CUSTOMER PART NUMBER", "FLEX PART NUMBER", "SIZE", "TYPE"
                            CustomerPN = data[0]
                            FlexPN = data[1]
                            Type = data[2]
                            Size = data[3]
                            Braids.append((TestCable, Braid, CustomerPN, FlexPN, Size, Type))
        Braids = sorted(Braids[1:])
        Braids = [("TEST CABLE", "BRAID", "CUSTOMER PART NUMBER", "FLEX PART NUMBER", "SIZE", "TYPE")] + Braids

        # GATHER DATA on TMS Inventory
        data = self.sc.load_csv(self.path_tms_inv)
        parm = self.sc.load_csv(self.path_tms_db)
        tms = {}
        parameters = {}
        for line in data[1:]:
            PartNumber = line[1]
            if PartNumber in tms:
                tms[PartNumber] += 1
            else:
                tms[PartNumber] = 1
        for line in parm:
            PartNumber = line[0]
            Size = line[1]
            Color = line[2]
            parameters[PartNumber] = (Size,Color)
        for PartNumber, qty in tms.items():
            if PartNumber in parameters:
                Size = parameters[PartNumber][0]
                Color = parameters[PartNumber][1]
                TMSInventory.append((PartNumber, Size, Color, qty))
        for PartNumber, values in parameters.items():
            if not PartNumber in tms and PartNumber != "PART NUMBER":
                TMSInventory.append((PartNumber, values[0], values[1], 0))

        # GATHER DATA on Missing connectors

        # CREATE EXCEL TABLES

        # END
        self.sc.print("Done!")
        self.sc.exit()

main()