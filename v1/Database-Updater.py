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

        # test all paths
        self.sc.test_path(self.path_programs)
        self.sc.test_path(self.path_braids)

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
        TMSInventory = []
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
                    for file in files:
                        if file == PartNumber+".csv":
                            csv_file = True
                        if file == PartNumber+".txt":
                            txt_file = True
                        if file == PartNumber+".html":
                            html_file = True
                    if csv_file and txt_file and html_file:
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
                if ".csv" in file:
                    data = self.sc.load_csv(root+"/"+file)
                    TestCable = file.replace(".csv", "")
                    for line in data[1:]:
                        Braid = line[-4]
                        CustomerPN = line[-3]
                        FlexPN = line[-2]
                        Type = line[-1]
                        Braids.append((TestCable, Braid, CustomerPN, FlexPN, Size, Type))
        # GATHER DATA on TMS Inventory
        # GATHER DATA on Missing connectors

        # CREATE EXCEL TABLES

        # END
        self.sc.print("Done!")
        self.sc.exit()

main()