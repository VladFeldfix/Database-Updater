# Download SmartConsole.py from: https://github.com/VladFeldfix/Smart-Console/blob/main/SmartConsole.py
from SmartConsole import *
from functions.Programs import Programs
from functions.Testing_Cables import TestingCables
from functions.Jigs import Jigs
from functions.Products import Products
from functions.Test_Logs import Test_Logs
from functions.Generate_Excel import Generate_Excel
import datetime
import time

class main:
    # constructor
    def __init__(self):
        # load smart console
        self.sc = SmartConsole("Database Updater", "3.1")

        # set-up main memu
        self.sc.add_main_menu_item("RUN", self.run)

        # get settings
        self.path_database = self.sc.get_setting("Database Location")
        self.path_programs = self.sc.get_setting("Programs Location")
        self.path_test_results = self.sc.get_setting("Test Results Location")
        self.path_test_cables = self.sc.get_setting("Test Cables")

        # test all paths
        self.sc.test_path(self.path_database)
        self.sc.test_path(self.path_programs)
        self.sc.test_path(self.path_test_results)
        self.sc.test_path(self.path_test_cables)
        self.sc.test_path(self.path_test_cables+"/JIGS.csv")

        # display main menu
        self.sc.start()
    
    def gettime(self):
        ts = time.time()
        return datetime.datetime.fromtimestamp(ts).strftime('%Y-%m-%d %H:%M:%S')

    def run(self):
        ts = self.gettime()
        self.sc.print(ts+" Processing...")
        self.sc.print("\nGATHERING DATA:")
        
        ts = self.gettime()
        self.sc.print(ts+" Gathering data on programs...")
        self.data_programs = Programs.get_data(self.path_programs, self.path_test_results)
        
        ts = self.gettime()
        self.sc.print(ts+" Gathering data on test cables...")
        self.data_testingcables = TestingCables.get_data(self.path_test_cables)
        
        ts = self.gettime()
        self.sc.print(ts+" Gathering data on jigs...")
        self.data_jigs = Jigs.get_data(self.path_test_cables+"/JIGS.csv")
        
        ts = self.gettime()
        self.sc.print(ts+" Gathering data on products...")
        self.data_products = Products.get_data(self.path_programs)
        
        ts = self.gettime()
        self.sc.print(ts+" Gathering data on test logs...")
        self.data_testlogs = Test_Logs.get_data(self.path_test_results)
        
        self.generate_files()
    
    def generate_files(self):
        self.sc.print("\nGENERATING DATA:")

        ts = self.gettime()
        self.sc.print(ts+" Generating Lua files...")
        TestingCables.makelua(self.path_test_cables)

        ts = self.gettime()
        self.sc.print(ts+" Generating Excel table...")
        xl = Generate_Excel()
        xl.setup_sheets(self.path_database)
        xl.write_data(self.data_programs, self.data_testingcables, self.data_jigs, self.data_products, self.data_testlogs)
        
        # finish
        ts = self.gettime()
        self.sc.print("")
        self.sc.good(ts+" Done!")
        
        # restart
        self.sc.restart()

        
main()