# this file gather information for the Test Cables spreadsheet
# the fields displayed will be the following:

# result[0]:
# 1. Test Cable ID              for example R1_001
# 2. Branch                     for example 2 (R2_004_2)
# 3. Plug Part Number           for example D38999/24D35SN
# 4. Flex Plug Part Number      for example 1120659429
# 5. Size                       The number of pins in the plug (not including GND)
# 7. Type                       Pin, Socket, Both, NA

# result[1]:
# 1. Test Cable ID              for example R1_001
# 2. Number of empty spaces     for exmaple 25

import os

class TestingCables:
    def get_data(test_cables):
        result = [ [("Test Cable ID","Branch","Plug Part Number","Flex Plug Part Number","Size","Type")] , [("Test Cable", "Empty Spaces")]]

        data = {} # R1_004: {branch: [pn, flex_pn, pintype, size]}
        empty_spaces_data = {} # R1_004: 15

        # read braids maps
        for root, dirs, files in os.walk(test_cables):
            for file in files:
                if ".csv" in file and not "jigs.csv" in file:
                    # set variables
                    test_cable_id = file.replace(".csv", "")
                    if not test_cable_id in data:
                        data[test_cable_id] = {}
                    branch = "1"
                    plug_pn = "None"
                    flex_pn = "None"
                    plug_type = "None"

                    # read file data
                    f = open(root+"/"+file, 'r')
                    lines = f.readlines()
                    f.close()
                    
                    # go over each line
                    for line in lines[1:]:
                        line = line.split(",")
                        if len(line) == 7:
                            branch = line[3]
                            plug_pn = line[4]
                            flex_pn = line[5]
                            plug_type = line[6].replace("\n", "")
                            if branch != "":
                                if not branch in data[test_cable_id]:
                                    data[test_cable_id][branch] = [plug_pn, flex_pn, plug_type, 0]
                        if len(line) > 2:
                            branch = line[1]
                            pin = line[2]
                            if branch != "":
                                if pin.upper() != "BODY" and pin != "":
                                    if branch in data[test_cable_id]:
                                        data[test_cable_id][branch][3] += 1
                        
                        # empty spaces
                        if len(line) > 2:
                            if line[0] != "" and line[1] == "":
                                if not test_cable_id in empty_spaces_data:
                                    empty_spaces_data[test_cable_id] = 1
                                else:
                                    empty_spaces_data[test_cable_id] += 1
        
        # save results
        for testcable_id, branches in data.items(): # R1_004: {branch: [pn, flex_pn, pintype, size]}
            for branch, details in branches.items(): # branch: [pn, flex_pn, pintype, size]
                # result = [("Test Cable ID","Branch","Plug Part Number","Flex Plug Part Number","Size","Type")]
                result[0].append((testcable_id, branch, details[0], details[1], details[3], details[2]))

        for testcable_id, empty_spaces in empty_spaces_data.items():
            result[1].append((testcable_id,empty_spaces))

        return result

    def makelua(path_test_cables):
        # read braids maps
        test_cables = [] # a list of TestCable objects
        for root, dirs, files in os.walk(path_test_cables):
            for file in files:
                if ".csv" in file and not "JIGS.csv" in file:
                    test_cable_name = file.replace(".csv", "")
                    if not test_cable_name in test_cables:
                        new = TestCable(test_cable_name)
                        test_cables.append(new)
                    f = open(root+"/"+file, 'r')
                    lines = f.readlines()
                    f.close()
                    
                    previous_pin = ""
                    for line in lines[1:]:
                        line = line.replace("\n", "")
                        line = line.split(",")
                        if len(line) > 2:
                            globalpoint = line[0]
                            plug = line[1]
                            pin = line[2]
                            kelvin = None
                            if pin == previous_pin:
                                kelvin = globalpoint
                            previous_pin = pin
                            if plug != "":
                                if not plug in new.branches:
                                    new.branches[plug] = {}
                                new.branches[plug][globalpoint] = (pin, kelvin)
        
        # make lua files
        for obj_tc in test_cables:
            # generate lua code
            output_filename = obj_tc.name+".lua"
            output_text = ["{"]
            output_text.append("\tversion = 1,")
            output_text.append("\tname = [["+obj_tc.name+"]],")
            output_text.append("\tconnector_list = ")
            output_text.append("\t{")
            for branch, data in obj_tc.branches.items():
                output_text.append("\t\t{")
                output_text.append("\t\t\tname = [["+obj_tc.name+"_"+branch+"]],")
                output_text.append("\t\t\tpins = ")
                output_text.append("\t\t\t{")
                for global_point, pins in data.items():
                    kelvin = ""
                    gp = global_point
                    if pins[1] != None:
                        kelvin = ",kelvin="+pins[1]
                        output_text.pop()
                        gp = int(global_point)-1
                    if pins[0] != "":
                        output_text.append("\t\t\t\t{"+str(gp)+",name=[["+pins[0]+"]]"+kelvin+"}")
                output_text.append("\t\t\t}")
                output_text.append("\t\t},")
            output_text.append("\t}")
            output_text.append("}")

            # save lua file
            if not os.path.isdir(path_test_cables+"/LUA"):
                os.makedirs(path_test_cables+"/LUA")
            if not os.path.isfile(path_test_cables+"/LUA/"+output_filename):
                f = open(path_test_cables+"/LUA/"+output_filename, 'w')
                for line in output_text:
                    f.write(line+"\n")
                f.close()
class TestCable:
    def __init__(self, name):
        self.name = name
        self.branches = {} # {1: { global point : (pin name, kelvin global point) } }