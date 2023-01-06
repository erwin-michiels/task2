import xlrd
import json

def find_all_device_interfaces(xlf):
    ### READ EXCEL FILE AND RETURN NUMBER OF ROWS
    wb = xlrd.open_workbook(xlf)
    sheet = wb.sheet_by_index(0)
    number_rows = sheet.nrows
    inventory_dict = {}
    dev_interfaces = []
    for r in range(number_rows):
        if r > 0: ### first row contains columns names
            COL_A =  sheet.cell_value(r, 0)  #### column A
            COL_B =  sheet.cell_value(r, 1)  #### column B
            COL_C =  sheet.cell_value(r, 2)  #### column C
            COL_D =  sheet.cell_value(r, 3)  #### column D
            COL_E =  sheet.cell_value(r, 4)  #### column E
            inventory_dict["device"]      = COL_A
            inventory_dict["role"]        = COL_B
            inventory_dict["interface"]   = COL_C 
            inventory_dict["ipaddress"]   = COL_D
            inventory_dict["subnetmask"]  = COL_E
            dev_interfaces.append(inventory_dict.copy()) # need to use copy()
    return dev_interfaces

def make_list_of_devices_and_roles(inventory):  
    dev_list  = []
    dev_dict  = {}
    mem       = {}
    for rec in inventory:
        dev_dict["dev_name"] = rec["device"]
        dev_dict["role"]     = rec["role"]
        if mem != dev_dict["dev_name"]:
            dev_list.append(dev_dict.copy())  # need to use copy()
        mem = dev_dict["dev_name"]
    #for rec in loc_g:    
        #print(rec)
    #del loc_g[0] ### if last item copied as first item
    return dev_list


def attach_interfaces_to_devices(dev_name, inventory):    
    intf_dict = {}
    intf_list = [intf_dict]
    for item in inventory:
        if item["device"] == dev_name:
            if item["device"] != None:
                intf_dict["interface"]   = item["interface"]
                intf_dict["ipaddress"]   = item["ipaddress"]
                intf_dict["subnetmask"]  = item["subnetmask"]
                intf_list.append(intf_dict.copy()) # need to use copy()
    del intf_list[0] ### if last item copied as first item
    return intf_list

def main(): 
    inventory_list = find_all_device_interfaces("ipdevices.xlsx")
    device_list    = make_list_of_devices_and_roles(inventory_list)
    dev_dict = []
    rack_struc = []      
    for device_rec in device_list:  
        intf_list  = attach_interfaces_to_devices(device_rec["dev_name"], inventory_list)
        dev_dict = { "device": device_rec , "interfaces": intf_list }
        rack_struc.append(dev_dict) # updated 18JAN2022
    js_rack = json.dumps(rack_struc)
    print(js_rack)

#### execute main() when called directly        
if __name__ == '__main__':
    main()