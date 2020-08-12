def regexCheck(str_text):
    import re
    if re.match("^[0-9]{8}-[0-9]{3}-[0-9]{2}", str_text) or re.match("^[0-9]{8}-[0-9]{5}-[0-9]{2}", str_text) or re.match("^[0-9]{8}-[0-9]{4}-[0-9]{2}", str_text):
        return True
    else:
        return False


def writeXLFile(holding_dict):
    import openpyxl
    xl_file = "C:\\Users\\Report.xlsx"
    wb = openpyxl.load_workbook(xl_file)
    sheet = wb.active
    used_range = sheet['D1'].value
    if used_range is None:
        used_range = 1
    for dict_key, dict_value in holding_dict.items():
        used_range = int(used_range) + 1
        sheet['A' + str(used_range)].value = dict_key
        sheet['B' + str(used_range)].value = dict_value + ".dat"
        sheet['D1'].value = used_range
    wb.save(xl_file)


def readXML(xml_file):
    import xml.dom.minidom as minidom
    str_file_header = "!~ParentID~RevID~ChildID~ChildRev~BL:catiaParentPartName~BL:catiaFileName~BL:catiaOccurrenceName~BL:bl_plmxml_occ_xform" + "\n"
    str_file_text = ""
    xml_doc = minidom.parse(xml_file)
    property_node = xml_doc.getElementsByTagName('property')
    properties_count = 0
    while property_node[properties_count].getAttribute('id') != "CatiaV5-Teilenummer":
        properties_count = properties_count + 1

    teilenummer = property_node[properties_count].getAttribute("value")
    dash_position = teilenummer.rfind("-", 0)
    parent_id = teilenummer[0:dash_position]
    parent_rev = teilenummer[dash_position + 1:]

    occurrence = xml_doc.getElementsByTagName('occurrence')
    node_count = 0
    unique_wrong_file_name_dict = {}
    while node_count < occurrence.length:
        catia_parent_part_name = occurrence[node_count].getAttribute('id')
        catia_occurrence_name = occurrence[node_count].getAttribute('name')
        if catia_parent_part_name == catia_occurrence_name:
            catia_parent_part_name = ""

        cadreference = occurrence[node_count].getElementsByTagName('cadreference')
        catia_file_name = cadreference[0].getAttribute('path')
        dot_position = catia_file_name.rfind(".", 0)
        regex_check_string = catia_file_name[0:dot_position]
        regex_check_val = regexCheck(regex_check_string)
        if regex_check_val:
            dash_position = catia_file_name.rfind("-", 0)
            child_id = catia_file_name[0:dash_position]
            child_rev = catia_file_name[dash_position + 1:]
        else:
            child_id = catia_file_name
            child_rev = "child_rev"
            if child_id not in unique_wrong_file_name_dict:
                unique_wrong_file_name_dict.update({child_id: xml_file.name})

        tmatrix = occurrence[node_count].getElementsByTagName('tmatrix')
        t_entry = tmatrix[0].getElementsByTagName('entry')
        t_nodes = 0
        t_dict = {}
        while t_nodes < t_entry.length:
            t_id = t_entry[t_nodes].getAttribute('id')
            t_value = t_entry[t_nodes].getAttribute('value')
            if t_id == "4" or t_id == "8" or t_id == "4":
                t_value = float(t_value) / 1000

            t_dict.update({t_id: str(t_value)})
            t_nodes = t_nodes + 1
        bl_plmxml_occ_xform = t_dict.get("1") + " " + t_dict.get("5") + " " + t_dict.get("9") + " " + t_dict.get("13") + " " + t_dict.get("2") + " " + t_dict.get("6") + " " + t_dict.get("10") + " " + t_dict.get("14") + " " + t_dict.get(
            "3") + " " + t_dict.get("7") + " " + t_dict.get("11") + " " + t_dict.get("15") + " " + t_dict.get("4") + " " + t_dict.get("8") + " " + t_dict.get("12") + " " + t_dict.get("16")

        str_file_text = str_file_text + parent_id + "~" + parent_rev + "~" + child_id + "~" + child_rev + "~" + catia_parent_part_name + "~" + catia_file_name + "~" + catia_occurrence_name + "~" + bl_plmxml_occ_xform + "\n"
        node_count = node_count + 1
    writeXLFile(unique_wrong_file_name_dict)
    return str_file_header + str_file_text


def write_dat_file(str_filename, str_text):
    with open(str_filename, "w") as dat_file:
        dat_file.write(str_text)
        dat_file.close()
        print(str_filename + " has been created.")


# This is the main function
def main():
    import os
    input_dir = 'C:\\Users\\Desktop'
    for filename in os.listdir(input_dir):
        if filename.endswith('.appinfo'):
            with open(input_dir + "\\" + filename, "r") as xml_file:
                str_dbom_content = readXML(xml_file)
                write_dat_file(input_dir + "\\" + filename + ".dat", str_dbom_content)


main()
