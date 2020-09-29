from config import *
import glob
import xmltodict
import os

def xml_to_list(xml_file):
    with open(xml_file, 'r') as o_file:
        lines = []
        for l in o_file:
            lines.append(l)
    recipe_lines = lines[0].replace("><",">zzzzzzzzz<").split('zzzzzzzzz')
    return recipe_lines

def xml_meas_param_dict(xml_file):
    xml_lines = xml_to_list(xml_file)
    name_index = [i for i in range(len(xml_lines)) if xml_lines[i] == "<Name>name</Name>"]
    #remove first instance which is name of measparam file.
    name_index.pop(0)
    meas_dict = {xml_lines[i+2][7:-8]:xml_lines[i+7][7:-8] for i in name_index}
    return meas_dict

def get_xml_dict(xml_file):
    with open(xml_file) as fd:
        doc = xmltodict.parse(fd.read())
    return doc

def get_xml_val(doc, path):
    for p in path:
        try:
            new_p = int(p)
            doc = doc[new_p]
        except ValueError:
            doc = doc[p]
    return doc

def scrape_meas_params(meas_file_path):

    file_paths = glob.glob(meas_file_path + "*.xml")
    meas_params = {}

    for file in file_paths:
        meas_dict = xml_meas_param_dict(file)
        ppl = os.path.basename(file)[:-4].split("__")
        xml_dict = get_xml_dict(file)
        meas_param_id = get_xml_val(xml_dict,['Descriptor', 'Value', 'DirectMapping', 'Field', '0', 'Value'])

        temp = [ppl[0], ppl[1], ppl[2], ppl[3], meas_dict]

        meas_params[meas_param_id] = temp

    return meas_params
