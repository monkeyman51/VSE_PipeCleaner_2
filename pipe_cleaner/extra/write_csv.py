from xlrd import open_workbook
# from simplejson import dumps
from json import loads
from json import dumps


gens = open_workbook(r'C:\Users\joe.ton\Documents\skudoc.xlsx').sheet_by_index(0)
path = '../skudoc/may_2020/'


supplier_dict = {}
generation_dict = {}
DIMM_16GB_dict = {}
DIMM_32GB_dict = {}
SATA_SSD_480GB_dict = {}
SATA_SSD_960GB_dict = {}
NVMe_960GB_dict = {}
SATA_HDD_4TB_dict = {}
SATA_HDD_6TB_dict = {}
SATA_HDD_10TB_dict = {}
SATA_HDD_12TB_dict = {}
NVMe_1_9TB_dict = {}
NVMe_1_92TB_dict = {}
NVMe_3_84TB_dict = {}
NVMe_3_840GB_M_2_dict = {}
NVMe_3_92TB_dict = {}
Internal_Storage_960GB_SSD_dict = {}
Internal_Storage_10TB_7_2K_dict = {}
Internal_Storage_12TB_7_2K_dict = {}
Internal_Storage_14TB_7_2K_dict = {}
NVDIMM_dict = {}

configuration_list = []
supplier_list = []
generation_list = []
DIMM_16GB_list = []
DIMM_32GB_list = []
SATA_SSD_480GB_list = []
SATA_SSD_960GB_list = []
NVMe_960GB_list = []
SATA_HDD_4TB_list = []
SATA_HDD_6TB_list = []
SATA_HDD_10TB_list = []
SATA_HDD_12TB_list = []
NVMe_1_9TB_list = []
NVMe_1_92TB_list = {}
NVMe_3_84TB_list = []
NVMe_3_840GB_M_2_list = []
NVMe_3_92TB_list = []
Internal_Storage_960GB_SSD_list = []
Internal_Storage_10TB_7_2K_list = []
Internal_Storage_12TB_7_2K_list = []
Internal_Storage_14TB_7_2K_list = []
NVDIMM_list = []

def get_skudoc():
    get_supplier(path, supplier_dict, 'supplier')
    get_generation(path, generation_dict, 'generation')
    get_DIMM_16GB(path, DIMM_16GB_dict, 'DIMM_16GB')
    get_DIMM_32GB(path, DIMM_32GB_dict, 'DIMM_32GB')
    get_SATA_SSD_480GB(path, SATA_SSD_480GB_dict, 'SATA_SSD_480GB')
    get_SATA_SSD_960GB(path, SATA_SSD_960GB_dict, 'SATA_SSD_960GB')
    get_NVMe_960GB(path, NVMe_960GB_dict, 'NVMe_960GB')
    get_SATA_HDD_4TB(path, SATA_HDD_4TB_dict, 'SATA_HDD_4TB')
    get_SATA_HDD_6TB(path, SATA_HDD_6TB_dict, 'SATA_HDD_6TB')
    get_SATA_HDD_10TB(path, SATA_HDD_10TB_dict, 'SATA_HDD_10TB')
    get_SATA_HDD_12TB(path, SATA_HDD_12TB_dict, 'SATA_HDD_12TB')
    get_NVMe_1_9TB(path, NVMe_1_9TB_dict, 'NVMe_1_9TB')
    get_NVMe_3_84TB(path, NVMe_3_84TB_dict, 'NVMe_3_84TB')
    get_NVMe_3_840GB_M_2(path, NVMe_3_840GB_M_2_dict, 'NVMe_3_840GB_M_2')
    get_NVMe_3_92TB(path, NVMe_3_92TB_dict, 'NVMe_3_92TB')
    get_Internal_Storage_960GB_SSD(path, Internal_Storage_960GB_SSD_dict, 'Internal_Storage_960GB_SSD')
    get_Internal_Storage_10TB_7_2K(path, Internal_Storage_10TB_7_2K_dict, 'Internal_Storage_10TB_7_2K')
    get_Internal_Storage_12TB_7_2K(path, Internal_Storage_12TB_7_2K_dict, 'Internal_Storage_12TB_7_2K')
    get_Internal_Storage_14TB_7_2K(path, Internal_Storage_14TB_7_2K_dict, 'Internal_Storage_14TB_7_2K')
    get_NVDIMM(path, NVDIMM_dict, 'NVDIMM')

def get_supplier(path, data, file):
    with open(f'{path}{file}.json', 'w') as f:
        for configuration in range(gens.nrows):
            key = str(gens.cell(configuration, 0))
            key = str(unnecessary_excel(key))
            value = str(gens.cell(configuration, 1))
            value = str(unnecessary_excel(value))
            data.update({key: value})
        f.write(dumps(data, indent=4))

def get_generation(path, data, component):
    with open(f'{path}{component}.json', 'w') as f:
        for configuration in range(gens.nrows):
            key = str(gens.cell(configuration, 0))
            key = str(unnecessary_excel(key))
            value = str(gens.cell(configuration, 2))
            value = str(unnecessary_excel(value))
            data.update({key: value})
        f.write(dumps(data, indent=4))

def get_DIMM_16GB(path, data, component):
    with open(f'{path}{component}.json', 'w') as f:
        for configuration in range(gens.nrows):
            key = str(gens.cell(configuration, 0))
            key = str(unnecessary_excel(key))
            value = str(gens.cell(configuration, 3))
            value = str(unnecessary_excel(value))
            data.update({key: value})
        f.write(dumps(data, indent=4))

def get_DIMM_32GB(path, data, component):
    with open(f'{path}{component}.json', 'w') as f:
        for configuration in range(gens.nrows):
            key = str(gens.cell(configuration, 0))
            key = str(unnecessary_excel(key))
            value = str(gens.cell(configuration, 4))
            value = str(unnecessary_excel(value))
            data.update({key: value})
        f.write(dumps(data, indent=4))

def get_SATA_SSD_480GB(path, data, component):
    with open(f'{path}{component}.json', 'w') as f:
        for configuration in range(gens.nrows):
            key = str(gens.cell(configuration, 0))
            key = str(unnecessary_excel(key))
            value = str(gens.cell(configuration, 5))
            value = str(unnecessary_excel(value))
            data.update({key: value})
        f.write(dumps(data, indent=4))

def get_SATA_SSD_960GB(path, data, component):
    with open(f'{path}{component}.json', 'w') as f:
        for configuration in range(gens.nrows):
            key = str(gens.cell(configuration, 0))
            key = str(unnecessary_excel(key))
            value = str(gens.cell(configuration, 6))
            value = str(unnecessary_excel(value))
            data.update({key: value})
        f.write(dumps(data, indent=4))

def get_NVMe_960GB(path, data, component):
    with open(f'{path}{component}.json', 'w') as f:
        for configuration in range(gens.nrows):
            key = str(gens.cell(configuration, 0))
            key = str(unnecessary_excel(key))
            value = str(gens.cell(configuration, 7))
            value = str(unnecessary_excel(value))
            data.update({key: value})
        f.write(dumps(data, indent=4))

def get_SATA_HDD_4TB(path, data, component):
    with open(f'{path}{component}.json', 'w') as f:
        for configuration in range(gens.nrows):
            key = str(gens.cell(configuration, 0))
            key = str(unnecessary_excel(key))
            value = str(gens.cell(configuration, 8))
            value = str(unnecessary_excel(value))
            data.update({key: value})
        f.write(dumps(data, indent=4))

def get_SATA_HDD_6TB(path, data, component):
    with open(f'{path}{component}.json', 'w') as f:
        for configuration in range(gens.nrows):
            key = str(gens.cell(configuration, 0))
            key = str(unnecessary_excel(key))
            value = str(gens.cell(configuration, 9))
            value = str(unnecessary_excel(value))
            data.update({key: value})
        f.write(dumps(data, indent=4))

def get_SATA_HDD_10TB(path, data, component):
    with open(f'{path}{component}.json', 'w') as f:
        for configuration in range(gens.nrows):
            key = str(gens.cell(configuration, 0))
            key = str(unnecessary_excel(key))
            value = str(gens.cell(configuration, 10))
            value = str(unnecessary_excel(value))
            data.update({key: value})
        f.write(dumps(data, indent=4))

def get_SATA_HDD_12TB(path, data, component):
    with open(f'{path}{component}.json', 'w') as f:
        for configuration in range(gens.nrows):
            key = str(gens.cell(configuration, 0))
            key = str(unnecessary_excel(key))
            value = str(gens.cell(configuration, 11))
            value = str(unnecessary_excel(value))
            data.update({key: value})
        f.write(dumps(data, indent=4))

def get_NVMe_1_9TB(path, data, component):
    with open(f'{path}{component}.json', 'w') as f:
        for configuration in range(gens.nrows):
            key = str(gens.cell(configuration, 0))
            key = str(unnecessary_excel(key))
            value = str(gens.cell(configuration, 12))
            value = str(unnecessary_excel(value))
            data.update({key: value})
        f.write(dumps(data, indent=4))

def get_NVMe_3_84TB(path, data, component):
    with open(f'{path}{component}.json', 'w') as f:
        for configuration in range(gens.nrows):
            key = str(gens.cell(configuration, 0))
            key = str(unnecessary_excel(key))
            value = str(gens.cell(configuration, 13))
            value = str(unnecessary_excel(value))
            data.update({key: value})
        f.write(dumps(data, indent=4))

def get_NVMe_3_840GB_M_2(path, data, component):
    with open(f'{path}{component}.json', 'w') as f:
        for configuration in range(gens.nrows):
            key = str(gens.cell(configuration, 0))
            key = str(unnecessary_excel(key))
            value = str(gens.cell(configuration, 14))
            value = str(unnecessary_excel(value))
            data.update({key: value})
        f.write(dumps(data, indent=4))

def get_NVMe_3_92TB(path, data, component):
    with open(f'{path}{component}.json', 'w') as f:
        for configuration in range(gens.nrows):
            key = str(gens.cell(configuration, 0))
            key = str(unnecessary_excel(key))
            value = str(gens.cell(configuration, 15))
            value = str(unnecessary_excel(value))
            data.update({key: value})
        f.write(dumps(data, indent=4))
def get_Internal_Storage_960GB_SSD(path, data, component):
    with open(f'{path}{component}.json', 'w') as f:
        for configuration in range(gens.nrows):
            key = str(gens.cell(configuration, 0))
            key = str(unnecessary_excel(key))
            value = str(gens.cell(configuration, 16))
            value = str(unnecessary_excel(value))
            data.update({key: value})
        f.write(dumps(data, indent=4))

def get_Internal_Storage_10TB_7_2K(path, data, component):
    with open(f'{path}{component}.json', 'w') as f:
        for configuration in range(gens.nrows):
            key = str(gens.cell(configuration, 0))
            key = str(unnecessary_excel(key))
            value = str(gens.cell(configuration, 17))
            value = str(unnecessary_excel(value))
            data.update({key: value})
        f.write(dumps(data, indent=4))

def get_Internal_Storage_12TB_7_2K(path, data, component):
    with open(f'{path}{component}.json', 'w') as f:
        for configuration in range(gens.nrows):
            key = str(gens.cell(configuration, 0))
            key = str(unnecessary_excel(key))
            value = str(gens.cell(configuration, 18))
            value = str(unnecessary_excel(value))
            data.update({key: value})
        f.write(dumps(data, indent=4))

def get_Internal_Storage_14TB_7_2K(path, data, component):
    with open(f'{path}{component}.json', 'w') as f:
        for configuration in range(gens.nrows):
            key = str(gens.cell(configuration, 0))
            key = str(unnecessary_excel(key))
            value = str(gens.cell(configuration, 19))
            value = str(unnecessary_excel(value))
            data.update({key: value})
        f.write(dumps(data, indent=4))

def get_NVDIMM(path, data, component):
    with open(f'{path}{component}.json', 'w') as f:
        for configuration in range(gens.nrows):
            key = str(gens.cell(configuration, 0))
            key = str(unnecessary_excel(key))
            value = str(gens.cell(configuration, 20))
            value = str(unnecessary_excel(value))
            data.update({key: value})
        f.write(dumps(data, indent=4))

def unnecessary_excel(string):
    new_string = str(string)
    if "text:" in new_string:
        new_string = new_string.replace("text:", '')
    if "number:" in new_string:
        new_string = new_string.replace("number:", '')
    if "empty:" in new_string:
        new_string = new_string.replace("empty:", '')
    if "'" in new_string:
        new_string = new_string.replace("'", '')
    if "\n  " in new_string:
        new_string = new_string.replace("\n  ", '')
    if "\\n" in new_string:
        new_string = new_string.replace("\\n", ' ')
    if "\\n  " in new_string:
        new_string = new_string.replace("\\n  ", '')
    if "\u00a0" in new_string:
        new_string = new_string.replace("\u00a0", '')
    if "\u00c2" in new_string:
        new_string = new_string.replace("\u00c2", '')
    if "\n  " in new_string:
        new_string = new_string.replace("\n  ", '')
    if "\\n  " in new_string:
        new_string = new_string.replace("\\n  ", '')
    if "\u00a0" in new_string:
        new_string = new_string.replace("\u00a0", '')
    if "\u00c2" in new_string:
        new_string = new_string.replace("\u00c2", '')
    return new_string

get_skudoc()

def get_skudoc():
    with open('../skudoc/may_2020/skudoc.csv', 'r') as f:
        for configuration in f:
            configuration_list.append(configuration)
        for supplier in f:
            supplier_list.append(supplier)
        for generation in f:
            generation_list.append(generation)
        for DIMM_16GB in f:
            DIMM_16GB_list.append(DIMM_16GB)
        for DIMM_32GB in f:
            DIMM_32GB_list.append(DIMM_32GB)
        for SATA_SSD_480GB in f:
            SATA_SSD_480GB_list.append(SATA_SSD_480GB)
        for SATA_SSD_960GB in f:
            SATA_SSD_960GB_list.append(SATA_SSD_960GB)
        for NVMe_960GB in f:
            NVMe_960GB_list.append(NVMe_960GB)
        for SATA_HDD_4TB in f:
            SATA_HDD_4TB_list.append(SATA_HDD_4TB)
        for SATA_HDD_6TB in f:
            SATA_HDD_6TB_list.append(SATA_HDD_6TB)
        for SATA_HDD_10TB in f:
            SATA_HDD_10TB_list.append(SATA_HDD_10TB)
        for SATA_HDD_12TB in f:
            SATA_HDD_12TB_list.append(SATA_HDD_12TB)
        for NVMe_1_9TB in f:
            NVMe_1_9TB_list.append(NVMe_1_9TB)
        for NVMe_1_92TB in f:
            NVMe_1_92TB_list.append(NVMe_1_92TB)
        for NVMe_3_84TB in f:
            NVMe_3_84TB_list.append(NVMe_3_84TB)
        for NVMe_3840GB_M_2 in f:
            NVMe_3_840GB_M_2_list.append(NVMe_3840GB_M_2)
        for NVMe_3_92TB in f:
            NVMe_3_92TB_list.append(NVMe_3_92TB)
        for Internal_Storage_960GB_SSD in f:
            Internal_Storage_960GB_SSD_list.append(Internal_Storage_960GB_SSD)
        for Internal_Storage_10TB_7_2K in f:
            Internal_Storage_10TB_7_2K_list.append(Internal_Storage_10TB_7_2K)
        for Internal_Storage_12TB_7_2K in f:
            Internal_Storage_12TB_7_2K_list.append(Internal_Storage_12TB_7_2K)
        for Internal_Storage_14TB_7_2K in f:
            Internal_Storage_14TB_7_2K_list.append(Internal_Storage_14TB_7_2K)
        for NVDIMM in f:
            NVDIMM_list.append(NVDIMM)