
trr_qcl: list = ['HYNIX HMA82GR7CJR4N-VK(#M1098844-001)', 'SKHYNIX HMA84GR7CJR4N-VK(#M1086479-001)',
                 'MMOD,16GB,SM,DDR4,3200,HMA82GR7DJR4N-XN', 'HMA82GR7DJR4N-XN', '14.25.8100',
                 'BOOTLOADER: 0004.0100IMAGEA: V010D.E84IMAGEB: V010D.E84',
                 'SAMSUNG MZ1LB960HAJQ-000MV (M1077133-001) /FW: EDB78M5Q', '1.1.18.6']

machine_dimm_part_numbers: list = ['HMA84GR7CJR4N-VK', 'HMA82GR7DJR4N-XN']

for version in machine_dimm_part_numbers:
    for qualified_component in trr_qcl:
        if version in qualified_component:
            print('YES')