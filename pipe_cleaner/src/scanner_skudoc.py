import os
from pipe_cleaner.extra.crd_scanner_1 import break_target_configuration, decode_processor, decode_system


def break_skuddoc_name(skudoc_path: str):

    for file_name in os.listdir(skudoc_path):
        print(file_name)


def break_components(name_to_target: dict):
    """
    Break Machine's Target Configuration (TRR), processor, and system into a List.
    :param name_to_target:
    :return:
    """
    for machine_name in name_to_target:
        broken_parts: list = []

        broken_parts = break_target_configuration(name_to_target[machine_name])
        processor = decode_processor(machine_name)
        system = decode_system(machine_name)
        broken_parts.append(processor)
        broken_parts.append(system)

        # Rid of Duplicates
        final_components = list(set(broken_parts))

        print(final_components)


def main_method(name_to_target: dict, name_to_ticket: dict):
    """
    Start scanning SKUDOCs Here
    :return:
    """
    skudoc_path = r'\\172.30.1.100\pxe\Kirkland_Lab\Microsoft_CSI\Documentation\SKUDOC'

    # for file in os.listdir(skudoc_path):
    #     print(file)

    break_components(name_to_target)
