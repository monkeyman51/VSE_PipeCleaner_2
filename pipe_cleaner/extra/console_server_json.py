import requests
from json import loads

host_id = 'C1545-206753100006'
path_info = '../../reports/'


# Need to be in 172.30.1.100 network for console server access
def access(host, path):
    data = {
        'action': 'get_json_data',
        'host_id': f'{host}'
    }
    response = requests.post(url=f'http://172.30.1.100/results/{host}.json', data=data)

    with open(f'{path}{host}.json', 'w') as f:
        f.write(response.text)
    read_json(host, path)


class Componenets():
    with open(f'{str(path)}{str(host)}.json') as f:
        json = loads(f.read())

    def bios():
        bios = str(json['dmi']['bios']['version'])
        return bios

    def bmc():
        bmc = str(json['bmc']['mc']['firmware']).replace('.', '')
        return bmc

    # Need to get CPLD on Console Server
    def cpld():
        cpld = 'empty'
        return cpld

    def os():
        os = str(json['platform']['version'])
        print(os)

    # Need to get Chipset Driver on Console Server
    def chipset_driver():
        chipset_driver = 'empty'
        return chipset_driver

    # Need to get partition volume on Console Server
    def partition_volume():
        partition_volume = 'empty'
        return partition_volume

    # Need to get boot drive on Console Server
    def boot_drive():
        boot_drive = 'empty'
        return boot_drive

    # Need to get motherboard pn on Console Server
    def motherboard_pn():
        motherboard_pn = 'empty'
        return motherboard_pn

    # Need to get processors pn on Console Server
    def processors():
        processors = 'empty'
        return processors

    def tpm():
        tpm = json['tpm']['version'][:4:]
        return tpm

    # Need to get FPGA release package on Console Server
    def fpga_release_package():
        fpga_release_package = 'empty'
        return fpga_release_package

    # Need to get FPGA Board PN on Console Server
    def fpga_board_pn():
        fpga_board_pn = 'empty'
        return fpga_board_pn

    # Need to get FPGA Active Image on Console Server
    def fpga_active_image():
        fpga_active_image = 'empty'
        return fpga_active_image

    # Need to get FPGA Inactive Images on Console Server
    def fpga_inactive_images():
        fpga_inactive_images = 'empty'
        return fpga_inactive_images

    def machine_name():
        machine_name = json['machine_name']
        return machine_name



access(host_id, path_info)



















