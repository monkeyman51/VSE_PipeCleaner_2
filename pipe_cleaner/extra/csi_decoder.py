
system_name = 'S1545'

# system_name = system_name[3::2]
# print(system_name)

def decode_system(string):
    first_character = string[:1:]
    second_character = string[1::4]
    third_character = string[2::3]
    fourth_character = string[3::2]
    fifth_character = string[4::1]

    if first_character == 'C':
        print(f'Building Block Function: {first_character} in {string} = Compute')
    if first_character == 'S':
        print(f'Building Block Function: {first_character} in {string} = Storage')
    if first_character == 'J':
        print(f'Building Block Function: {first_character} in {string} = JBOD')
    if first_character == 'E':
        print(f'Building Block Function: {first_character} in {string} = Enclosing')
    if first_character == 'F':
        print(f'Building Block Function: {first_character} in {string} = Flash')

    if second_character == '1':
        print(f'Platform/ Rack Infrastructure: {second_character} in {string} = 1st Gen of MFST Rack Architecture')
    if second_character == '2':
        print(f'Platform/ Rack Infrastructure: {second_character} in {string} = Compute')

    if third_character == '1':
        print(f'Variation of Base Architecture: {third_character} in {string} = 1 SKU of Architecture')
    if third_character == '2':
        print(f'Variation of Base Architecture: {third_character} in {string} = 2 SKUs of Architecture')
    if third_character == '3':
        print(f'Variation of Base Architecture: {third_character} in {string} = 3 SKUs of Architecture')
    if third_character == '4':
        print(f'Variation of Base Architecture: {third_character} in {string} = 4 SKUs of Architecture')
    if third_character == '5':
        print(f'Variation of Base Architecture: {third_character} in {string} = 5 SKUs of Architecture')
    if third_character == '6':
        print(f'Variation of Base Architecture: {third_character} in {string} = 6 SKUs of Architecture')
    if third_character == '7':
        print(f'Variation of Base Architecture: {third_character} in {string} = 7 SKUs of Architecture')
    if third_character == '8':
        print(f'Variation of Base Architecture: {third_character} in {string} = 8 SKUs of Architecture')
    if third_character == '9':
        print(f'Variation of Base Architecture: {third_character} in {string} = 9 SKUs of Architecture')

    if fourth_character == '4':
        print(f'Base Architecture (Processor Generation): {fourth_character} in {string} = Haswell')
    if fourth_character == '5':
        print(f'Base Architecture (Processor Generation): {fourth_character} in {string} = Broadwell')
    if fourth_character == '6':
        print(f'Base Architecture (Processor Generation): {fourth_character} in {string} = Skylake')
    if fourth_character == '7':
        print(f'Base Architecture (Processor Generation): {fourth_character} in {string} = Cascade Lake')
    if fourth_character == '8':
        print(f'Base Architecture (Processor Generation): {fourth_character} in {string} = Coffee Lake')

    if fifth_character == '0':
        print(f'System Vendors (Processor Generation): {fifth_character} in {string} = Microsoft')
    if fifth_character == '5':
        print(f'System Vendors (Processor Generation): {fifth_character} in {string} = ZT Systems')

decode_system(system_name)