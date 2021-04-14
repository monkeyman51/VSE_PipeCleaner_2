from random import choice


def get_joke():
    jokes = [
        '\tI hear that if you wear a mask and glasses,\n'
        '\tthat you may be eligible for condensation. - Jack Zornes',
        
        '\tWhy do elephants paint their toe nails red?\n'
        '\tTo hide in the apple tree...\n'
        '\tHave you ever seen an elephant in an apple tree?\n'
        '\tWorks doesnt it? - James Pinto',

        '\tHOW TO TRANSLATE WORK EMAILS:\n'
        '\t\t- "I have a question." actually means... "I have 56 questions."\n'
        '\t\t- "Do you have a minute?" actually means... "Do you have an hour?"\n'
        '\t\t- "Swap out a DIMM." actually means... "Blade suddenly not working. Debug the entire pipe."\n'
        '\t\t- "Whats wrong with this Gen 5 blade?" actually means... "Hope its not the Air Max connector."\n'
        '\t\t- "Troubleshoot this blade." actually means... "Restart the computer."\n'
        '\t\t- "What made the blade have a red dot in Console Server?" actually means... \n'
        '\t\tLiterally ANYTHING...',
        
        '\tI went to a restaurant that serves "Breakfast at Any Time".\n'
        '\tSo I ordered French Toast during the Renaissance. - Christopher Smith / Internet'
    ]
    return choice(jokes)
