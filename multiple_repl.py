import re

def multiple_repl(adict, text):
    def find_repl(mo):
        for i in adict:
            if re.search(i, mo.group(0)):
                return adict[i]
    regex = re.compile('|'.join(adict))
    return regex.sub(find_repl, text) 

if __name__ == '__main__':
    print('This script provides a function for multiple substitutions using regex')
