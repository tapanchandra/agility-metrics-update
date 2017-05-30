import sys
import re

def ensure_standard_format(sprint,format_code):
    pattern = 'DI_DCS_([\d.]+)' if(sprint.startswith('DI')) else '([\d.]+)' 
    matcher = re.search(pattern,sprint,flags=0)

    if(matcher is None):
        print('ERROR')
        return None

    sprint_num = matcher.group(1) if(matcher is not None) else None

    tokens = sprint_num.split('.')

    year = int(tokens[0]) if (len(tokens[0]) == 4 and  tokens[0].isdigit() and tokens[0].startswith('20')) else int('20' + tokens[0]) if (len(tokens[0]) == 2 and tokens[0].isdigit()) else -1
    sprint_iteration = int(tokens[1]) if (tokens[1].isdigit()) else -1
    padded_sprint_iteration = '0' + str(sprint_iteration) if (sprint_iteration < 10) else str(sprint_iteration)

    if(year == -1 or sprint_iteration == -1):
        return None
    else:
        return str(year) + '.' + padded_sprint_iteration if format_code == 1 else 'DI_DCS_' + str(year)[2:4] + '.' + str(sprint_iteration) if format_code == 0 else str(year) + '.' + str(sprint_iteration)


def flushed_print(msg):
    print(msg)
    sys.stdout.flush()

def sort_list_by_id(sprint_list):
    #TODO : Sort all entries by sprint id
    return sprint_list