from jira import JIRA
import base64
import os
import re
from oauth2client.service_account import ServiceAccountCredentials
import gspread
from jira import JIRA
from datetime import datetime as dt
import sys
import requests
from requests.packages.urllib3.exceptions import InsecureRequestWarning

#Suppress Insecure Request Warnings
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

agility_workbook = 'Fitnesse - JIRA'
agility_worksheet = 'Enterprise Arch'
team_name = 'IF - Frontrunners'
starting_sprint = 'DI_DCS_17.06'
data_workbook_name = 'Data'
data_worksheet_name = 'Sprints'
ftr_agility_metrics_workbook = 'ftr-agility-metrics-entries'
ftr_agility_metrics_worksheet = 'entries'
stories_from_last_sprint = 'project = \"IF Integration\" AND Sprint = \"{0}\" AND Sprint = \"{1}\" AND \"Dev team\" = \"System\"'
stories_committed_currentsprint = 'project = \"IF Integration\" AND Sprint = \"{0}\" AND \"Dev team\" = \"System\"'
stories_delivered_currentsprint = 'project = \"IF Integration\" AND Sprint = \"{0}\" AND \"Dev team\" = \"System\" AND status = \"Closed\"'


# TO DO find the current sprint  
# https://docs.google.com/spreadsheets/d/1n65r7nhj37TxV4rrRh-0sf0zm41vBacgSEIDimFdgiI/edit#gid=439243553 -> SPrints sheet 
# TO DO verify from starting_sprint values exist in the agility sheet and 
# append the new sprint based on sysdate between FromData and ToDate from above sheet; 
# get the Sprint and change it to 2017.xx format 


def get_jira_instance():
    username = base64.b64decode( os.environ['AUTH'].encode( 'utf-8' ) ).decode( 'utf-8' ).split(':')[0].strip() 
    password = base64.b64decode( os.environ['AUTH'].encode( 'utf-8' ) ).decode( 'utf-8' ).split(':')[1].strip()
    
    jira = JIRA(basic_auth=(username,password),options = {
    'server' : 'https://jira.jda.com',
    'verify' : False,
    })
    return jira

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
        return str(year) + '.' + padded_sprint_iteration if format_code == 1 else 'DI_DCS_' + str(year)[2:4] + '.' + str(sprint_iteration)

    # jira_format = 0
    # gsheet_format = 1
    # print(ensure_standard_format('DI_DCS_2017.06',gsheet_format))
    # print(ensure_standard_format('2017.06',gsheet_format))

    # print(ensure_standard_format('2017.6',jira_format))

def flushed_print(msg):
    print(msg)
    sys.stdout.flush()

def sort_list_by_id(sprint_list):
    #TODO : Sort all entries by sprint id
    return sprint_list

def compare_sprint_ids(first_sprint,second_sprint):
    first_sprint_num = float(first_sprint[7:])
    second_sprint_num = float(second_sprint[7:])
    
    return 0 if(first_sprint_num == second_sprint_num) else -1 if (first_sprint_num < second_sprint_num) else 1

def get_google_sheet_instance(workbook_name,worksheet_name):
    scope = ['https://spreadsheets.google.com/feeds']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('src\google_token.json',scope)
    gclient = gspread.authorize(credentials)
    workbook_instance = gclient.open(workbook_name)  
    worksheet_instance = workbook_instance.worksheet(worksheet_name)

    return workbook_instance,worksheet_instance

def get_required_sprints():
    
    data_workbook,data_worksheet = get_google_sheet_instance(data_workbook_name,data_worksheet_name)
    todays_date = dt.strptime(dt.now().strftime('%d-%b-%Y'),'%d-%b-%Y')

    sprint_list = []
    
    #Start from 2nd Row as 1st is a header
    max_iterations = 500
    row_index = 2

    while row_index > 0 and row_index < max_iterations:

        row_data = data_worksheet.row_values(row_index)
        sprint_id = row_data[0]
        from_date = dt.strptime(row_data[1],'%d-%b-%Y')
        to_date = dt.strptime(row_data[2],'%d-%b-%Y')
        
        if(compare_sprint_ids(sprint_id,starting_sprint) >= 0):
            sprint_list.append([row_data[0],row_data[1],row_data[2]])

        if from_date <= todays_date <= to_date:
            break

        row_index += 1
    
    print(sprint_list)
    #Append prefix to all sprint ids
    sprint_list = [[ensure_standard_format(x[0],1),x[1],x[2]] for x in sprint_list] 

    #Sort sprint list by sprint index   
    sprint_list = sort_list_by_id(sprint_list)
    print(sprint_list)
    exit()
    
    return sprint_list

def get_sprint_rows_by_team(metrics_sheet,team_name):

    #Map containing all existing sprint records and their row ids in the sheet
    sprints_row_map = {}

    #Get all records for the current team 
    team_records_list = metrics_sheet.findall(team_name)
    team_entries_row_numbers = sorted([record.row for record in team_records_list])

    iterationnum_cells = metrics_sheet.range('B' + str(team_entries_row_numbers[0]) + ':B' + str(team_entries_row_numbers[-1]))
    
    for cell in iterationnum_cells:
        sprints_row_map[cell.value] = cell.row
    
    return sprints_row_map

def insert_sprint_details_googlesheet(ftr_agility_sheet,details_to_insert,insert_at_row):

    global team_name

    # Insert the missing and current sprint details into google sheet
    # flushed_print([team_name,details_to_insert[0],details_to_insert[1],details_to_insert[2]])
    # metrics_sheet.insert_row([team_name,details_to_insert[0]],insert_at_row)
    
    iteration_start_date = dt.strptime(details_to_insert[1],'%d-%b-%Y').strftime('%m/%d/%Y')
    iteration_end_date =  dt.strptime(details_to_insert[2],'%d-%b-%Y').strftime('%m/%d/%Y')
    
    ftr_agility_sheet.append_row([insert_at_row,team_name,str(details_to_insert[0]),iteration_start_date,iteration_end_date])


def fetch_sprint_details(sprint_name,sprint_ids):

    jira = get_jira_instance()
    
    jira_current_sprint = ensure_standard_format('2017.',0)
    sprint_stories = jira.search_issues(stories_delivered_currentsprint.format(jira_current_sprint),maxResults=500)

    print(sprint_stories)


    # for story_id in sprint_stories:
    #     if(story_id.key == 'DCS-14746'):
    #         print('story points : ' + str(story_id.fields.customfield_10003))
    #         print('Story Status : ' + story_id.fields.status.name)
    #         sprint_list = story_id.fields.customfield_10006

    #         for sprint in sprint_list:
    #             print(sprint)

    #         exit()
    
    return sprint

def refresh_sheet(workbook,worksheet):
    #Delete the JIRA Problems sheet

    try:
        wbook,wsheet = get_google_sheet_instance(workbook,worksheet)
        if(wsheet is not None):
            wbook.del_worksheet(wsheet)    
    except WorksheetNotFound as ex:
        pass

    #Create new sheet
    return wbook.add_worksheet(worksheet,1,50)

def main():

    global team_name

    #Read agility metrics sheet
    metrics_workbook,metrics_sheet = get_google_sheet_instance(agility_workbook,agility_worksheet)
    flushed_print('got metrics sheet')

    #Get the sprint details till date starting from starting_sprint 
    sprint_list = get_required_sprints()
    flushed_print('Generated sprint_list')

    #Get all rows for my team from the agility metrics sheet    
    sprints_row_map = get_sprint_rows_by_team(metrics_sheet,team_name)
    flushed_print('Created sprints_row_map')

    #get the row number for the starting sprint in the sgility metrics sheet
    start_sprint_row_num = -1 if ensure_standard_format(starting_sprint,1) not in sprints_row_map.keys() else sprints_row_map[ensure_standard_format(starting_sprint,1)]
    flushed_print('Starting sprint row num is ' + str(start_sprint_row_num))

    #Abort if start sprint is not present in the sheet
    if(start_sprint_row_num == -1):
        print('[ERROR] : The specified start sprint was not present in the google sheet. The start sprint id should already have been present in the sheet.')
        exit()
    
    # Get the instance for the ftr-agility-metrics-entries workbook 
    ftr_agility_sheet = refresh_sheet(ftr_agility_metrics_workbook,ftr_agility_metrics_worksheet)

    #For each sprint in sprint_list, check if there exists a row at `sprint_row_num` row in agility metrics sheet
    sprint_row_num = start_sprint_row_num    
    for sprint in sprint_list:
        #verify if there is a record with this sprint id at this row number
        flushed_print('Looking for ' + sprint[0] + ' at ' + str(sprint_row_num))
        iteration_num = metrics_sheet.cell(sprint_row_num,2).value
        
        if iteration_num != sprint[0]:
            #If missing , add a row
            flushed_print('Inserting a row at ' + str(sprint_row_num))
            sprint_details = fetch_sprint_details(sprint[0])

            
            insert_sprint_details_googlesheet(ftr_agility_sheet,sprint_details,sprint_row_num)
            sprint_row_num -= 1

        
        #Check for next sprint in sprint_list
        sprint_row_num += 1

if __name__ == '__main__':
    # main()
    # fetch_sprint_details()
    get_required_sprints()