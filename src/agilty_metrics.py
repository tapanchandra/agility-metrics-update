from jira import JIRA
import base64
import os
import re
from oauth2client.service_account import ServiceAccountCredentials
import gspread
from jira import JIRA
from datetime import datetime as dt
import requests
from requests.packages.urllib3.exceptions import InsecureRequestWarning
from helper import sort_list_by_id,flushed_print,ensure_standard_format
from gspread.exceptions import WorksheetNotFound

#Suppress Insecure Request Warnings
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

resource_planning_workbook = 'Resource Planning'
agility_workbook = 'Agility Goal Tracking by Product Line'
agility_worksheet = 'Enterprise Arch'
agility_team_name = 'IF - Frontrunners'
jira_team_name = 'System'
starting_sprint = 'DI_DCS_17.06'
data_workbook_name = 'Data'
data_worksheet_name = 'Sprints'
ftr_agility_metrics_workbook_name = 'ftr-agility-metrics-entries'
ftr_agility_metrics_worksheet_name = 'entries'
stories_from_last_sprint = 'project = \"IF Integration\" AND Sprint = \"{curr_sprint}\" AND Sprint = \"{prev_sprint}\" AND \"Dev team\" = \"' + jira_team_name + '\" AND type = Story'
stories_committed_currentsprint = 'project = \"IF Integration\" AND Sprint = \"{curr_sprint}\" AND Sprint != \"{prev_sprint}\"  AND \"Dev team\" = \"' + jira_team_name + '\" AND type = Story'
stories_delivered_currentsprint = 'project = \"IF Integration\" AND Sprint = \"{curr_sprint}\" AND Sprint != \"{next_sprint}\"  AND \"Dev team\" = \"' + jira_team_name + '\" AND type = Story AND status = \"Closed\"'

jira_format = 0
data_sheet_format = 1
agility_sheet_format = 2

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

def get_googlesheet_workbook_instance(workbook_name):
    scope = ['https://spreadsheets.google.com/feeds']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('src\google_token.json',scope)
    gclient = gspread.authorize(credentials)
    workbook_instance = gclient.open(workbook_name)  

    return workbook_instance

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
        
        if todays_date <= to_date:
            break

        if(compare_sprint_ids(sprint_id,starting_sprint) >= 0):
            sprint_list.append([row_data[0],row_data[1],row_data[2]])

        
        row_index += 1
    #Append prefix to all sprint ids
    sprint_list = [[ensure_standard_format(x[0],data_sheet_format),x[1],x[2]] for x in sprint_list] 

    #Sort sprint list by sprint index   
    sprint_list = sort_list_by_id(sprint_list)
     
    return sprint_list

def get_sprint_rows_by_team(metrics_sheet,agility_team_name):

    #Map containing all existing sprint records and their row ids in the sheet
    sprints_row_map = {}

    #Get all records for the current team 
    team_records_list = metrics_sheet.findall(agility_team_name)
    team_entries_row_numbers = sorted([record.row for record in team_records_list])

    iterationnum_cells = metrics_sheet.range('B' + str(team_entries_row_numbers[0]) + ':B' + str(team_entries_row_numbers[-1]))
    
    for cell in iterationnum_cells:
        sprints_row_map[cell.value] = cell.row
    
    return sprints_row_map


def insert_sprint_details_googlesheet(ftr_agility_sheet,details_to_insert):

    global agility_team_name

    iteration_start_date = dt.strptime(details_to_insert[1],'%d-%b-%Y').strftime('%m/%d/%Y')
    iteration_end_date =  dt.strptime(details_to_insert[2],'%d-%b-%Y').strftime('%m/%d/%Y')
    
    ftr_agility_sheet.append_row(
        [agility_team_name,
        float(details_to_insert[0]),
        iteration_start_date,
        iteration_end_date,
        float(details_to_insert[1]),
        float(details_to_insert[2]),
        float(details_to_insert[3]),
        float(details_to_insert[4]),
        float(details_to_insert[5]),
        float(details_to_insert[6]),
        float(details_to_insert[7])])


def increment_sprint(current_sprint,sprint_list):
    if(current_sprint not in sprint_list):
        print('Sprint '+ current_sprint +' is not present in the sprint list')
        return None
    
    max_index = len(sprint_list) - 1
    current_sprint_index = sprint_list.index(current_sprint)
    return sprint_list[current_sprint_index + 1] if current_sprint_index < max_index else None

def decrement_sprint(current_sprint,sprint_list):
    if(current_sprint not in sprint_list):
        print('Sprint '+ current_sprint +' is not present in the sprint list')
        return None
    
    current_sprint_index = sprint_list.index(current_sprint)
    return sprint_list[current_sprint_index - 1] if current_sprint_index > 0 else None

def fetch_sprint_details(jira_current_sprint,sprint_list):

    jira_previous_sprint = decrement_sprint(jira_current_sprint,sprint_list)
    jira_next_sprint = increment_sprint(jira_current_sprint,sprint_list)

    #Get Stories carried over from last iteration
    jira_query = stories_from_last_sprint.format(curr_sprint=jira_current_sprint,prev_sprint=jira_previous_sprint)
    carried_over_stories = jira.search_issues(jira_query,maxResults=500)
    print('Carried over stories are : '+ ','.join([story.key for story in carried_over_stories]))
    no_of_carried_stories = len(carried_over_stories)
    

    #Get Stories committed this iteration
    jira_query = stories_committed_currentsprint.format(curr_sprint=jira_current_sprint,prev_sprint=jira_previous_sprint)
    committed_stories = jira.search_issues(jira_query,maxResults=500)
    print('Committed stories are : '+ ','.join([story.key for story in committed_stories]))
    no_of_committed_stories = len(committed_stories)
    total_points_committed = sum([ int(story.fields.customfield_10003) for story in committed_stories])

    #Get Stories delivered this iteration
    jira_query = stories_delivered_currentsprint.format(curr_sprint=jira_current_sprint,next_sprint=jira_next_sprint)
    delivered_stories = jira.search_issues(jira_query,maxResults=500)
    print('Delivered stories are : '+ ','.join([story.key for story in delivered_stories]))
    no_of_delivered_stories = len(delivered_stories)
    total_points_delivered = sum([ int(story.fields.customfield_10003) for story in delivered_stories])

    # Get the resource google sheet instance for the current sprint
    resource_current_sprint_worksheet = resource_workbook.worksheet(jira_current_sprint)

    # Get the planned staff of current iteration
    planned_staff = resource_current_sprint_worksheet.cell(32,4).value

    # Get the actual staff of current iteration
    actual_staff = resource_current_sprint_worksheet.cell(33,4).value

    print(jira_current_sprint,no_of_carried_stories,no_of_committed_stories,no_of_delivered_stories,total_points_committed,total_points_delivered,actual_staff,planned_staff)

    return [jira_current_sprint,no_of_carried_stories,no_of_committed_stories,no_of_delivered_stories,total_points_committed,total_points_delivered,actual_staff,planned_staff]

def recreate_sheet(workbook,worksheet):
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

    global agility_team_name

    #Get the sprint details till date starting from starting_sprint 
    sprint_list = get_required_sprints()
    flushed_print('Generated sprint_list')
    
    # Get the instance for the ftr-agility-metrics-entries workbook 
    ftr_agility_sheet = recreate_sheet(ftr_agility_metrics_workbook_name,ftr_agility_metrics_worksheet_name)

    # Fetch sprint list in jira format 
    jira_sprint_list = [ensure_standard_format(x[0],jira_format) for x in sprint_list]

    #For each sprint in sprint_list, generate a agility metrics row
    for sprint in sprint_list:

        sprint_num = ensure_standard_format(sprint[0],agility_sheet_format)
        
        flushed_print('Inserting a row at for ' + str(sprint_num))
        sprint_details = fetch_sprint_details(ensure_standard_format(sprint_num,jira_format),jira_sprint_list)
        
        insert_sprint_details_googlesheet(ftr_agility_sheet,sprint_details)
        


jira = get_jira_instance()
resource_workbook = get_googlesheet_workbook_instance(resource_planning_workbook)
if __name__ == '__main__':
    main()
    