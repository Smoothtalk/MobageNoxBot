#pull one map
import requests
import re
import os

from bs4 import BeautifulSoup
from pathlib import Path

worlds = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12']

events = ['Visitors_Dyed_in_Red']

RE_EVENT_MAP_PATTERN = "[A-Z]-[0-9]"

def retrieveNodeMap(node_map_raw):
    local_map_rows = []
    map_rows = node_map_raw.find_all('tr')

    for map_row in map_rows:
        local_map_row = []

        header_columns = map_row.find_all('th')
        data_columns = map_row.find_all('td')

        for header in header_columns:
            if(header.string == None):
                local_map_row.append('--')
            else:
                local_map_row.append(header.string)

        for data in data_columns:
            if('DarkGreen' in data['style']):
                local_map_row.append('#')
            elif(data.img != None):
                if(data.img['alt'] == 'Fleet spawn'):
                    local_map_row.append('~')
                elif(data.img['alt'] == 'Enemy spawn'):
                    local_map_row.append('*')
                elif(data.img['alt'] == 'Resupply node'):
                    local_map_row.append('+')
                elif(data.img['alt'] == 'Boss spawn'):
                    local_map_row.append('X')
                elif(data.img['alt']== 'Secret node'):
                    local_map_row.append('?')
                # print(data.img['alt'])
            else:
                local_map_row.append(' ')

        local_map_rows.append(local_map_row)
    return local_map_rows

def retrieveNormalMapLayout(mapID):
    requestURL = "https://azurlane.koumakan.jp/" + mapID + "#Node%20map"

    page = requests.get(requestURL)
    soup = BeautifulSoup(page.content, 'html.parser')
    node_map_raw = soup.find('div', class_='tabbertab', title="Node map").find('table', class_='wikitable')
    local_node_map = retrieveNodeMap(node_map_raw)
    return local_node_map

def retrieveSpecialMapsLayout(eventName):
    # check for normal and hard mode maps
    # count how many maps in each mode
    # make a map for each node

    requestURL = "https://azurlane.koumakan.jp/Events/" + eventName + "#Node%20map"

    event_maps = []

    page = requests.get(requestURL)
    soup = BeautifulSoup(page.content, 'html.parser')
    have_normal_maps = soup.find(id='Normal_Maps')
    have_hard_maps = soup.find(id='Hard_Maps')
    if(have_normal_maps != None): #we have normal maps in this event
        local_normal_maps = []
        event_maps.append(['Normal Maps'])
        tabber_normal = have_normal_maps.find_next('div')
        normal_maps = tabber_normal.find_all('div', class_='tabbertab')
        for normal_map in normal_maps:
            if(re.match(RE_EVENT_MAP_PATTERN, normal_map['title'])):
                map_title = normal_map['title']
            if(normal_map['title'] == "Node map"):
                table_map = normal_map.find('table', class_='wikitable')
                local_normal_map = retrieveNodeMap(table_map)
                if(map_title != None):
                    local_normal_maps.append('(' + map_title.strip() + ')')
                local_normal_maps.append(local_normal_map)
        event_maps.append(local_normal_maps)
    if(have_hard_maps != None):
        local_hard_maps = []
        event_maps.append(['Hard Maps'])
        tabber_hard = have_hard_maps.find_next('div')
        hard_maps = tabber_hard.find_all('div', class_='tabbertab')
        for hard_map in hard_maps:
            if(re.match(RE_EVENT_MAP_PATTERN, hard_map['title'])):
                map_title = hard_map['title']
            if(hard_map['title'] == "Node map"):
                table_map = hard_map.find('table', class_='wikitable')
                local_hard_map = retrieveNodeMap(table_map)
                if(map_title != None):
                    local_hard_maps.append('(' + map_title.strip() + ')')
                local_hard_maps.append(local_hard_map)
        event_maps.append(local_hard_maps)
    return event_maps

for world in worlds:
    fileName = 'World-' + world + '.txt'
    world_file_name = Path("./maps/" + fileName)
    if(world_file_name.is_file()):
        print (str(world_file_name) + ' exists')
    else:
        world_file = open(world_file_name, 'a')
        for i in range(1,5):
            map = world + '-' + str(i)
            local_node_map = retrieveNormalMapLayout(map)
            world_file.write('(' + map + ')\n')
            for row in local_node_map:
                world_file.write(str(row) + '\n')
        world_file.close()

for event in events:
    local_event_maps = retrieveSpecialMapsLayout(event)
    fileName = event.replace('_', ' ') + '.txt'
    event_file_name = Path("./maps/" + fileName)
    if(event_file_name.is_file()):
        print (str(event_file_name) + ' exists')
    else:
        event_file = open(event_file_name, 'a')
        local_event_map = retrieveSpecialMapsLayout(event)
        for event_map in local_event_maps:
            for row in event_map:
                if('Maps' in str(row) or re.match('(.' + RE_EVENT_MAP_PATTERN + '.)', str(row))):
                    event_file.write(str(row) + '\n')
                else:
                    for col in row:
                        event_file.write(str(col) + '\n')
