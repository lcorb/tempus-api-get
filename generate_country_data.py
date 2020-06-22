# import urllib.request
import statistics

import requests
import asyncio
from openpyxl import Workbook
from aiohttp import ClientSession

tempus = "https://tempus.xyz/api"
player_end = "/players/id"
rank_end = "/ranks/overall?start="
detailed_maps_end = "/maps/detailedList"
tt_end = "/maps/name"

valid_zones = ['map', 'course', 'bonus']

rank_increment = 50
rank_max = 10_000

id_increment = 200
id_max = 500_000

map_requests = {}
player_tts = {}
# author_names = {}
player_ids = []
player_stat_requests = []
country_stats = {}

def optional(f, arg, default):
    try:
        return f(arg)
    except:
        return default

def dict_optional(dictionary, attr, nested_attr, return_value, default = 0):

    if not attr in dictionary:
        return default

    return dictionary[attr][nested_attr][return_value] if nested_attr in dictionary[attr] else default

def get_rank_async(rank):
    try:
        r = requests.request(method='GET', url = f'{tempus}{rank_end}{rank}')      
        r.raise_for_status()
    except:
        return
    
    parsed = r.json()
    return parsed

async def get_player_async(id, session):
    try:
        r = await session.request(method='GET', url = f'{tempus}{player_end}/{id}/stats')
        r.raise_for_status()
    except:
        return
    
    parsed = await r.json()
    return parsed

def get_maps():
    try:
        r = requests.request(method='GET', url = f'{tempus}{detailed_maps_end}')
        r.raise_for_status()
    except:
        return
    
    return r.json()

async def get_tt_by_type_async(session, map_name, zone_type, zone_idx = 1):
    try:
        r = await session.request(method='GET', url = f'{tempus}{tt_end}/{map_name}/zones/typeindex/{zone_type}/{zone_idx}/records/list?limit=10')
        r.raise_for_status()
    except:
        return
    
    return await r.json()

async def parse_maps(map, session):
    for zone in map['zone_counts']:
        if zone in valid_zones:
            for i in range(1, map['zone_counts'][zone] + 1):
                times = await get_tt_by_type_async(session, map['name'], zone, i)
                for record in times['results']['soldier']:
                    if record['user_id'] not in player_tts:
                        player_tts[record['user_id']] = {
                            "map": {
                                "soldier": {"1": [],"2": [],"3": [],"4": [],"5": [],"6": []},
                                "demoman": {"1": [],"2": [],"3": [],"4": [],"5": [],"6": []}
                            },
                            "course": {
                                "soldier": {"1": [],"2": [],"3": [],"4": [],"5": [],"6": []},
                                "demoman": {"1": [],"2": [],"3": [],"4": [],"5": [],"6": []}
                            },
                            "bonus": {
                                "soldier": {"1": [],"2": [],"3": [],"4": [],"5": [],"6": []},
                                "demoman": {"1": [],"2": [],"3": [],"4": [],"5": [],"6": []}
                            }
                        }
                        player_tts[record['user_id']][times['zone_info']['type']]['soldier'][str(times['tier_info']['3'])] = [record['rank']]
                    else:
                        player_tts[record['user_id']][times['zone_info']['type']]['soldier'][str(times['tier_info']['3'])].append(record['rank'])

                for record in times['results']['demoman']:
                    if record['user_id'] not in player_tts:
                        player_tts[record['user_id']] = {
                            "map": {
                                "soldier": {"1": [],"2": [],"3": [],"4": [],"5": [],"6": []},
                                "demoman": {"1": [],"2": [],"3": [],"4": [],"5": [],"6": []}
                            },
                            "course": {
                                "soldier": {"1": [],"2": [],"3": [],"4": [],"5": [],"6": []},
                                "demoman": {"1": [],"2": [],"3": [],"4": [],"5": [],"6": []}
                            },
                            "bonus": {
                                "soldier": {"1": [],"2": [],"3": [],"4": [],"5": [],"6": []},
                                "demoman": {"1": [],"2": [],"3": [],"4": [],"5": [],"6": []}
                            }
                        }
                        player_tts[record['user_id']][times['zone_info']['type']]['demoman'][str(times['tier_info']['4'])] = [record['rank']]
                    else:
                        player_tts[record['user_id']][times['zone_info']['type']]['demoman'][str(times['tier_info']['4'])].append(record['rank'])

   
    # https://tempus.xyz/api/players/id/71541/author
    # >:(
    # for author in map.authors:
    #     if author in author_names:
    #         author_names[author] +=1
    #     else:
    #         author_names[author] = 1


async def parse_stats(stats):
        country = stats['player_info']['country']
        rank = stats['rank_info']['rank']
        points = stats['rank_info']['points']
        id = stats['player_info']['id']

        if id == 11020 or id == 29575:
            country = 'Australia'

        if country is None:
            return

        if points < 1:
            return

        # this is hell
        map_soldier_tt_tier_1_placement = []
        map_soldier_tt_tier_2_placement = []
        map_soldier_tt_tier_3_placement = []
        map_soldier_tt_tier_4_placement = []
        map_soldier_tt_tier_5_placement = []
        map_soldier_tt_tier_6_placement = []

        map_demoman_tt_tier_1_placement = []
        map_demoman_tt_tier_2_placement = []
        map_demoman_tt_tier_3_placement = []
        map_demoman_tt_tier_4_placement = []
        map_demoman_tt_tier_5_placement = []
        map_demoman_tt_tier_6_placement = []

        course_soldier_tt_tier_1_placement = []
        course_soldier_tt_tier_2_placement = []
        course_soldier_tt_tier_3_placement = []
        course_soldier_tt_tier_4_placement = []
        course_soldier_tt_tier_5_placement = []
        course_soldier_tt_tier_6_placement = []

        course_demoman_tt_tier_1_placement = []
        course_demoman_tt_tier_2_placement = []
        course_demoman_tt_tier_3_placement = []
        course_demoman_tt_tier_4_placement = []
        course_demoman_tt_tier_5_placement = []
        course_demoman_tt_tier_6_placement = []

        bonus_soldier_tt_tier_1_placement = []
        bonus_soldier_tt_tier_2_placement = []
        bonus_soldier_tt_tier_3_placement = []
        bonus_soldier_tt_tier_4_placement = []
        bonus_soldier_tt_tier_5_placement = []
        bonus_soldier_tt_tier_6_placement = []

        bonus_demoman_tt_tier_1_placement = []
        bonus_demoman_tt_tier_2_placement = []
        bonus_demoman_tt_tier_3_placement = []
        bonus_demoman_tt_tier_4_placement = []
        bonus_demoman_tt_tier_5_placement = []
        bonus_demoman_tt_tier_6_placement = []

        if id in player_tts:
            for zone in player_tts[id]:
                for tf2_class in player_tts[id][zone]:
                    for tier in player_tts[id][zone][tf2_class]:
                        for placement in player_tts[id][zone][tf2_class][tier]:
                            eval(f"{zone}_{tf2_class}_tt_tier_{tier}_placement").append(placement)

        map_pr_count = dict_optional(stats, 'pr_stats', 'map', 'count')
        course_pr_count = dict_optional(stats, 'pr_stats', 'course', 'count')
        bonus_pr_count = dict_optional(stats, 'pr_stats', 'bonus', 'count')

        map_pr_percentage = 0
        course_pr_percentage = 0
        bonus_pr_percentage = 0

        # 6 maps are impossible as soldier
        if 'map' in stats['pr_stats']:
            map_pr_percentage = (map_pr_count / ((stats['zone_count']['map']['count'] * 2) - 6)) * 100
        
        # 1 course
        if 'course' in stats['pr_stats']:
            course_pr_percentage = (course_pr_count / ((stats['zone_count']['course']['count'] * 2) - 1)) * 100
        
        # 7 (???) bonuses
        if 'bonus' in stats['pr_stats']:
            bonus_pr_percentage = (bonus_pr_count / ((stats['zone_count']['bonus']['count'] * 2) - 7)) * 100

        map_pr_points = dict_optional(stats, 'pr_stats', 'map', 'points')
        course_pr_points = dict_optional(stats, 'pr_stats', 'course', 'points')
        bonus_pr_points = dict_optional(stats, 'pr_stats', 'bonus', 'count')
     
        map_top_count = dict_optional(stats, 'top_stats', 'map', 'count')
        course_top_count = dict_optional(stats, 'top_stats', 'course', 'count')
        bonus_top_count = dict_optional(stats, 'top_stats', 'bonus', 'count')

        map_top_points = dict_optional(stats, 'top_stats', 'map', 'points')
        course_top_points = dict_optional(stats, 'top_stats', 'course', 'points')
        bonus_top_points = dict_optional(stats, 'top_stats', 'bonus', 'points')

        map_wr_count = dict_optional(stats, 'wr_stats', 'map', 'count')
        course_wr_count = dict_optional(stats, 'wr_stats', 'course', 'count')
        bonus_wr_count = dict_optional(stats, 'wr_stats', 'bonus', 'count')

        map_wr_points = dict_optional(stats, 'wr_stats', 'map', 'points')
        course_wr_points = dict_optional(stats, 'wr_stats', 'course', 'points')
        bonus_wr_points = dict_optional(stats, 'wr_stats', 'bonus', 'points')

        if country in country_stats:
            country_stats[country]['ranks'].append(rank)
            country_stats[country]['points'].append(points)
            country_stats[country]["map_pr_count"].append(map_pr_count)
            country_stats[country]["course_pr_count"].append(course_pr_count)
            country_stats[country]["bonus_pr_count"].append(bonus_pr_count)
            country_stats[country]["map_pr_percentage"].append(map_pr_percentage)
            country_stats[country]["course_pr_percentage"].append(course_pr_percentage)
            country_stats[country]["bonus_pr_percentage"].append(bonus_pr_percentage)
            country_stats[country]["map_pr_points"].append(map_pr_points)
            country_stats[country]["course_pr_points"].append(course_pr_points)
            country_stats[country]["bonus_pr_points"].append(bonus_pr_points)
            country_stats[country]["map_top_count"].append(map_top_count)
            country_stats[country]["course_top_count"].append(course_top_count)
            country_stats[country]["bonus_top_count"].append(bonus_top_count)
            country_stats[country]["map_top_points"].append(map_top_points)
            country_stats[country]["course_top_points"].append(course_top_points)
            country_stats[country]["bonus_top_points"].append(bonus_top_points)
            country_stats[country]["map_wr_count"].append(map_wr_count)
            country_stats[country]["course_wr_count"].append(course_wr_count)
            country_stats[country]["bonus_wr_count"].append(bonus_wr_count)
            country_stats[country]["map_wr_points"].append(map_wr_points)
            country_stats[country]["course_wr_points"].append(course_wr_points)
            country_stats[country]["bonus_wr_points"].append(bonus_wr_points)
            country_stats[country]['map_soldier_tt_tier_1_placement'] += map_soldier_tt_tier_1_placement
            country_stats[country]['map_soldier_tt_tier_2_placement'] += map_soldier_tt_tier_2_placement
            country_stats[country]['map_soldier_tt_tier_3_placement'] += map_soldier_tt_tier_3_placement
            country_stats[country]['map_soldier_tt_tier_4_placement'] += map_soldier_tt_tier_4_placement
            country_stats[country]['map_soldier_tt_tier_5_placement'] += map_soldier_tt_tier_5_placement
            country_stats[country]['map_soldier_tt_tier_6_placement'] += map_soldier_tt_tier_6_placement

            country_stats[country]['map_demoman_tt_tier_1_placement'] += map_demoman_tt_tier_1_placement
            country_stats[country]['map_demoman_tt_tier_2_placement'] += map_demoman_tt_tier_2_placement
            country_stats[country]['map_demoman_tt_tier_3_placement'] += map_demoman_tt_tier_3_placement
            country_stats[country]['map_demoman_tt_tier_4_placement'] += map_demoman_tt_tier_4_placement
            country_stats[country]['map_demoman_tt_tier_5_placement'] += map_demoman_tt_tier_5_placement
            country_stats[country]['map_demoman_tt_tier_6_placement'] += map_demoman_tt_tier_6_placement
    
            country_stats[country]['course_soldier_tt_tier_1_placement'] += course_soldier_tt_tier_1_placement
            country_stats[country]['course_soldier_tt_tier_2_placement'] += course_soldier_tt_tier_2_placement
            country_stats[country]['course_soldier_tt_tier_3_placement'] += course_soldier_tt_tier_3_placement
            country_stats[country]['course_soldier_tt_tier_4_placement'] += course_soldier_tt_tier_4_placement
            country_stats[country]['course_soldier_tt_tier_5_placement'] += course_soldier_tt_tier_5_placement
            country_stats[country]['course_soldier_tt_tier_6_placement'] += course_soldier_tt_tier_6_placement
    
            country_stats[country]['course_demoman_tt_tier_1_placement'] += course_demoman_tt_tier_1_placement
            country_stats[country]['course_demoman_tt_tier_2_placement'] += course_demoman_tt_tier_2_placement
            country_stats[country]['course_demoman_tt_tier_3_placement'] += course_demoman_tt_tier_3_placement
            country_stats[country]['course_demoman_tt_tier_4_placement'] += course_demoman_tt_tier_4_placement
            country_stats[country]['course_demoman_tt_tier_5_placement'] += course_demoman_tt_tier_5_placement
            country_stats[country]['course_demoman_tt_tier_6_placement'] += course_demoman_tt_tier_6_placement

            country_stats[country]['bonus_soldier_tt_tier_1_placement'] += bonus_soldier_tt_tier_1_placement
            country_stats[country]['bonus_soldier_tt_tier_2_placement'] += bonus_soldier_tt_tier_2_placement
            country_stats[country]['bonus_soldier_tt_tier_3_placement'] += bonus_soldier_tt_tier_3_placement
            country_stats[country]['bonus_soldier_tt_tier_4_placement'] += bonus_soldier_tt_tier_4_placement
            country_stats[country]['bonus_soldier_tt_tier_5_placement'] += bonus_soldier_tt_tier_5_placement
            country_stats[country]['bonus_soldier_tt_tier_6_placement'] += bonus_soldier_tt_tier_6_placement
    
            country_stats[country]['bonus_demoman_tt_tier_1_placement'] += bonus_demoman_tt_tier_1_placement
            country_stats[country]['bonus_demoman_tt_tier_2_placement'] += bonus_demoman_tt_tier_2_placement
            country_stats[country]['bonus_demoman_tt_tier_3_placement'] += bonus_demoman_tt_tier_3_placement
            country_stats[country]['bonus_demoman_tt_tier_4_placement'] += bonus_demoman_tt_tier_4_placement
            country_stats[country]['bonus_demoman_tt_tier_5_placement'] += bonus_demoman_tt_tier_5_placement
            country_stats[country]['bonus_demoman_tt_tier_6_placement'] += bonus_demoman_tt_tier_6_placement
        else:
            total_ranked = stats['country_rank_info']['total_ranked']
            country_stats[country] = {
                "total_ranked": total_ranked,
                "ranks": [rank],
                "points": [points],
                "map_pr_count": [map_pr_count],
                "course_pr_count": [course_pr_count],
                "bonus_pr_count": [bonus_pr_count],
                "map_pr_percentage": [map_pr_percentage],
                "course_pr_percentage": [course_pr_percentage],
                "bonus_pr_percentage": [bonus_pr_percentage],
                "map_pr_points": [map_pr_points],
                "course_pr_points": [course_pr_points],
                "bonus_pr_points": [bonus_pr_points],
                "map_top_count": [map_top_count],
                "course_top_count": [course_top_count],
                "bonus_top_count": [bonus_top_count],
                "map_top_points": [map_top_points],
                "course_top_points": [course_top_points],
                "bonus_top_points": [bonus_top_points],
                "map_wr_count": [map_wr_count],
                "course_wr_count": [course_wr_count],
                "bonus_wr_count": [bonus_wr_count],
                "map_wr_points": [map_wr_points],
                "course_wr_points": [course_wr_points],
                "bonus_wr_points": [bonus_wr_points],
                "map_soldier_tt_tier_1_placement": map_soldier_tt_tier_1_placement,
                "map_soldier_tt_tier_2_placement": map_soldier_tt_tier_2_placement,
                "map_soldier_tt_tier_3_placement": map_soldier_tt_tier_3_placement,
                "map_soldier_tt_tier_4_placement": map_soldier_tt_tier_4_placement,
                "map_soldier_tt_tier_5_placement": map_soldier_tt_tier_5_placement,
                "map_soldier_tt_tier_6_placement": map_soldier_tt_tier_6_placement,

                "map_demoman_tt_tier_1_placement": map_demoman_tt_tier_1_placement,
                "map_demoman_tt_tier_2_placement": map_demoman_tt_tier_2_placement,
                "map_demoman_tt_tier_3_placement": map_demoman_tt_tier_3_placement,
                "map_demoman_tt_tier_4_placement": map_demoman_tt_tier_4_placement,
                "map_demoman_tt_tier_5_placement": map_demoman_tt_tier_5_placement,
                "map_demoman_tt_tier_6_placement": map_demoman_tt_tier_6_placement,
        
                "course_soldier_tt_tier_1_placement": course_soldier_tt_tier_1_placement,
                "course_soldier_tt_tier_2_placement": course_soldier_tt_tier_2_placement,
                "course_soldier_tt_tier_3_placement": course_soldier_tt_tier_3_placement,
                "course_soldier_tt_tier_4_placement": course_soldier_tt_tier_4_placement,
                "course_soldier_tt_tier_5_placement": course_soldier_tt_tier_5_placement,
                "course_soldier_tt_tier_6_placement": course_soldier_tt_tier_6_placement,
        
                "course_demoman_tt_tier_1_placement": course_demoman_tt_tier_1_placement,
                "course_demoman_tt_tier_2_placement": course_demoman_tt_tier_2_placement,
                "course_demoman_tt_tier_3_placement": course_demoman_tt_tier_3_placement,
                "course_demoman_tt_tier_4_placement": course_demoman_tt_tier_4_placement,
                "course_demoman_tt_tier_5_placement": course_demoman_tt_tier_5_placement,
                "course_demoman_tt_tier_6_placement": course_demoman_tt_tier_6_placement,

                "bonus_soldier_tt_tier_1_placement": bonus_soldier_tt_tier_1_placement,
                "bonus_soldier_tt_tier_2_placement": bonus_soldier_tt_tier_2_placement,
                "bonus_soldier_tt_tier_3_placement": bonus_soldier_tt_tier_3_placement,
                "bonus_soldier_tt_tier_4_placement": bonus_soldier_tt_tier_4_placement,
                "bonus_soldier_tt_tier_5_placement": bonus_soldier_tt_tier_5_placement,
                "bonus_soldier_tt_tier_6_placement": bonus_soldier_tt_tier_6_placement,
        
                "bonus_demoman_tt_tier_1_placement": bonus_demoman_tt_tier_1_placement,
                "bonus_demoman_tt_tier_2_placement": bonus_demoman_tt_tier_2_placement,
                "bonus_demoman_tt_tier_3_placement": bonus_demoman_tt_tier_3_placement,
                "bonus_demoman_tt_tier_4_placement": bonus_demoman_tt_tier_4_placement,
                "bonus_demoman_tt_tier_5_placement": bonus_demoman_tt_tier_5_placement,
                "bonus_demoman_tt_tier_6_placement": bonus_demoman_tt_tier_6_placement
            }
    

def get_rank_wrapper(id):
    ranks = get_rank_async(id)
    if ranks is not None:
        for player in ranks['players']:
            player_ids.append(player['id'])

async def get_player_wrapper(id, session):
    r = await get_player_async(id, session)
    if r is not None:
        player_stat_requests.append(r)

# async def async_rank_request_wrapper(start, max):
#     async with ClientSession() as session:
#         await asyncio.gather(*[get_rank_wrapper(rank, session) for rank in range(start, max)])

async def async_stats_request_wrapper(start, max):
    async with ClientSession() as session:
        await asyncio.gather(*[get_player_wrapper(player_ids[id], session) for id in range(start, len(player_ids))])


async def async_stats_parse_wrapper():
    await asyncio.gather(*[parse_stats(stats) for stats in player_stat_requests])

async def async_tt_wrapper(maps):
    async with ClientSession() as session:
        await asyncio.gather(*[parse_maps(map, session) for map in maps])

def generate_data(country, stats):
    
    tt_stats = []

    for zone in ['map', 'course', 'bonus']:
        for tf2_class in ['soldier', 'demoman']:
            for tier in ['1', '2', '3', '4', '5', '6']:
                for stat in ['median', 'mean', 'mode']:
                    tt_stats.append(
                        eval(f'optional(statistics.{stat}, stats["{zone}_{tf2_class}_tt_tier_{tier}_placement"], 0)')
                    )
        
    # delete
    return [
        country,
        stats['total_ranked'],
        len(stats['ranks']),
        statistics.mean(stats['ranks']),
        statistics.median(stats['ranks']),
        min(stats['ranks']),
        sum(stats['points']),
        statistics.mean(stats['points']),
        sum(stats['map_pr_count']),
        statistics.mean(stats['map_pr_percentage']),
        sum(stats['map_pr_points']),
        statistics.mean(stats['map_pr_count']),
        statistics.mean(stats['map_pr_points']),
        sum(stats['course_pr_count']),
        statistics.mean(stats['course_pr_percentage']),
        sum(stats['course_pr_points']),
        statistics.mean(stats['course_pr_count']),
        statistics.mean(stats['course_pr_points']),
        sum(stats['bonus_pr_count']),
        statistics.mean(stats['bonus_pr_percentage']),
        sum(stats['bonus_pr_points']),
        statistics.mean(stats['bonus_pr_count']),
        statistics.mean(stats['course_pr_points']),
        sum(stats['map_top_count']),
        sum(stats['map_top_points']),
        statistics.mean(stats['map_top_count']),
        statistics.mean(stats['map_top_points']),
        sum(stats['course_top_count']),
        sum(stats['course_top_points']),
        statistics.mean(stats['course_top_count']),
        statistics.mean(stats['course_top_points']),
        sum(stats['bonus_top_count']),
        sum(stats['bonus_top_points']),
        statistics.mean(stats['bonus_top_count']),
        statistics.mean(stats['course_pr_points']),
        sum(stats['map_wr_count']),
        sum(stats['map_wr_points']),
        statistics.mean(stats['map_wr_count']),
        statistics.mean(stats['map_wr_points']),
        sum(stats['course_wr_count']),
        sum(stats['course_wr_points']),
        statistics.mean(stats['course_wr_count']),
        statistics.mean(stats['course_wr_points']),
        sum(stats['bonus_wr_count']),
        sum(stats['bonus_wr_points']),
        statistics.mean(stats['bonus_wr_count']),
        statistics.mean(stats['course_wr_points']),
    ] + tt_stats

def generate_xls():

    tt_stats_headers = []

    for zone in ['Map', 'Course', 'Bonus']:
        for tf2_class in ['Soldier', 'Demoman']:
            for tier in ['1', '2', '3', '4', '5', '6']:
                for stat in ['Median', 'Mean', 'Mode']:
                    tt_stats_headers.append(
                        f'{zone}_{tf2_class}_TT_Tier_{tier}_{stat}'
                    )

    headers = [
        'Country',
        'Total_Players',
        'Sample_Size',
        'Average_Rank',
        'Median_Rank',
        'Highest_Rank',
        'Total_Points',
        'Average_Points',

        'Total_Completion_Map_Count',
        'Average_Completion_Map_Percentage',
        'Total_Completion_Map_Points',
        'Average_Completion_Map_Count',
        'Average_Completion_Map_Points',

        'Total_Completion_Course_Count',
        'Average_Completion_Course_Percentage',
        'Total_Completion_Course_Points',
        'Average_Completion_Course_Count',
        'Average_Completion_Course_Points',

        'Total_Completion_Bonus_Count',
        'Average_Completion_Bonus_Percentage',
        'Total_Completion_Bonus_Points',
        'Average_Completion_Bonus_Count',
        'Average_Completion_Bonus_Points',


        'Total_Top_Map_Count',
        'Total_Top_Map_Points',
        'Average_Top_Map_Count',
        'Average_Top_Map_Points',

        'Total_Top_Course_Count',
        'Total_Top_Course_Points',
        'Average_Top_Course_Count',
        'Average_Top_Course_Points',

        'Total_Top_Bonus_Count',
        'Total_Top_Bonus_Points',
        'Average_Top_Bonus_Count',
        'Average_Top_Bonus_Points',


        'Total_Record_Map_Count',
        'Total_Record_Map_Points',
        'Average_Record_Map_Count',
        'Average_Record_Map_Points',

        'Total_Record_Course_Count',
        'Total_Record_Course_Points',
        'Average_Record_Course_Count',
        'Average_Record_Course_Points',

        'Total_Record_Bonus_Count',
        'Total_Record_Bonus_Points',
        'Average_Record_Bonus_Count',
        'Average_Record_Bonus_Points',
        ]

    headers += tt_stats_headers
    wb = Workbook()
    sheet = wb.active
    sheet.title = "Tempus Country Statistics"
    sheet.append(headers)

    for country in country_stats:
        sheet.append(generate_data(country, country_stats[country]))
    
    wb.save('country_data.xlsx')


def main():
    loop = asyncio.get_event_loop()

    map_requests = get_maps()
    loop.run_until_complete(async_tt_wrapper(map_requests))

    current_rank = 0
    current_max_rank = rank_increment

    for i in range(current_rank, rank_max, rank_increment):
        # should make this async too
        get_rank_wrapper(current_rank)
        loop.run_until_complete(async_stats_request_wrapper(current_rank, current_max_rank))

        current_rank += rank_increment
        current_max_rank =+ rank_increment

    loop.run_until_complete(async_stats_parse_wrapper())
    loop.close()
    generate_xls()


main()
