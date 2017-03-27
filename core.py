# -*- coding: utf-8 -*-
"""
Created on Tue Mar  7 20:19:44 2017

@author: beano
"""

import pandas as pd
import numpy as np
from openpyxl import load_workbook

def triangular(n):
    return sum(range(n+1))

# Define file variables
weekly_file = 'Squash_results.xlsx'
squash_path = 'C:\\Users\\beano\\Google Drive\\Squash\\'

# Define switches
week_num = 4
master = True

# Define columns for statistics
stats_columms = ['Name',\
                 'PL',\
                 'W',\
                 'L',\
                 'PF',\
                 'PA',\
                 'DIFF',\
                 'WB',\
                 'LB',\
                 'Points',\
                 'Normalised Score'
                 ]

# Determine if you are running for an individual week or for the all weeks
if master:
    # Read in all weeks results
    for week in range(1,week_num+1):
        tmp_games_df = pd.read_excel(squash_path + weekly_file, sheetname='Week{0}_games'.format(week))
        tmp_tables_df = pd.read_excel(squash_path + weekly_file, sheetname='Week{0}_table'.format(week))
        
        if week == 1:
            weekly_games_df = tmp_games_df
            tmp_tables_df['week'] = week
            weekly_tables_df = tmp_tables_df
        else:
            weekly_games_df = weekly_games_df.append(tmp_games_df)
            tmp_tables_df['week'] = week
            weekly_tables_df = weekly_tables_df.append(tmp_tables_df)
            #print(weekly_games_df)
else:
    # Read in weekly tab
    weekly_games_df = pd.read_excel(squash_path + weekly_file, sheetname='Week{0}_games'.format(week_num))

# Obtain list of players
players_tmp = weekly_games_df['Player 1'].unique().tolist() + weekly_games_df['Player 2'].unique().tolist()
players = list(set(players_tmp))

# Create empty dataframe with necessary columns
weekly_stats_df = pd.DataFrame(columns=stats_columms)

# Go through each player and retrieve their results
for player in players:
    # Determine games player participated in
    player_games_df = weekly_games_df.loc[(weekly_games_df['Player 1'] == player) | (weekly_games_df['Player 2'] == player)]
    
    # Determine games won by player
    player_wins_df = weekly_games_df.loc[((weekly_games_df['Player 1'] == player) &\
                                            (weekly_games_df['Score 1'] > weekly_games_df['Score 2']))\
                                            |
                                            ((weekly_games_df['Player 2'] == player) &\
                                            (weekly_games_df['Score 2'] > weekly_games_df['Score 1']))]
    
    # Obtain table of players weekly results
    player_tables_df = weekly_tables_df.loc[(weekly_tables_df['Name'] == player)]
    
    # Obtain points for and points against
    PF = 0
    PA = 0
    WB = 0
    LB = 0
    
    # Determine points
    for _, game in player_games_df.iterrows():
        # Points for and against
        sc1_tmp = game['Score 1']
        sc2_tmp = game['Score 2']
        if game['Player 1'] == player:
            PF += sc1_tmp
            PA += sc2_tmp
            diff_tmp = sc1_tmp - sc2_tmp
        elif game['Player 2'] == player:
            PF += sc2_tmp
            PA += sc1_tmp
            diff_tmp = sc2_tmp - sc1_tmp
        # Bonus points
        if diff_tmp >= 7:
            WB += 1
        elif diff_tmp >= -3 and diff_tmp < 0:
            LB += 1
    
    # Metrics for table
    PL = player_games_df.shape[0]
    W = player_wins_df.shape[0]
    L = PL - W
    DIFF = PF - PA
    BP = 0
    Points = W*3 + WB + LB
    
    # Determine whether to use normalised score (for weekly table) or weighted score (master table)
    normalised_score = 0
    if master == True:
        for _, weekly_result in player_tables_df.iterrows():
            wk = weekly_result['week']
            n_score = weekly_result['Normalised Score']            
            normalised_score += int(n_score*wk/triangular(week_num))
    else:
        normalised_score = int((Points/PL)*(100.0/4.0))
    
    # Create row for player data to be added to the dataframe    
    player_data = [player, PL, W, L, PF, PA, DIFF, WB, LB, Points, normalised_score]
    
    #Create dataframe for addition to main dataframe
    player_df = pd.DataFrame(np.array([player_data]), columns=stats_columms)
    weekly_stats_df = weekly_stats_df.append(player_df)       

# Convert all number columns to integers    
for col in weekly_stats_df.columns[1:]:
    weekly_stats_df[col] = weekly_stats_df[col].astype(int)
  
# Sort values by points
weekly_stats_df = weekly_stats_df.sort_values(['Normalised Score', 'DIFF'], ascending=False)
print(weekly_stats_df)

# Write player_database to excel document
book = load_workbook(squash_path + weekly_file)
writer = pd.ExcelWriter(squash_path + weekly_file, engine='openpyxl') 
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

if master:
    weekly_stats_df.to_excel(writer, sheet_name='Master_table', index=False)
else:
    weekly_stats_df.to_excel(writer, sheet_name='Week{0}_table'.format(week_num), index=False)

writer.save()