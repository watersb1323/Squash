# -*- coding: utf-8 -*-
"""
Created on Tue Mar  7 20:19:44 2017

@author: beano
"""

import pandas as pd
import numpy as np
from openpyxl import load_workbook

weekly_file = 'Squash_results.xlsx'
squash_path = 'C:\\Users\\beano\\Google Drive\\Squash\\'

# Read in weekly tab
weekly_scores_df = pd.read_excel(squash_path + weekly_file, sheetname='Week1_games')

players_tmp = weekly_scores_df['Player 1'].unique().tolist() + weekly_scores_df['Player 2'].unique().tolist()
players = list(set(players_tmp))

stats_columms = ['Name',\
                 'PL',\
                 'W',\
                 'L',\
                 'PF',\
                 'PA',\
                 'DIFF',\
                 'WB',\
                 'LB',\
                 'Points'
                 ]
week_num = '1'

weekly_stats_df = pd.DataFrame(columns=stats_columms)

# Go through each player and retrieve their results
for player in players:
    # Determine games player participated in
    player_games_df = weekly_scores_df.loc[(weekly_scores_df['Player 1'] == player) | (weekly_scores_df['Player 2'] == player)]
    
    # Determine games won by player
    player_wins_df = weekly_scores_df.loc[((weekly_scores_df['Player 1'] == player) &\
                                            (weekly_scores_df['Score 1'] > weekly_scores_df['Score 2']))\
                                            |
                                            ((weekly_scores_df['Player 2'] == player) &\
                                            (weekly_scores_df['Score 2'] > weekly_scores_df['Score 1']))]
    
    # Obtain points for and points against
    PF = 0
    PA = 0
    WB = 0
    LB = 0
    
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
    player_data = [player, PL, W, L, PF, PA, DIFF, WB, LB, Points]
    
    #Create dataframe for addition to main dataframe
    player_df = pd.DataFrame(np.array([player_data]), columns=stats_columms)
    weekly_stats_df = weekly_stats_df.append(player_df)       

# Convert all number columns to integers    
for col in weekly_stats_df.columns[1:]:
    weekly_stats_df[col] = weekly_stats_df[col].astype(int)
  
# Sort values by points
weekly_stats_df = weekly_stats_df.sort_values('Points', ascending=False)
print(weekly_stats_df)

# Write player_database to excel document
book = load_workbook(squash_path + weekly_file)
writer = pd.ExcelWriter(squash_path + weekly_file, engine='openpyxl') 
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

weekly_stats_df.to_excel(writer, sheet_name='Week' + week_num + '_table', index=False)
writer.save()