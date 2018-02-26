# -*- coding: utf-8 -*-
"""
Created on Tue Mar  7 20:19:44 2017

@author: Brian Waters
"""

import pandas as pd
import numpy as np
from openpyxl import load_workbook

def triangular(n):
    return sum(range(n+1))

# Define file variables
weekly_file = 'Squash_results_2018.xlsx'
# squash_path = 'C:\\Users\\beano\\Google Drive\\Squash\\'
squash_path = 'C:\\Users\\brwaters\\Documents\\GitHub\\Squash\\'

# Define switches
week_num = 1
master = 0

# Define Bonus points
win_points = 3
win_bonus = 7
lose_bonus = -3

# Set up excel Writer object
book = load_workbook(squash_path + weekly_file)
writer = pd.ExcelWriter(squash_path + weekly_file, engine='openpyxl') 
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

# Define columns for statistics
stats_columms = ['Name',
                 'PL',
                 'W',
                 'L',
                 'PF',
                 'PA',
                 'DIFF',
                 'WB',
                 'LB',
                 'Points',
                 'Percentage'
                 ]

# Define columns for head-to-head sheets
head_to_head_columms = ['Player',
                         'PL',
                         'W',
                         'L',
                         'PF',
                         'PA',
                         'DIFF',
                         'WB',
                         'LB',
                         'Points',
                         'Percentage'
                         ]

# Determine if you are running for an individual week or for the all weeks
if master:
    # Read in all weeks results
    for week in range(1, week_num+1):
        tmp_games_df = pd.read_excel(squash_path + weekly_file, sheet_name='Week{0}_games'.format(week))
        tmp_tables_df = pd.read_excel(squash_path + weekly_file, sheet_name='Week{0}_table'.format(week))
        
        if week == 1:
            weekly_games_df = tmp_games_df
            tmp_tables_df['Week'] = week
            weekly_tables_df = tmp_tables_df
        else:
            weekly_games_df = weekly_games_df.append(tmp_games_df)
            tmp_tables_df['Week'] = week
            weekly_tables_df = weekly_tables_df.append(tmp_tables_df)
            #print(weekly_games_df)
else:
    # Read in weekly tab
    weekly_games_df = pd.read_excel(squash_path + weekly_file, sheet_name='Week{0}_games'.format(week_num))

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
    
    if master:
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
        if diff_tmp >= win_bonus:
            WB += 1
        elif diff_tmp >= lose_bonus and diff_tmp < 0:
            LB += 1
    
    # Metrics for table
    PL = player_games_df.shape[0]
    W = player_wins_df.shape[0]
    L = PL - W
    DIFF = PF - PA
    Points = W*win_points + WB + LB
    
    # Determine whether to use normalised score (for weekly table) or weighted score (master table)
    normalised_score = 0
    if master == True:
        # Determine normalised score for master 
        for _, weekly_result in player_tables_df.iterrows():
            wk = weekly_result['Week']
            n_score = weekly_result['Percentage']            
            normalised_score += n_score*wk/triangular(week_num)
            # print(player, wk, n_score, normalised_score)
        normalised_score = int(normalised_score)
        
        # Create table of accumulative scores for head-to-heads
        # Create empty dataframe with necessary columns
        vs_player_df = pd.DataFrame(columns=head_to_head_columms)
        for vs_player in players:
            # Skip if it is the vs_player is the same as player
            if player == vs_player:
                continue
            
            # Determine games player participated in against this particular player
            vs_player_games_df = player_games_df.loc[(player_games_df['Player 1'] == vs_player) | (player_games_df['Player 2'] == vs_player)]
            
            # continue if player has never played against this player
            if vs_player_games_df.shape[0] == 0:
                continue
            
            # Determine games won by player
            vs_player_wins_df = player_games_df.loc[((player_games_df['Player 1'] == vs_player) &\
                                                    (player_games_df['Score 1'] > player_games_df['Score 2']))\
                                                    |
                                                    ((player_games_df['Player 2'] == vs_player) &\
                                                    (player_games_df['Score 2'] > player_games_df['Score 1']))]
                                                    
            # Obtain points for and points against
            vs_PF = 0
            vs_PA = 0
            vs_WB = 0
            vs_LB = 0
            
            # Determine points
            for _, vs_game in vs_player_games_df.iterrows():
                # Points for and against
                vs_sc1_tmp = vs_game['Score 1']
                vs_sc2_tmp = vs_game['Score 2']
                if vs_game['Player 1'] == player:
                    vs_PF += vs_sc1_tmp
                    vs_PA += vs_sc2_tmp
                    vs_diff_tmp = vs_sc1_tmp - vs_sc2_tmp
                elif vs_game['Player 2'] == player:
                    vs_PF += vs_sc2_tmp
                    vs_PA += vs_sc1_tmp
                    vs_diff_tmp = vs_sc2_tmp - vs_sc1_tmp
                # Bonus points
                if vs_diff_tmp >= win_bonus:
                    vs_WB += 1
                elif vs_diff_tmp >= lose_bonus and vs_diff_tmp < 0:
                    vs_LB += 1
            
            # Metrics for table
            vs_PL = vs_player_games_df.shape[0]
            vs_L = vs_player_wins_df.shape[0]
            vs_W = vs_PL - vs_L
            vs_DIFF = vs_PF - vs_PA
            vs_Points = vs_W*win_points + vs_WB + vs_LB
            total_points_achievable = win_points + win_bonus
            vs_normalised_score = int((vs_Points/vs_PL)*(100.0/4.0))
            
            # Create row for player data to be added to the dataframe    
            vs_player_data = [vs_player, vs_PL, vs_W, vs_L, vs_PF, vs_PA, vs_DIFF, vs_WB, vs_LB, vs_Points, vs_normalised_score]
            
            # Construct vs_player dataframe with data from each head to head
            vs_player_tmp_df = pd.DataFrame(np.array([vs_player_data]), columns=head_to_head_columms)
            vs_player_df = vs_player_df.append(vs_player_tmp_df)
        
        # Convert all number columns to integers    
        for col in vs_player_df.columns[1:]:
            vs_player_df[col] = vs_player_df[col].astype(int)
            
        # Sort values by points
        vs_player_df = vs_player_df.sort_values(['Percentage', 'DIFF'], ascending=False)
        # print(vs_player_df)
        
        # Write to excel document
        vs_player_df.to_excel(writer, sheet_name='{0}'.format(player), startrow=1, index=False)
        
        # --------------------------
        # Add combined weekly tables
        # --------------------------
        # Obtain table of players weekly results
        h2h_weekly_tables_df = weekly_tables_df.loc[(weekly_tables_df['Name'] == player)]
        
        # Obtain columns for this table
        tab_cols = list(h2h_weekly_tables_df)
        
        # Move the 'week' column to head of list using index, pop and 
        # insert and also remove the 'Name' column
        tab_cols.insert(0, tab_cols.pop(tab_cols.index('Week')))
        tab_cols.pop(tab_cols.index('Name'))
        h2h_weekly_tables_df = h2h_weekly_tables_df.ix[:, tab_cols]
        
        # Order the new table by the week number, to put emphasis on the most recent week
        h2h_weekly_tables_df = h2h_weekly_tables_df.sort_values(['Week'], ascending=False)
        
        # Add total row
        tot_row = h2h_weekly_tables_df.sum()
        tot_row = pd.DataFrame([tot_row], columns=h2h_weekly_tables_df.columns)
        # Cast entire dataframe as int and then cast first column as string
        tot_row['Week'] = tot_row['Week'].astype(str)
        for col in tot_row.columns[1:]:
            tot_row[col] = tot_row[col].astype(int)
        tot_row['Week'] = 'Total'

        # Add average row
        avg_row = h2h_weekly_tables_df.mean()
        avg_row = pd.DataFrame([avg_row], columns=h2h_weekly_tables_df.columns)
        # Cast entire dataframe as int and then cast first column as string
        avg_row['Week'] = avg_row['Week'].astype(str)
        for col in avg_row.columns[1:]:
            avg_row[col] = avg_row[col].astype(int)
        avg_row['Week'] = 'Averages'
        
        # Cast entire dataframe as int and then cast first column as string
        h2h_weekly_tables_df['Week'] = h2h_weekly_tables_df['Week'].astype(str)
        for col in h2h_weekly_tables_df.columns[1:]:
            h2h_weekly_tables_df[col] = h2h_weekly_tables_df[col].astype(int)
        
        h2h_weekly_tables_df = h2h_weekly_tables_df.append(tot_row, ignore_index=True) 
        h2h_weekly_tables_df = h2h_weekly_tables_df.append(avg_row, ignore_index=True)  
        
        # Write table to Writer
        h2h_weekly_tables_df.to_excel(writer, sheet_name='{0}'.format(player), startrow=14, index=False)
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
weekly_stats_df = weekly_stats_df.sort_values(['Percentage', 'DIFF'], ascending=False)
# print(weekly_stats_df)

if master:
    weekly_stats_df.to_excel(writer, sheet_name='Master_table', index=False)
else:
    weekly_stats_df.to_excel(writer, sheet_name='Week{0}_table'.format(week_num), index=False)

writer.save()