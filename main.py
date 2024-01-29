# main.py  Harrison Leduc CSCI 120 Final Project
import csv
import numpy as np
import pandas as pd #allows us to convert an xlsx file into a csv file
import matplotlib.pyplot as plt 


def convert_xlsx(filename):  
    """Converts a user-specified excel file and cleans up missing values with the creation of a list (player_data). Overall, this method will 
    create an overarching dictionary (data) that has sheet names (schools) as the keys, and the values point
    to another dictionary (d). 'd' has player names for the keys, and another dictionary (playerrow) for the values. 'playerrow' has column info
    for keys (i.e., goals, name, games played, etc.) and the associated data for the value (56, 'Ben Ten', 13...). Playerrow is in a for loop
    and is recreated for each player. 'playerrow' is created by iterating through 'player_data'.
    """
    wb = pd.ExcelFile(filename)
    df = pd.read_excel(wb, sheet_name=None)  #df is a dictionary, keys are the index of the sheets, values are the sheet names
                                            #the sheet names are a list of lists. The school name contains a list of every single player/row and their attributes
                                            # and each player/row is it's own list as well... explains the dictionary with values being a list of lists
    data = {}
    player_data = {}
    for school in df.keys():  #Goes through each sheet, giving the 'school' and related data the name of that specific sheet
        df[school].fillna(0, inplace=True) #all NaN's (not a number's) become 0, allows all columns to be read
        df[school].replace('-', '0', inplace=True) #makes dashes into zeros, allows all columns to be read
        player_data[school] = []   #opens a list where we will store each of the rows/player values in their own lists
        for index, row in df[school].iterrows(): 
            if index == 0 or index == 1: #skips the first two rows since they are garbage data
                continue
            new_row = '' #creates an empty string
            for col in list(df[school].columns): 
                new_row += str(row[col]) + ','  #adds each column's associated values to new_row for that row
            values = new_row.split(',')  #splits each value with a column

            number, player, games_played, games_started, \
                goals, assists, points, shots, shooting_percentage, \
                shots_on_goal, shots_on_goal_percentage, man_up_goals, man_down_goals, groundballs, \
                turnovers, caused_turnovers, faceoff_win_loss, \
                faceoff_percentage = values[:18]  #gives names for each of the first 18 columns in the data set (19-21 are garbage data)
       
         
            player = player.replace('.' , '') #removes periods following player names
            
            player_data[school].append([int(number), player, int(games_played), int(games_started), \
                                int(goals), int(assists), int(points), int(shots), \
                                float(shooting_percentage), int(shots_on_goal), float(shots_on_goal_percentage), \
                                int(man_up_goals), int(man_down_goals), int(groundballs), \
                                int(turnovers), int(caused_turnovers), faceoff_win_loss , \
                                float(faceoff_percentage)]) #creates a list of lists. each index is a list with a complete row's data
        
        playerrow = {}
        d = {}
        for row in player_data[school][:-2]: #dictionary that maps to a list of lists
            playerrow = {'Number': row[0], 'Name': row[1], 'Games Played': row[2], 'Games Started': row[3], 'Goals': row[4], 'Assists': row[5], \
                        'Points': row[6], 'Shots': row[7], 'Shooting Percentage': row[8], 'Shots On Goal': row[9], 'SOG Percentage': row[10], \
                            'Man Up Goals': row[11], 'Man Down Goals': row[12], 'Groundballs': row[13], 'Turnovers': row[14], \
                                'Caused Turnovers': row[15], 'Faceoff Win-Loss': row[16], 'Faceoff Percentage': row[17]}
            d |= {row[1]: playerrow}
        data[school] = d

    return data



def create_stats(filename):
        """Takes in the data dictionary from previous function, creates a separate list for the specified school's goals, assists, and points.
        Excludes any 0 values/players from the lists. Automatically calls the function to plot all three lists.
        """
        data = convert_xlsx(filename)     # should show d = {'Name': {goals: 55, assits = 44 ... }}
       
        
        for school in data:
            goals = []
            assists = []
            points = []
           
            for name in data[school]:
                if data[school][name].get('Goals') != 0:   #exclude players with 0 for the category
                    goals += [(data[school][name].get('Goals'), name)]         #should show goals = [(55, name1), (42, name2)]

                if data[school][name].get('Assists') != 0:
                    assists += [(data[school][name].get('Assists'), name)]
                if data[school][name].get('Points') != 0:
                    points += [(data[school][name].get('Points'), name)]


            goals = sorted(goals)  #makes 1 for the specific school. but in for loop so will end up doing all schools
            assists = sorted(assists)
            points = sorted(points)
            plot_goals(goals)
            plot_assists(assists)
            plot_points(points)


def plot_goals(goals):
    """Call function to create bar chart of player's goals for a university
    """
    x_labels = [val[1] for val in goals]
    y_labels = [val[0] for val in goals]
    plt.figure(figsize=(18,10))
    ax = pd.Series(y_labels).plot(kind='bar')
    ax.set_xticklabels(x_labels)

    rects = ax.patches

    for rect, label in zip(rects, y_labels):
        height = rect.get_height()
        ax.text(rect.get_x() + rect.get_width()/2, height + 0.1, label, ha='center', va='bottom') #sets height above bar for the # of goals text, distance from bottom for names
    plt.bar(range(len(goals)), [val[0] for val in goals], align='center')
    plt.xticks(range(len(goals)), [val[1] for val in goals])
    plt.xticks(rotation=70)
    plt.show()

def plot_assists(assists):
    """Call function to create bar chart of player's assists for a university
    """
    x_labels = [val[1] for val in assists]
    y_labels = [val[0] for val in assists]
    plt.figure(figsize=(18,10))
    ax = pd.Series(y_labels).plot(kind='bar')
    ax.set_xticklabels(x_labels)

    rects = ax.patches

    for rect, label in zip(rects, y_labels):
        height = rect.get_height()
        ax.text(rect.get_x() + rect.get_width()/2, height + 0.1, label, ha='center', va='bottom')
    plt.bar(range(len(assists)), [val[0] for val in assists], align='center')
    plt.xticks(range(len(assists)), [val[1] for val in assists])
    plt.xticks(rotation=70)
    plt.show()

    
def plot_points(points):
    """Call function to create bar chart of player's total points for a university
    """
    x_labels = [val[1] for val in points]
    y_labels = [val[0] for val in points]
    plt.figure(figsize=(18,10))
    ax = pd.Series(y_labels).plot(kind='bar')
    ax.set_xticklabels(x_labels)

    rects = ax.patches

    for rect, label in zip(rects, y_labels):
        height = rect.get_height()
        ax.text(rect.get_x() + rect.get_width()/2, height + 0.1, label, ha='center', va='bottom')
    plt.bar(range(len(points)), [val[0] for val in points], align='center')
    plt.xticks(range(len(points)), [val[1] for val in points])
    plt.xticks(rotation=70)
    plt.show()  



  


if __name__ == '__main__':
    create_stats('Excel_Stats.xlsx')