import os
import pandas as pd
from datetime import datetime, timedelta
import itertools
import random

def generate_schedule(teams, start_date):
    schedule = []

    # Generate combinations of teams for home and away matches
    matches = list(itertools.permutations(teams, 2))

    # Shuffle the matches to randomize the schedule
    random.shuffle(matches)

    match_dates = [start_date + timedelta(days=i) for i in range(len(matches))]

    for i, match in enumerate(matches):
        home_team, away_team = match
        match_date = match_dates[i]

        # Add the match to the schedule
        schedule.append((home_team, away_team, match_date))

    return schedule

def read_international_matches(filename):
    if not os.path.isfile(filename):
        print(f"File '{filename}' not found.")
        return []

    # Read Excel file using pandas
    df = pd.read_excel(filename)

    # Extract relevant data
    international_matches = []
    for index, row in df.iterrows():
        home_team, away_team = row['match'].split(' vs ')
        location = row['location']
        
        # Convert Timestamp to string
        date_str = str(row['date'])
        
        # Parse the string to datetime
        date = datetime.strptime(date_str, '%Y-%m-%d %H:%M:%S')  # Adjust the format based on the actual format in your DataFrame
        international_matches.append((home_team, away_team, location, date))

    return international_matches

def main():
    # Get the number of teams from the user
    num_teams = int(input("Enter the number of teams: "))

    # Get the names of the teams
    teams = [input(f"Enter the name of Team {i + 1}: ") for i in range(num_teams)]

    # Get the start date for the fixtures
    start_date_str = input("Enter the start date for fixtures (YYYY-MM-DD): ")
    start_date = datetime.strptime(start_date_str, '%Y-%m-%d')

    # Generate the initial schedule
    initial_schedule = generate_schedule(teams, start_date)
    print("\nInitial Schedule:")
    for match in initial_schedule:
        print(f"{match[0]} vs {match[1]} on {match[2].strftime('%Y-%m-%d')}")

    # Ask the user to input a file with international matches
    international_file = input("\nEnter the filename with international matches (or leave blank): ")

    if international_file:
        international_matches = read_international_matches(international_file)

        # Check and update the schedule to avoid conflicts with international matches
        for i, match in enumerate(initial_schedule):
            home_team, away_team, match_date = match

            # Check if either the home or away team is in the initial schedule
            if home_team in teams or away_team in teams:
                # Find the index of the teams in the initial schedule
                home_team_index = teams.index(home_team) if home_team in teams else None
                away_team_index = teams.index(away_team) if away_team in teams else None

                # Check if the date of the international match conflicts with the initial schedule
                for international_match in international_matches:
                    international_home_team, international_away_team, _, international_match_date = international_match

                    if (
                        ((international_home_team == home_team or international_away_team == home_team) and home_team_index is not None) or
                        ((international_home_team == away_team or international_away_team == away_team) and away_team_index is not None)
                    ) and international_match_date == match_date:
                        # Reschedule the initial match for the teams on a different date
                        new_date = match_date + timedelta(days=1)  # You can adjust the logic for rescheduling
                        initial_schedule[i] = (home_team, away_team, new_date)

                        print(f"Conflict: {home_team} vs {away_team} on {match_date.strftime('%Y-%m-%d')} already scheduled. Rescheduled to {new_date.strftime('%Y-%m-%d')}.")
                        break

        print("\nUpdated Schedule:")
        for match in initial_schedule:
            print(f"{match[0]} vs {match[1]} on {match[2].strftime('%Y-%m-%d')}")

        # Convert initial_schedule and updated_schedule to pandas DataFrame
        df_initial_schedule = pd.DataFrame(initial_schedule, columns=['Home Team', 'Away Team', 'Date'])
        df_updated_schedule = pd.DataFrame(initial_schedule, columns=['Home Team', 'Away Team', 'Date'])

        # Save initial_schedule and updated_schedule to Excel files
        initial_schedule_file = "initial_schedule.xlsx"
        updated_schedule_file = "updated_schedule.xlsx"

        # Set the date format and adjust column width for Excel output
        with pd.ExcelWriter(initial_schedule_file, engine='xlsxwriter') as writer:
            df_initial_schedule.to_excel(writer, sheet_name='Sheet1', index=False)
            worksheet = writer.sheets['Sheet1']
            date_format = writer.book.add_format({'num_format': 'yyyy-mm-dd'})
            worksheet.set_column('C:C', 12, date_format)

        with pd.ExcelWriter(updated_schedule_file, engine='xlsxwriter') as writer:
            df_updated_schedule.to_excel(writer, sheet_name='Sheet1', index=False)
            worksheet = writer.sheets['Sheet1']
            date_format = writer.book.add_format({'num_format': 'yyyy-mm-dd'})
            worksheet.set_column('C:C', 12, date_format)

    else:
        print("\nNo international matches added.")

if __name__ == "__main__":
    main()
