"""
Convert Excel files to CSV for use in the scheduling application
"""
import pandas as pd
import os

def convert_excel_to_csv():
    # Create Data directory if it doesn't exist
    os.makedirs('Data', exist_ok=True)
    
    # Convert Events.xlsx to CSV
    print("Converting Events.xlsx...")
    events_df = pd.read_excel('Data/Events.xlsx')
    events_df.to_csv('Data/events.csv', index=False)
    print(f"Created Data/events.csv with {len(events_df)} events")
    
    # Convert Event_Days.xlsx to CSV
    print("Converting Event_Days.xlsx...")
    days_df = pd.read_excel('Data/Event_Days.xlsx')
    # Remove the unnamed column
    days_df = days_df.drop(columns=[col for col in days_df.columns if 'Unnamed' in col])
    days_df.to_csv('Data/event_days.csv', index=False)
    print(f"Created Data/event_days.csv with {len(days_df)} event days")
    
    print("\nConversion complete!")

if __name__ == '__main__':
    convert_excel_to_csv()
