import pandas as pd

print("=== Events.xlsx ===")
try:
    events_df = pd.read_excel('Data/Events.xlsx')
    print(events_df.head(20))
    print(f"\nShape: {events_df.shape}")
    print(f"Columns: {events_df.columns.tolist()}")
except Exception as e:
    print(f"Error reading Events.xlsx: {e}")

print("\n\n=== Event_Days.xlsx ===")
try:
    days_df = pd.read_excel('Data/Event_Days.xlsx')
    print(days_df.head(20))
    print(f"\nShape: {days_df.shape}")
    print(f"Columns: {days_df.columns.tolist()}")
except Exception as e:
    print(f"Error reading Event_Days.xlsx: {e}")
