import requests
import json
import os

# --- Configuration ---
# Replace with your actual Personal Access Token (PAT)
# You can also put this in your .env file as WEBEX_PAT and load it.
# For this script, you can temporarily hardcode it if you're careful,
# but for your main project, always use .env.
YOUR_PAT = "MjVjMzJlNzktNWU2Mi00MTE4LThhYzItNDlkMDhlZjg4Y2FmNzA0NDg3ZWYtMmU1_PF84_1eb65fdf-9643-417f-9974-ad72cae0e10f"

# The exact title of the Webex Space where your bot is located
TARGET_ROOM_TITLE = "Services EA BPT" 

# --- Webex API Details ---
WEBEX_API_URL = "https://webexapis.com/v1/rooms"
HEADERS = {
    "Authorization": f"Bearer {YOUR_PAT}",
    "Content-Type": "application/json"
}
PARAMS = {
    "max": 100,  # Fetch up to 100 rooms
    "type": "group" # We're looking for a group space
}

print(f"Searching for room with title: '{TARGET_ROOM_TITLE}'...")

try:
    response = requests.get(WEBEX_API_URL, headers=HEADERS, params=PARAMS)
    response.raise_for_status() # Raise an exception for HTTP errors (4xx or 5xx)
    rooms_data = response.json()

    found_room_id = None
    if rooms_data and "items" in rooms_data:
        for room in rooms_data["items"]:
            if room.get('title') == TARGET_ROOM_TITLE:
                found_room_id = room.get('id')
                print("\n--- ROOM FOUND! ---")
                print(f"Title: {room.get('title')}")
                print(f"Room ID: {room.get('id')}")
                print(f"Type: {room.get('type')}")
                print(f"Last Activity: {room.get('lastActivity')}")
                print("-------------------\n")
                break
        
        if found_room_id:
            print(f"Please use this Room ID in your .env file for WEBEX_ROOM_ID: {found_room_id}")
        else:
            print(f"Room with title '{TARGET_ROOM_TITLE}' not found among the rooms you are in.")
            print("\n--- All Rooms You Are In (for debugging) ---")
            for room in rooms_data["items"]:
                print(f"Title: {room.get('title')}, ID: {room.get('id')}")
            print("------------------------------------------")
    else:
        print("No rooms found or unexpected response format.")

except requests.exceptions.RequestException as e:
    print(f"Error fetching Webex rooms: {e}")
    if hasattr(e, 'response') and e.response is not None:
        print(f"Response status code: {e.response.status_code}")
        print(f"Response content: {e.response.text}")
except Exception as e:
    print(f"An unexpected error occurred: {e}")