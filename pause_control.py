"""
P3D Pause Control Script
Allows pausing and unpausing Prepar3D through SimConnect.

Usage:
    python pause_control.py          # Toggle pause
    python pause_control.py pause    # Pause the sim
    python pause_control.py unpause  # Unpause the sim
"""

import sys
import time
import logging

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

# Add SimConnect to path
sys.path.insert(0, r'c:\prepar3dpywrapper2.0\Python-SimConnect')

from SimConnect.P3DSimConnect import P3DSimConnect


def main():
    # Determine action from command line
    action = "toggle"
    if len(sys.argv) > 1:
        arg = sys.argv[1].lower()
        if arg in ("pause", "on", "1"):
            action = "pause"
        elif arg in ("unpause", "off", "0"):
            action = "unpause"
        else:
            action = "toggle"

    print("Connecting to Prepar3D...")
    
    try:
        sc = P3DSimConnect()
        print("Connected!")
        
        # Give time for connection to fully establish
        time.sleep(0.5)
        
        if action == "toggle":
            # Map and send PAUSE_TOGGLE event
            evt_id = sc.map_to_sim_event("PAUSE_TOGGLE")
            sc.send_event(evt_id, 0)
            print("Pause toggled!")
            
        elif action == "pause":
            # PAUSE_SET with 1 = Pause the sim
            evt_id = sc.map_to_sim_event("PAUSE_SET")
            sc.send_event(evt_id, 1)
            print("Simulation PAUSED")
            
        elif action == "unpause":
            # PAUSE_SET with 0 = Unpause the sim
            evt_id = sc.map_to_sim_event("PAUSE_SET")
            sc.send_event(evt_id, 0)
            print("Simulation UNPAUSED")
        
        # Give more time for command to process before disconnecting
        time.sleep(1.0)
        
        # Clean exit
        sc.exit()
        print("Disconnected.")
        
    except FileNotFoundError as e:
        print(f"ERROR: {e}")
        print("Make sure the P3D SimConnect DLL is in the correct location.")
        
    except ConnectionError as e:
        print(f"ERROR: {e}")
        print("Make sure Prepar3D is running with a flight loaded.")
        
    except Exception as e:
        print(f"ERROR: {e}")
        raise


if __name__ == "__main__":
    main()
