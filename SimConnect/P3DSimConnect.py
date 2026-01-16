import clr
import os
import sys
import time
import logging
from ctypes import *

# Default to searching for the DLL in common locations or define via environment variable
P3D_DLL_PATH = os.environ.get("P3D_SIMCONNECT_DLL_PATH")

LOGGER = logging.getLogger(__name__)

class P3DSimConnect:
    def __init__(self, auto_connect=True, dll_path=None):
        self.dll_path = dll_path or P3D_DLL_PATH
        self.Requests = {}
        self.Facilities = []
        self.connected = False
        self.quit = 0
        self.simconnect = None
        
        if not self.dll_path or not os.path.exists(self.dll_path):
            # Try to find it in the 'managed' folders if we are in the dev environment
            # This is a fallback specific to this user's workspace structure
            possible_paths = [
                r"C:\prepa3D_versions_managed\managed_v6\LockheedMartin.Prepar3D.SimConnect.dll",
                r"C:\prepa3D_versions_managed\managed_v5\LockheedMartin.Prepar3D.SimConnect.dll",
                r"C:\prepa3D_versions_managed\managed_v4\LockheedMartin.Prepar3D.SimConnect.dll",
            ]
            for p in possible_paths:
                if os.path.exists(p):
                    self.dll_path = p
                    break
        
        if not self.dll_path or not os.path.exists(self.dll_path):
             raise FileNotFoundError("Could not find LockheedMartin.Prepar3D.SimConnect.dll. Please set P3D_SIMCONNECT_DLL_PATH.")

        # Load the Managed DLL
        try:
            clr.AddReference(self.dll_path)
            from LockheedMartin.Prepar3D.SimConnect import SimConnect as P3D_SC
            from LockheedMartin.Prepar3D.SimConnect import SimConnectConstants
            self.sc_cls = P3D_SC
            self.constants = SimConnectConstants
        except Exception as e:
            LOGGER.error(f"Failed to load P3D DLL: {e}")
            raise

        if auto_connect:
            self.connect()

    def connect(self):
        try:
            # Managed SimConnect constructor acts as Open()
            # Signature: SimConnect(string szName, IntPtr hWnd, uint UserEventWin32, WaitHandle hEventHandle, uint ConfigIndex)
            # We use the windowless handle overload if available or standard
            self.simconnect = self.sc_cls("Python-SimConnect-Wrapper", 0, 0, 0, 0)
            
            # Subscribe to events
            self.simconnect.OnRecvOpen += self.on_recv_open
            self.simconnect.OnRecvQuit += self.on_recv_quit
            self.simconnect.OnRecvException += self.on_recv_exception
            self.simconnect.OnRecvSimobjectData += self.on_recv_simobject_data
            
            LOGGER.info("Connected to Prepar3D Managed SimConnect.")
            self.connected = True
            
        except Exception as e:
            LOGGER.error(f"Connection failed: {e}")
            raise

    def on_recv_open(self, simconnect, data):
        LOGGER.info(f"SimConnect Open: App {data.szApplicationName} {data.dwApplicationVersionMajor}.{data.dwApplicationVersionMinor}")

    def on_recv_quit(self, simconnect, data):
        LOGGER.info("SimConnect Quit")
        self.connected = False
        self.quit = 1

    def on_recv_exception(self, simconnect, data):
        LOGGER.warning(f"SimConnect Exception: {data.dwException}")

    def on_recv_simobject_data(self, simconnect, data):
        req_id = data.dwRequestID
        if req_id in self.Requests:
            # We need to parse the data. In managed code, data is often an object.
            # Depending on how the request was made (AddToClientDataDefinition vs AddToDataDefinition)
            # For now, simplistic handling assuming double/float for basic telemetry
             self.Requests[req_id].outData = data.dwData

    def request_data(self, _Request):
        # Implementation of request_data using Managed API
        # This requires mapping the definition and request IDs
        if not self.simconnect:
            return

        # In a real implementation, we need to handle the Enum mapping from the existing library
        # to the uint/Enum expected by the Managed DLL.
        # This is a complex mapping task.
        
        # Example dummy call to show structure
        # self.simconnect.RequestDataOnSimObjectType(...)
        pass

    def exit(self):
        if self.simconnect:
            self.simconnect.Dispose()
            self.simconnect = None
        self.quit = 1

    def IsHR(self, hr, value):
        # Managed API usually throws exceptions instead of returning HRESULTs, 
        # but if we get one, 0 is S_OK.
        return hr == value
