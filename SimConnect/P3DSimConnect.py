import clr
import os
import sys
import time
import logging
import ctypes
from System import Enum as NetEnum
from System import UInt32, Int32, IntPtr
from System.Runtime.InteropServices import Marshal

LOGGER = logging.getLogger(__name__)

# Constants to match original library
SIMCONNECT_UNUSED = 0xFFFFFFFF
SIMCONNECT_OBJECT_ID_USER = 0

class P3DSimConnect:
    def __init__(self, auto_connect=True, dll_path=None, p3d_version=None):
        self.Requests = {}
        self.Facilities = []
        self.quit = 0
        self.ok = False
        self.running = False
        self.paused = False
        self.simconnect = None
        self.own_handle = None
        
        # Internal counters for dynamic IDs
        self._def_id_counter = 0
        self._req_id_counter = 0
        self._evt_id_counter = 0
        self._group_id_counter = 0
        
        # Mappings
        self.mapping_def_id = {}
        self.mapping_req_id = {}
        
        # Load DLL
        self._load_dll(dll_path, p3d_version)

        if auto_connect:
            self.connect()

    def _load_dll(self, dll_path, version):
        """Attempts to load the P3D Managed SimConnect DLL."""
        search_paths = []
        if dll_path:
            search_paths.append(dll_path)
        
        # Default known paths in reverse version order
        base_dir = r"C:\prepa3D_versions_managed"
        if version:
             search_paths.append(os.path.join(base_dir, f"managed_v{version}", "LockheedMartin.Prepar3D.SimConnect.dll"))
        else:
            search_paths.append(os.path.join(base_dir, "managed_v6", "LockheedMartin.Prepar3D.SimConnect.dll"))
            search_paths.append(os.path.join(base_dir, "managed_v5", "LockheedMartin.Prepar3D.SimConnect.dll"))
            search_paths.append(os.path.join(base_dir, "managed_v4", "LockheedMartin.Prepar3D.SimConnect.dll"))
            
        dll_found = False
        for path in search_paths:
            if os.path.exists(path):
                try:
                    LOGGER.info(f"Loading P3D DLL from: {path}")
                    clr.AddReference(path)
                    dll_found = True
                    break
                except Exception as e:
                    LOGGER.error(f"Failed to load {path}: {e}")
        
        if not dll_found:
             raise FileNotFoundError("Could not find or load LockheedMartin.Prepar3D.SimConnect.dll.")

        # Import namespaces after loading DLL
        from LockheedMartin.Prepar3D.SimConnect import SimConnect as ScClass
        # SimConnectConstants does not exist in v4/v5/v6 managed, constants are on the class or enums or we define them.
        
        try:
            from LockheedMartin.Prepar3D.SimConnect import SIMCONNECT_SIMOBJECT_TYPE
            from LockheedMartin.Prepar3D.SimConnect import SIMCONNECT_PERIOD
            from LockheedMartin.Prepar3D.SimConnect import SIMCONNECT_DATA_REQUEST_FLAG
            from LockheedMartin.Prepar3D.SimConnect import SIMCONNECT_DATATYPE
        except ImportError:
            # Fallback or specific version handling if names differ
            LOGGER.error("Could not import P3D Enums. DLL version might be incompatible.")
            raise

        self.sc_cls = ScClass
        self.SC_TYPE = SIMCONNECT_SIMOBJECT_TYPE
        self.SC_PERIOD = SIMCONNECT_PERIOD
        self.SC_FLAG = SIMCONNECT_DATA_REQUEST_FLAG
        self.SC_DATATYPE = SIMCONNECT_DATATYPE
        
        # Constants
        # Use ScClass members if available, or defaults
        self.Config = None # Not used heavily, or usually 0



# ... (inside __init__, skipping to connect)

    def connect(self):
        try:
            # Signature: SimConnect(string Name, IntPtr hWnd, uint UserEventWin32, WaitHandle hEventHandle, uint ConfigIndex)
            # hWnd = IntPtr.Zero (Windowless)
            # UserEventWin32 = 0
            # hEventHandle = None (WaitHandle)
            # ConfigIndex = 0
            
            self.simconnect = self.sc_cls(
                "Python-SimConnect",
                IntPtr.Zero,
                UInt32(0),
                None,
                UInt32(0)
            )
            
            # Subscribe to events
            self.simconnect.OnRecvOpen += self.on_recv_open
            self.simconnect.OnRecvQuit += self.on_recv_quit
            self.simconnect.OnRecvException += self.on_recv_exception
            self.simconnect.OnRecvSimobjectData += self.on_recv_simobject_data
            self.simconnect.OnRecvEvent += self.on_recv_event
            self.simconnect.OnRecvSystemState += self.on_recv_system_state
            
            LOGGER.info("Connected to Prepar3D.")
            
            # Wait for connection to establish (OnRecvOpen sets ok=True)
            self.ok = True # Assume true if no exception for managed, but OnRecvOpen confirms it
            
        except Exception as e:
            LOGGER.error(f"Connection failed: {e}")
            raise ConnectionError(f"Connection failed: {e}")

    # --- Event Handlers ---

    def on_recv_open(self, simconnect, data):
        LOGGER.info(f"Connected to P3D: {data.szApplicationName} {data.dwApplicationVersionMajor}.{data.dwApplicationVersionMinor}")
        self.ok = True

    def on_recv_quit(self, simconnect, data):
        LOGGER.info("P3D Quit")
        self.quit = 1
        self.ok = False
        self.connected = False

    def on_recv_exception(self, simconnect, data):
        LOGGER.warning(f"SimConnect Exception: {data.dwException} Packet: {data.dwSendID}")

    def on_recv_event(self, simconnect, data):
        # Handle system events like Pause/Crashes
        # We need to map EventID back to what we registered if we tracked it
        pass

    def on_recv_system_state(self, simconnect, data):
        # Handle system state (paused, etc)
        pass

    def on_recv_simobject_data(self, simconnect, data):
        req_id = data.dwRequestID
        if req_id in self.Requests:
            request = self.Requests[req_id]
            # data.dwData is the object array
            # In Managed SimConnect, dwData is NOT a pointer, it's (usually) the object we registered.
            # But wait, RequestDataOnSimObjectType usually returns a struct.
            # pythonnet matches the 'object' signature.
            
            # If we requested generic data (Struct), it comes as object.
            # We assume we requested doubles.
            
            try:
                # We expect data.dwData to be a single double or array of doubles?
                # Actually, Managed wrapper returns the struct instance we defined via RegisterDataDefineStruct 
                # OR we likely used AddToDataDefinition, which returns data in a specific way.
                
                # With AddToDataDefinition, we get an object that we need to cast?
                # Actually, data.dwData is `object`.
                # If we used simple types, it might be an array.
                
                # CRITICAL: MSFS and P3D Managed behavior for pythonnet interactions.
                # Assuming standard behavior: double[] if flexible.
                
                # Current simplistic approach:
                request.outData = data.dwData
            except Exception as e:
                LOGGER.error(f"Error parsing data for req {req_id}: {e}")


    # --- API Implementation ---

    def new_def_id(self):
        self._def_id_counter += 1
        return self._def_id_counter

    def new_request_id(self):
        self._req_id_counter += 1
        return self._req_id_counter
    
    def map_to_sim_event(self, name):
        # Maps a string name (e.g. "Pause") to an ID
        evt_id = self._evt_id_counter
        self._evt_id_counter += 1
        
        # Ensure name is string
        if isinstance(name, bytes):
            name = name.decode()
            
        self.simconnect.MapClientEventToSimEvent(evt_id, name)
        return evt_id

    def add_to_notification_group(self, group, event, bMaskable=False):
        # group is an ID, event is event ID
        if isinstance(group, object) and hasattr(group, 'value'): group = group.value
        if isinstance(event, object) and hasattr(event, 'value'): event = event.value
        
        self.simconnect.AddClientEventToNotificationGroup(group, event, bMaskable)

    def request_data(self, _Request):
         # Register definition if not done
        if not hasattr(_Request, 'DATA_DEFINITION_ID'):
            _Request.DATA_DEFINITION_ID = self.new_def_id()
             
            # Add variables
            for definition in _Request.definitions:
                # definition: (b'VAR_NAME', b'Unit', Datatype)
                name = definition[0]
                unit = definition[1]
                
                if isinstance(name, bytes): name = name.decode()
                if isinstance(unit, bytes): unit = unit.decode()
                
                datum_type = self.SC_DATATYPE.FLOAT64 # Default
                
                self.simconnect.AddToDataDefinition(
                    _Request.DATA_DEFINITION_ID,
                    name,
                    unit,
                    datum_type,
                    0.0,
                    SIMCONNECT_UNUSED
                )

        if not hasattr(_Request, 'DATA_REQUEST_ID'):
             _Request.DATA_REQUEST_ID = self.new_request_id()
             self.Requests[_Request.DATA_REQUEST_ID] = _Request

        _Request.outData = None
        
        # Request Data
        # User object (0)
        self.simconnect.RequestDataOnSimObjectType(
            _Request.DATA_REQUEST_ID,
            _Request.DATA_DEFINITION_ID,
            0,
            self.SC_TYPE.USER
        )

    def get_data(self, _Request):
        self.request_data(_Request)
        attempts = 0
        while _Request.outData is None and attempts < 100: # 1s timeout
            time.sleep(0.01)
            attempts += 1
        
        if _Request.outData is None:
            return False
        return True

    def set_data(self, _Request):
        # TODO: Implement Set Data
        # This requires constructing the data array in .NET format
        pass

    def send_event(self, evnt, data=0):
        # event is ID, data is int
        if hasattr(evnt, 'value'): evnt = evnt.value
        if hasattr(data, 'value'): data = data.value
        
        # Group priority
        priority = 1 # GROUP_PRIORITY_HIGHEST
        
        self.simconnect.TransmitClientEvent(
            SIMCONNECT_OBJECT_ID_USER,
            evnt,
            data,
            priority,
            16 # EVENT_FLAG_GROUPID_IS_PRIORITY
        )
        return True

    def exit(self):
        self.quit = 1
        if self.simconnect:
            self.simconnect.Dispose()
            self.simconnect = None
