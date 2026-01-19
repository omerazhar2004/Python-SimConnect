import clr
import os
import sys
import time
import logging
import ctypes
from System import Enum as NetEnum
from System import UInt32, Int32, IntPtr, Enum, AppDomain
from System.Runtime.InteropServices import Marshal
from System.Reflection import TypeAttributes, AssemblyName
from System.Reflection.Emit import AssemblyBuilderAccess

LOGGER = logging.getLogger(__name__)

# Constants to match original library
SIMCONNECT_UNUSED = 0xFFFFFFFF
SIMCONNECT_OBJECT_ID_USER = 0

class ID:
    def __init__(self, val):
        self.value = val
    def __repr__(self):
        return str(self.value)


class P3DSimConnect:
    # Class-level shared dynamic enum type (created once)
    _EventEnumType = None
    
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
        
        # Create dynamic enum type for event/group IDs (once per class)
        if P3DSimConnect._EventEnumType is None:
            P3DSimConnect._EventEnumType = self._create_event_enum()
        
        if p3d_version is None:
            p3d_version = self._detect_p3d_version()
            if p3d_version:
                LOGGER.info(f"Auto-detected P3D Version: {p3d_version}")

        # Load DLL
        self._load_dll(dll_path, p3d_version)

        if auto_connect:
            self.connect()
    
    def _create_event_enum(self):
        """Create a dynamic .NET Enum type for event IDs."""
        try:
            aName = AssemblyName("P3DEventEnums")
            ab = AppDomain.CurrentDomain.DefineDynamicAssembly(aName, AssemblyBuilderAccess.Run)
            mb = ab.DefineDynamicModule("EventEnumModule")
            eb = mb.DefineEnum("ClientEventID", TypeAttributes.Public, UInt32)
            
            # Pre-define 100 enum values for flexibility
            for i in range(100):
                eb.DefineLiteral(f"EVENT_{i}", UInt32(i))
            
            return eb.CreateType()
        except Exception as e:
            LOGGER.error(f"Failed to create dynamic enum: {e}")
            return None

    def _detect_p3d_version(self):
        """Attempts to detect running P3D version via Named Pipes."""
        try:
            # Check for P3D v6
            if os.path.exists(r"\\.\pipe\Lockheed Martin Prepar3D v6\SimConnect"):
                return 6
            # Check for P3D v5
            if os.path.exists(r"\\.\pipe\Lockheed Martin Prepar3D v5\SimConnect"):
                return 5
            # Check for P3D v4
            if os.path.exists(r"\\.\pipe\Lockheed Martin Prepar3D v4\SimConnect"):
                return 4
        except Exception as e:
            LOGGER.warning(f"Failed to check pipes for version detection: {e}")
        return None

    def _load_dll(self, dll_path, version):
        """Attempts to load the P3D Managed SimConnect DLL."""
        search_paths = []
        if dll_path:
            search_paths.append(dll_path)
        
        # Look for standard SDK paths
        sdk_base = r"C:\Program Files\Lockheed Martin"
        # Search for any P3D SDK folders
        try:
            if os.path.exists(sdk_base):
                 for item in os.listdir(sdk_base):
                      if "SDK" in item:
                           item_path = os.path.join(sdk_base, item, "lib", "SimConnect", "managed", "LockheedMartin.Prepar3D.SimConnect.dll")
                           if os.path.exists(item_path):
                                # Tag with version if possible
                                if "v4" in item: search_paths.append((4, item_path))
                                elif "v5" in item: search_paths.append((5, item_path))
                                elif "v6" in item: search_paths.append((6, item_path))
                                else: search_paths.append((None, item_path))
        except:
             pass

        # Custom paths
        base_dir = r"C:\prepa3D_versions_managed"
        if os.path.exists(base_dir):
            for v in [6, 5, 4]:
                p = os.path.join(base_dir, f"managed_v{v}", "LockheedMartin.Prepar3D.SimConnect.dll")
                if os.path.exists(p):
                    search_paths.append((v, p))

        # Re-order based on requested version
        ordered_paths = []
        if version:
             # FOR P3D v5: Prioritize v4 Legacy DLL because v5.4 SDK DLL is often incompatible with 5.3 server
             if version == 5:
                  for v, p in search_paths:
                       if v == 4: ordered_paths.append(p)
             
             # Then try exact version
             for v, p in search_paths:
                  if v == version and p not in ordered_paths: ordered_paths.append(p)
             
             # Finally the rest
             for v, p in search_paths:
                  if p not in ordered_paths: ordered_paths.append(p)
        else:
             # No version, just add all in DESC order of version
             for v, p in sorted(search_paths, key=lambda x: x[0] if x[0] else 0, reverse=True):
                  ordered_paths.append(p)

        dll_found = False
        loaded_path = None
        for path in ordered_paths:
            try:
                LOGGER.info(f"Checking for DLL at: {path}")
                # We can't really "unload" in clr easily, but we can avoid re-loading
                clr.AddReference(path)
                LOGGER.info(f"Loaded P3D DLL from: {path}")
                dll_found = True
                loaded_path = path
                break
            except Exception as e:
                LOGGER.debug(f"Failed to load {path}: {e}")
        
        if not dll_found:
             raise FileNotFoundError("Could not find or load LockheedMartin.Prepar3D.SimConnect.dll in any of the expected locations.")

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
            from System.Threading import AutoResetEvent
            self.hEvent = AutoResetEvent(False)
            
            self.simconnect = self.sc_cls(
                "Python-SimConnect",
                IntPtr.Zero,
                UInt32(0),
                self.hEvent,
                UInt32(0)
            )
            
            # Dummy handle for compatibility
            self.hSimConnect = None

            
            # Subscribe to events
            self.simconnect.OnRecvOpen += self.on_recv_open
            self.simconnect.OnRecvQuit += self.on_recv_quit
            self.simconnect.OnRecvException += self.on_recv_exception
            self.simconnect.OnRecvSimobjectData += self.on_recv_simobject_data
            self.simconnect.OnRecvEvent += self.on_recv_event
            self.simconnect.OnRecvSystemState += self.on_recv_system_state
            
            LOGGER.info("Connected to Prepar3D.")
            
            # Wait for connection to establish (OnRecvOpen sets ok=True)
            LOGGER.debug("Waiting for OnRecvOpen...")
            timeout = 3.0
            start = time.time()
            while not self.ok and (time.time() - start) < timeout:
                try:
                    self.simconnect.ReceiveMessage()
                except Exception as e:
                    # LOGGER.debug(f"ReceiveMessage throw: {e}")
                    pass
                time.sleep(0.1)
            
            if not self.ok:
                LOGGER.warning(f"Connection established but OnRecvOpen not received within {timeout}s.")
                # We set ok=True anyway as some versions might not send Open immediately or at all in some configs
                # self.ok = True 
            else:
                 LOGGER.info("Connection confirmed by OnRecvOpen.")
            
        except Exception as e:
            LOGGER.error(f"Connection failed during constructor or early pump: {e}")
            raise ConnectionError(f"Connection failed: {e}")

    # --- Event Handlers ---

    def on_recv_open(self, simconnect, data):
        LOGGER.info(f"OnRecvOpen: Connected to {data.szApplicationName} version {data.dwApplicationVersionMajor}.{data.dwApplicationVersionMinor}")
        self.ok = True

    def on_recv_quit(self, simconnect, data):
        LOGGER.info("OnRecvQuit: P3D closed connection")
        self.quit = 1
        self.ok = False
        self.connected = False

    def on_recv_exception(self, simconnect, data):
        LOGGER.warning(f"OnRecvException: Exception={data.dwException} Packet={data.dwSendID} Index={data.dwIndex}")

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
        return ID(self._def_id_counter)

    def new_request_id(self):
        self._req_id_counter += 1
        return ID(self._req_id_counter)
    
    def map_to_sim_event(self, name):
        # Maps a string name (e.g. "PAUSE_TOGGLE") to an ID
        evt_id = self._evt_id_counter
        self._evt_id_counter += 1
        
        # Ensure name is string
        if isinstance(name, bytes):
            name = name.decode()
        
        # Use our dynamically created enum type
        if P3DSimConnect._EventEnumType is not None:
            evt_enum = Enum.ToObject(P3DSimConnect._EventEnumType, UInt32(evt_id))
        else:
            # Fallback (shouldn't happen)
            evt_enum = Enum.ToObject(self.SC_PERIOD, evt_id)
            
        self.simconnect.MapClientEventToSimEvent(evt_enum, name)
        return evt_id

    def add_to_notification_group(self, group, event, bMaskable=False):
        # group is an ID, event is event ID
        if isinstance(group, object) and hasattr(group, 'value'): group = group.value
        if isinstance(event, object) and hasattr(event, 'value'): event = event.value
        
        # Use our dynamically created enum type
        if P3DSimConnect._EventEnumType is not None:
            group_enum = Enum.ToObject(P3DSimConnect._EventEnumType, UInt32(group))
            event_enum = Enum.ToObject(P3DSimConnect._EventEnumType, UInt32(event))
        else:
            group_enum = Enum.ToObject(self.SC_PERIOD, group)
            event_enum = Enum.ToObject(self.SC_PERIOD, event)

        self.simconnect.AddClientEventToNotificationGroup(group_enum, event_enum, bMaskable)

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
                    _Request.DATA_DEFINITION_ID.value,
                    name,
                    unit,
                    datum_type,
                    0.0,
                    SIMCONNECT_UNUSED
                )

        if not hasattr(_Request, 'DATA_REQUEST_ID'):
             _Request.DATA_REQUEST_ID = self.new_request_id()
             self.Requests[_Request.DATA_REQUEST_ID.value] = _Request

        _Request.outData = None
        
        # Request Data
        # User object (0)
        if P3DSimConnect._EventEnumType is not None:
             def_enum = Enum.ToObject(P3DSimConnect._EventEnumType, UInt32(_Request.DATA_DEFINITION_ID.value))
             req_enum = Enum.ToObject(P3DSimConnect._EventEnumType, UInt32(_Request.DATA_REQUEST_ID.value))
        else:
             def_enum = Enum.ToObject(self.SC_PERIOD, _Request.DATA_DEFINITION_ID.value)
             req_enum = Enum.ToObject(self.SC_PERIOD, _Request.DATA_REQUEST_ID.value)

        self.simconnect.RequestDataOnSimObjectType(
            req_enum,
            def_enum,
            UInt32(0),
            self.SC_TYPE.USER
        )

    def get_data(self, _Request):
        self.request_data(_Request)
        attempts = 0
        while _Request.outData is None and attempts < 100: # 1s timeout
            try:
                self.simconnect.ReceiveMessage()
            except:
                pass
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
        
        # Priority/Group ID
        priority = 1 # GROUP_PRIORITY_HIGHEST
        
        # Use our dynamically created enum type
        if P3DSimConnect._EventEnumType is not None:
            evt_enum = Enum.ToObject(P3DSimConnect._EventEnumType, UInt32(evnt))
            group_enum = Enum.ToObject(P3DSimConnect._EventEnumType, UInt32(priority))
        else:
            # Fallback
            evt_enum = Enum.ToObject(self.SC_PERIOD, evnt)
            group_enum = Enum.ToObject(self.SC_PERIOD, priority)

        try:
             from LockheedMartin.Prepar3D.SimConnect import SIMCONNECT_EVENT_FLAG
             flag_enum = SIMCONNECT_EVENT_FLAG.GROUPID_IS_PRIORITY
        except:
             flag_enum = Enum.ToObject(self.SC_PERIOD, 16) # 16 = GROUPID_IS_PRIORITY
        
        self.simconnect.TransmitClientEvent(
            UInt32(SIMCONNECT_OBJECT_ID_USER),
            evt_enum,
            UInt32(data),
            group_enum,
            flag_enum
        )
        return True

    def exit(self):
        self.quit = 1
        if self.simconnect:
            self.simconnect.Dispose()
            self.simconnect = None

    # --- Compatibility Helpers for RequestList.py ---
    
    def IsHR(self, hr, value):
        # Managed calls throw exceptions, so if we get here, it succeeded.
        return True

    def add_data_definition(self, define_id, datum_name, units, datum_type, epsilon, datum_id):
        # Wrapper for AddToDataDefinition
        # define_id: ID object or int
        if hasattr(define_id, 'value'): define_id = define_id.value
        
        # Cast define_id to an Enum (we use SC_DATATYPE as a proxy type)
        try:
            define_id = Enum.ToObject(self.SC_DATATYPE, int(define_id))
        except:
            pass # Keep as is if fails
        
        if isinstance(datum_name, bytes): datum_name = datum_name.decode()
        if isinstance(units, bytes): units = units.decode()
        
        # Mapping datum_type (ctypes int) to Managed Enum
        # Original: FLOAT64 = 4 (usually)
        # Managed: SIMCONNECT_DATATYPE.FLOAT64
        
        m_type = self.SC_DATATYPE.FLOAT64
        # Basic mapping logic
        if str(datum_type) == "4": # SIMCONNECT_DATATYPE_FLOAT64
             m_type = self.SC_DATATYPE.FLOAT64
        elif hasattr(datum_type, 'real') and int(datum_type) == 4:
             m_type = self.SC_DATATYPE.FLOAT64
             
        try:
            self.simconnect.AddToDataDefinition(
                define_id,
                datum_name,
                units,
                m_type,
                epsilon,
                SIMCONNECT_UNUSED
            )
            return 0 # S_OK
        except Exception as e:
            LOGGER.error(f"AddToDataDefinition Failed: {e}")
            return -1

    def clear_data_definition(self, define_id):
        if hasattr(define_id, 'value'): define_id = define_id.value
        try:
            define_id = Enum.ToObject(self.SC_DATATYPE, int(define_id))
            self.simconnect.ClearDataDefinition(define_id)
        except:
            pass
            
    def get_last_sent_packet_id(self):
        # Fake it
        return 0
