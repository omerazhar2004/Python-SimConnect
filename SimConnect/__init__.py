from .P3DSimConnect import P3DSimConnect as SimConnect, SIMCONNECT_UNUSED as DWORD
def millis(): return 0 # mock
from .RequestList import AircraftRequests, Request
from .EventList import AircraftEvents, Event
from .FacilitiesList import FacilitiesRequests, Facilitie


def int_or_str(value):
	try:
		return int(value)
	except TypeError:
		return value


__version__ = "0.4.26"
VERSION = tuple(map(int_or_str, __version__.split(".")))

__all__ = ["SimConnect", "Request", "Event", "millis", "DWORD", "AircraftRequests", "AircraftEvents", "FacilitiesRequests"]
