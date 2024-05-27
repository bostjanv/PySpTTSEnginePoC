# https://pythonhosted.org/comtypes/server.html

# generate wrapper code for the type library, this needs
# to be done only once (but also each time the IDL file changes)
from comtypes.client import GetModule
GetModule("pysapi.tlb")
