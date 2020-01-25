###########################################################################################################################################
# listOutlookFolders
#
#   -Simple script to go through the outlook and find the different sections
#    within Microsoft outlook (ie inbox,outbox,drafts, etc.)
#   -Used as a testing method of obtaining the needed folder to perform my task
#
#
###########################################################################################################################################
from win32com.client import constants
from win32com.client.gencache import EnsureDispatch as Dispatch

outlook = Dispatch("Outlook.Application")
mapi = outlook.GetNamespace("MAPI")

class OutLookIndex():
    def __init__(self, outlook_object):
        self._obj = outlook_object

    def items(self):
        array_size = self._obj.Count
        for item_index in range(1,array_size+1):
            yield (item_index, self._obj[item_index])

    def prop(self):
        return sorted( self._obj._prop_map_get_.keys() )

for inx, folder in OutLookIndex(mapi.Folders).items():
    # iterate all Outlook folders (top level)
    print ("-"*70)
    print (folder.Name)

    for inx,subfolder in OutLookIndex(folder.Folders).items():
        print ("(%i)" % inx, subfolder.Name,"=> ", subfolder)


