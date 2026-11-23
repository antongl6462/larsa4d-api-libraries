import win32com.client
from win32com.client import gencache

application = win32com.client.Dispatch("Larsa2000.Application")
# Generate LARSA constants through the running app and keep element objects late-bound.
gencache.EnsureDispatch(application.CreateObject("LarsaElements", "clsMemberLoad")[0])
gencache.GetClassForCLSID = lambda clsid: None

def CreateObject(class_name):
    library_name, class_name = class_name.split(".")
    return application.CreateObject(library_name, class_name)[0]

project = application.project

joint1 = CreateObject("LarsaElements.clsJoint")
joint1.location.setCoordinates(0, 0, 0)
project.joints.append(joint1)

joint2 = CreateObject("LarsaElements.clsJoint")
joint2.location.setCoordinates(10, 0, 0)
project.joints.append(joint2)

member = CreateObject("LarsaElements.clsMember")
member.Setjoint(1, joint1)
member.Setjoint(2, joint2)
project.members.append(member)

lc = CreateObject("LarsaElements.clsLoadCase")
project.primaryLoadCases.append(lc)
mload = CreateObject("LarsaElements.clsMemberLoad")
lc.MemberLoads.append(mload)
mload.member = member
mload.loadType = win32com.client.constants.UniformForce
mload.startW = 10

application.DataHasChanged(win32com.client.constants.DCM_UPDATE, None)
