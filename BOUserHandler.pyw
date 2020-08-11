#! python3
# BOUserHandler.pyw - contains the UserProfile class as well as several functions used for handling
# user profiles for generateBO.py

import os, shelve

class UserProfile:
    
    # This class contains attributes for user preferences, namely:
    # - username
    # - paths for BO the input file, area assignment file and the BO folder
    # - zpp_mpl view name
    # - whether the BO report should be run by PSP or material indices
    
    def __init__(self, username, inputFilePath, areaFilePath, targFolder, viewName, byMaterial):
        self.username = username
        self.inputFilePath = inputFilePath
        self.areaFilePath = areaFilePath
        self.targFolder = targFolder
        self.viewName = viewName
        self.byMaterial = byMaterial

        self.deleteE = False
        self.deleteEmptySpm = True
        self.copyDeliveryDates = True
        self.progressBar = False
        self.notifyAfterCompletion = True
        
    def __str__(self):
        return "BO profile: %s" %(self.username)

    
    def updateSecondaryParams(self, _deleteE, _deleteEmptySpm, _copyDeliveryDates, _progressBar, _notifyAfterCompletion):
        self.deleteE = _deleteE
        self.deleteEmptySpm = _deleteEmptySpm
        self.copyDeliveryDates = _copyDeliveryDates
        self.progressBar = _progressBar
        self.notifyAfterCompletion = _notifyAfterCompletion


def initializeUsers():
    # Checks if user profile data file exists in '.\BOuserData' - if not, creates it
    # Reads a list of existing user profiles and returns it
    # If no users exists yet, returns []

    if not os.path.exists(r'.\BOuserdata'):
        os.makedirs(r'.\BOuserdata')
    if not os.path.exists(r'.\BOuserdata\userdata.dat'):
        shelfFile = shelve.open(r'.\BOuserdata\userdata')
        shelfFile.close()
    shelfFile = shelve.open(r'.\BOuserdata\userdata')
    if 'users' in shelfFile:
        return shelfFile['users']
    else:
        return []


def isValidUser(u):
    stringParams = [u.username, u.inputFilePath, u.areaFilePath, u.targFolder, u.viewName]
    pathParams = [u.inputFilePath, u.areaFilePath, u.targFolder]
    if any([not isinstance(param, str) or not len(param)>0 for param in stringParams]):
        return False
    if any([not os.path.exists(param) for param in pathParams]):
        return False
    else:
        return True


def saveUser(userProfile):
    # Saves a new user profile to the shelf file
    # Handles the file not existing, this being the first user profile etc.
    
    if not os.path.exists(r'.\BOuserdata'):
        os.makedirs(r'.\BOuserdata')
    if not os.path.exists(r'.\BOuserdata\userdata.dat'):
        shelfFile = shelve.open(r'.\BOuserdata\userdata')
        shelfFile.close()
    shelfFile = shelve.open(r'.\BOuserdata\userdata')
    if 'users' in shelfFile:
        userList = shelfFile['users']
        userList.append(userProfile)
        shelfFile['users'] = userList
    else:
        userList = []
        userList.append(userProfile)
        shelfFile['users'] = userList
        

def deleteUserByName(userProfile):                            
    # Deletes a user profile with the same name from the shelf file
    # Returns True if a profile was deleted, otherwise returns False
    if not os.path.exists(r'.\BOuserdata'):
        os.makedirs(r'.\BOuserdata')
    if not os.path.exists(r'.\BOuserdata\userdata.dat'):
        shelfFile = shelve.open(r'.\BOuserdata\userdata')
        shelfFile.close()
    shelfFile = shelve.open(r'.\BOuserdata\userdata')
    if 'users' in shelfFile:
        usernames = [user.username for user in shelfFile['users']]
        if userProfile.username in usernames:
            userList = shelfFile['users'][:]    # This copies the list similar to list.copy() or list(old_list)
            del userList[usernames.index(userProfile.username)]
            shelfFile['users'] = userList
        shelfFile.close()
        return True
    else:
        shelfFile['users'] = []
        shelfFile.close()
        return False

