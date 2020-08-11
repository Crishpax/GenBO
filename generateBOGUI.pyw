#! python3
# generateBOGUI.py - a GUI for generateBO.py made using PySimpleGui

#====================================================== USING ======================================================


import generateBO as BO
import PySimpleGUI as sg
import BOUserHandler as uh
import os

#============================================== FUNCTION DEFINITIONS ==============================================


def formatPath(path):
    while True:
        try:
            index = path.index('/')
            path = path[:index]+'\\'+path[index+1:]
        except ValueError:
            break
    return path

#============================================== BODY ==============================================


sg.change_look_and_feel('Reddit')
userList = uh.initializeUsers()
userNames = [user.username for user in userList]
newUserTxt = 'Nowy użytkownik'
userNames.append(newUserTxt)



while True:
    userSelectLayout = [
                [sg.Txt('Wybierz profil użytkownika', size=(30,2), auto_size_text=False)],
                [sg.DropDown(userNames, default_value=userNames[0], size=(30,6), readonly=True)],
                [sg.OK(), sg.OK('Usuń')]
                ]
    window = sg.Window('GenBO').Layout(userSelectLayout)
    event, values = window.Read()

    if event in ['Exit', None]:
        currentUserName = None
        window.close()
        break
    if event == 'OK':
        currentUserName = values[0]
        window.close()
        break

    elif event is 'Usuń':
        currentUserName = values[0]
        if currentUserName != newUserTxt:
            userObj = userList[userNames.index(currentUserName)]
            print(userObj)
            confirmLayout = [
                        [sg.Txt('Usunąć użytkownika "%s"?' %(userObj.username), size=(15,2), auto_size_text=False)],
                        [sg.Yes('Tak'), sg.No('Nie')]
                        ]
            window1 = sg.Window('GenBO').Layout(confirmLayout)
            event1, values = window1.Read()
            print(event1)
            window1.close()
            
            
            if event1 in ['Yes', 'Tak']:
                uh.deleteUserByName(userObj)
                userList = uh.initializeUsers()
                userNames = [user.username for user in userList]
                userNames.append(newUserTxt)
        
    window.close()


configFile, areaFile, targFolder, viewName, userName = '', '', '', '', '' # Initialize core variables as empty strings

while True:
    if currentUserName == newUserTxt:                   # New user profile setup

        matPSPLayout = [
                    [sg.Txt('Wykonujesz zrzut BO po PSP projektów, czy po indeksach materiałów?', size=(30,3), auto_size_text=False)],
                    [sg.Radio('Po PSP', 'RADIO1', default=True, size=(10,3)),
                     sg.Radio('Po materiale', 'RADIO1', size=(10,3))],
                    [sg.OK()]
                    ]

        window = sg.Window('GenBO').Layout(matPSPLayout)

        event, values = window.Read()
    
        if values[0] == True:
            byMaterial = False
            fieldText = 'Plik PSP/SPM'
            matPSPDesc = r'PSP projektów, dla których wykonujesz zrzut (np. 118/542734400/*)' 
        if values[1] == True:
            byMaterial = True
            fieldText = 'Plik materiał/SPM'
            matPSPDesc = r'indeksy materiałów produkowanych, dla których wykonujesz zrzut' 

        descText = 'Wybierz plik .xlsx, który zawiera:\n - w kolumnie A {}\n - w kolumnie B SpM produktów (np. T2E*)'.format(matPSPDesc)
        window.close()

        # Select input config file
    
        setupLayout1 = [
            [sg.Txt(descText, size=(65, 4), auto_size_text=False)],
            [sg.Text(fieldText, size=(15, 1), auto_size_text=False, justification='right'),      
             sg.InputText(configFile), sg.FileBrowse('Arkusz .xlsx')],
                  [sg.OK()]]
                        
                
        window = sg.Window('GenBO').Layout(setupLayout1)
        while True:
            event, values = window.Read()
            if event is 'OK':
                if os.path.exists(values[0]) and os.path.splitext(values[0])[1] in ['.xlsx', '.xls', 'xlm']:
                    configFile = values[0]
                    break
                if not os.path.exists(values[0]):
                    sg.Popup('Wybrany plik nie istnieje', title='')
                if not os.path.splitext(values[0])[1] in ['.xlsx', '.xls', 'xlm']:
                    sg.Popup('Niepoprawne rozszerzenie pliku')
            if event in [None, 'Exit']:
                break

        window.close()
    

        # Select area file
    
        descText = 'Wybierz plik .xlsx, który zawiera:\n - w kolumnie A indeksy materiałów\n - w kolumnie B przypisanie materiału do obszaru'
        fieldText = 'Plik materiał/obszar'
    
        setupLayout2 = [
        [sg.Txt(descText, size=(65, 4), auto_size_text=False)],
        [sg.Text(fieldText, size=(15, 1), auto_size_text=False, justification='right'),      
         sg.InputText(areaFile), sg.FileBrowse('Arkusz .xlsx')],
              [sg.OK()]]
    
        if os.path.exists(configFile):
            window = sg.Window('GenBO').Layout(setupLayout2)
        while True:
            event, values = window.Read()
            if event is 'OK':
                if os.path.exists(values[0]) and os.path.splitext(values[0])[1] in ['.xlsx', '.xls', 'xlm']:
                    areaFile = values[0]
                    break
                if not os.path.exists(values[0]):
                    sg.Popup('Wybrany plik nie istnieje', title='')
                if not os.path.splitext(values[0])[1] in ['.xlsx', '.xls', 'xlm']:
                    sg.Popup('Niepoprawne rozszerzenie pliku')
            if event in [None, 'Exit']:
                break    

        window.close()

        # Select target BO folder

        descText = 'Wybierz folder, w którym znajdują się pliki BO'
        fieldText = 'Folder BO'
    
        folderSetupLayout = [
        [sg.Txt(descText, size=(65, 2), auto_size_text=False)],
        [sg.Text(fieldText, size=(10, 1), auto_size_text=False, justification='right'),      
         sg.InputText(targFolder), sg.FolderBrowse('Folder BO')],
              [sg.OK()]]

        if os.path.exists(areaFile):
            window = sg.Window('GenBO').Layout(folderSetupLayout)
        while True:
            event, values = window.Read()
            if event is 'OK':
                if os.path.isdir(values[0]):
                    targFolder = values[0]
                    break
                else:
                    sg.Popup('Wybrany folder nie istnieje', title='')
            if event in [None, 'Exit']:
                break    
        window.close()

    # Put in zpp_mpl view name

        descText = 'Wybierz układ dla ZPP_MPL'
        fieldText = 'Nazwa układu'
        defView = r'/BO_PL08'
        textSetupLayout = [
        [sg.Txt(descText, size=(25, 2), auto_size_text=False)],
        [sg.Radio(r'Standardowy (%s)' %defView, 'RADIO2', default = True, size=(18,2))],
        [sg.Radio('Inny:', 'RADIO2', size = (5,2)), sg.InputText(size=(13,2))],
              [sg.OK()]]

        if os.path.exists(targFolder):
            window = sg.Window('GenBO').Layout(textSetupLayout)
        event, values = window.Read()
        if event is 'OK':
            if values[0]:
                viewName = defView
            else:
                viewName = values[2]
            print(viewName)

    # Put in user profile name

        configFile = formatPath(configFile)
        areaFile = formatPath(areaFile)
        targFolder = formatPath(targFolder)
        window.close()

        descText = 'Wprowadź nazwę profilu dla tych ustawień'
        fieldText = 'Nazwa profilu'
        nameSetupLayout = [
        [sg.Txt(descText, size=(50, 2), auto_size_text=False)],
        [sg.Text(fieldText, size=(10, 1), auto_size_text=False, justification='right'),      
         sg.InputText(userName)],
              [sg.OK()]
                        ]
        if viewName is not '':
            window = sg.Window('GenBO').Layout(nameSetupLayout)
        warned = False
        while True:
            event, values = window.Read()
            if event is 'OK':
                if values[0] not in userNames or warned == True:
                    userName = values[0]
                    break
                elif warned == False:
                    sg.Popup('Istnieje już taki profil - dane zostaną nadpisane', title='')
                    warned = True
            if event in [None, 'Exit']:
                break
        currentUser = uh.UserProfile(userName, configFile, areaFile, targFolder,viewName, byMaterial)
        userText = 'Ustawienia dla profilu {}:'.format(currentUser.username)
        window.close()
    

    else:                                   # Grabbing values from an existing user profile
        if currentUserName == None:
            break
        currentUser = userList[userNames.index(currentUserName)]
        configFile = currentUser.inputFilePath
        areaFile = currentUser.areaFilePath
        targFolder = currentUser.targFolder
        viewName = currentUser.viewName
        byMaterial = currentUser.byMaterial
        userText = 'Ustawienia dla profilu {}:'.format(currentUser.username)


    if byMaterial:
        byMatTxt = 'Zrzut po: materiałach'
        filePathTxt = 'Ścieżka pliku mat/SPM: '
    else:
        byMatTxt = 'Zrzut po: PSP'
        filePathTxt = 'Ścieżka pliku PSP/SPM: '

    standardEndDate = BO.getSAPDateFormat(BO.calculateFutureBTDate())
    
    confirmationLayout =[
                    [sg.Txt(userText, size=(80,3), justification='center')],
                    [sg.Txt(byMatTxt[:9], size=(30,2)), sg.Txt(byMatTxt[10:], size=(50,2))],
                    [sg.Txt(filePathTxt, size=(30,2)), sg.Txt(configFile, size=(50,2))],
                    [sg.Txt('Ścieżka pliku obszarów: ', size=(30,2)), sg.Txt(areaFile, size=(50,2))],
                    [sg.Txt('Ścieżka folderu BO: ', size=(30,2)), sg.Txt(targFolder, size=(50,2))],
                    [sg.Txt('Nazwa układu: ', size=(30,3)), sg.Txt(viewName, size=(50,2))],
                    [sg.Checkbox('Usuń linie materiałów z nabyciem "E"', default=currentUser.deleteE)],
                    [sg.Checkbox('Usuń linie bez SpM', default=currentUser.deleteEmptySpm)],
                    [sg.Checkbox('Kopiuj daty dostaw z poprzedniego pliku', default=currentUser.copyDeliveryDates)],
                    [sg.Checkbox('Pokaż pasek zaawansowania raportu', default=currentUser.progressBar)],
                    [sg.Checkbox('Powiadom po wygenerowaniu pliku', default=currentUser.notifyAfterCompletion)],
                    [sg.Radio('2 tygodnie w przód (do %s)' %standardEndDate, 'RADIO2', default = True, size=(30,2))],
                    [sg.Radio('Do innej daty:', 'RADIO2', size = (10,2)), sg.InputText(size=(15,2))],
                    [sg.Txt('Wygenerować BO zgodnie z ustawieniami?', justification='center')],
                    [sg.Yes('Tak'), sg.No('Zmień'), sg.No('Zamknij')]
                    ]

    window = sg.Window('GenBO').Layout(confirmationLayout)
    event, values = window.Read()
    deleteE = values[0]
    deleteEmptySpm = values[1]
    deliveryDates = values[2]
    showProg = values[3]
    showPopup = values[4]
    customDate = values[6]
    print(values)
    if customDate == True:
        customDateObj = BO.getDatetimeFromSAPDate(values[7])    # If the user's custom date is not in the right format, getDatetimeFromSAPDate() will return None
                                                                # which - when passed to generateBO() - will run the report with the standard date
    window.close()

   
    if event in ['Tak', 'Yes', 'Zamknij', 'Exit', None] and uh.isValidUser(currentUser): # TODO poprawić
        if currentUser.username not in [user.username for user in userList]:
            currentUser.updateSecondaryParams(deleteE, deleteEmptySpm, deliveryDates, showProg, showPopup)
            uh.saveUser(currentUser)
        else:
            for user in userList:
                if user.username == currentUser.username:
                    currentUser.updateSecondaryParams(deleteE, deleteEmptySpm, deliveryDates, showProg, showPopup)
                    uh.deleteUserByName(user)
            uh.saveUser(currentUser)
        break

    if event in ['Zamknij', 'Exit', None] and not uh.isValidUser(currentUser):
        break
    else:
        if event in ['Tak', 'Yes'] and not uh.isValidUser(currentUser):
            sg.Popup('Wybrane pliki nie istnieją lub nie wprowadzono danych użytkownika', title='')
        currentUserName = newUserTxt
        userName = currentUser.username

    
    
if event in ['Tak', 'Yes']:
    failed = False
    if showProg == True:
    
        progressLayout = [[sg.Text('Generowanie raportu BO... Może to zająć kilka minut.', size=(50,2))],
                        [sg.Text('Start programu...', size=(50,2), key='progressText')],
                      [sg.ProgressBar(1, orientation='h', size=(50, 20), key='progress')],
                  ]
        window = sg.Window('GenBO', progressLayout, grab_anywhere=True).Finalize()
        progress_bar = window.FindElement('progress')
        progress_text = window.FindElement('progressText')
    
        progress_bar.UpdateBar(0, 7)
        progress_text.update('Pobieranie danych wejściowych...')
    try:    
        boInput = BO.getBOInput(configFile)
    except:
        if failed == False:
            sg.Popup('Błąd przy pobieraniu danych z pliku. Program zostanie zamknięty', title='Error')
            failed = True
        if showProg == True:
            window.close()

    if showProg == True:

        progress_bar.UpdateBar(1, 7)
        progress_text.update('Pobieranie przypisania materiałów do obszarów...')
        
    try:
        areaDict = BO.getMaterialArea(areaFile)
    except:
        if failed == False:
            sg.Popup('Błąd przy pobieraniu danych obszaru. Program zostanie zamknięty', title='Error')
            failed = True
        if showProg == True:
            window.close()
    
    if showProg == True:

        progress_bar.UpdateBar(2, 7)
        progress_text.update("Wykonywanie zrzutu 'zpp_mpl' i oczekiwanie na dane z SAP...")
        
    try:
        windows = BO.getSapWnd()
    except:
        if failed == False:
            sg.Popup('Błąd podczas przydzielania okna SAP. Program zostanie zamknięty', title='Error')
            failed = True
        if showProg == True:
            window.close()
    
    if showProg == True:

        progress_bar.UpdateBar(3, 7)
        progress_text.update('Wykonywanie zrzutu zpp_mpl i przetwarzanie danych z SAP...')

    if customDate is False:
        try:
            boData = BO.generateBO(windows[0], boInput[0], boInput[1], byMaterial, True, None, None, viewName, deleteE, deleteEmptySpm)
        except:
            if failed == False:
                sg.Popup('Błąd podczas raportu zpp_mpl. Program zostanie zamknięty', title='Error')
                failed = True
        if showProg == True:
            window.close()
    else:
        try:
            boData = BO.generateBO(windows[0], boInput[0], boInput[1], byMaterial, True, None, customDateObj, viewName, deleteE, deleteEmptySpm)
        except:
            if failed == False:
                sg.Popup('Błąd podczas raportu BO. Program zostanie zamknięty', title='Error')
                failed = True
            if showProg == True:
                window.close()
        
    if showProg == True:
        
        progress_bar.UpdateBar(4, 7)
        progress_text.update('Generowanie arkusza BO...')
        
    try:
        rawBOSheet = BO.createBOSheet(boData[0], boData[1])
    except:
        if failed == False:
            sg.Popup('Błąd przy tworzeniu arkusza BO. Program zostanie zamknięty', title='Error')
            failed = True
        if showProg == True:
            window.close()

    if showProg == True:
        
        progress_bar.UpdateBar(5, 7)
        progress_text.update("Pobieranie danych z coois...")
        
    try:
        cooisThread = boData[2][0]
        cooisThread.join()
        orderDict = boData[2][1].get()
    except:
        if failed == False:
            sg.Popup('Błąd podczas raportu coois. Program zostanie zamknięty', title='Error')
            failed = True
        if showProg == True:
            window.close()

    if showProg == True:
        
        progress_bar.UpdateBar(6, 7)
        progress_text.update('Formatowanie arkusza BO...')
        
    try:
        formattedBOSheet = BO.formatBO(rawBOSheet[0], rawBOSheet[1], orderDict, areaDict)
    except:
        if failed == False:
            sg.Popup('Błąd podczas formatowania arkusza. Program zostanie zamknięty', title='Error')
            failed = True
        if showProg == True:
            window.close()

    if showProg == True:
        
        progress_bar.UpdateBar(7, 7)
        progress_text.update('Zapisywanie arkusza BO...')
        
    try:
        generatedFile = BO.finishSaveBO(formattedBOSheet[0], formattedBOSheet[1], targFolder, deliveryDates)
    except:
        if failed == False:
            sg.Popup('Błąd podczas zapisywania arkusza. Program zostanie zamknięty', title='Error')
            failed = True
        if showProg == True:
            window.close()
    
    if showProg == True:
    
        window.close()

    if showPopup == True:
    
        finishedLayout = [
                        [sg.Txt(r'Wygenerowano:', size=(50, 2), auto_size_text=False, justification='left')],
                        [sg.Txt(r'%s' %(generatedFile), size=(50, 5), auto_size_text=False, justification='left')],
                        [sg.Txt(r'Otworzyć utworzony plik?', size=(50, 3), auto_size_text=False, justification='left')],
                        [sg.Yes('Tak'), sg.No('Nie')]
                        ]

        window = sg.Window('GenBO - gotowe!').Layout(finishedLayout)

        event, values = window.Read()
        if event in ['Tak', 'Yes']:
            os.startfile(generatedFile)
        window.close()

