import os
import glob
import shutil

def teardown():

    print('Start Teardown')

    windowsUser = os.getlogin()

    #Had to add the \\ after C:.  os.path.join wasn't inserting a \ between C: and Users for some reason.
    for path in glob.iglob(os.path.join('C:\\', 'Users', windowsUser, 'AppData', 'Local', 'Temp', 'scoped_dir*')):
        print(path)
        shutil.rmtree(path)

    for path in glob.iglob(os.path.join('C:\\', 'Users', windowsUser, 'AppData', 'Local', 'Temp', 'chrome_BITS_*')):
        print(path)
        shutil.rmtree(path)

    print('Teardown complete')
