
from all_imports import *

csv_path = f'D:\Plag\plag({today}).csv'


def uploadfile_long(text):
    print('[Warning]: Long Method is starting')
    sleep(1)
    i = int(text)+1
    mtq = getVars(0)
    ask = mtq.split('|')[1]
    while i < 100:
        if len(os.listdir('D:\Plag\plagdir')) != 0:
            try:
                try:
                    ap = int(''.join(x for x in getCsv(0) if x.isdigit()))
                    ln = int(''.join(x for x in getCsv(-1) if x.isdigit()))+1
                except Exception:
                    ap = 0
                    ln = -2
                if ln > int(getVars(3)):
                    ln = 1
                if i > int(text):
                    if ap == ln:
                        if getVars(2) == 'y':
                            sleep(2)
                            Download_plag()
                        elif getVars(2) == 'n':
                            driver.refresh()
                            report()
                        i = 0
                    elif i > int(getVars(3))+1:  # Restart when all slots are full
                        i = 0
                else:
                    if i > int(getVars(3))+1:  # Restart when all slots are full
                        i = 0
                tab(0)
                WebDriverWait(driver, 300).until(EC.element_to_be_clickable(
                    (By.XPATH, f'/html/body/div[2]/div[4]/div[2]/div[2]/div[3]/table/tbody/tr[{i}]/td[5]/a[1]')))
                driver.find_element(
                    by=By.XPATH, value=f'/html/body/div[2]/div[4]/div[2]/div[2]/div[3]/table/tbody/tr[{i}]/td[5]/a[1]').click()
                # /html/body/div[2]/div[4]/div[2]/div[2]/div[3]/table/tbody/tr[3]/td[5]/a[1]
                # /html/body/div[2]/div[4]/div[2]/div[2]/div[3]/table/tbody/tr[4]/td[5]/a[1]
                try:
                    Alert(driver).accept()
                except Exception:
                    pass
                driver.find_element(
                    by=By.XPATH, value=f'//*[@id="choose-file-btn"]').click()
                path = os.listdir('D:\Plag\plagdir')[0]
                writenme('open', path)
                wait('//*[@id="selected-file-name"]')
                file = driver.find_element(
                    by=By.XPATH, value=f'//*[@id="selected-file-name"]').get_attribute('innerText')
                wait('//*[@id="title"]')
                driver.find_element(
                    by=By.XPATH, value=f'//*[@id="title"]').send_keys(file.removesuffix('.docx'))
                sleep(2)
                wait('//*[@id="upload-btn"]')
                driver.find_element(
                    by=By.XPATH, value=f'//*[@id="upload-btn"]').click()

                try:
                    wait('//*[@id="submission-metadata-assignment"]')
                except Exception:
                    break

                ###################################################
                Slot_title = driver.find_element(
                    by=By.XPATH, value=f'//*[@id="submission-metadata-assignment"]').get_attribute('innerText').capitalize()
                submit_title = driver.find_element(
                    by=By.XPATH, value=f'//*[@id="submission-metadata-title"]').get_attribute('innerText')
                a = os.listdir('D:\Plag\plagdir')
                print(f'{submit_title},{Slot_title},Remaining_file: {len(a)}')
                csvwr(f'{submit_title},{Slot_title}')
                wrreplace('D:\Plag\Vars.txt', getVars(1), Slot_title)
                ###################################################

                WebDriverWait(driver, 250).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm-btn"]')))
                driver.find_element(
                    by=By.XPATH, value=f'//*[@id="confirm-btn"]').click()
                os.remove(f"D:\Plag\plagdir\{path}")
                WebDriverWait(driver, 250).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="close-btn"]')))
                driver.find_element(
                    by=By.XPATH, value=f'//*[@id="close-btn"]').click()
            except Exception as e:
                print('[Long method error - {}]: '.format(sys.exc_Info()
                      [-1].tb_lineno), type(e).__name__, e)
        else:
            print('[Info]: NO FILES ARE IN PLAGDIR')
            if getVars(2) == 'y':
                driver.refresh()
                sleep(2)
                Download_plag()
            elif getVars(2) == 'n':
                driver.refresh()
                report()
            break
        i += 1


def renameFile():  # Remove ( and , from plag files
    try:
        mks = os.listdir(r'D:\Plag\plagdir')
        for i in range(len(mks)):
            try:
                mdir = os.listdir('D:\Plag\plagdir')[i]
                path = f'D:\Plag\plagdir\\{mdir}'
                fixer = f'D:\Plag\plagdir\\{mdir.replace("(","").replace(")","").replace(",","").replace("‐","" "")}'
                os.rename(path, fixer)
                print('Renaming: ', path, fixer)
            except Exception:
                pass
        print('[Info]: File Renamed Successfully')
    except Exception:
        print('[Info]: File can"t be renamed')


def readcon():  # Reads CSV file
    with open(csv_path, 'r') as f:
        mnt = (f.read())
        f.close()
        print(mnt)
    return mnt


def wrreplace(cpth, search_text, replace_text):
    with open(cpth, 'r+') as f:
        file = f.read()
        file = re.sub(search_text, replace_text, file, 1)
        f.seek(0)
        f.write(file)
        f.truncate()


def getSlotnum():
    r = getVars(1)
    k = int(''.join(x for x in r if x.isdigit()))
    if k > 79:
        print(k)
        k = 0
    return k


def csvwr(cont):  # Updating the vars and Slots
    with open(csv_path, 'a', encoding='utf-8') as f:
        try:
            f.write(f'{cont}\n')
        except:
            pass
        f.close()


def wait(xpth):  # Selenium wait method
    return WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, xpth)))


def report():
    stats = None
    for x in range(2):

        ans = []
        i = 0  # USE FOR WEbsite
        j = 0  # USE FOR Excel
        while i < int(getVars(3))+2:
            try:
                percentage = driver.find_element(
                    by=By.XPATH, value=f'/html/body/div[2]/div[4]/div[2]/div[2]/div[3]/table/tbody/tr[{i}]/td[4]/span/a/span[1]').get_attribute('innerText')
                slot = driver.find_element(
                    by=By.XPATH, value=f'/html/body/div[2]/div[4]/div[2]/div[2]/div[3]/table/tbody/tr[{i}]/td[1]/div').get_attribute('innerText').capitalize()
                num = j
                for x in getCsv(num):
                    if x == "%":
                        stats = 'Already Percentage'
                        i -= 1
                        j += 1
                if getCsv(num) == slot:
                    ans.append('1')
                    if int(percentage.removesuffix('%')) < 20:
                        formUla = 'Accepted'
                    if int(percentage.removesuffix('%')) >= 20:
                        formUla = 'Rejected'
                    alternate = f'{percentage},{formUla}'
                    wrreplace(csv_path, getCsv(num), alternate)
                    print(f'[Replace]: from {slot} to {percentage}')
                else:
                    print(
                        f'[csvstatus - {getCsv(num)}]: Slot: {slot} and Status: {stats}')
                    j -= 1
                j += 1
            except Exception:
                pass
            i += 1
    try:
        if len(ans) == get_number_oflines(csv_path):
            pass

        else:
            print('Checked Files = {0}'.format(len(ans)))
            driver.refresh()
    except Exception:
        pass


def writenme(cys, nme):
    time.sleep(2)
    # This is for window na
    # me
    # win32gui.GetWindowText (win32gui.GetForegroundWindow())
    shell = win32com.client.Dispatch('WScript.Shell')
    hWnd = win32gui.FindWindow(None, cys)

    try:
        win32gui.ShowWindow(hWnd, 5)
        win32gui.SetForegroundWindow(hWnd)
    except Exception:
        pass
    shell.SendKeys(nme)
    shell.SendKeys('{Enter}')


def minize(cls):
    # win32gui.GetWindowText(win32gui.GetForegroundWindow())
    Minimize = win32gui.FindWindow(None, cls)
    win32gui.ShowWindow(Minimize, win32con.SW_MINIMIZE)


def Download_plag():
    i = 1  # Loop
    j = 1  # Iter in excel
    for x in range(2):
        while i < int(getVars(3))+2:
            try:
                # Open plag.csv and get percentage or Slot of file
                with open(csv_path) as f:
                    serch = f.readlines()[j-1].split(',')[1]
                    f.close()
                if serch == 'Downloaded':
                    stats = 'Checked'
                else:
                    stats = 'Not checked'
                if stats != 'Checked':
                    search = int(''.join(x for x in serch if x.isdigit()))
                    tab(0)
                    file_sot = driver.find_element(
                        by=By.XPATH, value=f'/html/body/div[2]/div[4]/div[2]/div[2]/div[3]/table/tbody/tr[{search+1}]/td[1]').get_attribute('innerText').capitalize()
                    # Iter in TUrnitin of slots bvalue only
                    file_slot = int(
                        ''.join(x for x in file_sot if x.isdigit()))
                    if search != 0:
                        if search == file_slot:
                            actions = ActionChains(driver)
                            biew = driver.find_element(
                                by=By.XPATH, value=f'/html/body/div[2]/div[4]/div[2]/div[2]/div[3]/table/tbody/tr[{search+1}]/td[5]/a[2]')
                            actions.key_down(Keys.CONTROL).click(biew).key_up(
                                Keys.CONTROL).perform()  # Clicks on view to download
                            sleep(1)
                            tab(1)
                            WebDriverWait(driver, 100).until(EC.presence_of_element_located(
                                (By.XPATH, '/html/body/div/div[1]/aside/div[1]/section[4]/div/div[1]')))

                            driver.find_element(
                                by=By.XPATH, value='/html/body/div/div[1]/aside/div[1]/section[4]/div/div[1]').click()
                            WebDriverWait(driver, 100).until(EC.presence_of_element_located(
                                (By.XPATH, '/html/body/div[3]/div[2]/div[2]/div/div[1]/div[2]')))
                            sleep(4)
                            findxpth(
                                '/html/body/div[3]/div[2]/div[2]/div/div[1]/div[2]').click()
                            nmeoffile = driver.find_element(
                                by=By.XPATH, value='/html/body/div/header/h1/span[2]').get_attribute('innerHTML')
                            try:
                                percentag = driver.find_element(
                                    by=By.XPATH, value='/html/body/div/div[1]/aside/div[1]/section[2]/div[2]/div[1]/label').get_attribute('innerHTML')
                                percentage = '{0}%'.format(percentag)
                                if len(percentag) != 0:
                                    wrreplace(csv_path, file_sot,
                                              f'Downloaded,{file_sot},{percentage}')
                                    WebDriverWait(driver, 100).until(EC.invisibility_of_element(
                                        (By.XPATH, '/html/body/div[3]/div[2]/div[1]/div')))
                                    sleep(1)
                            except Exception:
                                percentage = 'Processing'
                            if percentag == '':
                                percentage = 'Processing'
                            print(
                                f'[Info]: Downloading({search}): {nmeoffile} - {percentage}')
                            j += 1
                        else:
                            print(f'[Info]: Slot Target: {search} {file_slot}')
                    else:
                        j += 1
                else:
                    j += 1
            except Exception as e:
                print(e)
                if len(driver.window_handles) == 2:
                    driver.close()
                break
            i += 1


# After downloading files we try to make archive of it but fst we move all old downloaded files to diffrent folder

    try:
        print('[Info]: Files will zipped in 8secs')
        sleep(8)
        try:
            # Move Zip (if avail) to Rbin
            move(r'D:\Plag\Download',r'D:\Plag\Backup\Rbin\Download.zip')
        except Exception:
            pass

        #if files are more than 3 then ZIP them
        if len(os.listdir(r'D:\Plag\Download Plag')) > 3:
            with zipfile.ZipFile(r'D:\Plag\Download\Download.zip', 'w', zipfile.ZIP_DEFLATED) as zipf:
                zipdir('D:\Plag\Download', zipf)
                print("Files are zipped !")
        else:
            n_fil = len(os.listdir(r'D:\Plag\Download'))
            print(f'[Can"t Zip]: {n_fil}')
    except Exception:
        print("Files can't be zipped!")
        pg.alert('Is Download Folder in dirc is presenet or not..?')



def login():
    try:
        print('[Info]: Logging in..')
        try:
            driver.find_element(
                by=By.XPATH, value='/html/body/header/div[2]/section[2]/div/div[2]/div/a[2]').click()
        except Exception:
            pass
        try:
            Wait('//*[@id="ibox_form_body"]/div[3]/input')
            driver.find_element(
                by=By.XPATH, value='//*[@id="ibox_form_body"]/div[3]/input').click()
        except Exception:
            pass
        sleep(1)
        try:
            for k in range(1, 5):
                try:
                    ntp = driver.find_element(
                        by=By.XPATH, value=f'/html/body/div[2]/form/div[2]/div[3]/div/div[2]/table/tbody/tr[{k}]/td[2]/a').get_attribute('innerText')
                    print('[Opening]:', ntp)
                    if ntp == getVars(4):
                        break
                    Wait(
                        f'/html/body/div[2]/form/div[2]/div[3]/div/div[2]/table/tbody/tr[{k}]/td[2]/a')
                except Exception:
                    pass
            driver.find_element(
                by=By.XPATH, value=f'/html/body/div[2]/form/div[2]/div[3]/div/div[2]/table/tbody/tr[{k}]/td[2]/a').click()
        except Exception:
            pass
    except Exception:
        pass


def uploadfile(text):
    print('[Info]: Short Method')
    idd = int(pageid)
    i = int(text)
    mtq = getVars(0)
    ask = mtq.split('|')[1]
    renameFile()
    minz = mtq.split('|')[0]

    try:
        readcon()
    except Exception:
        pass
    sleep(0.5)
    # //*[@id="assignment_122763293"]/td[1]/div
    tab(0)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.XPATH, f'/html/body/div[2]/div[4]/div[2]/div[2]/div[3]/table/tbody/tr[{i}]/td[5]/a[1]')))
    driver.find_element(
        by=By.XPATH, value=f'/html/body/div[2]/div[4]/div[2]/div[2]/div[3]/table/tbody/tr[{i}]/td[5]/a[1]').click()
    try:
        Alert(driver).accept()
    except Exception:
        pass
    lnk = driver.current_url
    lnkid = lnk.split('aid=')[1].split('&')[0]
    print(lnkid)
    while i < int(getVars(3))+1:
        if os.listdir('D:\Plag\plagdir') != []:
            try:
                tab(0)
                sleep(1)
                driver.find_element(
                    by=By.XPATH, value=f'//*[@id="choose-file-btn"]').click()
                path = os.listdir(r'D:\Plag\plagdir')[0]
                try:
                    writenme('open', path)
                except Exception:
                    print('[Info]: Retry..')
                    writenme('open', path)
                wait('//*[@id="selected-file-name"]')
                file = driver.find_element(
                    by=By.XPATH, value=f'//*[@id="selected-file-name"]').get_attribute('innerText')
                wait('//*[@id="title"]')
                driver.find_element(
                    by=By.XPATH, value=f'//*[@id="title"]').send_keys(file.removesuffix('.docx'))
                sleep(2)
                if minz == 'Yes':
                    minize('Turnitin - Google Chrome')
                wait('//*[@id="upload-btn"]')
                driver.find_element(
                    by=By.XPATH, value=f'//*[@id="upload-btn"]').click()
                wait('//*[@id="submission-metadata-assignment"]')

                ###################################################
                Slot_title = driver.find_element(
                    by=By.XPATH, value=f'//*[@id="submission-metadata-assignment"]').get_attribute('innerText').capitalize()
                submit_title = driver.find_element(
                    by=By.XPATH, value=f'//*[@id="submission-metadata-title"]').get_attribute('innerText')
                remaint = os.listdir('D:\Plag\plagdir')
                try:
                    print(
                        f'[Info]: {submit_title},{Slot_title},Remaining_file: {len(remaint)}')
                except Exception:
                    pass
                csvwr(f'{submit_title},{Slot_title}')
                wrreplace('D:\Plag\Vars.txt', getVars(1), Slot_title)
                ###################################################

                wait('//*[@id="confirm-btn"]')
                driver.find_element(
                    by=By.XPATH, value=f'//*[@id="confirm-btn"]').click()
                move(f"D:\Plag\plagdir\\{path}",
                     f"D:\Plag\Backup\Rbin\\{path}")
                WebDriverWait(driver, 250).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="close-btn"]')))
                sleep(1)
                if i == int(getVars(3)):
                    i = 0
                if len(remaint) == 1:
                    driver.find_element(
                        by=By.XPATH, value=f'//*[@id="close-btn"]').click()
                else:
                    try:
                        lnp = int(idd + i)
                        lnt = lnk.replace(str(lnkid), str(lnp))
                        driver.get(lnt)
                    except Exception as e:
                        print('Can"t Move to another', e)
                        break
            except Exception as e:
                print(
                    'UPLOAD ERROR IN LINE- [{}]: '.format(sys.exc_Info()[-1].tb_lineno), type(e).__name__, e)
                i += 1
                break
        else:
            print('[Info]: NO Files Are in Plagdir')
            if getVars(2) == 'y':
                driver.refresh()
                sleep(2)
                Download_plag()
            else:
                driver.refresh()
                report()
            break
        i += 1


def findg(search):
    for i in range(2, 100):
        try:
            kml = driver.find_element(
                by=By.XPATH, value=f'/html/body/div[2]/div[4]/div[2]/div[2]/div[3]/table/tbody/tr[{i}]/td[5]/span/ul/li[1]/script').get_attribute('innerHTML')
            it = (kml.find('fn'))
            ot = kml.find('type')
            fn = kml[it+5:ot-3]
            try:
                slotpe = driver.find_element(
                    by=By.XPATH, value=f'/html/body/div[2]/div[4]/div[2]/div[2]/div[3]/table/tbody/tr[{i}]/td[4]/span/a/span[1]').get_attribute('innerText')
                slotper = '{0}'.format(slotpe)
            except Exception:
                slotper = 'processing'
            file_name = fn
            file_slot = driver.find_element(
                by=By.XPATH, value=f'/html/body/div[2]/div[4]/div[2]/div[2]/div[3]/table/tbody/tr[{i}]/td[1]').get_attribute('innerText')

            if search.replace(" ", "_").capitalize() == file_name[0:len(search)].capitalize().removesuffix('.docx'):
                print(
                    f'[FOUND]: found[{i-1}]: {search} and file_name: {file_name[0:len(search)]},{file_slot},{slotper}')
                csvwr(f'{file_name.removesuffix(".docx")},{file_slot},{slotper}')
            else:
                # print(f'[Info]: search[{i-1}]: {search} and file_name: {file_name[0:len(search)]},{file_slot},{slotper}')
                pass
            with open(r'D:\Plag\plag_Info.csv', 'a') as k:
                k.write(
                    f'{file_name.removesuffix(".docx")},{file_slot},{slotper}\n')
                k.close()
        except Exception as e:
            print(e)
            print('[Info]: Searched Successfully')
            break


def rord():
    try:
        for x in (os.listdir(r'D:\Plag\Raw')):
            if int(x.split('.')[0]):
                if getVars(2) == 'y':
                    wrreplace(r"D:\Plag\Vars.txt", 'y', 'n')
                    print('[Info]: ReadMode On')
    except Exception as e:
        if getVars(2) == 'n':
            wrreplace(r"D:\Plag\Vars.txt", 'n', 'y')
            print('[Info]: DownloadingMode On')


rord()
try:
    move_all_doc()
    remove_emty_folder()
except Exception as e:
    print(e)


login()
if os.listdir('D:\Plag\plagdir') != []:
    if getVars(6) == '143':
        sl = getVars(6)
    else:
        sl = 'Plag Check'
elif getVars(2) == '':
    sl = pg.confirm(text=f'Choose Job[{int(getSlotnum())+1}]', title='Choose Job: ',
                    buttons=['Plag Check', 'Report', 'Download', 'Find Files'])
else:
    sl = getVars(2)
try:
    pagei = driver.find_element(
        by=By.XPATH, value='/html/body/div[2]/div[4]/div[2]/div[2]/div[3]/table/tbody/tr[2]/td[2]/button').get_attribute('id')
    pageid = ''.join(x for x in pagei if x.isdigit()
                     )  # This is for integer only
except Exception:
    print('[Info]: Not on main Page')
    pagei = driver.current_url
    lnkid = pagei.split('=')[1].split('&')[0]
    pageid = lnkid - getSlotnum()

if sl == 'Plag Check':
    uploadfile(getSlotnum()+1)

elif sl == 'Report' or sl == 'n':
    x = 'THIS IS REPORT CODE'
    print(str(x).rjust(10))
    report()

elif sl == 'Download' or sl == 'y':
    Download_plag()
    pg.alert('All Files are downloaded')

elif sl == '143':
    uploadfile_long(getSlotnum()+1)

else:
    search = sl
    findg(search)
