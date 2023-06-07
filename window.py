# When Button is submit
# https://www.turnitin.com/t_submit.asp?r=82.0021475055917&svr=55&lang=en_us&aid=131445715

# When Button is resubmit
# https://www.turnitin.com/t_submit.asp?aid=131445713&svr=45&session-id=c6ebd78911ab4ae28be0cd4516c6542f&lang=en_us

from all_imports import *

csv_path = f'D:\Plag\plag({today}).csv'


def uploadfile(text):
    try:
        move(csv_path,f'D:\Plag\Backup\plag({today}).csv')
    except:
        pass

    # WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, f'/html/body/div[2]/div[4]/div[2]/div[2]/div[4]/table/tbody/tr[2]/td[5]/a[1]')))

    print('[Warning]: Normal Method is starting')

    box = []
    for i in range(1, int(getVars(3))+1):
        try:
            Uniqid = driver.find_element(
                by=By.XPATH, value=f'/html/body/div[2]/div[4]/div[2]/div[2]/div[4]/table/tbody/tr[{i+1}]/td[5]/a[1]').get_attribute('href').split(",")[-1].removesuffix(");")
            try:
                box.append(int(Uniqid))
            except:
                box.append(int(Uniqid.split("aid=")[-1]))
                # print(int(Uniqid.split("aid=")[-1]))

        except:
            break

    # print(f'[Link]: https://www.turnitin.com/t_submit.asp?r=82.0021475055917&svr=55&lang=en_us&aid={UNP}')

    # Now we OPEN The links
    tab(0)
    i = int(text)
    print('[Tab]:', text)


    while i <= int(getVars(3)):
        if len(os.listdir('D:\Plag\plagdir')) != 0:
            try:
                # Here i means the Section number

                # print(i,int(getVars(3)))
                Unid = int(box[i-1])

                driver.get(f'https://www.turnitin.com/t_submit.asp?r=82.0021475055917&svr=55&lang=en_us&aid={Unid}')
                sleep(3)

                # Selecting a file
                driver.find_element(
                    by=By.XPATH, value=f'//*[@id="choose-file-btn"]').click()
                path = os.listdir('D:\Plag\plagdir')[0]

                writenme('Open', path)
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
                print(
                    f'[Info]: {submit_title},{Slot_title},Remaining_file: {len(a)}')
                csvwr(f'{submit_title},{Slot_title}')
                wrreplace('D:\Plag\Vars.txt', getVars(1), Slot_title)
                try:
                    with open('report.txt', 'w') as f:
                        f.write(
                            f'{submit_title},{Slot_title},Remaining_file: {len(a)}')
                        f.close()
                except:
                    pass
                ###################################################
                WebDriverWait(driver, 250).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm-btn"]')))
                driver.find_element(
                    by=By.XPATH, value=f'//*[@id="confirm-btn"]').click()
                os.remove(f"D:\Plag\plagdir\{path}")
                sleep(1)
                WebDriverWait(driver, 25).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="close-btn"]')))
                if i == int(getVars(3)):
                    i = 0
            except:
                print('[Error in line ({})]: '.format(
                    sys.exc_info()[-1].tb_lineno), type(e).__name__, e)

        else:
            try:
                sleep(2)
                driver.find_element(
                    by=By.XPATH, value=f'//*[@id="close-btn"]').click()
            except:
                pass
            sleep(2)
            if getVars(2) == 'y':
                try:
                    with open('report.txt', 'w') as f:
                        f.write('File Uploaded now downloading them')
                        f.close()
                except:
                    pass
                try:
                    for i in range(getVars(3)):
                        per = driver.find_elements(
                            by=By.CLASS_NAME, value=f'or-percentage')[i]
                        try:
                            if int(per):
                                pass
                        except:
                            driver.refresh()
                            break
                except:
                    pass
                sleep(2)
                Download_plag()
            elif getVars(2) == 'n':
                try:
                    for i in range(getVars(3)):
                        per = driver.find_elements(
                            by=By.CLASS_NAME, value=f'or-percentage')[i]
                        try:
                            if int(per):
                                break
                        except:
                            driver.refresh()
                            break
                except:
                    pass
                driver.refresh()
                # report()
            break

        i += 1


def uploadfile_long(text):
    print('[Warning]: Long Method is starting')
    sleep(1)
    i = int(text)+1
    min = getVars(0)
    while i < 100:
        if len(os.listdir('D:\Plag\plagdir')) != 0:
            try:
                # This is For coordination
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
                            # report()
                        i = 0
                    elif i > int(getVars(3))+1:  # Restart when all slots are full
                        i = 0
                else:
                    if i > int(getVars(3))+1:  # Restart when all slots are full
                        i = 0

                # Main code start from here
                tab(0)
                WebDriverWait(driver, 300).until(EC.element_to_be_clickable(
                    (By.XPATH, f'/html/body/div[2]/div[4]/div[2]/div[2]/div[4]/table/tbody/tr[{i}]/td[5]/a[2]')))
                driver.find_element(
                    by=By.XPATH, value=f'/html/body/div[2]/div[4]/div[2]/div[2]/div[4]/table/tbody/tr[{i}]/td[5]/a[2]').click()

                try:
                    Alert(driver).accept()
                except Exception:
                    pass

                # Selecting a file
                hWndT = win32gui.FindWindow(None, 'Turnitin - Google Chrome')

                try:
                    win32gui.ShowWindow(hWndT, 5)
                    win32gui.SetForegroundWindow(hWndT)
                except Exception:
                    pass
                driver.find_element(
                    by=By.XPATH, value=f'//*[@id="choose-file-btn"]').click()
                path = os.listdir('D:\Plag\plagdir')[0]

                writenme('Open', path)
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
                print(
                    f'[Info]: {submit_title},{Slot_title},Remaining_file: {len(a)}')
                csvwr(f'{submit_title},{Slot_title}')
                wrreplace('D:\Plag\Vars.txt', getVars(1), Slot_title)
                try:
                    with open('report.txt', 'w') as f:
                        f.write(
                            f'{submit_title},{Slot_title},Remaining_file: {len(a)}')
                        f.close()
                except:
                    pass
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
                # report()
            break
        i += 1


def Download_plag():
    for x in range(2):
        i = 1  # Loop
        j = 1  # Iter in excel
        driver.implicitly_wait(7)
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
                        by=By.XPATH, value=f'/html/body/div[2]/div[4]/div[2]/div[2]/div[4]/table/tbody/tr[{search+1}]/td[1]').get_attribute('innerText').capitalize()
                    # Iter in TUrnitin of slots bvalue only
                    file_slot = int(
                        ''.join(x for x in file_sot if x.isdigit()))
                    if search != 0:
                        if search == file_slot:
                            actions = ActionChains(driver)
                            biew = driver.find_element(
                                by=By.XPATH, value=f'/html/body/div[2]/div[4]/div[2]/div[2]/div[4]/table/tbody/tr[{search+1}]/td[5]/a[2]')
                            actions.key_down(Keys.CONTROL).click(biew).key_up(
                                Keys.CONTROL).perform()  # Clicks on view to download
                            tab(1)
                            sleep(2)
                            try:
                                swait('/html/body/div[6]/div[1]/aside/div[1]/section[4]/div/div[1]')
                                driver.find_element(
                                    by=By.XPATH, value='/html/body/div[6]/div[1]/aside/div[1]/section[4]/div/div[1]').click()
                            except:
                                try:
                                    swait('/html/body/div[5]/div[1]/aside/div[1]/section[4]/div/div[1]')
                                    driver.find_element(
                                        by=By.XPATH, value='/html/body/div[5]/div[1]/aside/div[1]/section[4]/div/div[1]').click()
                                except:
                                    driver.close()
                                    break
                                

                            sleep(4)
                            try:
                                findxpth(
                                    '/html/body/div[8]/div[2]/div[2]/div/div[1]').click()
                            except:
                                findxpth('/html/body/div[7]/div[2]/div[2]/div/div[1]').click()
                            # driver.implicitly_wait(0)

                            try:
                                nmeoffile = driver.find_element(by=By.XPATH, value='/html/body/div[6]/header/h1/span[2]').get_attribute('innerHTML')
                            except Exception:
                                try:
                                    nmeoffile = driver.find_element(by=By.XPATH, value='/html/body/div[5]/header/h1/span[2]').get_attribute('innerHTML')
                                except Exception:
                                    nmeoffile = "UNKNOWN"
                                
                            try:
                                sleep(4)
                                # WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[5]/div[1]/aside/div[1]/section[2]/div[2]/div[1]/label')))
                                try:
                                    percentag = driver.find_element(by=By.ID,value=f'sc5044-label').get_attribute('innerHTML')
                                except Exception:
                                    percentag = driver.find_element(by=By.XPATH, value='/html/body/div[5]/div[1]/aside/div[1]/section[2]/div[2]/div[1]/label').get_attribute('innerHTML')
                                percentage = '{0}%'.format(percentag)
                                if len(percentag) != 0:
                                    wrreplace(csv_path, file_sot,f'Downloaded,{file_sot},{percentage}')
                                    sleep(1)
                                    WebDriverWait(driver, 100).until(EC.invisibility_of_element(
                                        (By.XPATH, '/html/body/div[8]/div[2]/div[1]/div')))
                                    sleep(4)
                            except Exception as e:
                                print(f"Error while downloading {e}")
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
        sleep(20)
        try:
            # Move Zip (if avail) to Rbin
            move(r'D:\Plag\Download\Download.zip',r'D:\Plag\Backup\Download.zip')
        except Exception:
            pass

        # if files are more than 3 then ZIP them
        if len(os.listdir(r'D:\Plag\Download')) > 0:
            try:
                with zipfile.ZipFile(r'D:\Plag\Download\Download.zip', 'w', zipfile.ZIP_DEFLATED) as zipf:
                    zipdir('D:\Plag\Download', zipf)

                sleep(1)
                try:
                    with open('report.txt', 'w') as f:
                        f.write('File checked please download them')
                        f.close()
                except:
                    pass
                # send2trash("D:\Plag\Download\Download.zip")
            except:
                pass

    except Exception:
        print("[Warning]: Files can't be zipped!")
    

def login():
    driver.implicitly_wait(1)

    try:
        print('[Info]: Logging in..')
        try:
            driver.find_element(
                by=By.XPATH, value='/html/body/header/div[2]/section[2]/div/div[2]/div/a[2]').click()
        except Exception:
            pass
        try:

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
                except Exception:
                    pass
            try:
                driver.find_element(
                    by=By.XPATH, value=f'/html/body/div[2]/form/div[2]/div[3]/div/div[2]/table/tbody/tr[{k}]/td[2]/a').click()
            except:
                try:
                    driver.find_element(
                        by=By.XPATH, value=f'/html/body/div[2]/form/div[2]/div[3]/div/div[2]/table/tbody/tr[2]/td[2]/a').click()
                except:
                    driver.find_element(
                        by=By.XPATH, value=f'/html/body/div[2]/form/div[2]/div[3]/div/div[2]/table/tbody/tr[1]/td[2]/a').click()

        except Exception:
            pass
    except Exception:
        pass


try:
    # Move all doc also rename your all files
    move_all_doc()
    remove_emty_folder()
except Exception as e:
    print(e)
try:
    toal = int(len(os.listdir(r'D:\Plag\plagdir')))
except:
    pass

login()

try:
    # If Plagdir is empty then sl = 'Plag Check'
    if os.listdir('D:\Plag\plagdir') != []:
        if getVars(6) == '143':
            sl = getVars(6)
        else:
            sl = 'Plag Check'

    else:
        sl = getVars(2)


    # try:
    #     pagei = driver.find_element(
    #         by=By.XPATH, value='/html/body/div[2]/div[4]/div[2]/div[2]/div[4]/table/tbody/tr[2]').get_attribute('id')
    #     pageid = ''.join(x for x in pagei if x.isdigit()
    #                     )  # This is for integer only

    #     print(pageid)
    # except Exception:
    #     print('[Info]: Not on main Page')
    #     pagei = driver.current_url
    #     try:
    #         lnkid = pagei.split('=')[1].split('&')[0]
    #         pageid = lnkid - getSlotnum()
    #     except:
    #         pass
        
    # Now checking sl value
    if sl == 'Plag Check':
        uploadfile(getSlotnum()+1)
        
    elif sl == 'Download' or sl == 'y':
        print("[Info]: Direct Download")
        # for i in range(2):
        Download_plag()
        # pg.alert('All Files are downloaded')

    elif sl == '143':
        uploadfile_long(getSlotnum()+1)

    else:
        search = sl
        # findg(search)
except:
    pass
