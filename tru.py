from all_imports import *

csv_path = f'D:\Plag\plag({today}).csv'

# i starts from 2
j=1
k=0
last=2
kd=0
# Click on my grades
driver.find_element(by=By.XPATH, value=f'/html/body/div[2]/div[4]/div[1]/ul/li[2]/a').click()
sleep(1)
# Loop starts from here
for i in range(2,100):
    tab(0)
    try:
        with open(csv_path) as f:
            serch = f.readlines()[k].split(',')[1]
            f.close()

        if serch == 'Downloaded':
            stats = 'Checked'
        else:
            stats = 'Not checked'
            search = int(''.join(x for x in serch if x.isdigit()))

            # Choose tab
            view_element = driver.find_element(
                by=By.XPATH, value=f'/html/body/div[2]/div[4]/div[2]/div[3]/form/div/div[2]/table/tbody/tr[{search+1}]/td[8]/a')
            actions = ActionChains(driver)
            actions.key_down(Keys.CONTROL).click(view_element).key_up(Keys.CONTROL).perform()  # Clicks on view to download
            print(f"Tab: {search}")
            sleep(0.12)
            driver.execute_script("window.scrollBy(0, 100);")

            last+=1
    except:
        break
    k+=1
for i in range(2,100):
    try:
        with open(csv_path) as f:
            serch = f.readlines()[kd].split(',')[1]
            f.close()
        if serch != "Downloaded":
            search = int(''.join(x for x in serch if x.isdigit()))
            file_sot = serch.replace("\n","")
            try:
                tab(j)
                # print("Tab:",j)
                j+=1
            except:
                break
            # Wait and click download button but after confirming Processing

            try:
                WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[5]/div[1]/aside/div[1]/section[2]/div[2]/div[1]/label')))
                # print("BREAKING")
                percentag = driver.find_element(by=By.XPATH, value='/html/body/div[5]/div[1]/aside/div[1]/section[2]/div[2]/div[1]/label').get_attribute('innerText')
                percentage = '{0}%'.format(percentag)
                if len(percentag) != 0:
                    wrreplace(csv_path, file_sot,f'Downloaded,{file_sot},{percentage}')
                    # WebDriverWait(driver, 100).until(EC.invisibility_of_element((By.XPATH, '/html/body/div[7]/div[2]/div[1]/div')))
                    sleep(0.5)

            except Exception as e:
                print('[Error in line ({})]: '.format(sys.exc_info()[-1].tb_lineno))
                percentage = 'Processing'
                
                continue

            if percentag == '':
                percentage = 'Processing'

            WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div[1]/aside/div[1]/section[4]/div/div[1]')))
            driver.find_element(by=By.XPATH, value='/html/body/div/div[1]/aside/div[1]/section[4]/div/div[1]').click()

            # Choose current view
            WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[7]/div[2]/div[2]/div/div[1]/div[2]')))
            sleep(0.5)
            findxpth('/html/body/div[7]/div[2]/div[2]/div/div[1]/div[2]').click()

            # Name of file
            nmeoffile = driver.find_element(by=By.XPATH, value='/html/body/div/header/h1/span[2]').get_attribute('innerHTML')

            # Written work
            print(f'[Info]: Downloading [{search}]: {nmeoffile} - {percentage}')

            try:
                with open('report.txt', 'w') as f:
                    f.write(f'Downloading [{search}]: {nmeoffile} - {percentage}')
                    f.close()
            except:
                pass

    except:
        break
    kd+=1
# After downloading files we try to make archive of it but fst we move all old downloaded files to diffrent folder
try:
    sleep(20)
    try:
        # Move Zip (if avail) to Rbin
        move(r'D:\Plag\Download\Download.zip',r'D:\Code\Acade Projects\Django Project\Plagzip\plagzip\static\Download.zip')
    except Exception:
        pass

    # if files are more than 3 then ZIP them
    if len(os.listdir(r'D:\Plag\Download')) > 0:
        try:
            with zipfile.ZipFile(r'D:\Plag\Download\Download.zip', 'w', zipfile.ZIP_DEFLATED) as zipf:
                zipdir('D:\Plag\Download', zipf)

                
            sleep(1)
            try:
                with open('report.txt','w') as f:
                    f.write('File checked please download them')
                    f.close()
            except:
                pass
            # send2trash("D:\Plag\Download\Download.zip")
        except:
            pass

except Exception:
    print("[Warning]: Files can't be zipped!")
