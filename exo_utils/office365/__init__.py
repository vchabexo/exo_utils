from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys

import os
import time
import logging
import base64
from pathlib import Path

#TODO better logg

class Bot:
    """This class is a browser bot based on selenium
    """
    def __init__(self, headless=True, download_dir=None):
        """This class is a browser bot based on selenium
            headless     : (bool) Tells wether to open window in headless or not. Default True
            download_dir: (str) outputdir for the sheet. If empty downlaod in script dir 
        """    
        #param of chrome
        chrome_options = Options()

        # manage option to lunch the driver in headless
        if headless:
            chrome_options.add_argument("--headless")

        #option to set if file in headless
        options = {
            "download.prompt_for_download": False,
            'w3c': False
        }
        if download_dir:
            options['download.default_directory'] = download_dir
        chrome_options.add_experimental_option("prefs", options)
        self.downlaod_dir = download_dir
        self.chrome_options = chrome_options

    def open(self):
        """Open selenium bot
        """
        self.driver       = webdriver.Chrome(options=self.chrome_options)
        self.wait         = WebDriverWait(self.driver, 10)
        self.actionChains = ActionChains(self.driver)

    def close(self):
        """close selenium bot
        """
        self.driver.close()
    
    def get(self, url):
        """Go to a website
            url : (str) url of site
        """
        self.driver.get(url)

    def enterTextByxPath(self, textBoxxPath, text):
        """Enter information in a text field
            textBoxxPath  : (str) xPatrh of the field
            text          : (str) text to enter
        """
        textBox =  self.wait.until(
            EC.presence_of_element_located((By.XPATH, textBoxxPath))
        )
        
        textBox.clear()
        textBox.send_keys(text)

    def enterLongTextByxPath(self, textBoxxPath, text, maxWait=10):
        """Enter a long text filed in a bax
            textBoxxPath  : (str) xPatrh of the field
            text          : (str) text to enter
            maxWait       : (int) max wait in second. Default is 10
        """
        box = self.wait.until(
            EC.presence_of_element_located((By.XPATH, textBoxxPath))
        )
        self.driver.execute_script(f"document.evaluate('{textBoxxPath}', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.value='{text}';")
        #wait till execution
        t_init = time.time()
        while (not box.get_attribute('value') == text) and ((time.time()-t_init)<= maxWait):
            pass
        if ((time.time()-t_init) >= maxWait):
           raise RuntimeError(f"Could not fill box{textBoxxPath}")  
        box.send_keys(Keys.ENTER)
                


    def clickButtonByxPath(self, buttonxPath):
        """Click on a button
            buttonxPath : (str) xPatrh of the button
        """
        button =  self.wait.until(
            EC.element_to_be_clickable((By.XPATH, buttonxPath))
        )
        button.click()

    def rightClickOnElement(self, element):
        """Right click on an element
            element : (WebElement) element to right click on
        """
        self.actionChains.context_click(element).perform()
    
    def hooverOnElement(self, element):
        """Hoover mouse on an element
            element : (WebElement) element to hoover
        """
        self.actionChains.move_to_element(element).perform()

    def extractTextFromElement (self, elementxPath):
        """Returns text of an element
            elementxPath : (str) xPath of the element
        Returns (str) text of element
        """
        elt = self.wait.until(
            EC.presence_of_element_located((By.XPATH, elementxPath))
        )
        return elt.text

    def waitForElementText (self, elementxPath, textList, maxWait=30):
        """Wait tile an element text is in a specified texts list
            elementxPath : (str) xPath of the element
            textList     : [(str)] list of text element text must be in
            maxWait      : (int) max wait in second. Default is 10
        Returns true if element is in text list else false
        """
        t_init = time.time()
        isInList = False
        while (not isInList) and ((time.time()-t_init)<= maxWait):
            try:
                text = self.extractTextFromElement(elementxPath)
                if text in textList:
                    isInList = True
                else:
                    time.sleep(0.1)
            except:
                pass

        if not isInList:
            raise RuntimeError(f"Could not find text {textList} in {elementxPath}")  

    def connectToOffice(self, username, password):
        """Connect itself to an office page. This is specific for exo pages
            username : (str) username used for connect (must be like john_doe@rtm.quebec)
            password : (str) password used for connection
        """
        #enter username
        self.enterTextByxPath('//*[@id="i0116"]', username)
        #click on next
        self.clickButtonByxPath('//*[@id="idSIButton9"]')
        #enter password
        self.enterTextByxPath('//*[@id="passwordInput"]', password)
        #click submit
        self.clickButtonByxPath('//*[@id="submitButton"]')
        #click stay connected
        self.clickButtonByxPath('//*[@id="idSIButton9"]')

    def downloadFromButton(self, buttonXpath, excpected_nb_of_files=1, isWebElement=False):
        """Downloads files from clicking one a button
        WARNING !!!!!note that this is not a good practice for downloading see https://www.selenium.dev/documentation/en/worst_practices/file_downloads/ for more info
            buttonXpath             : (str) path of button to click on. Can also be a web element. In this case isWebElement msut be true.
            excpected_nb_of_files   : (int) Number of files expected to be download. Default is 1
            isWebElement            : (bool) Must be true if buttonXpath is a web eleemnt
        Returns list of download files if download folder is provided else an empty list
        """
        #to track if downloading conmplete we look for a new file in download folder
        if self.downlaod_dir:
            beforeFilesSet = set(os.listdir(self.downlaod_dir))

        if isWebElement == False:
            self.clickButtonByxPath(buttonXpath)
        else:
            buttonXpath.click()

        #loop for one minute max until new file appears in downloads
        #note that if no download files provided wait one minute by default
        t0 = time.time()
        hasNewFile = False
        while not hasNewFile and time.time()-t0 < 60:
            if self.downlaod_dir:
                currentFilesSet = set(os.listdir(self.downlaod_dir))
                diff = currentFilesSet - beforeFilesSet
                #filtter only fully donwload files
                diff = set(filter(lambda name: (name.split('.')[-1]!='crdownload') and (name.split('.')[-1]!='tmp'), diff))
                if len(diff) == excpected_nb_of_files:
                    hasNewFile = True
                    
            downlaodFilesNames = list(diff) if self.downlaod_dir else []
            #we wait .5 s between each loop to not looping like a made dog!!!!
            time.sleep(.5)

        #if still no file after one minutes we throw an error
        if (self.downlaod_dir) and (len(downlaodFilesNames) == 0):
            url = self.driver.current_url
            raise RuntimeError(f"Could not donwload data from url:{url}") 
        
        t1 = time.time()
        logging.info(f"Successfuly downloaded files {downlaodFilesNames} - duration: {t1 - t0:.2f}s")
        donwlaod_dir_path = self.downlaod_dir if self.downlaod_dir[-1] == '\\' else self.downlaod_dir + '\\'
        downlaodFilesNames = list(map(lambda name: donwlaod_dir_path+name, downlaodFilesNames)) 
        return downlaodFilesNames


def downloadPlannerSheet(plannerUrl, username, password, download_dir=None):
    """Download datsheet from a plannner
        plannerUrl  : (str) url of planner
        username    : (str) username to access planer
        password    : (str) password to access
        download_dir: (str) outputdir for the sheet. If empty downlaod in script dir 
    Returns file name if download_dir is specified else returns void
    """
    #go to planner
    bot = Bot(download_dir=download_dir, headless=True)
    bot.open()
    bot.get(plannerUrl)
    bot.connectToOffice(username, password)
    logging.info(f"Connection to planner {plannerUrl} successful")
    #download data
    #go to option menu
    # if planner has changes since last connection pop up showing changes might appears
    # in this case we must first close the pop-up to access the menu
    try:
        optionXPath = '//*[@id="planner-main-content"]/div/div[2]/div/div[2]/div/div[2]/div/div[2]/div/span/button/i'
        bot.clickButtonByxPath(optionXPath)
    except:
        #close pop up
        closePopUpButtonxPath = '/html/body/div[2]/div/div/div/div/div[1]/div[1]/button/span/i'
        bot.clickButtonByxPath (closePopUpButtonxPath)
        # then open option menu
        optionXPath = '//*[@id="planner-main-content"]/div/div[2]/div/div[2]/div/div[2]/div/div[2]/div/span/button/i'
        bot.clickButtonByxPath(optionXPath)
    # click on export excel    
    exportExcelName = 'Exporter le plan vers Excel'
    try:
        exportExcelButton = bot.driver.find_element_by_name(exportExcelName)
    except:
        time.sleep(1)
        exportExcelButton = bot.driver.find_element_by_name(exportExcelName)

    excelFile = bot.downloadFromButton(exportExcelButton, 1, isWebElement=True)
    bot.close()
    return excelFile


def downloadOnlineFile(fileUrl, username, password, document_type='excel', download_dir=None):
    """Download datsheet from a plannner
        fileUrl          : (str) url of planner
        username         : (str) username to access planer
        password         : (str) password to access
        document_type    : (str) type of document (excel, powerpoint). Default is excel
        download_dir     : (str) outputdir for the sheet. If empty downlaod in script dir 
    Returns file name if download_dir is specified else returns void
    """
    if not document_type in ['excel', 'powerpoint']:
        raise ValueError(f"type {document_type} is not supported")
    if not isinstance(download_dir, str):
        download_dir = str(download_dir)
    #go to file
    bot = Bot(download_dir=download_dir, headless=True)
    bot.open()
    bot.get(fileUrl)
    bot.connectToOffice(username, password)
    logging.info(f"Connection to file {fileUrl} successful")
    #switch frame to go to wac_frame
    bot.wait.until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="WebApplicationFrame"]'))
        )
    if document_type == 'excel':
        bot.driver.switch_to.frame("WebApplicationFrame")
    elif document_type == 'powerpoint':
        bot.driver.switch_to.frame("wac_frame") 

    #wait for page to load
    #To infer end of load we look for doc status text. If doc Status is 'Enregistré' 
    if document_type == 'excel': 
        documentStatusXPath = '//*[@id="documentTitle"]/span/div[2]/span/span/span[2]'
    elif document_type == 'powerpoint':
        documentStatusXPath = '//*[@id="documentTitle"]/span/div[2]/div/span/span[2]'
    bot.waitForElementText(documentStatusXPath, ['Enregistré'])
    
    #go to file menu
    fileXPath = '//*[@id="FileMenuLauncherContianer"]/button'
    bot.clickButtonByxPath(fileXPath)

    #click on save as
    if document_type == 'excel': 
        saveAsXPath = "//*[contains(text(), 'Enregistrer sous')]"
    elif document_type == 'powerpoint':
        saveAsXPath = '//*[@id="PptJewel.DownloadAs-Menu32"]'
    bot.clickButtonByxPath(saveAsXPath)  

    # click on download a copie
    if document_type == 'excel': 
        downloadXpath = "//*[contains(text(), 'Télécharger une copie')]"
        filePath = bot.downloadFromButton(downloadXpath, 1)
    elif document_type == 'powerpoint':
         # click on download a copie
        downlaodOptionXPath = '//*[@id="PptJewel.DownloadAs.DownloadACopy-Menu48"]'
        bot.clickButtonByxPath(downlaodOptionXPath)  

        #downlaod
        downloadXpath = '//*[@id="WACDialogActionButton"]'
        filePath = bot.downloadFromButton(downloadXpath, 1)
    bot.close()
    return filePath

def downlaodSharepointList(sharepointListUrl, username, password, download_dir=None):
    """Download a file from sharepoint
        sharepointListUrl    : (str) url of sharepoint list
        username         : (str) username to access planer
        password         : (str) password to access
        download_dir     : (str) outputdir for the sheet. If empty downlaod in script dir 
    Returns file name if download_dir is specified else returns void
    """

    #go to sharepoint
    bot = Bot(download_dir=download_dir, headless=True)
    bot.open()
    bot.get(sharepointListUrl)
    bot.connectToOffice(username, password)
    logging.info(f"Connection to sharepoint list {sharepointListUrl} successful")

    #clicking on export button

    menuXpath = '//*[@role="menuitem"]'
    bot.wait.until(
        EC.presence_of_element_located((By.XPATH, menuXpath))
    )
    menu = bot.driver.find_elements_by_xpath(menuXpath)
    #this loop is used to let time for menu to load with a max of 10 sec
    accessMenu = False
    t_init = time.time()
    while (accessMenu == False) and ((time.time() - t_init) <= 10):
        try:
            #if menu is not loaded we get an erro when searching for text
            menu[0].text
            accessMenu = True
        except:
            menu = bot.driver.find_elements_by_xpath(menuXpath)

    downloadButton = None
    exportButton = None
    showMoreButton = None
    for item in menu:
        if 'Exporter' in item.text:
            exportButton = item
        if '\ue712' in item.text:
            showMoreButton = item

    # fi export is not visible, we click on show more then export
    if exportButton is None:
        showMoreButton.click()
        #when clicking on showMore, new element in the drop down menu alos have a role of menuitem
        menu = bot.driver.find_elements_by_xpath(menuXpath)
        accessMenu = False
        t_init = time.time()
        while (accessMenu == False) and ((time.time() - t_init) <= 10):
            try:
                #if menu is not loaded we get an erro when searching for text
                menu[0].text
                accessMenu = True
            except:
                menu = bot.driver.find_elements_by_xpath(menuXpath)

        for item in menu:
            if 'Exporter' in item.text:
                exportButton = item
    exportButton.click()
    downloadXpath = '//*[@data-automationid="exportListCommand"]'
    downloadOptions = bot.driver.find_elements_by_xpath(downloadXpath)
    for item in downloadOptions:
        if 'Exporter dans un fichier CSV' in item.text:
            downloadButton = item

    filePath = bot.downloadFromButton(downloadButton, excpected_nb_of_files=1, isWebElement=True)
    bot.close()
    return filePath


def downlaodSharepointFile(sharepointUrl, filesName, username, password, download_dir=None):
    """Download a file from sharepoint
        sharepointUrl    : (str) url of sharepoint folder
        fileName         : [(str)] list of files names to download
        username         : (str) username to access planer
        password         : (str) password to access
        download_dir     : (str) outputdir for the sheet. If empty downlaod in script dir 
    Returns file name if download_dir is specified else returns void
    """

    #go to sharepoint
    bot = Bot(download_dir=download_dir, headless=True)
    bot.open()
    bot.get(sharepointUrl)
    bot.connectToOffice(username, password)
    logging.info(f"Connection to sharepoint {sharepointUrl} successful")
    #find file and select them
    filesTableXPath = './/div[@class="ms-List-cell"]'
    bot.wait.until(EC.presence_of_element_located((By.XPATH, filesTableXPath)))
    filesTable = bot.driver.find_elements_by_xpath(filesTableXPath)
    fileFound = []
    for fileDiv in filesTable:
        # since we can't have wait in an element we use try except with sleep of 1 sec
        try:
            elt = fileDiv.find_element_by_xpath('.//div/div/div/div[2]/div[2]/div/div[1]/div[1]/span/span/button')
        except: 
            time.sleep(1)
            elt = fileDiv.find_element_by_xpath('.//div/div/div/div[2]/div[2]/div/div[1]/div[1]/span/span/button')
        
        name = elt.text
        if name in filesName:
            # select file
            fileDiv.find_element_by_xpath('.//div/div/div/div[1]/div/div/i[2]').click()
            fileFound.append(name)
    if  len(set(filesName) - set(fileFound))>0:
        raise ValueError(f"Could not find file(s) {filesName} in page {sharepointUrl}")


    #Download file
    #look for tools in tool bar
    menuXpath = '//*[@role="menuitem"]'
    bot.wait.until(
        EC.presence_of_element_located((By.XPATH, menuXpath))
    )
    menu = bot.driver.find_elements_by_xpath(menuXpath)
    #this loop is used to let time for menu to load with a max of 10 sec
    accessMenu = False
    t_init = time.time()
    while (accessMenu == False) and ((time.time() - t_init) <= 10):
        try:
            #if menu is not loaded we get an erro when searching for text
            menu[0].text
            accessMenu = True
        except:
            menu = bot.driver.find_elements_by_xpath(menuXpath)

        
    downloadButton = None
    showMoreButton = None
    for item in menu:
        if 'Télécharger' in item.text:
            downloadButton = item
        if '\ue712' in item.text:
            showMoreButton = item

    if downloadButton is None:
        showMoreButton.click()
        #when clicking on showMore, new element in the drop down menu alos have a role of menuitem
        menu = bot.driver.find_elements_by_xpath(menuXpath)
        accessMenu = False
        t_init = time.time()
        while (accessMenu == False) and ((time.time() - t_init) <= 10):
            try:
                #if menu is not loaded we get an erro when searching for text
                menu[0].text
                accessMenu = True
            except:
                menu = bot.driver.find_elements_by_xpath(menuXpath)

        for item in menu:
            if 'Télécharger' in item.text:
                downloadButton = item

    filePath = bot.downloadFromButton(downloadButton, excpected_nb_of_files=len(filesName), isWebElement=True)
    bot.close()
    return filePath

def uploadSharpointFile(sharepointUrl, fileLocalPath, username, password):
    """Upload file to a sharepoint file
    Note that user must have access to the powerapp sherepointUploader to use this function
    WARNING !!!!! This functino does not return an error in SP adress or during upload
        sharepointUrl    : (str) url of sharepoint folder
        fileLocalPath    : (str) local path to file to upload
        username         : (str) username to access planer
        password         : (str) password to access
    Returns void
    """
    #go to sharepoint to extract folder path
    bot = Bot(headless=True)
    bot.open()
    bot.get(sharepointUrl)
    bot.connectToOffice(username, password)

    nomenclatureDivXPath = '//*[@id="appRoot"]/div[1]/div[2]/div[3]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[1]/div'
    nomenclatureText = bot.extractTextFromElement(nomenclatureDivXPath)
    nomenclature = nomenclatureText.split("\n\ue76c\n")
    #replace Documents with Shared Documents if in list
    if nomenclature[0] == 'Documents':
        nomenclature[0] = 'Shared Documents'
    
    sharepointPath = '/' + '/'.join(nomenclature)

    #get file name
    if isinstance(fileLocalPath, str):
        fileLocalPath = Path(fileLocalPath)

    #extract SP root site
    root = sharepointUrl.split('sites/', 1)
    root = root[0] +'sites/' + root[1].split('/')[0]

    #read and encode file
    with open(fileLocalPath, 'rb') as f:
        encoded_str = base64.b64encode(f.read())
    encoded_str = encoded_str.decode("utf-8")


    #connect to powerapps uploader
    bot.get('https://apps.powerapps.com/play/cd065e84-4a98-4ae4-9954-3eee5d9f8053?tenantId=a382d496-c5bd-4725-a0fd-46f2b72047e6')

    #authorized sharepoint connection
    #we use try be3cause doesn't always appears
    try:
        authorizedButtonXpath = '/html/body/div[1]/div[2]/div[2]/div/button[1]'
        bot.clickButtonByxPath(authorizedButtonXpath)
    except:
        pass

    #switch frame to go to wac_frame
    bot.wait.until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="fullscreen-app-host"]'))
        )
    bot.driver.switch_to_frame("fullscreen-app-host")

    #fill site
    sharepointSiteXPath = '//*[@id="publishedCanvas"]/div/div[2]/div/div/div[1]/div/div/div/div/input'
    bot.enterTextByxPath(sharepointSiteXPath, root)
    #fill folder
    sharepointFolderXPath = '//*[@id="publishedCanvas"]/div/div[2]/div/div/div[3]/div/div/div/div/input'
    bot.enterTextByxPath(sharepointFolderXPath, sharepointPath)
    #fill file name
    fileNameXPath = '//*[@id="publishedCanvas"]/div/div[2]/div/div/div[4]/div/div/div/div/input'
    bot.enterTextByxPath(fileNameXPath, fileLocalPath.name)
    #fill file content
    fileContentXPath = '//*[@id="publishedCanvas"]/div/div[2]/div/div/div[5]/div/div/div/div/textarea'
    bot.enterLongTextByxPath(fileContentXPath, encoded_str)

    #click on upload
    uploadButtonXPath = '//*[@id="publishedCanvas"]/div/div[2]/div/div/div[6]/div/div/div/div/button'
    bot.clickButtonByxPath(uploadButtonXPath)
    bot.wait.until(
            EC.element_to_be_clickable((By.XPATH, uploadButtonXPath))
        )
    logging.info(f"Successfuly exported file {fileLocalPath} into sharepoint {sharepointUrl}")
    #making sure powerapps has time to load
    time.sleep(1)
    bot.close()
    return
