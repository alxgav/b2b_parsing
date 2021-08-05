import os, glob, csv
from selenium import webdriver
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities

path = os.path.dirname(os.path.realpath(__file__))


def deleteAllFiles():
    files = glob.glob(f'{path}/download/*.xlsx')
    for f in files:
        try:
            os.remove(f)
        except OSError as e:
            print ("Error: %s : %s" % (f, e.strerror))



def read_csv(filename, encoding='utf-8', sep=';', header = None):
    data = []
    line = 0
    with open(filename, encoding=encoding) as csv_file:    #'ISO-8859-1'
        csv_reader = csv.reader(csv_file, delimiter=sep)
        if header != None:
            field = next(csv_reader).index(header)
        for row in csv_reader:
            if header == None:
                if line !=0:
                    data.append(row)
                line +=1
            else:
                data.append(row[field])
    return data


def unique_list(data):
    unique = []
    for i in data:
        if i in unique:
            continue
        else:
            unique.append(i)
    return unique




def getExcelFileName():
    return glob.glob(f'{path}/download/*')


def browser():
    capabilities = DesiredCapabilities.CHROME
    capabilities["goog:loggingPrefs"] = {"performance": "ALL"}
    options = webdriver.ChromeOptions()
    options.add_argument('user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.66 Safari/537.36')
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_argument("window-size=1920,1080")
    options.add_argument('--headless')
    # options.add_argument('--no-sandbox')
    options.add_argument('--disable-gpu')
    options.add_experimental_option("prefs", {
        "download.default_directory": f"{path}/download",
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing_for_trusted_sources_enabled": False,
        "safebrowsing.enabled": False
})
    return  webdriver.Chrome(executable_path='/home/alxgav/myproject/fiverr/parcer1/selenium/chrome/chromedriver', options=options)

