import setting as s
import mailto as email
from pprint import pprint
import time, csv, re, json, os, math
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from loguru import logger
import pandas as pd
path = os.path.dirname(os.path.realpath(__file__))
logger.add(f'{path}/error.log', format= '{time} {level} {message}', level='DEBUG', serialize=False)

config_data = json.load(open(f'{path}/config.json', 'r'))

sleep = 5
browser = s.browser()
browser.delete_all_cookies()
# browser.implicitly_wait(10)
browser.get(config_data['url'])
time.sleep(sleep)
start_time = time.time()


@logger.catch
def download():
    s.deleteAllFiles()
    print("downloading ....")
    try:
        browser.find_element_by_xpath('/html/body/bianchi-customer/acx-base/div/div/div/div[2]/div[2]/om-content/div/div[2]/page-loader/div/sdk-frame-section/sdk-frame-section[2]/sdk-frame-section/div/ordv-order-list-grid/div/div[1]/div[1]/div[1]/sdk-icon/div/i').click()
        time.sleep(sleep)
        # download file
        browser.find_element_by_xpath('/html/body/bianchi-customer/acx-base/div/div/div/div[2]/div[2]/om-content/div/div[2]/page-loader/div/sdk-frame-section/sdk-frame-section[3]/sdk-frame-section/div/ordv-order-list-button-list2/div/div/sdk-button-panel/div[2]/sdk-button/button').click()
        time.sleep(10)
    except Exception as ex:
        print(ex)
    finally:
        print("downloaded")

def login():
    user_input = browser.find_element_by_xpath('//*[@id="loginForm:user"]')
    user_input.clear()
    user_input.send_keys(config_data['site'][0])
    pass_input = browser.find_element_by_xpath('//*[@id="loginForm:passwd"]')
    pass_input.clear()
    pass_input.send_keys(config_data['site'][1])
    pass_input.send_keys(Keys.ENTER)
    time.sleep(sleep)
    browser.find_element_by_xpath('//*[@id="newMenuForm:j_id98"]/div/div/div[2]/ul/li[4]/a').click()
    print('1')
    time.sleep(sleep)
    browser.find_element_by_xpath('//*[@id="newMenuForm:j_id98"]/div/div/div[2]/ul/li[4]/ul/li/a/span[2]').click()
    print('2')
    time.sleep(sleep)
    browser.find_element_by_xpath('//*[@id="newMenuForm:j_id98"]/div/div/div[2]/ul/li[4]/ul/li/ul/li/a').click()
    print('3')
    time.sleep(sleep)
    download()
    browser.find_element_by_xpath('/html/body/bianchi-customer/acx-base/div/div/div/div[2]/div[2]/om-content/div/div[2]/page-loader/div/sdk-frame-section/sdk-frame-section[1]/sdk-frame-section[1]/div/ordv-order-list-button-list1/div/div/div[1]/sdk-button/button').click()
    time.sleep(sleep)
    browser.find_element_by_xpath('/html/body/sdk-modal-panel/mz-modal/div/div[1]/div/mz-modal-content/reorders-choice/div/div/div[1]/img').click()
                                   
    time.sleep(sleep)

    
@logger.catch
def get_data():
    # login to page
    login()
    #  end login
    try:
        data = []
        id_sku = []
        num = 1
        id_num = 1
        
        totals = browser.find_element_by_xpath('/html/body/bianchi-customer/acx-base/div/div/div/div[2]/div[2]/om-content/div/div[2]/page-loader/div/sdk-frame-section/sdk-frame-section[1]/sdk-frame-section[2]/div/ordv-order-product-search-grid/div/div[1]/div[2]/div').text
        print (totals)
        totals = math.ceil(int(re.sub('\D', '', totals ))/40)
        browser.execute_script("arguments[0].click();", browser.find_element_by_xpath('//*[@id="nextPage"]'))
        time.sleep(sleep)
        for page in range(1, totals+1):
            for i in browser.find_elements_by_class_name('item-col-ld'):
                id = i.get_attribute('id').strip()
                print (id, '===========================================')
                if id[0:4] not in config_data['out_of_list'] and id[0:5] not in config_data['out_of_list']:
                     
                    link =''
                    if id not in id_sku:
                        print (id, id_num, page)

                        try:
                            browser.execute_script("arguments[0].click();", i.find_element_by_xpath(f'//*[@id="{id}"]/sdk-collection-item-ld/div/div[1]/div/div/div[2]/div[2]/sdk-item-attribute/sdk-single-action-attribute/sdk-button/button'))
                            time.sleep(sleep)
                            browser.switch_to.window(browser.window_handles[1])
                            link = browser.current_url
                            browser.close()
                            browser.switch_to.window(browser.window_handles[0])
                        except:
                            link =''
                            browser.execute_script("arguments[0].click();", i.find_element_by_xpath('//*[@id="toast-container"]/div[1]'))
                        try:
                            AVAILABLE = i.find_element_by_xpath(f'//*[@id="{id}"]/sdk-collection-item-ld/div/div[1]/div/div/div[2]/div[3]/sdk-item-attribute/sdk-data-attribute/div/div/div/div[3]/p').text 
                        except:
                            AVAILABLE = ''
                        try:
                            PRODUCT = i.find_element_by_xpath(f'//*[@id="{id}"]/sdk-collection-item-ld/div/div[1]/div/div/div[2]/div[5]/sdk-item-attribute/sdk-data-attribute/div/div/div/div[3]/p').text 
                        except:
                            PRODUCT = ''
                        try:
                            PRICE = i.find_element_by_xpath(f'//*[@id="{id}"]/sdk-collection-item-ld/div/div[1]/div/div/div[2]/div[6]/sdk-item-attribute/sdk-data-attribute/div/div/div/div[3]/p').text 
                        except:
                            PRICE = ''
                        try:
                            PRODUCT_LINK = link
                        except:
                            PRODUCT_LINK = ''
                        data.append({'SKU': id ,
                                        'AVAILABLE':AVAILABLE,
                                        'PRODUCT':PRODUCT,
                                        'PRICE':PRICE,
                                        # 'RETAIL PRICE':RETAIL_PRICE,
                                        'PRODUCT_LINK': PRODUCT_LINK })
                        # i.location_once_scrolled_into_view
                        id_sku.append(id)  
                        num +=1
                        id_num +=1
            browser.execute_script("arguments[0].click();", browser.find_element_by_xpath('//*[@id="nextPage"]'))
            time.sleep(sleep)
        print (len(data))
        df = pd.DataFrame(s.unique_list(data))
        df.to_csv('bikes.csv', sep=';', index=False)
        make_html(readDB(s.unique_list(data)), get_acessoria())
    except Exception as ex:
        print(ex)
    finally:
        # pass
        browser.close()
        browser.quit()



@logger.catch
def readDB(data):
    data_new = []
    detail_data = []
    # login()   
    for r in data:
        search_input = browser.find_element_by_xpath('/html/body/bianchi-customer/acx-base/div/div/div/div[2]/div[2]/om-content/div/div[2]/page-loader/div/sdk-frame-section/sdk-frame-section[1]/sdk-frame-section[2]/div/ordv-order-product-search-grid/div/div[1]/div[4]/div[1]/mz-input-container/div/input')
        search_input.clear()
        id = r['SKU'].strip()
        # id = r.strip()
        print (id)
        search_input.clear()
        search_input.send_keys(id)
        time.sleep(sleep)
        s_button = browser.find_element_by_xpath('/html/body/bianchi-customer/acx-base/div/div/div/div[2]/div[2]/om-content/div/div[2]/page-loader/div/sdk-frame-section/sdk-frame-section[1]/sdk-frame-section[2]/div/ordv-order-product-search-grid/div/div[1]/div[4]/div[3]/sdk-button/button/i')
        browser.execute_script("arguments[0].click();", s_button)
        time.sleep(sleep)
        
       
        
        browser.find_element_by_xpath(f'//*[@id="{id}"]/sdk-collection-item-ld/div/div[1]/div/div/div[1]/i').click()
        
        # try:   
        #     product = browser.find_element_by_xpath(f'//*[@id="{id}"]/sdk-collection-item-ld/div/div[1]/div/div/div[2]/div[5]/sdk-item-attribute/sdk-data-attribute/div/div/div[3]/p').text.strip()
        # except:
        #     product = ''
        browser.execute_script("arguments[0].click();", browser.find_element_by_xpath('/html/body/bianchi-customer/acx-base/div/div/div/div[2]/div[2]/om-content/div/div[2]/page-loader/div/sdk-frame-section/sdk-frame-section[1]/sdk-frame-section[2]/div/ordv-order-product-search-grid/div/div[1]/div[3]/sdk-button[4]/button'))
        time.sleep(sleep)
        
        # ============================================
        # color_data = []
        for i in browser.find_elements_by_class_name('item-row-ld'):
            try:
                product = browser.find_element_by_xpath('/html/body/sdk-modal-panel/mz-modal/div/div[1]/div/mz-modal-content/order-product-grid-modal/div/div[1]/div[3]/div[1]').text.replace('Produkt\n', '').replace(id, '').strip()
            except:
                product = r['PRODUCT']
            try:
                color = i.find_element_by_xpath(f'//*[@id="{i.get_attribute("id").strip()}"]/sdk-collection-item-ld/div/div[1]/div/div[1]/div/div[2]/sdk-item-attribute/sdk-data-attribute/div/div/p').text.strip()
                color2 = i.find_element_by_xpath(f'//*[@id="{i.get_attribute("id").strip()}"]/sdk-collection-item-ld/div/div[1]/div/div[1]/div/div[2]/sdk-item-attribute/sdk-data-attribute/div/div/p').text.strip().split('-')[0]
            except:
                color = ''
                color2 = ''
            try:
                price = i.find_element_by_xpath(f'//*[@id="{i.get_attribute("id").strip()}"]/sdk-collection-item-ld/div/div[1]/div/div[1]/div/div[5]/sdk-item-attribute/sdk-cell-input-advanced-attribute/div/div/div/p').text.strip()
            except:
                price = ''
            size_data = []
            # try:
                
            for k in i.find_elements_by_class_name('multi-scroll-item'):
                size = ''
                try:
                    availabel = 'YES'
                    size = k.find_element_by_class_name('not-editable-line-height').text.strip()
                    
                except:
                    size = ''
                try:
                    if k.find_element_by_class_name('price-not-editable-span'):
                        availabel = "NO"
                except:
                    availabel = "YES"
                try:                                           
                    deliver = k.find_element_by_xpath(f'//*[@id="{i.get_attribute("id").strip()}"]/sdk-collection-item-ld/div/div[1]/div/div[1]/div/div[3]/sdk-item-attribute/sdk-size-attribute-full/div/div/div[1]/b').text.strip()
                except:
                    deliver = ''
                # print ('deliver',f'**********************===={size}=={k.get_attribute("id").strip()}====={deliver}========{i.get_attribute("id").strip()}======***********************')
                if availabel == "YES" and size !='':
                    size_data.append(f'{size} / {deliver}\n')
                if color != '' and size !='':
                    sku = (id+size+color2).strip()
                    detail_data.append({
                        'SKU': sku,
                        'PRODUKT': product,
                        'COLOR': color,
                        'SIZE': size,
                        'PRICE': price,
                        'AVAILABEL': availabel,
                        
                                            })
                    pprint(detail_data)
                
                                
        
            data_new.append({'SKU':id, 
                                'PRODUCT':r['PRODUCT'], 
                                'COLOR': color, 
                                'SIZE': size_data, 
                                'PRICE': r['PRICE'],
                                'link': r['PRODUCT_LINK'],
                                })
            # print ('===================',id,r['PRODUCT'], color, size_data, r['PRICE'], r['PRODUCT_LINK'], '=====================')
        # ============================================
        # print (detail_data)
        # time.sleep(sleep)
        browser.find_element_by_xpath(f'/html/body/sdk-modal-panel/mz-modal/div/div[1]/div/mz-modal-content/order-product-grid-modal/div/div[1]/div[1]/div[2]/sdk-button/button/i').click()
        time.sleep(sleep)
        browser.find_element_by_xpath(f'//*[@id="{id}"]/sdk-collection-item-ld/div/div[1]/div/div/div[1]/i').click()

    # html = make_html(new_data)
    df = pd.DataFrame(detail_data)
    df.to_csv('detail_bike.csv', sep=';', index=False)
    df_new = pd.DataFrame(data_new)
    df_new.to_csv('new_bike.csv', sep=';', index=False)
    # pprint(data_new)
    return data_new
    
    

# make html to send

@logger.catch
def make_html(data, table_acess):
    print ('generate email ==========================================')
    SUBJECT = 'Neue Fahrräder'
    FROM = 'alxgav@gmail.com'
    table_header ='''
        <table style="width:100%; border-collapse: collapse; border: 1px solid black;">
             <tr style="background: #D9D7D7;">
                <th style="border-collapse: collapse; border: 1px solid black;">CODE</th>
                <th style="border-collapse: collapse; border: 1px solid black;">PRODUCT</th>
                <th style="border-collapse: collapse; border: 1px solid black;">COLOR</th>
                <th style="border-collapse: collapse; border: 1px solid black;">SIZE/DELIVER</th>
                <th style="border-collapse: collapse; border: 1px solid black;">PRICE</th>
                <th style="border-collapse: collapse; border: 1px solid black;">LINK</th>
             </tr>
                  '''
    # f = open ('new_data.json', 'r', encoding='utf-8')
    code_data = []
    code = []
    # data = json.load(f)
    
    for i in data:
        code_data.append([i['SKU'], i['PRODUCT'], i['PRICE'], i['link']])
    code = s.unique_list(code_data)
    
    content = []
    
    for i in code:
        code = i[0]
        product = i[1]
        price = i[2]
        link = i[3]
        table_td = []
        for j in data:
            if code == j['SKU']:
                table_td.append(f'''
                    <td style="border-collapse: collapse; border: 1px solid black; text-align:center;">{j['COLOR']}</td>
                    <td style="border-collapse: collapse; border: 1px solid black; text-align:center;">{', '.join(j['SIZE'])}</td>
                '''.replace('\n', '').strip())
        content.append({'code':code, 'product':product, 'price': price, 'link':link, 'content':table_td})


    data = ''
    
    # n = 0
    for i in content:
        size  = len(i['content'])
        print (size, 'size ====================', i['code'])
        if size > 1:

            data += f'''
               
                <tr>
                       <td rowspan="{size}" style="border-collapse: collapse; border: 1px solid black; text-align:center;">{i['code']}</td>
                       <td rowspan="{size}" style="border-collapse: collapse; border: 1px solid black; text-align:center;">{i['product']}</td>
                       {i['content'][0]}
                       <td rowspan="{size}" style="border-collapse: collapse; border: 1px solid black; text-align:center;">{i['price']}</td>
                       <td rowspan="{size}" style="border-collapse: collapse; border: 1px solid black; text-align:center;">{i['link']} </td>
                     </tr>
                     <tr>
                     
                        { '<tr>'.join(i['content'][1:])}
                </tr>
                </tr>
                <tr>
                '''
        else:
            data += f'''
                <tr>
                       <td rowspan="{size}" style="border-collapse: collapse; border: 1px solid black; text-align:center;">{i['code']}</td>
                       <td rowspan="{size}" style="border-collapse: collapse; border: 1px solid black; text-align:center;">{i['product']}</td>
                       {i['content'][0]}
                       <td rowspan="{size}" style="border-collapse: collapse; border: 1px solid black; text-align:center;">{i['price']}</td>
                       <td rowspan="{size}" style="border-collapse: collapse; border: 1px solid black; text-align:center;">{i['link']} </td>
                     </tr>
                     <tr>
                '''
    print ('send email')
    if table_acess == None:
        table_acess = ''
    if len(code) > 0:
        email.send_email(SUBJECT, FROM, config_data['email'], f'''<html><body>{table_header}{data}</table><br>{table_acess}</body></html>''')
    else:
        if table_acess != '':
            email.send_email(SUBJECT, FROM, config_data['email'], f'''<html><body>{table_acess}</body></html>''')



#  acessories

def get_acessoria():
    try:
        browser.find_element_by_xpath('//*[@id="desktop-menu-container"]/mz-collapsible/ul/li[4]/div[1]').click()
        print('first menu')
        time.sleep(sleep)
        
        browser.find_element_by_xpath('//*[@id="desktop-menu-container"]/mz-collapsible/ul/li[4]/div[2]/mz-collapsible-item-body/div/mz-collapsible/ul/li/div[1]').click()
        print('second menu')
        time.sleep(sleep)
        browser.find_element_by_xpath('//*[@id="desktop-menu-container"]/mz-collapsible/ul/li[4]/div[2]/mz-collapsible-item-body/div/mz-collapsible/ul/li/div[2]/mz-collapsible-item-body/div/div/div').click()
        print('theart menu')
        time.sleep(sleep)
        browser.find_element_by_xpath('/html/body/bianchi-customer/acx-base/div/div/div/div[2]/div[2]/om-content/div/div[2]/page-loader/div/sdk-frame-section/sdk-frame-section[1]/sdk-frame-section[1]/div/ordv-order-list-button-list1/div/div/div[1]/sdk-button/button').click()
        time.sleep(sleep)
        browser.find_element_by_xpath('/html/body/sdk-modal-panel/mz-modal/div/div[1]/div/mz-modal-content/reorders-choice/div/div/div[2]/img').click()
                                    
        time.sleep(sleep)
        id_acess =[]
        data = []
        new_data = []
        num = 1
        totals = browser.find_element_by_xpath('/html/body/bianchi-customer/acx-base/div/div/div/div[2]/div[2]/om-content/div/div[2]/page-loader/div/sdk-frame-section/sdk-frame-section[1]/sdk-frame-section[2]/div/ordv-order-product-search-grid/div/div[1]/div[2]/div').text
        totals = math.ceil(int(re.sub('\D', '', totals ))/40)
        print (totals)
        click =1
        browser.execute_script("arguments[0].click();", browser.find_element_by_xpath('//*[@id="nextPage"]'))
        time.sleep(sleep)
        for i in range(totals+1):
            num = 1
            # print(len(browser.find_elements_by_class_name('item-col-ld')))
            for i in browser.find_elements_by_class_name('item-col-ld'):
                id = i.get_attribute('id').strip()
                if id not in id_acess:

                    try:
                        AVAILABLE = i.find_element_by_xpath(f'//*[@id="{id}"]/sdk-collection-item-ld/div/div[1]/div/div/div[2]/div[2]/sdk-item-attribute/sdk-data-attribute/div/div/div/div[3]/p').text 
                    except:
                        AVAILABLE = ''
                    try:
                        PRODUCT = i.find_element_by_xpath(f'//*[@id="{id}"]/sdk-collection-item-ld/div/div[1]/div/div/div[2]/div[4]/sdk-item-attribute/sdk-data-attribute/div/div/div/div[3]/p').text 
                    except:
                        PRODUCT = ''
                    try:
                        PRICE = i.find_element_by_xpath(f'//*[@id="{id}"]/sdk-collection-item-ld/div/div[1]/div/div/div[2]/div[5]/sdk-item-attribute/sdk-data-attribute/div/div/div/div[3]/p').text 
                    except:
                        PRICE = ''
                    if id not in  s.read_csv(f'{path}/acessoria.csv', header='id'):
                        new_data.append({'id': id ,
                                                'AVAILABLE':AVAILABLE,
                                                'PRODUCT':PRODUCT,
                                                'PRICE':PRICE,
                                            #  'RETAIL PRICE':RETAIL_PRICE
                                                })
                        print ('new data ===========================')
                        pprint(new_data)   
                        print ('new data ===========================') 

                    data.append({'id': id ,
                                'AVAILABLE':AVAILABLE,
                                'PRODUCT':PRODUCT,
                                'PRICE':PRICE,
                                # 'RETAIL PRICE':RETAIL_PRICE
                                })
                    id_acess.append(id)
                    num +=1
            browser.execute_script("arguments[0].click();", browser.find_element_by_xpath('//*[@id="nextPage"]'))
            time.sleep(sleep)
            print(click, 'click', len(data), num)
            click +=1
        df = pd.DataFrame(s.unique_list(data))
        df.to_csv('acessoria.csv', sep=';', index=False)
        nd = pd.DataFrame(s.unique_list(new_data))
        print ('acessoria data ',len(new_data))
        return nd.to_html( index=False)
    except Exception as ex:
        print (ex)


if __name__ in "__main__":
    get_data()

    print("--- %s seconds ---" % (time.time() - start_time))
