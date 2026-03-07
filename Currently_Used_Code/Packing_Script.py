import psycopg2, pprint
import time, sys, openpyxl, re, os, ctypes, requests,json, subprocess, keyboard, pyperclip
from psycopg2 import OperationalError
from rich.console import Console
from requests.auth import HTTPBasicAuth
from send2trash import send2trash
try:
    from PIL import Image
except:
    from pil import Image
from rich.panel import Panel
from time import gmtime, strftime, localtime
import os
from ebaysdk.trading import Connection as Trading
import zeep
from zeep.transports import Transport
from zeep.plugins import HistoryPlugin
import traceback

class Mail:
    def __init__(self):
        # considering the same user and pass for smtp an imap
        self.mail_user = 'kyal@scarlettmusic.com.au'
        self.mail_pass = 'redd8484'
        self.mail_host = 'mail.scarlettmusic.com.au'


    def send_email(self, to, subject, body, path, attach):
        message = MIMEMultipart()
        message["From"] = self.mail_user
        message["To"] = to
        message["Subject"] = subject
        message.attach(MIMEText(body, "plain"))

        ##### CODE FOR ATTACHMENT ########
        if attachment_file != '':
            with open(path + '\\' + attach, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
            encoders.encode_base64(part)

            part.add_header(
                "Content-Disposition",
                "attachment; filename= \"" + attach + "\"",
            )
            message.attach(part)
        # message.attach(part)

        text = message.as_string()

        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(self.mail_host, 465, context=context) as server:
            result = server.login(self.mail_user, self.mail_pass)
            r1 = server.sendmail(self.mail_user, to, text)

        imap = imaplib.IMAP4_SSL(self.mail_host, 993)
        imap.login(self.mail_user, self.mail_pass)
        p1 = imap.append('Inbox.Sent', '\\Seen', imaplib.Time2Internaldate(time.time()), text.encode('utf8'))
        imap.logout()


os.chdir(r'\\SERVER\Python')
from Online_orders import online_order

wsdl = 'http://neptune.alliedexpress.com.au:8080/ttws-ejb/TTWS?wsdl'
# wsdl = 'http://triton.alliedexpress.com.au:8080/ttws-ejb/TTWS'
allied_failed = ''

try:
    allied_client = zeep.Client(wsdl=wsdl, transport=transport, plugins=[history])

    allied_client.transport.session.proxies = {
        # Utilize for all http/https connections
        'http': 'http://neptune.alliedexpress.com.au:8080/ttws-ejb/TTWS', }
    allied_account = allied_client.service.getAccountDefaults('755cf13abb3934695f03bd4a75cfbca7', "SCAMUS", "VIC",
                                                              "AOE")

except:
    allied_failed = 'true'

updateorderheaders = {'Content-Type': 'application/json', 'Accept': 'application/json', 'NETOAPI_ACTION': 'UpdateOrder',
                  'NETOAPI_USERNAME': 'kyalAPI', 'NETOAPI_KEY': 'HSZ3EyI9ki6sseqRIkwy6tL6yFNPEUyM'}

console = Console(force_terminal=True)
now = str(strftime("%Y-%m-%d %H:%M:%S", localtime()))

CSIDL_PERSONAL = 5  # My Documents
SHGFP_TYPE_CURRENT = 0  # Get current, not default value
buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_PERSONAL, None, SHGFP_TYPE_CURRENT, buf)
documents_folder = buf.value

documentsaddress = documents_folder
imagelocations = documents_folder + r"\Python"

api = Trading(appid="KyalScar-Awaiting-PRD-aca833663-f54c920e", timeout=1000, config_file=None,
              devid="5c97c2a8-eae7-49cc-ae81-c3ed1bbab335", certid="PRD-ca8336639ad2-75af-45bc-992a-d0bc",
              token="v^1.1#i^1#r^1#f^0#I^3#p^3#t^Ul4xMF84OkUzOEU3NzVFQzA3NDUwNjUyMkY0NTc4QjlENjRFRDVFXzFfMSNFXjI2MA==")


while True:

    try:
        response = api.execute('GetMyeBaySelling', {"SoldList" : {"Include" : True, "IncludeNotes" : 'true', "DurationInDays" : 10, "OrderStatusFilter" : "AwaitingShipment" , "Pagination" : {"EntriesPerPage" : 200}}})
        unformatteddic = response.dict()
        paginationresult = int(unformatteddic['SoldList']['PaginationResult']['TotalNumberOfPages'])
        break

    except Exception as e:
        print(e)
        continue

SoldList = {}

soldlistnumber = 0
for yy in range(int(paginationresult)):

    pass_condition_list = ''

    while True:

        try:

            if pass_condition_list == 'true':
                break

            response = api.execute('GetMyeBaySelling', {
                "SoldList": {"Include": True, "IncludeNotes": 'true', "DurationInDays": 10,
                             "OrderStatusFilter": "AwaitingShipment",
                             "Pagination": {"EntriesPerPage": 200, "PageNumber": int(yy) + 1}}})
            unformatteddic = response.dict()
            #pprint.pprint(unformatteddic['SoldList']['OrderTransactionArray'])

            for t in unformatteddic['SoldList']['OrderTransactionArray']['OrderTransaction']:

                SoldList.update({soldlistnumber: t})
                soldlistnumber = soldlistnumber+1

            pass_condition_list = 'true'

            break

        except:
            continue

start = time.time()


def create_connection(db_name, db_user, db_password, db_host, db_port):
    connection = None
    try:
        connection = psycopg2.connect(
            database=db_name,
            user=db_user,
            password=db_password,
            host=db_host,
            port=db_port,
            keepalives=1,
            keepalives_idle=30,
            keepalives_interval=10,
            keepalives_count=5
        )
        # print("Connection to PostgreSQL DB successful")
    except OperationalError as e:
        print(f"The error '{e}' occurred")
    return connection


connection = create_connection(
    "postgres", "kyal", "123456789Kyal", "143.244.166.150", "5432"
)

try:
    cursor = connection.cursor()
except:
    connection = create_connection("postgres", "kyal", "123456789Kyal", "143.244.166.150", "5432")
    cursor = connection.cursor()
cursor.execute(
    "SELECT ORDER_ID, Purchase_Order,Sales_Record_Number,Sales_Channel,Name,Company,Address1,Address2,City,State,Postcode,Country,Phone,Email,Date_Placed,Date_Completed,Item_Sku,Item_number_eBay,Item_Quantity,Item_Name,Image_URL,Multiple_Orders,Postage_Service,Tracking_Number,Misc1,Misc2,Misc3,Misc4,Misc5,Misc6,Misc7,Misc8,Misc9,Misc10 FROM dailyorders;")
results = cursor.fetchall()
pprint.pprint(results)

console.print('''[bold green]\n\n\n\n\n\n\n\n\n\n\nWELCOME PEON

TO THE SCARLETT MUSIC PACKING SCRIPT
[/]''')  # 66

while True:

    try:

        running_time = time.time() - start

        if float(running_time) > 600:

            print('Please wait... Refreshing Sticky Notes')

            SoldList = {}

            soldlistnumber = 0
            for yy in range(int(paginationresult)):

                pass_condition_list = ''

                while True:

                    try:

                        if pass_condition_list == 'true':
                            break

                        response = api.execute('GetMyeBaySelling', {
                            "SoldList": {"Include": True, "IncludeNotes": 'true', "DurationInDays": 10,
                                         "OrderStatusFilter": "AwaitingShipment",
                                         "Pagination": {"EntriesPerPage": 200, "PageNumber": int(yy) + 1}}})
                        unformatteddic = response.dict()
                        # pprint.pprint(unformatteddic['SoldList']['OrderTransactionArray'])

                        for t in unformatteddic['SoldList']['OrderTransactionArray']['OrderTransaction']:
                            SoldList.update({soldlistnumber: t})
                            soldlistnumber = soldlistnumber + 1

                        pass_condition_list = 'true'

                        break
                    except:
                        continue

            start = time.time()

        orderfailed = ''



        order_number_packing = console.input('''\n\n\nWhat order number are you packing?
        
        Enter Response:''')

        order_number_packing = order_number_packing.strip()
        try:
            cursor = connection.cursor()
        except:
            connection = create_connection("postgres", "kyal", "123456789Kyal", "143.244.166.150", "5432")
            cursor = connection.cursor()
        cursor.execute(
            f"SELECT ORDER_ID, Purchase_Order,Sales_Record_Number,Sales_Channel,Name,Company,Address1,Address2,City,State,Postcode,Country,Phone,Email,Date_Placed,Date_Completed,Item_Sku,Item_number_eBay,Item_Quantity,Item_Name,Image_URL,Multiple_Orders,Postage_Service,Tracking_Number,Misc1,Misc2,Misc3,Misc4,Misc5,Misc6,Misc7,Misc8,Misc9,Misc10 FROM dailyorders WHERE Misc1 ='{order_number_packing}' OR Purchase_Order = '{order_number_packing.upper()}' OR Sales_Record_Number = '{order_number_packing}' ;")
        retrieved_order = cursor.fetchall()
        # pprint.pprint(retrieved_order)

        order_id = retrieved_order[0][0]
        if '+' in order_id:
            order_id = order_id.split(' + ')

        purchase_order = retrieved_order[0][1]
        if '+' in purchase_order:
            purchase_order = purchase_order.split(' + ')

        if isinstance(purchase_order, list):
            for x in range(len(purchase_order)):
                purchase_order[x] = purchase_order[x].replace('+', '')
                purchase_order[x] = purchase_order[x].replace(' ', '')

        try:
            sales_record_number = retrieved_order[0][2]
            if '+' in sales_record_number:
                sales_record_number = sales_record_number.split(' + ')
        except:
            sales_record_number = purchase_order

        sales_channel = retrieved_order[0][3]
        name = retrieved_order[0][4]
        company = retrieved_order[0][5]
        address1 = retrieved_order[0][6]
        address2 = retrieved_order[0][7]
        city = retrieved_order[0][8]
        state = retrieved_order[0][9]
        postcode = retrieved_order[0][10]
        country = retrieved_order[0][11]
        phone = retrieved_order[0][12]
        email = retrieved_order[0][13]
        date_placed = retrieved_order[0][14]
        postage_service = retrieved_order[0][22]
        tracking_number = retrieved_order[0][23]

        un_string_order_number_packing = retrieved_order[0][24]
        string_order_number_packing = retrieved_order[0][25]
        string_order_number_packing = string_order_number_packing.split(' ')
        if str(un_string_order_number_packing) != string_order_number_packing[0]:
            print('Uh oh, bug hit. Will need to do this one manually.')
            time.sleep(1)
            continue

        item_sku = []
        item_name = []
        item_quantity = []
        ebay_item_number = []
        image_url = []
        sticky_notes = []

        for x in retrieved_order:
            item_sku.append(x[16])
            ebay_item_number.append(x[17])
            item_quantity.append(x[18])
            item_name.append(x[19])
            image_url.append(f'https://www.scarlettmusic.com.au/assets/full/{x[16]}.jpg')

        sku_location = []


        #vvvvvvvvvvvvvvvvvvvvvvvvvvv CHECKING THAT THE ORDER HASN'T BEEN MARKED AS SENT ALREADYvvvvvvvvvvvvvvvvvvvv

        if sales_channel.lower() == 'website' or sales_channel.lower() == 'kogan' or sales_channel.lower() == 'mydeal' or sales_channel.lower() == 'ozsale' or sales_channel.lower() == 'amazon au' or sales_channel.lower() == 'amazon seller a' or sales_channel.lower() == 'catch' or sales_channel.lower() == 'everydaymarket' or sales_channel.lower() == 'bigw':
            data = {'Filter': {'OrderID': order_id,
                               'OutputSelector': ["ID", "OrderStatus", "GrandTotal", "StickyNotes"]}}

            headers = {'Content-Type': 'application/json', 'Accept': 'application/json', 'NETOAPI_ACTION': 'GetOrder',
                       'NETOAPI_USERNAME': 'kyalAPI', 'NETOAPI_KEY': 'HSZ3EyI9ki6sseqRIkwy6tL6yFNPEUyM', }

            r = requests.post('https://www.scarlettmusic.com.au/do/WS/NetoAPI', headers=headers, json=data)

            response = r.text
            response = json.loads(response)

            grand_total = response['Order'][0]['GrandTotal']

            for orderstatus in response['Order']:

               if orderstatus['OrderStatus'] != 'Pick':
                   print(f'Order {order_id} has already been marked as sent. Returning to start.')
                   time.sleep(2)
                   orderfailed = 'true'
                   continue
        elif sales_channel.lower() == 'ebay':
            getordersapi = api.execute('GetOrders', {'OrderIDArray': {'OrderID': purchase_order},
                                                     'Pagination': {'EntriesPerPage': 50, 'PageNumber': 1}})
            getordersresponse = getordersapi.dict()

            for t in getordersresponse['OrderArray']['Order']:
                if 'ShippedTime' in t:
                    print(f'Order {purchase_order} has already been marked as sent. Returning to start.')
                    orderfailed = 'true'
                    time.sleep(2)
                    continue #####

                if t['OrderStatus'] == 'Cancelled':
                    print(f'Order {purchase_order} is a cancelled order. Returning to start.')
                    orderfailed = 'true'
                    time.sleep(2)
                    continue  #####


        #^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ CHECKING ORDER HASN'T BEEN MARKED AS SENT ALREADY^^^^^^^^^^^^^^^
        if orderfailed == 'true':
            continue
        for x in range(len(item_sku)):

            sku = item_sku[x]

            sku = sku.strip()
            sku = sku.replace("/", "")
            sku = sku.replace("\\", "")


            img_data = requests.get(f'https://www.scarlettmusic.com.au/assets/full/{sku}.jpg').content

            image_address = f'{documentsaddress}\\{sku}'
            with open(f'{image_address}.jpg',
                      'wb') as handler:  #############WILL NEED TO CHANGE TO DOCUMENTS FOLDER
                handler.write(img_data)
            filesize = os.path.getsize(f'{image_address}.jpg')

            if filesize <= 1000:
                if os.path.isfile(f'{image_address}.jpg') == True:
                    send2trash(f'{image_address}.jpg')
                pass

            else:

                sku_location.append(f'{image_address}.jpg')
                continue

            if filesize < 1000:

                img_data = requests.get(f'https://www.scarlettmusic.com.au/assets/full/{sku}.png').content
                with open(f'{image_address}.png',
                          'wb') as handler:  #############WILL NEED TO CHANGE TO DOCUMENTS FOLDER
                    handler.write(img_data)

                    filesize = os.path.getsize(f'{image_address}.png')

            if filesize < 1000:

                if os.path.isfile(f'{image_address}.png') == True:
                    send2trash(f'{image_address}.png')
                pass

            else:

                sku_location.append(f'{image_address}.png')
                continue

            if filesize < 1000:

                try:
                    item_number = ebay_item_number[x]
                    if '-' in item_number:
                        item_number = item_number.split('-')[0]
                    api = Trading(appid="KyalScar-Awaiting-PRD-aca833663-f54c920e", timeout=120, config_file=None,
                                  devid="5c97c2a8-eae7-49cc-ae81-c3ed1bbab335",
                                  certid="PRD-ca8336639ad2-75af-45bc-992a-d0bc",
                                  token="v^1.1#i^1#r^1#f^0#I^3#p^3#t^Ul4xMF84OkUzOEU3NzVFQzA3NDUwNjUyMkY0NTc4QjlENjRFRDVFXzFfMSNFXjI2MA==")
                    response = api.execute('GetItem', {"ItemID": item_number})
                    unformatteddic = response.dict()

                    #pprint.pprint(unformatteddic)

                    pre_picture = unformatteddic['Item']['PictureDetails']['PictureURL']

                    if isinstance(pre_picture, list):
                        pre_picture = pre_picture[0]


                    picture_url = pre_picture.split('?')[0]

                    picture_url = picture_url.replace('JPG', 'jpg')

                    img_data = requests.get(picture_url).content
                    with open(f'{image_address}.jpg',
                              'wb') as handler:  #############WILL NEED TO CHANGE TO DOCUMENTS FOLDER
                        handler.write(img_data)

                        filesize = os.path.getsize(f'{image_address}.jpg')
                        sku_location.append(f'{image_address}.jpg')
                        continue
                except:
                    img_data = requests.get(f'https://www.scarlettmusic.com.au/assets/full/{sku}.jpg').content

                    image_address = f'{documentsaddress}\\{sku}'
                    with open(f'{image_address}.jpg',
                              'wb') as handler:  #############WILL NEED TO CHANGE TO DOCUMENTS FOLDER
                        handler.write(img_data)
                    filesize = os.path.getsize(f'{image_address}.jpg')
                    sku_location.append(f'{image_address}.jpg')
                    continue


        images = [Image.open(x) for x in sku_location]
        widths, heights = zip(*(i.size for i in images))

        total_width = sum(widths)
        max_height = max(heights)

        new_im = Image.new('RGB', (total_width, max_height))

        x_offset = 0
        for im in images:
            new_im.paste(im, (x_offset, 0))
            x_offset += im.size[0]

        new_im.save(f'{documentsaddress}\\Combined.jpg')

        OpenIt = subprocess.Popen(['C:\Python\JPEGView64\JPEGView.exe', f'{documentsaddress}\\Combined.jpg'])
        print('\n\n\n\n\n\n\n\n\n\n\n')

        if sales_channel.lower() == 'website' or sales_channel.lower() == 'kogan' or sales_channel.lower() == 'mydeal' or sales_channel.lower() == 'ozsale' or sales_channel.lower() == 'amazon au' or sales_channel.lower() == 'amazon seller a' or sales_channel.lower() == 'catch' or sales_channel.lower() == 'everydaymarket' or sales_channel.lower() == 'bigw':
            console.print(Panel.fit(f'ORDER: {order_id} - {name} (GRAND TOTAL: {grand_total}), Platform: [cyan]{sales_channel.upper()}[/]'))

            for orders in response['Order']:



                try:
                    for suborders in orders['StickyNotes']:
                        #sticky_notes.append(suborders["Description"])

                        console.print(f'\n[yellow]Sticky Note\nTitle: {suborders["Title"]}\nDescription: {suborders["Description"]}[/]\n')

                except:
                    console.print(f'\n[yellow]Sticky Note\nTitle: {orders["StickyNotes"]["Title"]}\nDescription: {orders["StickyNotes"]["Description"]}[/]\n')


        else:
            console.print(Panel.fit(f'ORDER: {order_id} - {name}, Platform: [cyan]{sales_channel.upper()}[/]'))


#####vvvvvvvvvvvvvvvvv   Getting eBay item notes vvvvvvvvvvvvvvvvvvvvvvvvvv
            print("\n")

            for item_number in ebay_item_number:

                for transactions in SoldList:

                    passcondition = 'failed'

                    Stickydic = {}

                    # print(transactions)
                    if 'Transaction' in SoldList[transactions]:

                        try:
                            orderlinelen = 1

                            ebayorderline = SoldList[transactions]['Transaction'][
                                'OrderLineItemID']  ###  Checking Single Order Lines
                            # print(ebayorderline)

                            if item_number == ebayorderline:
                                try:
                                    ItemNote = SoldList[transactions]['Transaction']['Item']['PrivateNotes']
                                    #print(f'Sticky Note: {ItemNote}')
                                    sticky_notes.append(ItemNote)



                                except (KeyError, TypeError):
                                    sticky_notes.append('')
                                    pass
                        except:
                            pass
                    elif 'Order' in SoldList[transactions]:

                        try:

                            orderlinelen = len(SoldList[transactions]['Order']['TransactionArray']['Transaction'])


                            for multitransactions in SoldList[transactions]['Order']['TransactionArray'][
                                'Transaction']:  ###  Checking Multi Order Lines

                                ebayorderline = multitransactions['OrderLineItemID']

                                if item_number == ebayorderline:

                                    try:
                                        ItemNote = multitransactions['Item']['PrivateNotes']
                                        #print(f'Sticky Note: {ItemNote}')
                                        sticky_notes.append(ItemNote)
                                    except (KeyError, TypeError):
                                        sticky_notes.append('')

                                        pass
                        except(TypeError):
                            pass

#####^^^^^^^^^^^^^^^   Getting eBay item notes ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
        console.print('\n\n[underline]The following items have been ordered[/]\n\n\n')


        for lines in range(len(retrieved_order)):

            if int(item_quantity[lines]) > 1:

                if sales_channel.lower() == 'ebay':

                    console.print(Panel.fit(f'[red](QTY: {item_quantity[lines]})[/red] \n\n[white]{item_name[lines]} \n(SKU: [/white][cyan]{item_sku[lines]}[/cyan])\n\n[yellow]Sticky Note: {sticky_notes[lines]}[/]\n'))


                else:
                    console.print(Panel.fit(f'[red](QTY: {item_quantity[lines]})[/red] \n\n[white]{item_name[lines]} \n(SKU: [/white][cyan]{item_sku[lines]}[/cyan])\n'))

            else:

                if sales_channel.lower() == 'ebay':

                    console.print(f'[white]{item_name[lines]} \n(SKU: [/white][cyan]{item_sku[lines]}[/cyan])\n\n[yellow]Sticky Note: {sticky_notes[lines]}[/]\n')

                else:
                    console.print(f'[white]{item_name[lines]} \n(SKU: [/white][cyan]{item_sku[lines]}[/cyan])\n')



        console.print('''\n\n\nCan this order be sent?
        
        [white]If YES - Pack the order up and press the [/][bold green]RIGHT ARROW KEY[/]
        
        [white]If NO - Press[/] [bold red]ESC[/][white][/]
        
        Otherwise, press the [blue]LEFT ARROW KEY[/] to skip this order\n''')

        while True:

            event = keyboard.read_event()
            if event.event_type == keyboard.KEY_DOWN:
                key = event.name

                if key == 'right':
                    sent_answer = 'yes'
                    print('\nPlease wait...\n')
                    break

                if key == 'esc':
                    sent_answer = 'no'
                    break

                if key == 'left':
                    sent_answer = 'skip'
                    break

        total_order = []

        if sent_answer == 'skip':
            print('\nReturning to Start...\n')
            OpenIt.terminate()
            continue


        if sent_answer == 'yes':

            ######################### IF the postage service is 'label', it will allow either sending untracked, or
                                    # creating a tracked label, to which it will then save the postage servce + tracking number #############

            if postage_service.lower() == 'label':

                print('''
                
        Press the RIGHT ARROW KEY to send this untracked (use existing label)
        
        Press the PLUS key (+) to create tracked label.''')

                while True:


                    event = keyboard.read_event()
                    if event.event_type == keyboard.KEY_DOWN:
                        key = event.name

                        if key == 'right':

                            postage_service = 'untracked'

                            break

                        if key == '+':

                            if sales_channel.lower() == 'ebay':

                                if isinstance(purchase_order, list):
                                    postage_service, tracking_number = online_order(purchase_order[0], 'true')
                                    break

                                else:

                                    postage_service, tracking_number = online_order(purchase_order, '')
                                    break

                            else:

                                if sales_channel.lower() == 'everydaymarket' or sales_channel.lower() == 'bigw':
                                    postage_service = "Australia Post"
                                    tracking_number = phone

                                else:

                                    postage_service = 'Untracked Post'

                                if isinstance(purchase_order, list):
                                    postage_service, tracking_number = online_order(purchase_order[0], 'true')
                                    break
                                else:
                                    postage_service, tracking_number = online_order(order_id, '')
                                    break




            if sales_channel.lower() == 'website' or sales_channel.lower() == 'kogan' or sales_channel.lower() == 'mydeal' or sales_channel.lower() == 'ozsale' or 'amazon' in sales_channel.lower()  or sales_channel.lower() == 'catch' or sales_channel.lower() == 'everydaymarket' or sales_channel.lower() == 'bigw':
                for lines in range(len(retrieved_order)):

                    if postage_service == 'Fastway' or postage_service == 'Freightster':
                        postage_service = 'Aramex'
                    if postage_service.lower() == 'tnt':
                        postage_service = 'Road Express'

                    if postage_service == 'Couriers Please':
                        postage_service = "Courier's Please"

                    if postage_service == 'CBDExpress':
                        postage_service = "CBD Express"

                    if postage_service.lower() == 'label' or postage_service.lower() == 'minilope' or postage_service.lower() == 'devilope':
                        postage_service = 'Untracked Post'
                        tracking_number = 'Untracked'

                    if postage_service == 'untracked':
                        postage_service = 'Untracked Post'
                        tracking_number = 'Untracked'


                    total_order.append({
                "SKU": item_sku[lines],
                "TrackingDetails": {
                  "ShippingMethod": postage_service,
                  "TrackingNumber": tracking_number,
                  "DateShipped": now
                }
              })


                #vvvvvvvvvvvvvvvvvvvvvvv UPDATING THE MARKETPLACES THAT NEED TO BE DONE MANUALLYvvvv

                if sales_channel.lower() == 'kogan' or sales_channel.lower() == 'ozsale':

                    print(f'Sales channel {sales_channel} detected.')
                    print(f'Taking you there now to manually enter tracking number (copied to clipboard)')
                    print(f'Return here once completed')
                    pyperclip.copy(tracking_number)
                    time.sleep(1)

                    if sales_channel.lower() == 'ozsale':
                        os.system(f"start \"\" https://www.mysalemarketplace.com/login/#/")

                    if sales_channel.lower() == 'kogan':
                        os.system(f"start \"\" https://dispatch.aws.kgn.io/Manage")

                    input(f'Press ENTER once tracking has been uploaded to {sales_channel} manually')

                if isinstance(order_id, str):

                    updateorderdata = {"Order": {
                        "OrderID": order_id,
                        "OrderStatus": "Dispatched",
                        "SendOrderEmail": "tracking",
                        "OrderLine": total_order}}

                    r = requests.post('https://www.scarlettmusic.com.au/do/WS/NetoAPI', headers=updateorderheaders,
                                      json=updateorderdata)

                    response = r.text
                    response = json.loads(response)

                    if response['Ack'].lower() == 'success':
                        print(f'Order {order_id} has been marked as sent on the website.')
                        try:
                            cursor = connection.cursor()
                        except:
                            connection = create_connection("postgres", "kyal", "123456789Kyal", "143.244.166.150",
                                                           "5432")
                            cursor = connection.cursor()
                        postage_service = postage_service.replace("'", "")
                        cursor.execute(
                            fr"INSERT INTO savedorders(ORDER_ID,Purchase_Order,Sales_Record_Number,Sales_Channel,Name,Address1,Address2,City,State,Postcode,Country,Phone,Email,Date_Placed,Date_Completed,Item_Sku,Item_number_eBay,Item_Quantity,Item_Name,Postage_Service,Tracking_Number) VALUES ('{order_id}', '{order_id}', '{sales_record_number}', '{sales_channel}', '{name}', '{address1[0:50]}', '{address2[0:50]}', '{city[0:50]}', '{state}', '{postcode}', 'AU', '{phone[0:50]}', '{email}', '{date_placed}','{now}' , '{item_sku[lines]}', '', '{item_quantity[lines]}', '{item_name[lines][0:50]}', '{postage_service}','{tracking_number}'); COMMIT;")

                        time.sleep(1)
                        OpenIt.terminate()
                        continue

                else:
                    for t in order_id:
                    # vvvvvvvvvvvvvvvvvvvvvv NOW UPDATING NETO  vvvvvvvvvvvvvvvvv
                        updateorderdata = {"Order": {
                            "OrderID": t,
                            "OrderStatus": "Dispatched",
                            "SendOrderEmail": "tracking",
                            "OrderLine": total_order}}

                        r = requests.post('https://www.scarlettmusic.com.au/do/WS/NetoAPI', headers=updateorderheaders,
                                          json=updateorderdata)

                        response = r.text
                        response = json.loads(response)

                        if response['Ack'].lower() == 'success':
                            print(f'Order {t} has been marked as sent on the website.')
                            try:
                                cursor = connection.cursor()
                            except:
                                connection = create_connection("postgres", "kyal", "123456789Kyal", "143.244.166.150",
                                                               "5432")
                                cursor = connection.cursor()
                            postage_service = postage_service.replace("'", "")
                            cursor.execute(
                                fr"INSERT INTO savedorders(ORDER_ID,Purchase_Order,Sales_Record_Number,Sales_Channel,Name,Address1,Address2,City,State,Postcode,Country,Phone,Email,Date_Placed,Date_Completed,Item_Sku,Item_number_eBay,Item_Quantity,Item_Name,Postage_Service,Tracking_Number) VALUES ('{order_id}', '{t}', '{sales_record_number}', '{sales_channel}', '{name}', '{address1[0:50]}', '{address2[0:50]}', '{city[0:50]}', '{state}', '{postcode}', 'AU', '{phone[0:50]}', '{email}', '{date_placed}','{now}' , '{item_sku[lines]}', '', '{item_quantity[lines]}', '{item_name[lines][0:50]}', '{postage_service}','{tracking_number}'); COMMIT;")

                            time.sleep(1)
                            OpenIt.terminate()
                            continue


                        else:
                            print('Possible failure, see below response, maybe grab Kyal.\n')
                            pprint.pprint(response)
                            input('\nPress Enter to continue.')
                            OpenIt.terminate()
                            continue
                if len(retrieved_order) > 3:
                    input(f'Please double check orders {order_id} have been sent manually, multi-orders have not been doing so hot, press Enter when done')

            if sales_channel.lower() == 'ebay':

                api = Trading(appid="KyalScar-Awaiting-PRD-aca833663-f54c920e", timeout=120, config_file=None,
                              devid="5c97c2a8-eae7-49cc-ae81-c3ed1bbab335", certid="PRD-ca8336639ad2-75af-45bc-992a-d0bc",
                              token="v^1.1#i^1#r^1#f^0#I^3#p^3#t^Ul4xMF84OkUzOEU3NzVFQzA3NDUwNjUyMkY0NTc4QjlENjRFRDVFXzFfMSNFXjI2MA==")

                if 'label' in postage_service.lower() or 'minilope' in postage_service.lower()  or 'devilope' in postage_service.lower():
                    postage_service = 'untracked'


                    if isinstance(purchase_order, str):

                        # response = api.execute('CompleteSale', {'OrderID': purchase_order, 'Shipped': True})

                        response = api.execute('CompleteSale', {'OrderID': purchase_order, 'Shipped': 1})

                        unformatteddic = response.dict()
                        if unformatteddic['Ack'].lower() == 'success':
                            print(f'Order {purchase_order} has been marked as sent on eBay.')

                            try:
                                cursor = connection.cursor()
                            except:
                                connection = create_connection("postgres", "kyal", "123456789Kyal", "143.244.166.150",
                                                               "5432")
                                cursor = connection.cursor()
                            postage_service = postage_service.replace("'", "")

                            cursor.execute(
                                fr"INSERT INTO savedorders(ORDER_ID,Purchase_Order,Sales_Record_Number,Sales_Channel,Name,Address1,Address2,City,State,Postcode,Country,Phone,Email,Date_Placed,Date_Completed,Item_Sku,Item_number_eBay,Item_Quantity,Item_Name,Postage_Service,Tracking_Number) VALUES ('{order_id}', '{purchase_order}', '{sales_record_number}', '{sales_channel}', '{name}', '{address1[0:50]}', '{address2[0:50]}', '{city[0:50]}', '{state}', '{postcode}', 'AU', '{phone[0:50]}', '{email}', '{date_placed}','{now}' , '{item_sku[lines]}', '', '{item_quantity[lines]}', '{item_name[lines][0:50]}', '{postage_service}','{tracking_number}'); COMMIT;")

                            time.sleep(1)
                            OpenIt.terminate()
                            continue

                        else:
                            print('Possible failure, see below response, maybe grab Kyal.\n')
                           # pprint.pprint(unformatteddic)
                            input('\nPress Enter to continue.')
                            OpenIt.terminate()
                            continue

                    else:
                        for t in purchase_order:
                            response = api.execute('CompleteSale', {'OrderID': t, 'Shipped': 1})

                            unformatteddic = response.dict()
                            if unformatteddic['Ack'].lower() == 'success':
                                print(f'Order {t} has been marked as sent on the eBay.')

                                try:
                                    cursor = connection.cursor()
                                except:
                                    connection = create_connection("postgres", "kyal", "123456789Kyal",
                                                                   "143.244.166.150", "5432")
                                    cursor = connection.cursor()
                                postage_service = postage_service.replace("'", "")

                                cursor.execute(
                                    fr"INSERT INTO savedorders(ORDER_ID,Purchase_Order,Sales_Record_Number,Sales_Channel,Name,Address1,Address2,City,State,Postcode,Country,Phone,Email,Date_Placed,Date_Completed,Item_Sku,Item_number_eBay,Item_Quantity,Item_Name,Postage_Service,Tracking_Number) VALUES ('{retrieved_order[0][0]}', '{t}', '{retrieved_order[0][2]}', '{sales_channel}', '{name}', '{address1[0:50]}', '{address2[0:50]}', '{city[0:50]}', '{state}', '{postcode}', 'AU', '{phone[0:50]}', '{email}', '{date_placed}','{now}' , '{item_sku[lines]}', '', '{item_quantity[lines]}', '{item_name[lines][0:50]}', '{postage_service}','{tracking_number}'); COMMIT;")

                                time.sleep(1)
                                OpenIt.terminate()
                                continue

                            else:
                                print('Possible failure, see below response, maybe grab Kyal.\n')
                                pprint.pprint(unformatteddic)
                                input('\nPress Enter to continue.')
                                OpenIt.terminate()
                                continue
                    if len(retrieved_order) > 3:
                        input(f'Please double check orders {purchase_order} have been sent manually, multi-orders have not been doing so hot, press Enter when done')
                else:

                    ebay_postage_service = postage_service

                    # go here for courier codes https://developer.ebay.com/devzone/xml/docs/reference/ebay/types/ShippingCarrierCodeType.html

                    if postage_service.lower() == 'fastway' or postage_service.lower() == 'freightster':
                        ebay_postage_service = 'ARAMEX'

                    if postage_service == 'Allied Express':
                        ebay_postage_service = 'ALLIEDEXPRESS'

                    if postage_service == 'Australia Post':
                        ebay_postage_service = 'AustraliaPost'

                    if postage_service == 'Mailplus':
                        ebay_postage_service = 'Toll'

                    if postage_service == 'Dai Post':
                        ebay_postage_service = 'DAIPost'

                    if isinstance(purchase_order, str):

                        if postage_service.lower() == 'untracked':

                            response = api.execute('CompleteSale', {'OrderID': purchase_order, 'Shipped': 1})


                        else:

                            response = api.execute('CompleteSale', {'OrderID': purchase_order, "Shipment": {
                                "ShipmentTrackingDetails": {'ShipmentTrackingNumber': tracking_number,
                                                            'ShippingCarrierUsed': ebay_postage_service, 'Shipped': True}}})

                        unformatteddic = response.dict()
                        if unformatteddic['Ack'].lower() == 'success':
                            print(f'Order {t["OrderID"]} has been marked as sent on the eBay.')

                            try:
                                cursor = connection.cursor()
                            except:
                                connection = create_connection("postgres", "kyal", "123456789Kyal", "143.244.166.150",
                                                               "5432")
                                cursor = connection.cursor()
                            postage_service = postage_service.replace("'", "")

                            cursor.execute(
                                fr"INSERT INTO savedorders(ORDER_ID,Purchase_Order,Sales_Record_Number,Sales_Channel,Name,Address1,Address2,City,State,Postcode,Country,Phone,Email,Date_Placed,Date_Completed,Item_Sku,Item_number_eBay,Item_Quantity,Item_Name,Postage_Service,Tracking_Number) VALUES ('{retrieved_order[0][0]}', '{purchase_order}', '{retrieved_order[0][2]}', '{sales_channel}', '{name}', '{address1[0:50]}', '{address2[0:50]}', '{city[0:50]}', '{state}', '{postcode}', 'AU', '{phone[0:50]}', '{email}', '{date_placed}','{now}' , '{item_sku[lines]}', '', '{item_quantity[lines]}', '{item_name[lines][0:50]}', '{postage_service}','{tracking_number}'); COMMIT;")

                            time.sleep(1)
                            OpenIt.terminate()
                            continue


                    else:

                        for t in purchase_order:

                            if postage_service.lower() == 'untracked':

                                response = api.execute('CompleteSale', {'OrderID': t, 'Shipped': 1})


                            else:

                                response = api.execute('CompleteSale', {'OrderID': t, "Shipment": {
                                    "ShipmentTrackingDetails": {'ShipmentTrackingNumber': tracking_number,
                                                                'ShippingCarrierUsed': ebay_postage_service,
                                                                'Shipped': True}}})

                            unformatteddic = response.dict()
                            if unformatteddic['Ack'].lower() == 'success':
                                print(f'Order {t} has been marked as sent on the eBay.')

                                try:
                                    cursor = connection.cursor()
                                except:
                                    connection = create_connection("postgres", "kyal", "123456789Kyal",
                                                                   "143.244.166.150", "5432")
                                    cursor = connection.cursor()
                                postage_service = postage_service.replace("'", "")

                                cursor.execute(
                                    fr"INSERT INTO savedorders(ORDER_ID,Purchase_Order,Sales_Record_Number,Sales_Channel,Name,Address1,Address2,City,State,Postcode,Country,Phone,Email,Date_Placed,Date_Completed,Item_Sku,Item_number_eBay,Item_Quantity,Item_Name,Postage_Service,Tracking_Number) VALUES ('{retrieved_order[0][0]}', '{t}', '{retrieved_order[0][2]}', '{sales_channel}', '{name}', '{address1[0:50]}', '{address2[0:50]}', '{city[0:50]}', '{state}', '{postcode}', 'AU', '{phone[0:50]}', '{email}', '{date_placed}','{now}' , '{item_sku[lines]}', '', '{item_quantity[lines]}', '{item_name[lines][0:50]}', '{postage_service}','{tracking_number}'); COMMIT;")

                                time.sleep(1)
                                OpenIt.terminate()
                                continue

                            else:
                                print('Possible failure, see below response, maybe grab Kyal.\n')
                                pprint.pprint(unformatteddic)
                                OpenIt.terminate()
                                input('\nPress Enter to continue.')
                                continue

        if sent_answer == 'no':

            cancel_answer = 'no'

            if 'label' not in postage_service.lower() and 'minilope' not in postage_service.lower() and 'devilope' not in postage_service.lower() and 'untracked' not in postage_service.lower() and 'mailplus' not in postage_service.lower() and 'fastway' not in postage_service.lower():

                console.print(f'''\n\n\n
                [white]Press[/] [bold red]ESC[/][white] to cancel your {postage_service} label[/]
                
                [white]Press [/][bold green]BACKSLASH[/][white] (AKA \) to keep partially send order with {postage_service} label[/]
        
                ''')

                while True:

                    event = keyboard.read_event()
                    if event.event_type == keyboard.KEY_DOWN:
                        key = event.name

                        if key == '\\':
                            cancel_answer = 'no'
                            break

                        if key == 'esc':
                            cancel_answer = 'yes'
                            break

            if cancel_answer == 'yes':

                #vvvvvvvvvvvvvvvvvvvvvv Cancelling congisnment if order cannot be fulfilled#######

                if postage_service.lower() == 'sendle':
                    print(f'Taking you to {postage_service} now to manually cancel label')
                    print(f'Return here once completed')
                    time.sleep(2)
                    os.system(f"start \"\" https://app.sendle.com/dashboard")

                elif postage_service.lower() == 'australia post':
                    accountnumber = '0007805312'
                    headers = {'Content-Type': 'application/json', 'Accept': 'application/json',
                               'Account-Number': accountnumber}
                    account = accountnumber
                    username = '966afe1c-c07b-4902-b1d7-77ad1aac0915'
                    secret = 'x428429fabeb99f420f1'

                    r = requests.get(
                        'https://digitalapi.auspost.com.au/shipping/v1/shipments?offset=0&number_of_shipments=1000&status=Created',
                        headers=headers, auth=HTTPBasicAuth(username, secret))

                    response = r.text
                    response = json.loads(response)
                    for orders in response['shipments']:
                        iterating_tracking_number = orders['items'][0]['tracking_details']['consignment_id']
                        if iterating_tracking_number == tracking_number:
                            shipment_id = orders['shipment_id']
                            break

                   # pprint.pprint(response)

                    r = requests.delete(
                        f'https://digitalapi.auspost.com.au/shipping/v1/shipments/{shipment_id}',
                        headers=headers, auth=HTTPBasicAuth(username, secret))

                    response = r.text
                    print(f'{postage_service} label deleted!')

                elif postage_service.lower() == 'freightster':
                    try:
                        cursor = connection.cursor()
                    except:
                        connection = create_connection("postgres", "kyal", "123456789Kyal", "143.244.166.150", "5432")
                        cursor = connection.cursor()
                    cursor.execute(f"DELETE FROM Freightster WHERE tracking_number = '{tracking_number}'; COMMIT;")
                    print(f'{postage_service} label deleted!')

                elif postage_service.lower() == 'toll':
                    try:
                        cursor = connection.cursor()
                    except:
                        connection = create_connection("postgres", "kyal", "123456789Kyal", "143.244.166.150", "5432")
                        cursor = connection.cursor()
                    cursor.execute(
                        f"UPDATE toll SET manifest_number = 'DELETED' WHERE ShipmentID = '{order_id}'; COMMIT;")
                    print('Success!')
                    print(f'{postage_service} label deleted!')

                elif postage_service.lower() == 'dai post':
                    payload = {
                        "cancelshipment": {'tracknbr': str(tracking_number)}
                    }

                    r = requests.post('https://daiglobaltrack.com/test/serviceconnect',
                                      auth=HTTPBasicAuth('ScarlettMusic', 'D5es4stu!'), json=payload)

                    response = r.text

                    dai_response = json.loads(response)

                    print(dai_response)
                    try:
                        cursor = connection.cursor()
                    except:
                        connection = create_connection("postgres", "kyal", "123456789Kyal", "143.244.166.150", "5432")
                        cursor = connection.cursor()
                    cursor.execute(f"DELETE FROM dai_post WHERE tracking_number = '{tracking_number}'; COMMIT;")

                    print('Success!')
                    print(f'{postage_service} label deleted!')
                    
                elif postage_service.lower() == 'cbdexpress':

                    reciever_email = 'kyal@scarlettmusic.com.au'

                    subject = f'Delete Consignment {tracking_number}'

                    body = f'''Hey Jarrod,

                        Could we please cancel consignment {tracking_number} please?

                        Sorry about that!

                        Cheers,
                        Kyal'''

                    path_to_documents = ''
                    m = Mail()
                    m.send_email(reciever_email, subject, body, path_to_documents, attachment_file)

                    try:
                        cursor = connection.cursor()
                    except:
                        connection = create_connection("postgres", "kyal", "123456789Kyal", "143.244.166.150", "5432")
                        cursor = connection.cursor()
                    cursor.execute(
                        f"UPDATE CDBExpress SET name = 'DELETED' WHERE consignment_number = '{order_id}'; COMMIT;")
                    print('Success!')
                    print(f'{postage_service} label deleted!')





                elif 'allied' in postage_service.lower():
                    cancel = allied_client.service.cancelDispatchJob('755cf13abb3934695f03bd4a75cfbca7', tracking_number,
                                                                     postcode)
                    if str(cancel) == '0':
                        print(f'Label successfully cancelled for {postage_service}')
                        time.sleep(2)
                        sys.exit()

                    elif str(cancel) == '-6':
                        print('This has already been cancelled dingus')
                        time.sleep(2)
                        sys.exit()

                    else:
                        print("Fuck bro not sure if this cancelled, maybe tell Kyal?")
                        time.sleep(2)
                        sys.exit()

                elif postage_service.lower() == 'couriers please':
                    cp_auth = '113153878:CBBC1CAD2F335C9CEEA1D0BC63C7056ADBC9C36DD54A9A4530A4380D9AAC5FE0'
                    ### This is for sandbox, when ready, replace 2nd half with CBBC1CAD2F335C9CEEA1D0BC63C7056ADBC9C36DD54A9A4530A4380D9AAC5FE0

                    cp_auth_encoded = cp_auth.encode()
                    cp_auth_encoded = base64.b64encode(cp_auth_encoded)
                    cp_auth_encoded = cp_auth_encoded.decode()

                    cp_headers = {'Host': 'api.couriersplease.com.au',
                                  'Accept': 'application/json',
                                  'Content-Type': 'application/json',
                                  'Authorization': f'Basic {cp_auth_encoded}',
                                  'Content-Length': '462'}

                    cp_url = 'https://api.couriersplease.com.au/v1/domestic/shipment/cancel'
                    cp_body = {"consignmentCode": tracking_number}

                    r = requests.post(cp_url, headers=cp_headers, json=cp_body)
                    response = r.text
                    response = json.loads(response)
                    #pprint.pprint(response)
                    time.sleep(2)

            while True:

                note = input('''What note would you like to leave on the order?
                
        1) ON PO
        2) Dropship
        3) Cannot find stock
        4) Custom note
        5) No note''')

                if note == '1':
                    note = 'ON PO'
                    break
                if note == '2':
                    note = 'Dropship'
                    break
                if note == '3':
                    note = 'Cannot find stock'
                    break
                if note == '4':
                    note = input('''What note would you like to leave?
                    
                    Enter Response:''')
                    break
                if note == '5':
                    break

                else:
                    print('Valid answer not found. Try Again.')
                    time.sleep(1)
                    continue
            if note == '5':
                continue

            if sales_channel.lower() == 'ebay':

                updateorderheaders = {'Content-Type': 'application/json', 'Accept': 'application/json',
                                      'NETOAPI_ACTION': 'UpdateOrder',
                                      'NETOAPI_USERNAME': 'kyalAPI', 'NETOAPI_KEY': 'HSZ3EyI9ki6sseqRIkwy6tL6yFNPEUyM'}

                if isinstance(order_id, str):
                    updateorderdata = {
                        "Order": {
                            "OrderID": order_id,
                            "StickyNotes": {
                                "StickyNote": [{
                                    "Title": note,
                                    "Description": note
                                }]

                            }
                        }}

                    r = requests.post('https://www.scarlettmusic.com.au/do/WS/NetoAPI', headers=updateorderheaders,
                                      json=updateorderdata)

                    response = r.text
                    response = json.loads(response)

                    if response['Ack'].lower() == 'success':
                        print(f'Note added to Order {order_id}')
                        time.sleep(1)


                    else:
                        print('Possible failure, see below response, maybe grab Kyal.\n')
                        pprint.pprint(response)
                        input('\nPress Enter to continue.')


                else:
                    for t in order_id:
                        updateorderdata = {
                            "Order": {
                                "OrderID": t,
                                "StickyNotes": {
                                    "StickyNote": [{
                                        "Title": note,
                                        "Description": note
                                    }]

                                }
                            }}

                        r = requests.post('https://www.scarlettmusic.com.au/do/WS/NetoAPI', headers=updateorderheaders,
                                          json=updateorderdata)

                        response = r.text
                        response = json.loads(response)

                        if response['Ack'].lower() == 'Success':
                            print(f'Note added to Order {order_id}')
                            OpenIt.terminate()
                            continue

                        else:
                            print('Possible failure, see below response, maybe grab Kyal.\n')
                            pprint.pprint(response)
                            input('\nPress Enter to continue.')
                            OpenIt.terminate()
                            continue

                pyperclip.copy(note)
                print('Note copied to clipboard. Taking you to eBay now so note can be manually entered')
                print(f'Return here once completed')

                if isinstance(purchase_order, str):
                    os.system(f"start \"\" https://www.ebay.com.au/sh/ord/details?orderid={purchase_order}")
                else:
                    for t in purchase_order:
                        os.system(f"start \"\" https://www.ebay.com.au/sh/ord/details?orderid={t}")

                input(f'Press ENTER once done.')

            else:
                updateorderheaders = {'Content-Type': 'application/json', 'Accept': 'application/json',
                                      'NETOAPI_ACTION': 'UpdateOrder',
                                      'NETOAPI_USERNAME': 'kyalAPI', 'NETOAPI_KEY': 'HSZ3EyI9ki6sseqRIkwy6tL6yFNPEUyM'}

                if isinstance(order_id, str):
                    updateorderdata = {
                        "Order": {
                            "OrderID": order_id,
                            "StickyNotes": {
                                "StickyNote": [{
                                    "Title": note,
                                    "Description": note
                                }]

                            }
                        }}

                    r = requests.post('https://www.scarlettmusic.com.au/do/WS/NetoAPI', headers=updateorderheaders,
                                      json=updateorderdata)

                    response = r.text
                    response = json.loads(response)

                    if response['Ack'].lower() == 'success':
                        print(f'Note added to Order {order_id}')
                        time.sleep(1)


                    else:
                        print('Possible failure, see below response, maybe grab Kyal.\n')
                        pprint.pprint(response)
                        input('\nPress Enter to continue.')


                else:
                    for t in order_id:
                        updateorderdata = {
                            "Order": {
                                "OrderID": t,
                                "StickyNotes": {
                                    "StickyNote": [{
                                        "Title": note,
                                        "Description": note
                                    }]

                                }
                            }}

                        r = requests.post('https://www.scarlettmusic.com.au/do/WS/NetoAPI', headers=updateorderheaders,
                                          json=updateorderdata)

                        response = r.text
                        response = json.loads(response)

                        if response['Ack'].lower() == 'Success':
                            print(f'Note added to Order {order_id}')
                            OpenIt.terminate()
                            continue

                        else:
                            print('Possible failure, see below response, maybe grab Kyal.\n')
                            pprint.pprint(response)
                            input('\nPress Enter to continue.')
                            OpenIt.terminate()
                            continue




        OpenIt.terminate()

        try:
            if os.path.isfile(f'{image_address}.jpg') == True:
                send2trash(f'{image_address}.jpg')
            if os.path.isfile(f'{image_address}.png') == True:
                send2trash(f'{image_address}.png')

        except:
            pass
        continue

    except Exception as e:

        print('Error found. \n\n')

        print(traceback.format_exc())

        print(repr(e))

        input('\n\n Press Enter to return to the start.')

        continue













