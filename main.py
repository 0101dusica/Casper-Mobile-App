import json
import kivy
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import openpyxl
from openpyxl.styles import *
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.gridlayout import GridLayout
from kivy.uix.label import Label
from kivy.uix.scrollview import ScrollView
from kivy.uix.popup import Popup
from kivy.uix.textinput import TextInput
from kivy.config import Config
from kivy.core.window import Window

kivy.require('2.0.0')
Config.set('graphics', 'mode', 'rgb')

def clear_data():
    # Clear all the text inputs
    app.table.clear_widgets()

def save_data(info = 0):
    # TODO, ispraviti pretragu, potrebno je izvuci podatke iz tabele
    global app
    for i in range(len(app.table.labels)):
        name = app.table.labels[i].text
        popis =  app.table.text_inputs[i]
    
        with open("drinks.json", "r") as f:
            data = json.load(f)
        
        for section in data:
            for item in data[section]:
                if item["naziv"].lower() == name.lower():
                    item["popis"] = int(popis) + item["popis"]

        jsonString = json.dumps(data)
        
        with open("drinks.json", "w") as f:
            f.write(jsonString)

    # Show a message box to confirm the save
    if(info ==0):
        app.table.labels.clear()
        app.table.text_inputs.clear()
        clear_data()
        show_confirmation_popup(1)

def restart_popis():
    with open("drinks.json", "r") as f:
        data = json.load(f)
        
    for section in data:
        for item in data[section]:
            item["popis"] = 0

    jsonString = json.dumps(data)
    
    with open("drinks.json", "w") as f:
        f.write(jsonString)

def submit_data():
    global app
    save_data(1)
    data = []
    # Get the data from the table
    with open("drinks.json", "r") as f:
            input = json.load(f)
        
    for section in input:
        for item in input[section]:
            name = item["naziv"].lower()
            popis =  item["popis"]
            data.append({"naziv":name, "popis":popis})

    # create a new workbook
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))
    
    grey_fill = PatternFill(start_color='d3d3d3', end_color='d3d3d3', fill_type='solid')

    worksheet['A1'] = 'Naziv Artikla'
    worksheet['A1'].alignment = Alignment(horizontal='center', vertical='center')
    worksheet['A1'].font = Font(bold=True)
    worksheet['A1'].border = thin_border
    worksheet['A1'].fill = grey_fill
    worksheet['B1'] = 'Popisano Stanje Artikla'
    worksheet['B1'].alignment = Alignment(horizontal='center', vertical='center')
    worksheet['B1'].font = Font(bold=True)
    worksheet['B1'].border = thin_border
    worksheet['B1'].fill = grey_fill

    
    worksheet.column_dimensions['A'].width = 35
    worksheet.column_dimensions['B'].width = 20

    count = 1
    for i in range(len(data)):
        count = count + 1
        worksheet['A' + str(count)] = data[i]["naziv"].capitalize()
        worksheet['A'+ str(count)].alignment = Alignment(horizontal='center', vertical='center')
        worksheet['A'+ str(count)].font = Font(bold=True)
        worksheet['A'+ str(count)].border = thin_border
        worksheet['B' + str(count)] = data[i]["popis"]
        worksheet['B'+ str(count)].alignment = Alignment(horizontal='center', vertical='center')
        worksheet['B'+ str(count)].font = Font(bold=True)
        worksheet['B'+ str(count)].border = thin_border
    
    workbook.save('popis.xlsx')

    # set up the SMTP server
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    smtp_username = 'trbovicdusica@gmail.com'
    smtp_password = 'yzerjqneezzkmdey'
    sender_email = smtp_username
    receiver_email = 'trbovicdusica@gmail.com'

    # create the message object
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = 'Popisano stanje artikala - ' + str(datetime.now().day) + "." + str(datetime.now().month) + "." + str(datetime.now().year)

    # open the file in bynary
    binary_file = open('popis.xlsx', 'rb')

    # create the attachment object
    payload = MIMEBase('application', 'octate-stream', Name='popis.xlsx')
    payload.set_payload((binary_file).read())

    # encode the attachment in base64
    encoders.encode_base64(payload)

    # add header with pdf name
    payload.add_header('Content-Decomposition', 'attachment', filename='popis.xlsx')
    msg.attach(payload)

    # create SMTP session
    session = smtplib.SMTP(smtp_server, smtp_port)

    # start TLS for security
    session.starttls()

    # login to the email server
    session.login(smtp_username, smtp_password)

    # send the email
    text = msg.as_string()
    session.sendmail(sender_email, receiver_email, text)
    
    # end the session
    session.quit()

    app.table.labels.clear()
    app.table.text_inputs.clear()
    clear_data()
    restart_popis()

    #popup_sumbit = Popup(title='Obavestenje!', content=Label(text='Email je Uspesno Poslat!'), size_hint=(None, None), size=(400, 200))
    #popup_sumbit.open()

    show_confirmation_popup()

def confirm_action():
    global app
    app.stop()
    app = MyApp()
    app.run()
    
def show_confirmation_popup(check = 0):
    # Create the layout for the popup
    layout = BoxLayout(orientation='vertical')
    if(check == 0):
        layout.add_widget(Label(text='Email je uspesno Poslat!'))
    else:        
        layout.add_widget(Label(text='Svi podaci su uspesno Sacuvani!'))

    # Create the "OK" button, just for user to get message
    yes_button = Button(text='OK')

    # Bind the buttons to their respective actions
    yes_button.bind(on_press=lambda x: [popup.dismiss(), confirm_action()])

    # Add the buttons to the layout
    button_layout = BoxLayout()
    button_layout.add_widget(yes_button)
    layout.add_widget(button_layout)

    # Create the popup and open it
    popup = Popup(title='Confirmation', content=layout, size_hint=(None, None), size=(400, 200))
    popup.open()
    
class Table(GridLayout):
    def __init__(self, **kwargs):
        super(Table, self).__init__(**kwargs)
        
        self.labels = []
        self.text_inputs = []

        self.background_color = (0.62, 0.55, 0.37, 1)  # set background color to white
        self.header_color = (0.12, 0.12, 0.12, 1)  # set header color to blue
        self.border_width = (1, 1, 1, 1)  # set border width to 1 pixel
        self.cols = 3
        self.row_force_default = True
        self.row_default_height = 40
        self.size_hint_y = None
        self.add_widget(MyLabel(text='Naziv Artikla'))
        self.add_widget(MyLabel(text='Popisano Stanje'))
        self.add_widget(MyLabel(text='Dodaj Novu Kolicinu'))
        with open('drinks.json') as f:
            data = json.load(f)
        
        self.items = []
        self.quantities = []
        self.popis = []
        for section in data:
            for item in data[section]:
                if(app.search_box.text == ""):
                    self.items.append(item['naziv'])
                    self.quantities.append(item['popis'])
                    self.popis.append(0)
            
        for i in range(len(self.items)):
            label = MyLabel(self.items[i], i)
            self.labels.append(label)
            self.add_widget(label)
            label = MyLabel(str(self.quantities[i]), i)
            self.add_widget(label)
            text_input = MyTextInput(i)
            self.text_inputs.append(self.popis[i])
            self.add_widget(text_input)

class MyLabel(Label):
    def __init__(self, text, id = -1, **kwargs):
        super(MyLabel, self).__init__(**kwargs)
        self.id = id
        self.text = text
        self.background_color = (0.12, 0.12, 0.12, 1)  # set text color to black
        self.color = (0.62, 0.55, 0.37, 1)  # set background color to white

class MyButton(Button):
    def __init__(self, text, **kwargs):
        super().__init__(**kwargs)
        self.text = text
        self.background_normal = '' # remove the default background image
        self.background_color = (0.12, 0.12, 0.12, 1)  # set text color to black
        self.color = (0.62, 0.55, 0.37, 1)  # set background color to white

    def on_press(self):
        if(self.text.lower() == "save"):    
            save_data()
        else:
            # Create the layout for the popup
            layout = BoxLayout(orientation='vertical')
            layout.add_widget(Label(text='Da li ste gotovi sa popisom?\n Svi podaci ce biti prosledjeni na emal: casper.trbovic@gmail.com', halign='center', font_size=14, size_hint_y=0.6, text_size=(600, None)))

            # Create the "Yes" and "No" buttons
            yes_button = Button(text='Potvrdi')
            no_button = Button(text='Odustani')

            # Bind the buttons to their respective actions
            yes_button.bind(on_press=lambda x: [popup.dismiss(), submit_data()])
            no_button.bind(on_press=lambda x: popup.dismiss())

            # Add the buttons to the layout
            button_layout = BoxLayout()
            button_layout.add_widget(yes_button)
            button_layout.add_widget(no_button)
            layout.add_widget(button_layout)

            # Create the popup and open it
            popup = Popup(title='Confirmation', content=layout, size_hint=(0.6, 0.3), title_align = 'center')
            popup.open()
    
class MyTextInput(TextInput):
    def __init__(self, id = -1, **kwargs):
        super(MyTextInput, self).__init__(**kwargs)
        self.id = id
        self.background_color = (0.62, 0.55, 0.37, 1)  # set background color to white
        self.color = (0.12, 0.12, 0.12, 1)  # set text color to black
    
    def on_text(self, instance, value):
        # this method will be called every time the text changes
        app.table.text_inputs[self.id] = value

class MyChangeTextInput(TextInput):
    def __init__(self, **kwargs):
        super(MyChangeTextInput, self).__init__(**kwargs)
        self.background_color = (0.62, 0.55, 0.37, 1)  # set background color to white
        self.color = (0.12, 0.12, 0.12, 1)  # set text color to black
    
    def on_text(self, instance, value):
        # filter the items to be displayed based on the user input
        filtered_items = [(item, quantity) for item, quantity in zip(app.table.items, app.table.quantities) if value.lower() in item.lower()]
        
        # update the table with the filtered items
        app.table.clear_widgets()
        for i, (item, quantity) in enumerate(filtered_items):
            label = MyLabel(item, i)
            app.table.labels.append(label)
            app.table.add_widget(label)
            label = MyLabel(str(quantity), i)
            app.table.add_widget(label)
            text_input = MyTextInput(i)
            app.table.text_inputs.append(0)
            app.table.add_widget(text_input)

class MyApp(App):
    def build(self):
        Window.clearcolor = (0.12, 0.12, 0.12, 1)
        self.window = GridLayout()
        self.window.cols = 1
        self.search_box = MyChangeTextInput(size_hint_y=None, height='40dp', hint_text = "Unesite Artikal Koji Pretrazujete..", hint_text_color = (0.12, 0.12, 0.12, 1))
        self.table = Table()
        
        self.scroll_view = ScrollView()
        self.scroll_view.add_widget(self.table)
        
        self.window.add_widget(self.search_box)
        self.window.add_widget(self.scroll_view)

        self.buttons_layout = BoxLayout(size_hint_y=None, height='30dp')
        self.save_button = MyButton(text='Save')
        self.submit_button = MyButton(text='Submit')
        self.buttons_layout.add_widget(self.save_button)
        self.buttons_layout.add_widget(self.submit_button)
        
        self.window.add_widget(self.buttons_layout)

        return self.window

# TODO Uraditi Step by Step Buildozer Installation, Sve prebaciti u Ubuntu, U isti folder dodati icon, i open pop-up image

if __name__ == '__main__':
    app = MyApp()
    app.run()
