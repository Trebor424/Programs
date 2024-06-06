

import SQL_Command
import Math_source

import json
import datetime
import tkinter as tk
from tkinter import ttk 
from tkcalendar import DateEntry
from datetime import datetime, timedelta
import matplotlib.pyplot as plt
import os

dict_chart_result = {}

class MainForm:

    def __init__(self, root):
        
        self.root = root
        self.configure_window()
        self.init_forms_combobox()
        self.init_proces_combobox()
        self.init_time_select()
        self.init_treeviews_manufacture()
        self.init_treeviews_tests()
        self.init_host_comp_select()
        self.init_search_button()
        self.create_toggle_switch()

    def configure_window(self):
        self.root.geometry("1500x800")
        self.root.title("Statistica")
        
    def create_toggle_switch(self):
        # Plik to save switch status
        self.STATE_FILE = 'toggle_switch_state.json'

        # Function to change status of switch
        def save_state(state):
            with open(self.STATE_FILE, 'w') as f:
                json.dump(state, f)

        # Function to read of status switch
        def load_state():
            if os.path.exists(self.STATE_FILE):
                with open(self.STATE_FILE, 'r') as f:
                    return json.load(f)
            return {'toggle_state': False}

        # Function to switch status
        def toggle_switch():
            self.toggle_state = not self.toggle_state
            save_state({'toggle_state': self.toggle_state})
            update_switch_text()

        # Function to check text basic on status of switch
        def update_switch_text():
            if self.toggle_state:
                toggle_button.config(text="Wylaczone")
            else:
                toggle_button.config(text="Wlaczone")

        # load plik status
        state = load_state()
        self.toggle_state = state['toggle_state']

        # Create switch
        style = ttk.Style()
        style.configure('TButton', font=('Helvetica', 16))

        toggle_button_label = tk.Label(self.root, text="Wyswietlanie Grup", font=("Arial", 12)) 
        toggle_button_label.grid(row=9, column=0, padx=10, pady=10)
        
        toggle_button = ttk.Button(self.root, text="", command=toggle_switch, style='TButton')
        toggle_button.grid(row=10, column=0, padx=10, pady=10)  # Umieœæ prze³¹cznik na siatce w g³ównym oknie

        # Load the first status 
        update_switch_text()

    def init_search_button(self):
        search_button = tk.Button(self.root, text="Wyszukaj", command=self.on_search_button_click)
        search_button.grid(row=1, column=7, padx=10, pady=10)
        

    def init_forms_combobox(self):
        server_label = tk.Label(self.root, text="Wybierz Serwer ", font=("Arial", 12)) 
        server_label.grid(row=1, column=0, padx=10, pady=10)

        server_combobox_values = ["serwerName"]
        self.server_combobox = ttk.Combobox(self.root, values=server_combobox_values)
        self.server_combobox.set("serwerName")
        self.server_combobox.grid(row=2, column=0, padx=10, pady=10)

        database_label = tk.Label(self.root, text="Wybierz Baze Danych: ", font=("Arial", 12)) 
        database_label.grid(row=3, column=0, padx=10, pady=10)
        

        database_combobox_values = ["Nazwa1","Nazwa2", "Nazwa3"]
        #database_combobox_values =SQL_Command.database_init(self.get_selected_server)
        self.database_combobox = ttk.Combobox(self.root, values=database_combobox_values)
        #self.database_combobox.set("PCCI-kopia")
        self.database_combobox.grid(row=4, column=0, padx=10, pady=10)

    def init_proces_combobox(self): 
        process_label = tk.Label(self.root, text="Wybierz Proces Testowania", font=("Arial", 12))
        process_label.grid(row=1, column=2, padx=10, pady=10)

        process_combobox_values = ["FPT", "ICT", "FFT"]
        self.process_combobox = ttk.Combobox(self.root, values=process_combobox_values)
        self.process_combobox.grid(row=1, column=3, padx=10, pady=10)
        
         # Add action to combobox when bing him
        self.process_combobox.bind("<<ComboboxSelected>>", self.on_combobox_selected)
        
    def init_host_comp_select(self):
        machine_label = tk.Label(self.root, text="Wybierz Hosta", font=("Arial", 12))
        machine_label.grid(row=1, column=4, padx=10, pady=10)

        self.machine_combobox = ttk.Combobox(self.root)
        self.machine_combobox.grid(row=1, column=5, padx=10, pady=10)

        # Add action after bing of combobox element
        self.machine_combobox.bind("<<ComboboxSelected>>", self.on_combobox_selected)
            
    def init_time_select(self):
        start_date_label = tk.Label(self.root, text="Wybierz date od kiedy: ", font=("Arial", 12))
        start_date_label.grid(row=5, column=0, padx=10, pady=10)

        self.start_date_entry = DateEntry(self.root, width=14, background="darkblue", foreground="white", borderwidth=4, date_pattern="dd/mm/yy")
        self.start_date_entry.grid(row=6, column=0, padx=10, pady=10)
        self.start_date_entry.set_date(datetime.now() - timedelta(days=60))

        end_date_label = tk.Label(self.root, text="Wybierz date do kiedy: ", font=("Arial", 12))
        end_date_label.grid(row=7, column=0, padx=10, pady=10)

        self.end_date_entry = DateEntry(self.root, width=14, background="darkblue", foreground="white", borderwidth=4 , date_pattern="dd/mm/yy")
        self.end_date_entry.grid(row=8, column=0, padx=10, pady=10)
    
    def init_treeviews_manufacture(self):
        # Replace the Listbox with a Treeview
        self.treeview1 = ttk.Treeview(self.root,height = 35 , columns=("Name","PanelPrefix", "SerialNumberPrefix"), show="headings")
        
        self.treeview1.heading("Name", text="Name",command=lambda: self.sort_treeview_string(self.treeview1, "Name"))
        self.treeview1.heading("PanelPrefix", text="PanelPrefix",command=lambda: self.sort_treeview_string(self.treeview1, "PanelPrefix"))
        self.treeview1.heading("SerialNumberPrefix", text="SerialNumberPrefix",command=lambda: self.sort_treeview_string(self.treeview1, "SerialNumberPrefix"))

        # Set column widths
        self.treeview1.column("Name", width=100)
        self.treeview1.column("PanelPrefix", width=100)
        self.treeview1.column("SerialNumberPrefix", width=200)
        self.treeview1.place(x=200, y=40)
        self.treeview1.bind("<Double-Button-1>", self.on_treeviews_manufacture_click)
        
        scrollbar_treeview1 = ttk.Scrollbar(self.root, orient="vertical", command=self.treeview1.yview)
        scrollbar_treeview1.place(x=605, y=40, height=725)
        self.treeview1.configure(yscrollcommand=scrollbar_treeview1.set)

          
    def init_treeviews_tests(self):
        self.treeview2 = ttk.Treeview(self.root, height=35, columns=("Fixture", "Test_Name", "Group", "CP", "CPK", "PP", "PPK", "O_stand", "Limit_Change"), show="headings")
        self.treeview2.heading("Fixture", text="Fixture", command=lambda: self.sort_treeview_string(self.treeview2, "Fixture"))
        self.treeview2.heading("Test_Name", text="Test_Name", command=lambda: self.sort_treeview_string(self.treeview2, "Test_Name"))
        self.treeview2.heading("Group", text="Group", command=lambda: self.sort_treeview_value(self.treeview2, "Group"))
        self.treeview2.heading("CP", text="CP", command=lambda: self.sort_treeview_value(self.treeview2, "CP"))
        self.treeview2.heading("CPK", text="CPK", command=lambda: self.sort_treeview_value(self.treeview2, "CPK"))
        self.treeview2.heading("PP", text="PP", command=lambda: self.sort_treeview_value(self.treeview2, "PP"))
        self.treeview2.heading("PPK", text="PPK", command=lambda: self.sort_treeview_value(self.treeview2, "PPK"))
        self.treeview2.heading("O_stand", text="O_stand", command=lambda: self.sort_treeview_value(self.treeview2, "O_stand"))
        self.treeview2.heading("Limit_Change", text="Limit_Change", command=lambda: self.sort_treeview_string(self.treeview2, "Limit_Change"))

        # Set column widths
        self.treeview2.column("Fixture", width=50)
        self.treeview2.column("Test_Name", width=150)
        self.treeview2.column("Group", width=50)
        self.treeview2.column("CP", width=50)
        self.treeview2.column("CPK", width=50)
        self.treeview2.column("PP", width=50)
        self.treeview2.column("PPK", width=50)
        self.treeview2.column("O_stand", width=100)
        self.treeview2.column("Limit_Change", width=100)

        self.treeview2.bind("<<TreeviewSelect>>", self.chart_charts_create)
        self.treeview2.place(x=640, y=40)
        
        scrollbar_treeview2 = ttk.Scrollbar(self.root, orient="vertical", command=self.treeview2.yview)
        scrollbar_treeview2.place(x=1295, y=40 , height=725)
        self.treeview2.configure(yscrollcommand=scrollbar_treeview2.set)

    def chart_charts_create(self, event):
        plt.close()
        
        selected_items = self.treeview2.selection()
                
        if selected_items:
            item_id = selected_items[0]
            item_index = self.treeview2.index(item_id)
            item_values = self.treeview2.item(item_id)
            test_name = item_values['values'][1]
        
            selected_items2 = self.treeview1.selection()
            item_id2 = selected_items2[0]
            item_values2 = self.treeview1.item(item_id2)
            board_name = item_values2['values'][0]
            
 
            if self.toggle_state is True:
                with open('calculate_magazine.txt', 'r') as file:
                    for num , line in enumerate(file,1):
                        if test_name in line:
                            Dane=line.split(';')
                            print(Dane , self)
                            Math_source.Chart_create(Dane)
                    Dane=[]
  
            elif  self.toggle_state is False:
                print("Nie skonczyles")
                test_name=test_name[:-2]
                with open('calculate_magazine.txt', 'r') as file:
                    dict_values={}
                    i=0
                    for num , line in enumerate(file,1):
                        if test_name in line:
                            i=i+1
                            Dane=line.split(';')
                            print(Dane , self)
                            #list_values.append(Dane)
                            dict_values[f"List_{i}"]=Dane
                    Math_source.Chart_create_group(dict_values)
                        
        selected_items = self.treeview2.selection()
        if selected_items:
            # Get the ID of the selected row
            item_id = selected_items[0]
            item_index = self.treeview2.index(item_id)
            item_values = self.treeview2.item(item_id)
            test_name = item_values['values'][1]
            
            selected_items2 = self.treeview1.selection()
            item_id2 = selected_items2[0]
            item_values2 = self.treeview1.item(item_id2)
            board_name = item_values2['values'][0]
            
 
      
    def add_to_treeview1_manufacture(self, values):
        # Insert values into the Treeview
        self.treeview1.insert("", tk.END, values=values)
        
    def add_to_treeview2_testresult(self, values):
        # Insert values into the Treeview
        self.treeview2.insert("", tk.END, values=values)
        
    def clear_to_treeview_manufacture(self):
        # Delete all items from the Treeview
        self.treeview1.delete(*self.treeview1.get_children())
        
    def clear_to_treeview_tests(self):
        # Delete all items from the Treeview
        self.treeview2.delete(*self.treeview2.get_children())
        
    def add_to_treeview_test(self, values):
        # Insert values into the Treeview
        self.treeview2.insert("", tk.END, values=values)        
    
    
    def on_treeviews_manufacture_click(self, event):
        
        # Clear secound treeview2 
        self.clear_to_treeview_tests()

        selected_items = self.treeview1.selection()
        
        if selected_items:
            item_id = selected_items[0]

            Test_Proces_Values = self.treeview1.item(item_id, "values")
        serial_number_prefix=Test_Proces_Values[2]
        panel_number_prefix=Test_Proces_Values[1]
        Name=Test_Proces_Values[0]
        
        test_databases = SQL_Command.test_databases(self.get_selected_server(),self.get_selected_database(), self.get_selected_process(), self.get_selected_host(),self.get_selected_start_date(), self.get_selected_end_date(),serial_number_prefix,panel_number_prefix) 
        
        compare_tests=Math_source.Compare_tests(test_databases)
        
        del test_databases

        test_results = Math_source.calculate_cp_cpk_pp_ppk_o_stand(compare_tests, self.get_selected_database(), Name)
        del compare_tests
        
        for element in test_results:
            try:
                self.add_to_treeview_test(
                    [element.fixture,
                    element.testname,
                    element.group,
                    element.cp,
                    element.cpk,
                    element.pp,
                    element.ppk,
                    element.o_stand,
                    element.limitchange])
            except:
                NameError()
                

    def sort_treeview_string(self, tree, col, reverse=False):
        data = [(tree.set(child, col), child) for child in tree.get_children('')]
        data.sort(reverse=reverse)
        for index, (val, child) in enumerate(data):
            tree.move(child, '', index)
        tree.heading(col, command=lambda: self.sort_treeview_string(tree, col, not reverse)) 
        
    def sort_treeview_value(self, tree, col, reverse=False):
        data = [(float(tree.set(child, col)), child) for child in tree.get_children('')]
        data.sort(reverse=reverse)
        for index, (val, child) in enumerate(data):
            tree.move(child, '', index)
        tree.heading(col, command=lambda: self.sort_treeview_value(tree, col, not reverse)) 
    
    def get_selected_server(self):
        return self.server_combobox.get()
    
    def get_selected_host(self):
        return self.machine_combobox.get()

    def get_selected_database(self):
        return self.database_combobox.get()

    def get_selected_process(self):
        return self.process_combobox.get()

    def get_selected_start_date(self):
        start_date_str = self.start_date_entry.get()

        
        try:
            start_date = datetime.strptime(start_date_str, "%d/%m/%y")
            return start_date
        except ValueError:
            print("Blad: Nieprawidlowy format daty")
            return None

    def get_selected_end_date(self):
        end_date_str = self.end_date_entry.get()

        
        try:
            end_date = datetime.strptime(end_date_str, "%d/%m/%y")
            return end_date
        except ValueError:
            print("Blad: Nieprawidlowy format daty")
            return None
    
    #Function whose take value of comboboxa
    def on_combobox_selected(self, event):
        
        selected_server = self.get_selected_server()
        print("Wybrany Serwer:", selected_server),

        selected_database = self.get_selected_database()
        print("Wybrana Baza Danych:", selected_database)

        selected_process = self.get_selected_process()
        print("Wybrany Proces Testowania:", selected_process)

        start_date = self.get_selected_start_date()
        print("Data poczatkowa:", start_date)

        end_date = self.get_selected_end_date()
        print("Data koncowa:", end_date)

        self.clear_to_treeview_manufacture()
        
        Tester_database=SQL_Command.tester_init(self.get_selected_server(),self.get_selected_database(), self.get_selected_process())
        
        values_to_add=[]
        for element in Tester_database:
            values_to_add.append(list(element.values())[1])
        self.machine_combobox['values'] = values_to_add

    def on_search_button_click(self):
        
        self.clear_to_treeview_manufacture()
        
        Tester_database=SQL_Command.tester_init(self.get_selected_server(),self.get_selected_database(), self.get_selected_process())   
        
        for element in Tester_database:
            if element['host'] == self.get_selected_host():
                print(f"{element['host']}                                           {self.get_selected_host()}")
                selected_host_id = element['id']
        
        product_database2 = SQL_Command.Products_init_new(self.get_selected_server(),self.get_selected_database(), self.get_selected_process(), selected_host_id,self.get_selected_start_date(), self.get_selected_end_date()) 
        for product in product_database2:
            Serialnumber=list(product.values())[1]
            self.add_to_treeview1_manufacture([list(product.values())[2],list(product.values())[4],Serialnumber]) #[:-2]]
            
    
    def calculate_button_click(self):
        
        print("Wejscie w calculate button click")
        self.clear_to_treeview_tests()
        
        selected_items = self.treeview1.selection()
        if selected_items:
            item_id = selected_items[0]
            
            Test_Proces_Values = self.treeview1.item(item_id, "values")

            
        

