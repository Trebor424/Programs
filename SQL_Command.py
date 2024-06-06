# -*- coding: utf-8 -*-
import pyodbc
import Project_Classes
import os

SQL_Host_Product_Search_init="""

"""

SQL_Host_Product_Search="""

"""
 
 
SQL_Product_Search=""" 

	"""
SQL_FPT_Testers_List="""

  """
SQL_ICT_Testers_List=""" 

  """
  
SQL_FFT_Testers_List="""

    """
SQL_FPT_ICT_FFT_Product_List="""

  """


SQL_serialnumber_masterid_search_FPT_board="""

"""

SQL_serialnumber_masterid_search_FPT_panels="""

"""

SQL_serialnumber_masterid_search_ICT_singleboard="""
  
"""
SQL_serialnumber_masterid_search_ICT_panels="""

"""


SQL_serialnumber_masterid_search_FFT="""

"""

### Main SQL FUNCTIONS
Designators = {
    'T' : 1000000000000,
    'G' : 1000000000,
    'M' : 1000000,
    'k' : 1000 ,
    'm' : 0.001 ,
    'u' : 0.000001 ,
    'n' : 0.000000001 ,
    'p' : 0.000000000001,
}
###


def database_init(selected_server :str):
    conn_str = (
        f'DRIVER={{SQL Server}};'
        f'SERVER={selected_server};'
        'Trusted_Connection=yes'
    )

    databases = []

    try:
        # Connect to the SQL Server
        with pyodbc.connect(conn_str) as connection:
            # Create a cursor from the connection
            cursor = connection.cursor()

            # Execute the query to retrieve the list of databases
            query = "SELECT name FROM sys.databases"
            cursor.execute(query)

            # Fetch the results
            databases = [row[0] for row in cursor.fetchall()]

            # Filter and add databases with "-" in their name
            selected_server.filtered_databases = [db for db in databases if "-" in db]
            print(selected_server.filtered_databases)

    except Exception as e:
        print(f"Error connecting to the database: {e}")

    return databases

def tester_init(selected_server :str , selected_database :str, selected_process :str) -> list[str]:
    conn_str = (
        f'DRIVER={{SQL Server}};'
        f'DATABASE={selected_database};'
        f'SERVER={selected_server};'
        'Trusted_Connection=yes'
        )
    
    Tester_database=[]
    
    try:
    # Connect to the SQL Server
        connection = pyodbc.connect(conn_str)
                
        # Create a cursor from the connection
        cursor = connection.cursor()
             
        if selected_process == "ICT" :
            query = SQL_ICT_Testers_List
            print("ICT")
            
        elif selected_process == "FPT" :
            query = SQL_FPT_Testers_List
            print("FPT")
            
        elif selected_process == "FFT" :
            query = SQL_FFT_Testers_List
            print("FFT")
        
        # Execute the query to retrieve the list of databases
        
        cursor.execute(query)
             
        # Fetch the results
        for row in cursor.fetchall():
            product_dict = {
                'id': row[0],
                'host': row[1],
                }
            Tester_database.append(product_dict)
            #print(f"Tester database - {product_dict}")
        
                    
    except Exception as e:
        print(f"Error connecting to the database: {e}")
    finally:
        # Close the connection in the 'finally' block to ensure it's always closed
        if connection:
            connection.close()
    return Tester_database
            
def Products_init_new(selected_server :str, selected_database:str, selected_process:str, selected_host:str, start_time:str, end_time:str) -> list[str]:  

    conn_str = (
        f'DRIVER={{SQL Server}};'
        f'DATABASE={selected_database};'
        f'SERVER={selected_server};'
        'Trusted_Connection=yes'
        )
    
    Product_database=[]
    Product_database_verify_listy=[]
    
    try:
        query=SQL_Product_Search
   
    # Connect to the SQL Server
        connection = pyodbc.connect(conn_str)
        cursor = connection.cursor()
        cursor.execute(query)
        print(query)
        
        for row in cursor.fetchall():
            Product_database_dict = {
                'skidprefix': row[0],
                'serialnumberprefix': row[1],
                'lasermarking': row[2],
                'productpartnumber': row[3],
                'panelnumberprefix': row[4],
                'numberofboards': row[5],
                'valid': row[6],
                'productpartnumber': row[7],
            }
            Product_database.append(Product_database_dict)

            
        query2=f'''
        select distinct scannednumber 
        from vREG_OF_PROCESS
        '''
        query2+=f" where timestamp between '{start_time}' AND '{end_time}' AND hostid like {selected_host}"
        
        cursor2 = connection.cursor()
        cursor2.execute(query2)    
        print(query2)
        
        for row in cursor2.fetchall():
            Product_database_verify_listy.append(row[0])
        #print(Product_database_verify_listy)
        
        for element in Product_database:
            print(f"{element['lasermarking']};{element['serialnumberprefix']};{element['panelnumberprefix']}")

        verified_products = []
        for element in Product_database:
            for element_product_database in Product_database_verify_listy:
                if element['panelnumberprefix'] in element_product_database:
                    if element not in verified_products:
                        verified_products.append(element)
                elif element['serialnumberprefix'] in element_product_database:
                    if element not in verified_products:
                        verified_products.append(element)
                
                #print(f"{element['panelnumberprefix']}               {element['serialnumberprefix']}")
                
                
                
        with open(os.path.join(os.path.dirname(__file__), "Nazwy_wyrobow", f"{selected_database}.txt"), 'r') as file:
            zawartosc = file.readlines() 
            
         
        for line in zawartosc:
            try:
                name, serial_number, panel_number = line.strip().split(';')
                # Check if the line contains at least three tab-separated values
                if len(name) > 0 and len(serial_number) > 0 and len(panel_number) > 0:
                    for element in verified_products:
                        if (len(element['lasermarking'])==0 and (panel_number in element['panelnumberprefix']) or (serial_number in element['serialnumberprefix'])):
                            element['lasermarking'] = name

                else:
                    print("Incomplete line:", line)
            except ValueError:
                print("Invalid line format:", line)
                  
    except Exception as e:
        print(f"Error connecting to the database: {e}")
    finally:
        if connection:
            connection.close()
    return verified_products


def test_databases(selected_server:str , selected_database:str, selected_process:str, selected_host:str, start_time:str, end_time:str,serial_number_prefix:str, panel_number_prefix:str) -> list[str]:
    conn_str = (
        f'DRIVER={{SQL Server}};'
        f'DATABASE={selected_database};'
        f'SERVER={selected_server};'
        'Trusted_Connection=yes'
        )
    
    Pan_Ser_numberdict=[]
    
    #print(f"{selected_server} , {selected_database}, {selected_process}, {selected_host}, {start_time}, {end_time}")
    

    # Connect to the SQL Server
    connection = pyodbc.connect(conn_str)
    cursor = connection.cursor()
    
    if selected_process == "ICT" :
        
        query = SQL_serialnumber_masterid_search_ICT_singleboard
        query+=f" where REG_OF_TEST_ICT.timestamp between '{start_time}' AND '{end_time}'  AND serialnumber like '{serial_number_prefix}%' "
        query+=f"AND maxvalue not like 'NULL' AND minvalue not like 'NULL' AND measuredvalue not like 'NULL'"
        print(query)
        cursor.execute(query)

        Zmienna=cursor.fetchall()
        
        if not len(Zmienna) == 0:
        
            for row in Zmienna:
                
                Wyrob = Project_Classes.Wyrob(
                        row[0],  # masterid
                        row[1],  # panel/serialnumber
                        row[2],  # boardsnumber
                        row[3],  # parts
                        "",  # aux
                        "",  # value
                        "",  # loc
                        "",  # el
                        row[5],  # reference
                        row[6],  # +%
                        row[7],  # -%
                        row[9],        # timestamp
                        row[8], #
                        "ICT"
                    )
                
                #print(f"{Wyrob.parts} + {Wyrob.value}")
                #print(f"-{Wyrob.tolerance_minus} = {Wyrob.value} = +{Wyrob.tolerance_plus}")
                Pan_Ser_numberdict.append(Wyrob)
            
        else:
                
            query = SQL_serialnumber_masterid_search_ICT_panels
            query+=f" where REG_OF_TEST_ICT.timestamp between '{start_time}' AND '{end_time}'  AND panelorserialnumber like '{panel_number_prefix}%' "
            query+=f"AND maxvalue not like 'NULL' AND minvalue not like 'NULL' AND measuredvalue not like 'NULL'"
            print(query)
            cursor.execute(query)
            
            for row in cursor.fetchall():

                Wyrob = Project_Classes.Wyrob(
                    row[0],  # masterid
                    row[1],  # panelorseialnumber
                    row[2],  # boardsnumber
                    row[3],  # parts
                    "",  # aux
                    "",  # value
                    "",  # loc
                    "",  # el
                    row[5],  # reference
                    row[6],  # +%
                    row[7],  # -%
                    row[9],        # timestamp
                    row[8], #
                    "ICT"
                )
                
                #print(f"{Wyrob.parts} + {Wyrob.value}")
                #print(f"-{Wyrob.tolerance_minus} = {Wyrob.value} = +{Wyrob.tolerance_plus}")
                Pan_Ser_numberdict.append(Wyrob)
            
    
    elif selected_process == "FPT" :
        query = SQL_serialnumber_masterid_search_FPT_board
        query+=f" where timestamp between '{start_time}' AND '{end_time}'  AND serialnumber like '{serial_number_prefix}%'  AND judge not like 'SKIP'"
        
        cursor.execute(query)
        Zmienna=cursor.fetchall()
        #print(cursor.fetchall())
        
        if not len(Zmienna) == 0:
            
            query = SQL_serialnumber_masterid_search_FPT_board
            query+=f""" where timestamp between '{start_time}' AND '{end_time}'  AND serialnumber like '{serial_number_prefix}%'  
            AND judge NOT LIKE 'SKIP' 
            AND loc in ('KELV','RES','CAP','IND')
            AND [-%] not like '0' 
            AND [+%] not like '0'
            AND judge = 'PASS'
            """
            cursor.execute(query)
            print(query)
            
            for row in cursor.fetchall():
                # print(row)
                # i=0
                # for element in row:
                #     print(f"{i}  ==== {row[i]}")
                #     i+=1
                
                Valuereference=row[10]
                for element in Designators:
                    if element in row[11]:  # Corrected if condition
                        Valuereference *= Designators[element]
                
                Valuevariable = 0
                if row[18] =='' :
                    Valuevariable = row[15]   # testValue
                    #print(f"INIT Valuevariable ---------> {Valuevariable}")
                    for element in Designators:
                        if element in row[16]:  # Corrected if condition
                            Valuevariable *= Designators[element]
                else:
                    Valuevariable = row[17]   # testValue
                    for element in Designators:
                        if element in row[18]:  # Corrected if condition
                            Valuevariable *= Designators[element]
                            #print(f"------------> Valuevariable {Valuevariable}  Element {element} Row {row[15]}")
                
                tolerance_plus=Valuereference+(row[13]/50)*Valuereference
                tolerance_minus=Valuereference-(row[14]/50)*Valuereference
                
                Wyrob = Project_Classes.Wyrob(
                    row[0],  # masterid
                    row[21],  # panel/serialnumber
                    row[1],  # boardsnumber
                    row[3],  # parts
                    row[4],  # aux
                    row[5],  # value
                    row[8],  # loc
                    row[9],  # el
                    row[10],  # reference
                    tolerance_plus,  # +%
                    tolerance_minus,  # -%
                    row[24],       # timestamp
                    Valuevariable,
                    "FPT_BOARD"
                )

                #print(f"{Wyrob.parts} + {Valuevariable}")
                #print(f"-{Wyrob.tolerance_minus} = {Valuereference} = +{Wyrob.tolerance_plus}")
                Pan_Ser_numberdict.append(Wyrob)

        else:
            query = SQL_serialnumber_masterid_search_FPT_panels
            query += f""" WHERE timestamp BETWEEN '{start_time}' AND '{end_time}' AND panelnumber LIKE '{panel_number_prefix}%' 
            AND judge NOT LIKE 'SKIP' 
            AND loc in ('KELV','RES','CAP','IND')
            AND [-%] not like '0' 
            AND [+%] not like '0'
            AND judge = 'PASS'
            """
            cursor.execute(query)
            #print(panel_number_prefix)
            print(query)     
                    
            for row in cursor.fetchall():
                
                Valuereference=row[8]
                for element in Designators:
                    if element in row[9]:  
                        Valuereference *= Designators[element]
                
                Valuevariable = 0
                if row[16] == '':
                    Valuevariable = row[13]  
                    for element in Designators:
                        if element in row[14]: 
                            Valuevariable *= Designators[element]
                else:
                    Valuevariable = row[15]   # testValue
                    for element in Designators:
                        if element in row[16]:  
                            Valuevariable *= Designators[element]
                
                tolerance_plus=Valuereference+(row[11]/50)*Valuereference
                tolerance_minus=Valuereference-(row[12]/50)*Valuereference
                
                Wyrob = Project_Classes.Wyrob(
                    row[0],  # masterid
                    row[18],  # panel/serialnumber
                    row[1],  # boardsnumber
                    row[3],  # parts
                    row[4],  # aux
                    row[5],  # value
                    row[6],  # loc
                    row[7],  # el
                    row[8],  # reference
                    tolerance_plus,  # +%
                    tolerance_minus,  # -%
                    row[22],         # timestamp
                    Valuevariable,
                    "FPT_PANEL"
                )

                #print(f"{Wyrob.parts} + {Valuevariable}")
                #print(f"-{Wyrob.tolerance_minus} = {Valuereference} = +{Wyrob.tolerance_plus}")
                Pan_Ser_numberdict.append(Wyrob)
            
        
    elif selected_process == "FFT" :
        query = SQL_serialnumber_masterid_search_FFT
        query+=f" where timestamp between '{start_time}' AND '{end_time}'  AND serialnumber like '{serial_number_prefix}%' AND teststep not like '%ontaz%' "
        #monta¿, do wyjebania
        
        print(query)
        cursor.execute(query)
        
        for row in cursor.fetchall():
            #print(row)

            i=0
            for element in row:
                #print(f"{i}  ==== {row[i]}")
                i+=1
            
            Wyrob = Project_Classes.Wyrob(
                    row[0],  # masterid
                    row[1],  # panel/serialnumber
                    1,  # boardsnumber
                    row[2],  # parts
                    1,  # aux
                    1,  # value
                    1,  # loc
                    1,  # el
                    1,  # reference
                    row[4],  # +%
                    row[3],  # -%
                    row[6],         # timestamp
                    row[5],
                    "FFT"
                )
            
            #print(f"{Wyrob.parts} + {Wyrob.value}")
            #print(f"-{Wyrob.tolerance_minus} = {Wyrob.value} = +{Wyrob.tolerance_plus}")
            Pan_Ser_numberdict.append(Wyrob)

    return Pan_Ser_numberdict
