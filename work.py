import pandas as pd

# Укажите путь к вашему Excel-файлу
file_path = 'laptop.xlsx'

# Прочитайте нужный лист (можно указать имя или индекс, например 0)
df = pd.read_excel(file_path, sheet_name='Sheet1')

# Извлеките столбец по имени (например, 'price') и преобразуйте в список
column_array_price = df['Price'].dropna().tolist()  # dropna() убирает пустые ячейки (опционально)

column_array_model= df['Model'].dropna().tolist()

column_array_SSD= df['SSD'].dropna().tolist()

column_array_Ram= df['Ram'].dropna().tolist()

column_array_Display= df['Display'].dropna().tolist()

column_array_Display1= df['Display'].dropna().tolist()

column_array_Display2= df['Display'].dropna().tolist()

column_array_Graphics= df['Graphics'].dropna().tolist()


# ------------------------------------------
i=0
for box in column_array_Graphics:
    if (('RTX' in box) or ('RX' in box) or ('GTX' in box)):
       column_array_Graphics[i]= 1
    else:
        column_array_Graphics[i] = 0
    i+=1


# If lengths don't match, handle accordingly
if len(column_array_Graphics) != len(df):
    # Example: Pad with None or empty strings
    if len(column_array_Graphics) < len(df):
        column_array_Graphics.extend([''] * (len(df) - len(column_array_Graphics)))
    else:
        # Alternatively, truncate the list
        column_array_Graphics = column_array_Graphics[:len(df)]

df['Gaming_status'] = column_array_Graphics
output_path = '/home/vadim/my_venv/new_laptop.xlsx'
df.to_excel(output_path, sheet_name='Sheet1', index=False)

# ------------------------------------------


# ------------------------------------------
i=0
for box in column_array_Display2:
    if ',' in box:
        o=box.split(',')
        o[1] = ''.join(c for c in o[1] if c in '0123456789.')
        box1=o[1]
        new_string=''
        flag=0
        for char in box1:
            if (flag>=4):
                new_string += char
            flag+=1
        new_string = ''.join(c for c in new_string if c in '0123456789.')
    
    column_array_Display2[i]= new_string
    i+=1



# If lengths don't match, handle accordingly
if len(column_array_Display2) != len(df):
    # Example: Pad with None or empty strings
    if len(column_array_Display2) < len(df):
        column_array_Display2.extend([''] * (len(df) - len(column_array_Display2)))
    else:
        # Alternatively, truncate the list
        column_array_Display2 = column_array_Display2[:len(df)]

df['pixels_width'] = column_array_Display2
output_path = '/home/vadim/my_venv/new_laptop.xlsx'
df.to_excel(output_path, sheet_name='Sheet1', index=False)


# ------------------------------------------
i=0
for box in column_array_Display1:
    if ',' in box:
        o=box.split(',')
        o[1] = ''.join(c for c in o[1] if c in '0123456789.')
        box1=o[1]
        new_string=''
        flag=0
        for char in box1:
            if (flag==4):
                break
            new_string += char
            flag+=1
        new_string = ''.join(c for c in new_string if c in '0123456789.')
    
    column_array_Display1[i]= new_string
    i+=1



# If lengths don't match, handle accordingly
if len(column_array_Display1) != len(df):
    # Example: Pad with None or empty strings
    if len(column_array_Display1) < len(df):
        column_array_Display1.extend([''] * (len(df) - len(column_array_Display1)))
    else:
        # Alternatively, truncate the list
        column_array_Display1 = column_array_Display1[:len(df)]

df['pixels_len'] = column_array_Display1
output_path = '/home/vadim/my_venv/new_laptop.xlsx'
df.to_excel(output_path, sheet_name='Sheet1', index=False)

# ------------------------------------------

# ------------------------------------------
i=0
for box in column_array_Display:
    if ',' in box:
        o=box.split(',')
        box1=o[0]
        new_string=''
        flag=0
        for char in box1:
            if ((char == ' ') and (flag==1)):
                break
            new_string += char
            flag=1
    else:
        new_string=''
    
    new_string = ''.join(c for c in new_string if c in '0123456789.')
    
    column_array_Display[i]= new_string
    i+=1



# If lengths don't match, handle accordingly
if len(column_array_Display) != len(df):
    # Example: Pad with None or empty strings
    if len(column_array_Display) < len(df):
        column_array_Display.extend([''] * (len(df) - len(column_array_Display)))
    else:
        # Alternatively, truncate the list
        column_array_Display = column_array_Display[:len(df)]

df['Inches'] = column_array_Display
output_path = '/home/vadim/my_venv/new_laptop.xlsx'
df.to_excel(output_path, sheet_name='Sheet1', index=False)

# ------------------------------------------

# ------------------------------------------
i=0
for box in column_array_Ram:
    new_string=''
    for char in box:
        if char == ' ':
            break
        new_string += char
 
    new_string = ''.join(c for c in new_string if c in '0123456789')
    column_array_Ram[i]= new_string
    i+=1

df['Valume_Ram'] = column_array_Ram
output_path = '/home/vadim/my_venv/new_laptop.xlsx'
df.to_excel(output_path, sheet_name='Sheet1', index=False)

# ------------------------------------------
# ------------------------------------------
i=0
for box in column_array_SSD:
    new_string=''
    for char in box:
        if char == ' ':
            break
        new_string += char
 
    new_string = ''.join(c for c in new_string if c in '0123456789')
    column_array_SSD[i]= new_string
    i+=1

# If lengths don't match, handle accordingly
if len(column_array_SSD) != len(df):
    # Example: Pad with None or empty strings
    if len(column_array_SSD) < len(df):
        column_array_SSD.extend([''] * (len(df) - len(column_array_SSD)))
    else:
        # Alternatively, truncate the list
        column_array_SSD = column_array_SSD[:len(df)]


df['Volume_SSD'] = column_array_SSD
output_path = '/home/vadim/my_venv/new_laptop.xlsx'
df.to_excel(output_path, sheet_name='Sheet1', index=False)

# ------------------------------------------
i=0
for box in column_array_model:
    new_string=''
    for char in box:
        if char == ' ':
            break
        new_string += char
    column_array_model[i]= new_string
    i+=1

df['Manufacturer'] = column_array_model
output_path = '/home/vadim/my_venv/new_laptop.xlsx'
df.to_excel(output_path, sheet_name='Sheet1', index=False)

# ------------------------------------------

i=0
for box in column_array_price:
    column_array_price[i] = box.replace(',','').replace('₹','')
    i+=1
    

df['Price'] = column_array_price
output_path = '/home/vadim/my_venv/new_laptop.xlsx'
df.to_excel(output_path, sheet_name='Sheet1', index=False)

# ------------------------------------------