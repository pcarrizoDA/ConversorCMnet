import pandas as pd
df=pd.read_excel('C:/Users/pcarrizo/Documents/CM/valor_pagados.xlsx')
#Elimino columnas en blanco
df = df.drop(['Unnamed: 1', 'Unnamed: 2','Unnamed: 6','Unnamed: 13'],axis=1)
# Eliminar filas que contienen palabras específicas
keywords = ["Proveedor","Total de Pagos","Resumen por Forma de Pago","Subtotal", "MANANTIAL DEL SILENCIO", "Página", "Valore Pagados", "Filtros Seleccionados", "Fecha Inicial"]
for keyword in keywords:
     df = df[~df.apply(lambda row: row.astype(str).str.contains(keyword).any(), axis=1)]   
#Completo con valor 0 datos en blanco de la primera columna
df['Unnamed: 0'] = df['Unnamed: 0'].fillna(0)
#Filtro en el data frame solo las columnas que tengan un valor distinto a 0
df = df[df['Unnamed: 0'] != 0]
band = 0
#recorro el data frame para asignar la fecha correcta
for i in range(len(df)):
    if band == 0 and df.iloc[i]['Unnamed: 0']=='Fecha de Asiento':
        fechaaux=df.iloc[i]['Unnamed: 4']
        band = 1
    if df.iloc[i]['Unnamed: 0'] !='Fecha de Asiento':
        df.iloc[i]['Unnamed: 4'] = fechaaux
    else:
        fechaaux=df.iloc[i]['Unnamed: 4']
# Eliminar filas que contienen palabras específicas
palabras = ["Fecha de Asiento"]
for keyword in palabras:
     df = df[~df.apply(lambda row: row.astype(str).str.contains(keyword).any(), axis=1)] 
#Cambiar nombres de columnas
dfinal = df.rename(columns={'Unnamed: 0':'Proveedor',
                            'Filtros Seleccionados':'Documento',
                            'Unnamed: 4':'Número',
                            'Unnamed: 5':'Tipo Doc',
                            'Unnamed: 7':'Sld. Bruto',
                            'Unnamed: 8':'Modificador',
                            'Unnamed: 9':'Otra Moneda',
                            'Unnamed: 10':'Valor Neto',
                            'Unnamed: 11':'En Abierto',
                            'Unnamed: 12':'Cuentas/C x Forma Pag',
                            'Unnamed: 14':'Historico'})
#convierto a string la columna documento para que me tome correctamente el valor en excel 
dfinal['Documento'] = dfinal['Documento'].astype(str)
dfinal.to_excel('C:/Users/pcarrizo/Documents/CM/valor_pagados_procesado.xlsx', index=False)