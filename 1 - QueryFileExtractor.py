from xml.dom import minidom
from xlutils.copy import copy 
import xlwt
import xlrd
import os
import os.path




## Leer directorio de archivos ##
listadoArchivos = os.listdir('.')
listadoLimpio=[]


## Excluir Archivos no deseados ##
for each in listadoArchivos:
    if ".dtsx" in each and "Carga" in each:     
        listadoLimpio.append(each)
        
for archivo in listadoLimpio:
    xmldoc = minidom.parse(archivo)
    archivo = archivo[10:-5]
    try:
        os.mkdir(archivo)
    except:
        pass
    cont=0
    #print("P: "+archivo)
    
    
    properties = xmldoc.getElementsByTagName('property')
    for properti in properties:
        
        try:
            query = properti.firstChild.nodeValue
            if properti.hasAttribute("description"):
                queryElement = properti.attributes['description'].value
                if (queryElement == "The SQL command to be executed."):
                    #f = open(os.path.join(archivo, archivo+" Query "+str(cont) + ".txt"), "a")
                    query = query.replace('WITH (NOLOCK)','')
                    if "Dataareaid" not in query:
                        print("Archivo: "+archivo)
                    #f.write(query)
                    cont+=1
                    #f.close()
        except:
            a=0
    
        
