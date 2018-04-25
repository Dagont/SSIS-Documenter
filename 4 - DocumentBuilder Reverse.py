from xml.dom import minidom 
import xlwt
import xlrd
import os
import sys
import gc
sys.setrecursionlimit(8000)

conta=0

listadoLimpio=[]
listadoArchivos = os.listdir(os.getcwd())
## Excluir Archivos no deseados ##

for each in listadoArchivos:
    if "." not in each:# and "DimEstructura" in each:
        listadoLimpio.append(each)

#print(listadoLimpio)
    

def obtenerOrigenCarga(value,indexY,ady,adx):
    global hojaOut
    global hojaIn
    
    tipoCaja = hojaIn.cell(indexY,2).value
    tipoArchivo=hojaIn.cell(indexY,0).value
    if (tipoCaja == "Origen" and tipoArchivo == "Carga"):
        if(value==hojaIn.cell(indexY,5).value):
           sigPaso=hojaIn.cell(indexY,4).value
           hojaOut.write(ady,adx,sigPaso)
           adx+=1
           
           return adx, sigPaso
    
    if ((indexY+2)> hojaIn.nrows):
        return adx, value
    return obtenerOrigenCarga(value,indexY+1,ady,adx)

def obtenerOrigenExtraccion(indexY,ady,adx):
    global hojaOut
    global hojaIn
    
    tipoCaja = hojaIn.cell(indexY,2).value
    tipoArchivo=hojaIn.cell(indexY,0).value
    if (tipoCaja == "Origen" and tipoArchivo == "Extraccion"):
        ady+=1
        #if(value==hojaIn.cell(indexY,4).value):
        sigPaso=hojaIn.cell(indexY,5).value
        hojaOut.write(ady,adx,sigPaso)
        #adx+=1
        
        obtenerSiguiente(sigPaso,ady,adx+1)
        return obtenerOrigenExtraccion(indexY+1,ady,adx)
    if ((indexY+2)> hojaIn.nrows):
        return
    return obtenerOrigenExtraccion(indexY+1,ady,adx)

## Obtener Merge ##
def obtenerMerge(modo,value,indexY,ady,adx):
    global hojaOut
    global conta
    global hojaIn
    #conta+=1
    #print(conta)
    tipoArchivo=hojaIn.cell(indexY,0).value
    tipoCaja = hojaIn.cell(indexY,2).value
    if (tipoCaja == "Merge Join" and tipoArchivo==modo):
        if(value==hojaIn.cell(indexY,5).value and hojaIn.cell(indexY,4).value!=hojaIn.cell(indexY,5).value):
            sigPaso=hojaIn.cell(indexY,4).value
            hojaOut.write(ady,adx,sigPaso)
            adx+=1
            try:
                return obtenerMerge(modo,sigPaso,1,ady,adx)
            except:
                return adx, sigPaso
    
    if ((indexY+2)> hojaIn.nrows):
        return adx, value
    return obtenerMerge(modo,value,indexY+1,ady,adx)

## Obtener Sort ##
def obtenerSort(modo,value,indexY,ady,adx):
    global hojaOut
    global conta
    global hojaIn
    #conta+=1
    #print(conta)
    tipoArchivo=hojaIn.cell(indexY,0).value
    tipoCaja = hojaIn.cell(indexY,2).value
    if (tipoCaja == "Sort" and tipoArchivo==modo):
        if(value==hojaIn.cell(indexY,5).value and hojaIn.cell(indexY,4).value!=hojaIn.cell(indexY,5).value):
            sigPaso=hojaIn.cell(indexY,4).value
            hojaOut.write(ady,adx,sigPaso)
            adx+=1
            try:
                return obtenerSort(modo,sigPaso,1,ady,adx)
            except:
                return adx, sigPaso
    
    if ((indexY+2)> hojaIn.nrows):
        return adx, value
    return obtenerSort(modo,value,indexY+1,ady,adx)

## Obtener Pivot ##
def obtenerPivot(modo,value,indexY,ady,adx):
    global hojaOut
    global conta
    global hojaIn
    #conta+=1
    #print(conta)
    tipoArchivo=hojaIn.cell(indexY,0).value
    tipoCaja = hojaIn.cell(indexY,2).value
    if (tipoCaja == "Pivot Column" and tipoArchivo==modo):
        if(value==hojaIn.cell(indexY,5).value and hojaIn.cell(indexY,4).value!=hojaIn.cell(indexY,5).value):
            sigPaso=hojaIn.cell(indexY,4).value
            hojaOut.write(ady,adx,sigPaso)
            adx+=1
            try:
                return obtenerPivot(modo,sigPaso,1,ady,adx)
            except:
                return adx, sigPaso
    
    if ((indexY+2)> hojaIn.nrows):
        return adx, value
    return obtenerPivot(modo,value,indexY+1,ady,adx)


## Obtener Lookups ##
def obtenerLookups(modo,value,indexY,ady,adx):
    global hojaOut
    global hojaIn
    tipoCaja = hojaIn.cell(indexY,2).value
    tipoArchivo=hojaIn.cell(indexY,0).value

    if (tipoCaja == "LookUp Column" and tipoArchivo==modo):
        if(value==hojaIn.cell(indexY,5).value and hojaIn.cell(indexY,4).value!=hojaIn.cell(indexY,5).value):
           sigPaso=hojaIn.cell(indexY,4).value
           hojaOut.write(ady,adx,sigPaso)
           adx+=1
           try:
                return obtenerLookups(modo,sigPaso,1,ady,adx)
           except:
                return adx, sigPaso
           
    
    if ((indexY+2)> hojaIn.nrows):
        return adx, value
    return obtenerLookups(modo,value,indexY+1,ady,adx)

## Obtener Derived ##
def obtenerDerived(modo,value,indexY,ady,adx):
    global hojaOut
    global hojaIn
    
    tipoCaja = hojaIn.cell(indexY,2).value
    tipoArchivo=hojaIn.cell(indexY,0).value
    if (tipoCaja == "Derived Column" and tipoArchivo==modo):
        #print("Comparando : "+hojaIn.cell(indexY,4).value+" == "+value)
        if(value==hojaIn.cell(indexY,5).value and hojaIn.cell(indexY,4).value!=hojaIn.cell(indexY,5).value):
           
           sigPaso=hojaIn.cell(indexY,4).value
           hojaOut.write(ady,adx,sigPaso)
           adx+=1
           try:
                return obtenerDerived(modo,sigPaso,1,ady,adx)
           except:
                return adx, sigPaso
           
    
    if ((indexY+2)> hojaIn.nrows):
        return adx, value
    return obtenerDerived(modo,value,indexY+1,ady,adx)

## Obtener DataConversion ##
def obtenerDataConversion(modo,value,indexY,ady,adx):
    global hojaOut
    global hojaIn
    
    tipoCaja = hojaIn.cell(indexY,2).value
    tipoArchivo=hojaIn.cell(indexY,0).value
    if (tipoCaja == "Data Conversion" and tipoArchivo==modo):
        if(value==hojaIn.cell(indexY,5).value and hojaIn.cell(indexY,4).value!=hojaIn.cell(indexY,5).value):
           sigPaso=hojaIn.cell(indexY,4).value
           #print("DC ADX: "+str(adx))
           hojaOut.write(ady,adx,sigPaso)
           adx+=1
           
           try:
                return obtenerDataConversion(modo,sigPaso,1,ady,adx)
           except:
                return adx, sigPaso
           
    
    if ((indexY+2)> hojaIn.nrows):
        return adx, value
    return obtenerDataConversion(modo,value,indexY+1,ady,adx)
   
## Obtener Destino Extraccion (Intermedio)#
def obtenerDestinoExtraccion(value,indexY,ady,adx):
    global hojaOut
    global hojaIn
    
    tipoCaja = hojaIn.cell(indexY,2).value
    tipoArchivo=hojaIn.cell(indexY,0).value
    if (tipoCaja == "Destino" and tipoArchivo=="Extraccion"):
        if(value==hojaIn.cell(indexY,5).value):
           sigPaso=hojaIn.cell(indexY,4).value
           hojaOut.write(ady,adx,sigPaso)
           adx+=1
           try:
                return obtenerDestinoExtraccion(sigPaso,indexY+1,ady,adx)
           except:
                return
           
    
    if ((indexY+2)> hojaIn.nrows):
        return adx, value
    return obtenerDestinoExtraccion(value,indexY+1,ady,adx)

def obtenerDestinoCarga(value,indexY,ady,adx):
    global hojaOut
    global hojaIn
    global hojaOut2
    
    tipoCaja = hojaIn.cell(indexY,2).value
    tipoArchivo=hojaIn.cell(indexY,0).value
    if (tipoCaja == "Destino" and tipoArchivo=="Carga"):
        if(value==hojaIn.cell(indexY,4).value):
           sigPaso=hojaIn.cell(indexY,5).value
           hojaOut.write(ady,adx,sigPaso)
           hojaOut2.write(ady,2,sigPaso)
           adx+=1
           return True
    if ((indexY+2)> hojaIn.nrows):
        return False
    return obtenerDestinoCarga(value,indexY+1,ady,adx)


def obtenerSiguiente(value,prevPaso,ady,adx):
    global hojaOut2
    test=adx
    values=[]

    ## BackTracking ##

    obtenerTablasFinales(value,prevPaso,1,ady,adx-3)
    values.append(value)

    
    newAdx, value = obtenerDataConversion("Carga",value,1,ady,adx)
    values.append(value)

    newAdx, value = obtenerDerived("Carga",value,1,ady,newAdx)
    values.append(value)

    newAdx, value = obtenerLookups("Carga",value,1,ady,newAdx)
    values.append(value)

    newAdx, value = obtenerMerge("Carga",value,1,ady,newAdx)
    values.append(value)

    newAdx, value = obtenerPivot("Carga",value,1,ady,newAdx)
    values.append(value)

    newAdx, value = obtenerSort("Carga",value,1,ady,newAdx)
    values.append(value)

    newAdx, value = obtenerOrigenCarga(value,1,ady,newAdx)
    values.append(value)

    newAdx, value = obtenerDestinoExtraccion(value,1,ady,newAdx)
    values.append(value)

    newAdx, value = obtenerDataConversion("Extraccion",value,1,ady,newAdx)
    values.append(value)

    newAdx, value = obtenerDerived("Extraccion",value,1,ady,newAdx)
    values.append(value)

    newAdx, value = obtenerLookups("Extraccion",value,1,ady,newAdx)
    values.append(value)

    newAdx, value = obtenerMerge("Extraccion",value,1,ady,newAdx)
    values.append(value)

    newAdx, value = obtenerPivot("Extraccion",value,1,ady,newAdx)
    values.append(value)
    
    newAdx, value = obtenerSort("Extraccion",value,1,ady,newAdx)
    values.append(value)

    values.reverse()

    enganchado=False
    for valueX in values:
        try:
            enganchado,newAdx, value = obtenerValoresIniciales(valueX,1,ady,newAdx)
            #obtenerTablasIniciales(valueX,1,ady,newAdx)
            #obtenerConexionesIniciales(valueX,1,ady,newAdx+1)
            if (enganchado):
                break
        except:
            pass
        
    if not (enganchado):
        #print("No enganche: "+value)
        hojaOut2.write(ady,2,values[0]+" ?")
    enganchado=False
    
    for valueX in values:
        enganchado = obtenerTablasIniciales(valueX,1,ady,newAdx)
        if (enganchado):
            #print("EOTI")
            break
        
    enganchado=False
    for valueX in values:
        enganchado = obtenerConexionesIniciales(valueX,1,ady,newAdx+1)
        if (enganchado):
            #print("EOCI")
            break
        
    return

def obtenerConexionesIniciales(value, indexY,ady,adx):
    #print("Hello")
    global hojaOut
    global hojaOut2
    global hojaIn
    #print(value)
    try:
        tipoCaja = hojaIn.cell(indexY,2).value
        if (tipoCaja == "Conexion Origen" or tipoCaja == "Excel Source Connection"):
            #print("Enganchando: "+value+" -> "+hojaIn.cell(indexY,5).value)
            if(value==hojaIn.cell(indexY,5).value):
                sigPaso=hojaIn.cell(indexY,4).value
                #hojaOut.write(ady,adx,sigPaso)
                hojaOut2.write(ady,0,sigPaso)
                return
        
        if ((indexY+1)> hojaIn.nrows):
            return
        return obtenerConexionesIniciales(value,indexY+1,ady,adx)
    except Exception as e:
        #print("ROTO: "+str(e))
        return

def obtenerTablasIniciales(value, indexY,ady,adx):
    global hojaOut
    global hojaOut2
    global hojaIn
    #print(value)
    try:
        tipoCaja = hojaIn.cell(indexY,2).value
        if (tipoCaja == "Query Source Table" or tipoCaja == "Excel Source Table"):
            #print("Enganchando: "+value+" -> "+hojaIn.cell(indexY,5).value)
            if(value==hojaIn.cell(indexY,5).value or hojaIn.cell(indexY,5).value=="*"):
                sigPaso=hojaIn.cell(indexY,4).value
                #hojaOut.write(ady,adx,sigPaso)
                hojaOut2.write(ady,1,sigPaso)
                #adx+=1
                #print("Escribo: "+sigPaso)
                return True
        
        if ((indexY+1)> hojaIn.nrows):
            return False
        return obtenerTablasIniciales(value,indexY+1,ady,adx)
    except Exception as e:
        #print("ROTO: "+str(e))
        return False

def obtenerTablasFinales(value,sigPaso,indexY,ady,adx):
    global hojaOut
    global hojaOut2
    global hojaIn
    #print(lastValue)
    #print(value)
    #print(hojaIn.cell(indexY,4).value)
    try:
        tipoArchivo=hojaIn.cell(indexY,0).value
        tipoCaja = hojaIn.cell(indexY,2).value
        if (tipoCaja == "Tabla Destino" and tipoArchivo=="Carga"):
            #print(value)
            #print(hojaIn.cell(indexY,4).value)
            if(value==hojaIn.cell(indexY,4).value or sigPaso==hojaIn.cell(indexY,4).value):
                origen=hojaIn.cell(indexY,5).value
                hojaOut.write(ady,adx,origen)
                hojaOut2.write(ady,4,origen)
                return 
        
        if ((indexY+1)> hojaIn.nrows):
            return
        return obtenerTablasFinales(value,sigPaso,indexY+1,ady,adx)
    except Exception as e:
        print("roto "+value)
        print(e)
        return


def obtenerValoresIniciales(value,indexY,ady,adx):
    global hojaOut
    global hojaOut2
    global hojaIn
    try:
        tipoCaja = hojaIn.cell(indexY,2).value
        if (tipoCaja == "Query Alias" or
            tipoCaja == "Excel"):
            #print("Enganchando: "+value+" -> "+hojaIn.cell(indexY,5).value)
            if(value==hojaIn.cell(indexY,5).value):
                #print("Exito :"+value)
                #print("EXCEL MATCH")
                sigPaso=hojaIn.cell(indexY,4).value
                hojaOut.write(ady,adx,sigPaso)
                hojaOut2.write(ady,2,sigPaso)
                adx+=1
                return True, adx, sigPaso
        
        if ((indexY+1)> hojaIn.nrows):
            return False, adx, value
        return obtenerValoresIniciales(value,indexY+1,ady,adx)
    except Exception as e:
        #print(e)
        return False, adx, value
        pass

## afy = Archivo Fuente eje Y3
## ady = Archivo Destino eje Y


def IT_obtenerOrigenExtraccion(indexY,ady,adx):
    global hojaOut
    global hojaIn
    
    for indexY in range(1,hojaIn.nrows):
        tipoCaja = hojaIn.cell(indexY,2).value
        tipoArchivo=hojaIn.cell(indexY,0).value
        
        if (tipoCaja == "Destino" and tipoArchivo == "Carga"):
            ady+=1
            #if(value==hojaIn.cell(indexY,4).value):
            
            prevPaso=hojaIn.cell(indexY,5).value
            value=hojaIn.cell(indexY,4).value
            hojaOut.write(ady,adx,prevPaso)
            hojaOut.write(ady,adx+1,value)
            hojaOut2.write(ady,3,prevPaso)
            #adx+=1
            #print(ady)
            obtenerSiguiente(value,prevPaso,ady,adx+2)

            
print("Iniciando Ejecución")
for archivo in listadoLimpio:
    for root, dirs, files in os.walk(archivo, topdown=False):
        for name in files:
            gc.collect()
            if ".xlsx" in name:
                print("F: "+name)
                libroIn=xlrd.open_workbook(os.path.join(root, name))
                libroOut = xlwt.Workbook()
                for sheet in range (0,libroIn.nsheets):
                    hojaIn = libroIn.sheet_by_index(sheet)
                    
                    #print("$$$ NROWS = "+str(hojaIn.nrows))
                    #print("HojaIn: "+libroIn.sheet_names()[sheet])
                    hojaOut = libroOut.add_sheet(libroIn.sheet_names()[sheet])
                    hojaOut2 = libroOut.add_sheet(libroIn.sheet_names()[sheet]+"_Clean")
                    ady=0
                    adx=1
                    afx=0
                    afy=0
                    hojaOut.write(0,0,"Nombres Alternativos")
                    hojaOut2.write(0,0,"Conexion")
                    hojaOut2.write(0,1,"Tabla Fuente")
                    hojaOut2.write(0,2,"Campo Fuente")
                    hojaOut2.write(0,3,"Campo Destino")
                    hojaOut2.write(0,4,"Tabla Destino")
                    hojaOut2.write(0,5,"Nombres Alternativos")
                    IT_obtenerOrigenExtraccion(1,ady,adx)
            
                libroOut.save(os.path.join(root, "(MFD) "+name[:-1]))




