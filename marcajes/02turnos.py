import openpyxl #proceso dos
from datetime import datetime,time

archivo='./marcajes septiembre/9. SEPTIEMBRE BEL_noDuplicados_1.xlsx'
libro = openpyxl.load_workbook(archivo)
hoja = libro['Hoja1']

limiteColumnas='J'
colIdEmpleado='B'
colChecada='E'
colHora='G'
colTurno='H'
limitefilas=15618+1

colTest='L'

entrada={
    'primMax':datetime.strptime('07:00:00', '%H:%M:%S').time(),
    'primMin':datetime.strptime('05:00:00', '%H:%M:%S').time(),
    'segMax':datetime.strptime('15:00:00', '%H:%M:%S').time(),
    'segMin':datetime.strptime('13:00:00', '%H:%M:%S').time(),
    'terMax':datetime.strptime('23:00:00', '%H:%M:%S').time(),
    'terMin':datetime.strptime('21:00:00', '%H:%M:%S').time(),

}


fechaChecadaEmpl=[]
emplFecha=[]
fechaHora=[]
turnoOk=[]
listaEmpleados=[]
#fechaBase=[]

for fila in range(1,limitefilas): #limitefilas
    coordEmpl=f'{colIdEmpleado}{fila}'
    empleado=hoja[coordEmpl].value
    coordChecada=f'{colChecada}{fila}'
    fechaChecada=hoja[coordChecada].value
    coordHora=f'{colHora}{fila}'
    horaChecada=hoja[coordHora].value
    empl=[empleado,coordEmpl,fechaChecada,fila,horaChecada]
    fechaChecadaEmpl.append(empl)
    if empleado is not None:
        listaEmpleados.append(empleado)


empleados=list(set(listaEmpleados))
empleados.sort()

fechaBase=fechaChecadaEmpl[3][0]
print(fechaBase)
for idEmpl in empleados:
    for linea in fechaChecadaEmpl:
        if linea[2] is not None and idEmpl==linea[0] and linea[4]is not None:
            for linea1 in fechaChecadaEmpl:
                if linea1[2] is not None and idEmpl==linea1[0] and linea[2]==linea1[2]:
                                  
                    if fechaBase!=linea[2]:
                        fechaBase=linea[2]
                        #print(linea[2],linea[3],linea1[2],linea1[3]) 
                        emplFecha.sort()

                        #print(emplFecha)

                        #print(f' la posicion: {emplFecha[0]}, corresponde a la hora {fechaChecadaEmpl[emplFecha[0]-1][4]}_ la salida esta en {emplFecha[-1]}_____')
                        posEntra=emplFecha[0]
                        posSale=emplFecha[-1]
                        horaEntra=fechaChecadaEmpl[emplFecha[0]-1][4]

                        if horaEntra!='HORA':
                                if linea1[4]is not None and entrada['primMin']<= horaEntra <=entrada['primMax']:
                                    #print(linea[4])
                                    primerTurno=[posEntra,posSale,1]
                                    #print(primerTurno)
                                    turnoOk.append(primerTurno)

                                elif linea1[4]is not None and entrada['segMin']<= horaEntra <=entrada['segMax']:
                                    #print('segundo turno')
                                    segundoTurno=[posEntra,posSale,2]
                                    turnoOk.append(segundoTurno)  

                                elif linea1[4]is not None and entrada['terMin']<= horaEntra <=entrada['terMax']:
                                    tercerTurno=[posEntra,posSale,3]
                                    #print('tercero turno')
                                    turnoOk.append(tercerTurno)




                        emplFecha.clear()
                        emplFecha.append(linea1[3]) 

                    elif fechaBase==linea1[2]:
                        emplFecha.append(linea1[3])  
                    """
                        if len(emplFecha)>0:
                            #print(fechaBase, 'la fecha se repite en:')
                            emplFecha.sort()
                            #print(emplFecha)
                            #print(f'la entrada esta en: {emplFecha[0]} y es: {linea1[4]}')
                            #print(f'la salida esta en: {emplFecha[-1]} y es: {linea1[4]}')
                            #print (linea1[4])
                            #print (type(linea1[4]))
                            if linea1[4]!='HORA':
                                if linea1[4]is not None and entrada['primMin']<= linea1[4] <=entrada['primMax']:
                                    print(linea[4])
                                    primerTurno=[emplFecha[0],emplFecha[-1],1]
                                    print(primerTurno)
                                    turnoOk.append(primerTurno)
                                elif linea1[4]is not None and entrada['segMin']<= linea1[4] <=entrada['segMax']:
                                    #print('segundo turno')
                                    segundoTurno=[emplFecha[0],emplFecha[-1],2]
                                    turnoOk.append(segundoTurno)                               
                                elif linea1[4]is not None and entrada['terMin']<= linea1[4] <=entrada['terMax']:
                                    tercerTurno=[emplFecha[0],emplFecha[-1],3]
                                    #print('tercero turno')
                                    turnoOk.append(tercerTurno)                               
                        emplFecha.clear()
                        emplFecha.append(linea1[3])
                        
                    elif fechaBase==linea1[2]:
                        emplFecha.append(linea1[3])

                        ##print(f'{idEmpl}, {linea1[2]}, {linea[3]}, {linea1[3]}')"""
"""
hoja['H5']=55
libro.save('./marcajes noviembre/11.NOVIEMBRE_TurnosOk.xlsx')

"""
for lineaT in turnoOk:
    #print(lineaT)
    coordEntra=f'{colTurno}{lineaT[0]}'
    coordSale=f'{colTurno}{lineaT[1]}'
    #print(f'{coordEntra}, turno:{lineaT[2]}')
    hoja[coordEntra]=lineaT[2]
    #print(f'{coordSale}, turno:{lineaT[2]}')
    hoja[coordSale]=lineaT[2]
    if coordEntra=='H4275' or coordSale=='H4275':
        print(lineaT)
libro.save('./marcajes septiembre/9. SEPTIEMBRE BEL_asignaTurno_1.xlsx')

#H4273"
