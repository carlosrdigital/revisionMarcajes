import openpyxl
from datetime import datetime,time

archivo='./marcajes noviembre/11. NOVIEMBRE.xlsx'
libro = openpyxl.load_workbook(archivo)
hoja = libro['noviembre']

limiteColumnas='J'
colIdEmpleado='B'
colChecada='F'
colHora='G'
colTurno='H'
salidas={
    'tercerMax':datetime.strptime('09:00:00', '%H:%M:%S').time(),
    'tercerMin':datetime.strptime('06:00:00', '%H:%M:%S').time(),
    'mixtoOpMax':datetime.strptime('22:00:00', '%H:%M:%S').time(),
    'mixtoOpMin':datetime.strptime('17:00:00', '%H:%M:%S').time()
}

entradas={
    'tercerMax':datetime.strptime('23:00:00', '%H:%M:%S').time(),
    'tercerMin':datetime.strptime('17:00:00', '%H:%M:%S').time(),
    'mixtoOpMax':datetime.strptime('11:00:00', '%H:%M:%S').time(),
    'mixtoOpMin':datetime.strptime('05:00:00', '%H:%M:%S').time()
}



limitefilas=15030+1
#test=hoja['g5'].value
#print(test)
#print(type(test))

for fila in range(6220,6230): #limitefilas
    coordTurno=f'{colTurno}{fila}'
    turno=hoja[coordTurno].value

    #print(f'{coordHora}, {hora}, checada')

    if turno==3:
        coordHora=f'{colHora}{fila}'
        hora=hoja[coordHora].value
        if hora is not None and salidas['tercerMin']<= hora <=salidas['tercerMax']:
            nuevaSalida=time(5,hora.minute,hora.second)
            print(f'{coordHora}, {hora}, salida tercer truno {nuevaSalida}')
            #print(type(nuevaSalida))
            #hoja[coordHora]=nuevaSalida ##cambia la hora en excel
    if turno=='M':
        coordHora=f'{colHora}{fila}'
        hora=hoja[coordHora].value
        if hora is not None and salidas['mixtoOpMin']<= hora <=salidas['mixtoOpMax']:
            nuevaSalida=time(16,hora.minute,hora.second)
            print(f'{coordTurno},{coordHora},{hora},salida mixto  opertativo {nuevaSalida}')







#libro.save(archivo) ##guarda los cambios
        

            
        

    

