import openpyxl #proceso dos
from datetime import datetime,time

archivo='./marcajes noviembre/11.NOVIEMBRE_sinDuplicados1.xlsx'
libro = openpyxl.load_workbook(archivo)
hoja = libro['noviembre']

limiteColumnas='J'
colIdEmpleado='B'
colChecada='E'
colHora='G'
colTurno='H'
limitefilas=14826+1

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
    coordTurno=f'{colTurno}{fila}'
    turno=hoja[coordTurno].value

    empl=[empleado,coordEmpl,fechaChecada,fila,horaChecada,turno]
    fechaChecadaEmpl.append(empl)
    if empleado is not None:
        listaEmpleados.append(empleado)


empleados=list(set(listaEmpleados))
empleados.sort()

fechaBase=fechaChecadaEmpl[3][2]

for idEmpl in empleados:
    for linea in fechaChecadaEmpl:
        if linea[5] is not None and linea[5]=='MC' and sabadosNoLab:


    
