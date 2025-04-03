import openpyxl,random #proceso tres
from datetime import datetime,time

archivo='./marcajes noviembre/11.NOVIEMBRE_QnaDos_3.xlsx'
libro = openpyxl.load_workbook(archivo)
hoja = libro['noviembre']

limiteColumnas='J'
colIdEmpleado='B'
colFecha='E'
colHora='G'
colTurno='H'
limitefilas=14826+1
colChecada='F'

colTest='L'

entrada={
    'primMax':datetime.strptime('07:00:00', '%H:%M:%S').time(),
    'primMin':datetime.strptime('05:00:00', '%H:%M:%S').time(),
    'segMax':datetime.strptime('15:00:00', '%H:%M:%S').time(),
    'segMin':datetime.strptime('13:00:00', '%H:%M:%S').time(),
    'terMax':datetime.strptime('23:00:00', '%H:%M:%S').time(),
    'terMin':datetime.strptime('21:00:00', '%H:%M:%S').time(),

}

horaEnt={
    'Prim':5,
    'Seg':13,
    'Ter':21,
    'M':8,
    'MC':8

}


#fechaBase=[]
turnos=[]

for fila in range(1,limitefilas): #limitefilas
    coordHora=f'{colHora}{fila}'
    hora=hoja[coordHora].value
    if hora is None:
        coordTurno=f'{colTurno}{fila}'
        turno=hoja[coordTurno].value
        coordFecha=f'{colFecha}{fila}'
        fecha=hoja[coordFecha].value
        if turno is not None:
            horaTurnoFila=[hora,turno,fila,fecha]    
            turnos.append(horaTurnoFila)

for valTurno in turnos:
    minutosTurnoEnt = random.randint(10, 30)
    segTurnoEnt = random.randint(30, 59)
    minutosMixEnt=random.randint(1,59)


    #print(valTurno)
    if valTurno[1]==1:
        #print(f'{colChecada}{valTurno[2]} - 14:{minutosTurnoEnt}:{segTurnoEnt}')
        hoja[f'{colHora}{valTurno[2]}']=datetime.strptime(f'14:{minutosTurnoEnt}:{segTurnoEnt}', '%H:%M:%S').time()
        
        hoja[f'{colChecada}{valTurno[2]}']=valTurno[3]

    if valTurno[1]==2:
        #print(f'{colChecada}{valTurno[2]} - 22:{minutosTurnoEnt}:{segTurnoEnt}')
        hoja[f'{colHora}{valTurno[2]}']=datetime.strptime(f'22:{minutosTurnoEnt}:{segTurnoEnt}', '%H:%M:%S').time()
        hoja[f'{colChecada}{valTurno[2]}']=valTurno[3]


    if valTurno[1]==3:
        print(f'{colChecada}{valTurno[2]} - 21:{minutosTurnoEnt}:{segTurnoEnt}')
        hoja[f'{colHora}{valTurno[2]}']=datetime.strptime(f'05:{minutosTurnoEnt}:{segTurnoEnt}', '%H:%M:%S').time()
        hoja[f'{colChecada}{valTurno[2]}']=valTurno[3]


    if valTurno[1]=='M':
        print(f'{colChecada}{valTurno[2]} - 08:{minutosMixEnt}:{segTurnoEnt}')
        hoja[f'{colHora}{valTurno[2]}']=datetime.strptime(f'04:{segTurnoEnt}:{segTurnoEnt}', '%H:%M:%S').time()
        hoja[f'{colChecada}{valTurno[2]}']=valTurno[3]

    
    if valTurno[1]=='MC':
        print(f'{colChecada}{valTurno[2]} - 08:{minutosMixEnt}:{segTurnoEnt}')
        hoja[f'{colHora}{valTurno[2]}']=datetime.strptime(f'05:{segTurnoEnt}:{segTurnoEnt}', '%H:%M:%S').time()
        hoja[f'{colChecada}{valTurno[2]}']=valTurno[3]

libro.save('./marcajes noviembre/11.NOVIEMBRE_QnaDos_4.xlsx')
