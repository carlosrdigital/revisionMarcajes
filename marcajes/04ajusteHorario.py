import openpyxl,random #
from datetime import datetime,time

archivo='./marcajes noviembre/11.NOVIEMBRE_addHoras5.xlsx'
libro = openpyxl.load_workbook(archivo)
hoja = libro['noviembre']

limiteColumnas='J'
colIdEmpleado='B'
colFecha='E'
colChecada='F'
colHora='G'
colTurno='H'
limitefilas=14631+1

limEntrada={
    'infPrim':datetime.strptime('05:00:00', '%H:%M:%S').time(),
    'supPrim':datetime.strptime('11:30:00', '%H:%M:%S').time(),
    'infseg':datetime.strptime('05:00:00', '%H:%M:%S').time(),
    'supseg':datetime.strptime('13:30:00', '%H:%M:%S').time(),
    'infter':datetime.strptime('13:30:00', '%H:%M:%S').time(),
    'supter':datetime.strptime('21:30:00', '%H:%M:%S').time(),
    'infmix':datetime.strptime('05:00:00', '%H:%M:%S').time(),
    'supmix':datetime.strptime('08:00:00', '%H:%M:%S').time()
}

limSalida={
    'infPrim':datetime.strptime('14:30:00', '%H:%M:%S').time(),
    'supPrim':datetime.strptime('22:30:00', '%H:%M:%S').time(),
    'infseg':datetime.strptime('15:00:00', '%H:%M:%S').time(),
    'supseg':datetime.strptime('23:30:00', '%H:%M:%S').time(),
    'infter':datetime.strptime('05:30:00', '%H:%M:%S').time(),
    'supter':datetime.strptime('13:29:00', '%H:%M:%S').time(),
    'infmixAdm':datetime.strptime('18:00:00', '%H:%M:%S').time(),
    'supmixAdm':datetime.strptime('23:00:00', '%H:%M:%S').time(),
    'infmixop':datetime.strptime('17:00:00', '%H:%M:%S').time(),
    'supmixop':datetime.strptime('23:00:00', '%H:%M:%S').time(),
}

turnos=[]
nuevosAjustes=[]

for fila in range(1,limitefilas): #limitefilas
    coordHora=f'{colHora}{fila}'
    hora=hoja[coordHora].value
    if hora is not None:
        coordTurno=f'{colTurno}{fila}'
        turno=hoja[coordTurno].value
        #coordFecha=f'{colFecha}{fila}'
        #fecha=hoja[coordFecha].value
        if turno is not None:
            horaTurnoFila=[hora,turno,fila]#fecha]    
            turnos.append(horaTurnoFila)

for linea in turnos:
    minutosSalida=random.randint(1, 30)
    minutosEntra=random.randint(31,59)
    if linea[1]==1 and limEntrada['infPrim']<linea[0]<limEntrada['supPrim']: #primero
        #print(linea[0])
        ajuste=time(5,minutosEntra,linea[0].second)
        prim=(ajuste,linea[2])
        nuevosAjustes.append(prim)

    if linea[1]==1 and limSalida['infPrim']<linea[0]<limSalida['supPrim']:
        #print(linea[0])
        ajuste=time(14,minutosSalida,linea[0].second)
        prim=(ajuste,linea[2])
        nuevosAjustes.append(prim)

    if linea[1]==2 and limEntrada['infseg']<linea[0]<limEntrada['supseg']: #segundo
        #print(linea[0])
        ajuste=time(13,minutosEntra,linea[0].second)
        prim=(ajuste,linea[2])
        nuevosAjustes.append(prim)

    if linea[1]==2 and limSalida['infseg']<linea[0]<limSalida['supseg']: #segundo
        #print(linea[0])
        ajuste=time(22,minutosSalida,linea[0].second)
        prim=(ajuste,linea[2])
        nuevosAjustes.append(prim)

    if linea[1]==3 and limEntrada['infter']<linea[0]<limEntrada['supter']: #primero
        #print(linea[0])
        ajuste=time(21,minutosEntra,linea[0].second)
        prim=(ajuste,linea[2])
        #print(prim)
        nuevosAjustes.append(prim)

    if linea[1]==3 and limSalida['infter']<linea[0]<limSalida['supter']: #primero
        #print(linea[0])
        ajuste=time(5,minutosSalida,linea[0].second)
        prim=(ajuste,linea[2])
        print(prim)
        nuevosAjustes.append(prim)

    if linea[1]=='M' and limEntrada['infmix']<linea[0]<limEntrada['supmix']: #primero
        #print(linea[0])
        ajuste=time(8,minutosEntra,linea[0].second)
        prim=(ajuste,linea[2])
        nuevosAjustes.append(prim)
    
    if linea[1]=='M' and limSalida['infmixop']<linea[0]<limSalida['supmixop']:
        #print(linea[0])
        ajuste=time(16,linea[0].minute,linea[0].second)
        prim=(ajuste,linea[2])
        nuevosAjustes.append(prim)

    if linea[1]=='MC' and limEntrada['infmix']<linea[0]<limEntrada['supmix']:
        #print(linea[0])
        ajuste=time(8,linea[0].minute,linea[0].second)
        prim=(ajuste,linea[2])
        nuevosAjustes.append(prim)

    if linea[1]=='MC' and limSalida['infmixAdm']<linea[0]<limSalida['supmixAdm']:
        #print(linea[0])
        ajuste=time(17,linea[0].minute,linea[0].second)
        prim=(ajuste,linea[2])
        nuevosAjustes.append(prim)


for ajus in nuevosAjustes:
    hoja[f'{colHora}{ajus[1]}']=ajus[0]
libro.save('./marcajes noviembre/11.NOVIEMBRE_QnaDos_dos.xlsx')


