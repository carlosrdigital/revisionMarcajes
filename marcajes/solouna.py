import openpyxl #proceso uno eliminar duplicados
from datetime import datetime,time

archivo='./marcajes noviembre/11.NOVIEMBRE_QnaDos_.xlsx'
libro = openpyxl.load_workbook(archivo)
hoja = libro['noviembre']

limiteColumnas='J'
colIdEmpleado='B'
colNombre='C'
colDepto='D'
colChecada='E'
colFechaChec='F'
colHora='G'
colTurno='H'
limitefilas=15618+1
ultimoRepetido=[]

listaEmpleados=[]
fechaChecadaEmpl=[] 



for fila in range(1,limitefilas): #limitefilas #extraer info
    coordEmpl=f'{colIdEmpleado}{fila}'
    empleado=hoja[coordEmpl].value
    coordChecada=f'{colChecada}{fila}'
    fechaChecada=hoja[coordChecada].value
    coordTurno=f'{colTurno}{fila}'
    turno=hoja[coordTurno].value
    coordNom=f'{colNombre}{fila}'
    nombre=hoja[coordNom].value
    coordDep=f'{colDepto}{fila}'
    depto=hoja[coordDep].value
    coordFechaChec=f'{colFechaChec}{fila}'
    fechaChec=hoja[coordFechaChec].value         

    empl=[empleado,coordEmpl,fechaChecada,fila,turno,nombre,depto,fechaChec]
    fechaChecadaEmpl.append(empl)
    if empleado is not None:
        listaEmpleados.append(empleado)


empleados=list(set(listaEmpleados))
empleados.sort()
noReps=[]
cuenta= None
for idEmpl in empleados:
    for linea in fechaChecadaEmpl:
        #print(linea)
        if linea[2] is not None and linea[4] is not None and idEmpl==linea[0]:
            reps=0
            for linea1 in fechaChecadaEmpl:
                if linea1[2] is not None and idEmpl==linea1[0]:
                    if linea[2]== linea1[2]:
                        reps+=1
                        #print (reps)
                        if reps==1:
                            unaCheck=1
                            cuenta=linea
                        elif reps>1:
                            unaCheck=0
                            cuenta=None
            if cuenta is not None:
                noReps.append(cuenta)


                            #print(f'{idEmpl}, {linea[2]}, {linea[3]}, repeticiones: {reps}')
                            #ultimoRepetido.append(linea1[3])
for noRepite in reversed(noReps):
    #print(noRepite)
    nuevaFila=noRepite[3]+1
    hoja.insert_rows(nuevaFila)
    hoja[f'B{nuevaFila}']=noRepite[0] #id
    hoja[f'C{nuevaFila}']=noRepite[5] #id
    hoja[f'D{nuevaFila}']=noRepite[6] #DEPTO
    hoja[f'E{nuevaFila}']=noRepite[2] #fecha
    hoja[f'F{nuevaFila}']=noRepite[7] #FECHACHEC
    #hoja[f'G{nuevaFila}']=noRepite[2] #HORA
    hoja[f'H{nuevaFila}']=noRepite[4] #fecha
    hoja[f'K{nuevaFila}']='+1'





    


libro.save('./marcajes noviembre/11.NOVIEMBRE_QnaDos_3.xlsx')

