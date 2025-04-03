import openpyxl #proceso uno eliminar duplicados
from datetime import datetime,time

archivo='./marcajes septiembre/9. SEPTIEMBRE BEL.xlsx'
libro = openpyxl.load_workbook(archivo)
hoja = libro['Hoja1']

limiteColumnas='J'
colIdEmpleado='B'
colChecada='F'
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
    empl=[empleado,coordEmpl,fechaChecada,fila]
    fechaChecadaEmpl.append(empl)
    if empleado is not None:
        listaEmpleados.append(empleado)


empleados=list(set(listaEmpleados))
empleados.sort()

for idEmpl in empleados:
    for linea in fechaChecadaEmpl:
        #print(linea)
        if linea[2] is not None and idEmpl==linea[0]:
            reps=0
            for linea1 in fechaChecadaEmpl:
                if linea1[2] is not None and idEmpl==linea1[0]:
                    if linea[2]== linea1[2]:
                        reps+=1
                        #print (reps)
                        if reps>2:
                            print(f'{idEmpl}, {linea[2]}, {linea[3]}, repeticiones: {reps}')
                            ultimoRepetido.append(linea1[3])

if ultimoRepetido is not None:
    eliminarFila=list(set(ultimoRepetido))
    eliminarFila.sort(reverse=True)
    print('se eliminaran los duplicados')

    for elimina in eliminarFila:
        print(f'eliminando fila con mas de dos checadas por dia: {elimina} ')
        hoja.delete_rows(elimina,1)
    libro.save('./marcajes septiembre/9. SEPTIEMBRE BEL_noDuplicados_1.xlsx')
else:
    print('no hay mas de dos checadas por dia')


     #       fechaXempleado.append(linea[2])#(linea[2])
    #        dataMan.append(linea)
    #print(f'empleado examinado{idEmpl}')
    #cantidadFechas=len(fechaXempleado)
    #for fecha1 in fechaXempleado:
    #    cuentaDiferencias=0
    #    for indi,fecha2 in enumerate(fechaXempleado):
           # if fecha1==fecha2:
          #      print(f'{fecha1} se repite en {dataMan[indi]}')
            #elif fecha1!=fecha2:
             #   cuentaDiferencias+=1
        #if cantidadFechas==cuentaDiferencias:
         #   print(f'la fecha {fecha1}, no se repite')
    #fechaXempleado.clear()
    #dataMan.clear()
    
"""    for indi,unaFecha in enumerate(fechaXempleado):
repeticiones=fechaXempleado.count(unaFecha)
#print(f'{unaFecha}, se repite: {repeticiones}')
if repeticiones==1:
print('agregar')
print(dataMan[indi])
elif repeticiones>2:
print(repeticiones)
print('eliminar excedente')
print(dataMan[indi])"""




        
