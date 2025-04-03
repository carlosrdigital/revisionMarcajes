from datetime import datetime,time
x=datetime.strptime('06:00:00', '%H:%M:%S').time()<datetime.strptime('05:30:00', '%H:%M:%S').time()
print(x)

""" 
listanum=[0,1,2,3,4,5,6]

for num in listanum:
    for num1 in listanum:
        print(num,num1)
        if num==3:
            listanum.pop(3)
            """