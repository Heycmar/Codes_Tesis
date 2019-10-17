import pymysql
import openpyxl, serial
import threading
from Tkinter import *
import numpy as np
from numpy.linalg import inv
from time import time
from time import sleep
import matplotlib.pyplot as plt
import matplotlib.animation as animation

def QPhild(H,f,A_cons,b):
    n1=len(A_cons)
    eta=-1*np.dot(inv(H),f)
    kk=0
    for i in range(0,n1):
        if (np.dot(A_cons[i,:],eta)>b[i]):
            kk=kk+1
        else:
            kk=kk+0
    if (kk==0):
        return eta
    P=np.dot(A_cons,np.dot(inv(H),A_cons.transpose()))
    d=np.dot(A_cons,np.dot(inv(H),f))+b
    n=len(d)
    m=len(d[0])
    x_ini=np.zeros((n,m))
    lambda1=x_ini    
    al=10
    for km in range(0,38):
        lambda_p=lambda1
        for i in range(0,n):
            w=np.dot(P[i,:],lambda1)-np.dot(P[i,i],lambda1[i,0])
            w=w+d[i,0]
            #print(w)
            la=-w/P[i,i]
            lambda1[i,0]=np.max([0,la])
        al=np.dot((lambda1-lambda_p).transpose(),lambda1-lambda_p)
        if (al<10e-8):
            break
    eta=np.dot(inv(-1*H),f)+np.dot(np.dot(inv(-1*H),A_cons.transpose()),lambda1)
    return eta

doc = openpyxl.load_workbook('Constantes.xlsx')
hoja = doc.get_sheet_by_name('Hoja1')
kij=np.zeros((9,9))
for f in range(9):
    g=0
    for c in 'ABCDEFGHI':
        celda=c+str(f+1)
        kij[f][g]=float(hoja[celda].value)
        g=g+1

doc1 = openpyxl.load_workbook('ConsPert.xlsx')
hoja1 = doc1.get_sheet_by_name('Hoja1')
kip=np.zeros((9,2))
for f in range(9):
    g=0
    for c in 'AB':
        celda=c+str(f+1)
        kip[f][g]=float(hoja1[celda].value)
        g=g+1

Ts=1#Muestreo

###################################################### PID ###################################
kp=0.15#Proporcional Continua
ki=0.01#Integral Continua
kd=0#Derivativa Continua

Ti=kp/ki
Td=kd/kp
Kpd=kp-kp*Ts/(2*Ti)
Kid=kp*Ts/Ti
Kdd=Td*kp/Ts
a=Kid+Kpd+Kdd
b=-Kpd-2*Kdd
c=Kdd

Amatriz=[a,a,a,a,a,a,a,a,a]
Bmatriz=[b,b,b,b,b,b,b,b,b]
Cmatriz=[c,c,c,c,c,c,c,c,c]

ematriz=np.zeros((9,3))
upid=np.zeros((9,3))

s=np.zeros((9,1))
y=np.zeros((9,1))

Refg=100
Ref=Refg*np.ones((9,1))

pert=np.zeros((2,1))

#################################################### Replicador #####################################
P=5000
f=np.zeros((10,1))

ya=np.zeros((9,1))
xa=10*np.ones((10,1))
x=100*np.ones((10,1))
x[9][0]=4100
uarep=np.zeros((10,1))

B=1000#800
beta=0.44#0.1
alpha=0.1#0.5

#####################################################################################################
#################################### MPC  #############################################

Ad=0.3679*np.eye(9)
Bd=np.zeros((9,9))
Bd[0][0]=11.8140
Bd[1][1]=24.0706
Bd[2][2]=15.1225
Bd[3][3]=24.7479
Bd[4][4]=11.3085
Bd[5][5]=9.1221
Bd[6][6]=13.9532
Bd[7][7]=15.2619
Bd[8][8]=18.5826

Cd=kij
Dd=np.zeros((9,9))

m1=len(Cd)
n1=len(Bd)

n_in=len(Bd[0])
An=np.eye(m1+n1)
An[0:n1,0:n1]=Ad
An[n1:n1+m1,0:n1]=np.dot(Cd,Ad)
Bn=np.zeros((n1+m1,n_in))
Bn[0:n1,:]=Bd
Bn[n1:n1+m1,:]=np.dot(Cd,Bd)
Cn=np.zeros((m1,n1+m1))
Cn[:,n1:n1+m1]=np.eye(m1)
n=len(Cn)
Np=10
Nc=4
F=np.zeros((Np*n,2*n_in))
F[0:n,:]=np.dot(Cn,An)
for i in range(0,Np-1):
    A1=An
    j=1
    while j<=i+1:
        A1=np.dot(A1,An)
        j=j+1
    F[n:n+n1,:]=np.dot(Cn,A1)
    n=n+n1
n=len(Cn)
m=len(Cn[0])
n2=n
Phi=np.zeros((Np*n,Nc*n))
Phi[0:n,0:n]=np.dot(Cn,Bn)
for i in range(0,Np-1):
    j=1
    A1=An
    while j<i+1:
        A1=np.dot(A1,An)
        j=j+1
    Prod=np.dot(Cn,A1)
    Phi[n:n+n1,0:n2]=np.dot(Prod,Bn)
    Phi[n:n+n1,n2:len(Phi[0])]=Phi[n-n1:n,0:n2*(Nc-1)]
    n=n+n1
Phi_Phi=np.dot(Phi.transpose(),Phi)
Phi_F=np.dot(Phi.transpose(),F)
Phi_R=Phi_F[:,len(Phi_F[0])-n2:len(Phi_F[0])]
xm=np.zeros((9,1))
xm_old=np.zeros((9,1))
Xf=np.zeros((2*n2,1))

u=np.zeros((9,1))
ua=np.zeros((9,1))

R=2100*np.eye(len(Phi_Phi))

Ref1=Refg*np.ones((9,1))

d=np.zeros((9,1))
d[0][0]=18.6895
d[1][0]=38.0792
d[2][0]=23.9234
d[3][0]=39.1506
d[4][0]=17.8898
d[5][0]=14.4309
d[6][0]=22.0736 
d[7][0]=24.1440
d[8][0]=29.3973


#########Restricciones####################
n_in=9
aux=9

A_cons=np.zeros((2*Nc*n_in,n_in*Nc))

A_cons[0:n_in,0:n_in]=np.eye(n_in)

for i in range(1,Nc):
    A_cons[aux:aux+n_in,n_in:len(A_cons[0])]=A_cons[aux-n_in:aux,0:n_in*(Nc-1)]
    A_cons[aux:aux+n_in,0:n_in]=np.eye(n_in)
    aux=aux+n_in

A_cons[aux:aux+n_in,0:n_in]=-1*np.eye(n_in)

A_cons[aux:aux+n_in,n_in:len(A_cons[0])]=np.zeros((n_in,n_in*(Nc-1)))
for i in range(0,Nc):
    A_cons[aux:aux+n_in,n_in:len(A_cons[0])]=A_cons[aux-n_in:aux,0:n_in*(Nc-1)]
    A_cons[aux:aux+n_in,0:n_in]=-1*np.eye(n_in)
    aux=aux+n_in

umin=0
umax=100
umpc=np.zeros((9,1))

DeltaU=np.zeros((Nc*9,1))
b=np.zeros((2*n_in*Nc,1))

def PID(stop):
    global y
    while True:
        start=time()
        
        if stop():
            break
        
        s1_1=100
        s2_1=100
        s3_1=100
        s4_1=100
        s5_1=100
        s6_1=100
        s7_1=100
        s8_1=100
        s9_1=100
        s10_1=100
        s11_1=100

        """arduino1=serial.Serial('/dev/rfcomm0',9600)
        arduino2=serial.Serial('/dev/rfcomm1',9600)
        arduino3=serial.Serial('/dev/rfcomm2',9600)
        arduino4=serial.Serial('/dev/rfcomm3',9600)
        arduino5=serial.Serial('/dev/rfcomm4',9600)
        arduino6=serial.Serial('/dev/rfcomm5',9600)
        arduino7=serial.Serial('/dev/rfcomm6',9600)

        s1_1=arduino1.readline()
        s2_1=arduino2.readline()
        s3_1=arduino3.readline()
        s4_1=arduino4.readline()
        s5_1=arduino5.readline()
        s6_1=arduino6.readline()
        s7_1=arduino7.readline()
        
        while s1_1.find('.')==-1:
            s1_1=arduino1.readline()
        
        while s2_1.find('.')==-1:
            s2_1=arduino2.readline()
        
        while s3_1.find('.')==-1:
            s3_1=arduino3.readline()
        
        while s4_1.find('.')==-1:
            s4_1=arduino4.readline()

        while s5_1.find('.')==-1:
            s5_1=arduino5.readline()
        
        while s6_1.find('.')==-1:
            s6_1=arduino6.readline()
        
        while s7_1.find('.')==-1:
            s7_1=arduino7.readline()

        arduino1.close()
        arduino2.close()
        arduino3.close()
        arduino4.close()
        arduino5.close()
        arduino6.close()
        arduino7.close()

        f8=open ('med8.txt','r')
        s8_1=f8.readline()
        f8.close()
        while s8_1.find('.')==-1:
            f8=open ('med8.txt','r')
            s8_1=f8.readline()
            f8.close()

        f9=open ('med9.txt','r')
        s9_1=f9.readline()
        f9.close()
        while s9_1.find('.')==-1:
            f9=open ('med9.txt','r')
            s9_1=f9.readline()
            f9.close()    

        f10=open ('med10.txt','r')
        s10_1=f10.readline()
        f10.close()
        while s10_1.find('.')==-1:
            f10=open ('med10.txt','r')
            s10_1=f10.readline()
            f10.close()

        f11=open ('med11.txt','r')
        s11_1=f11.readline()
        f11.close()
        while s11_1.find('.')==-1:
            f11=open ('med11.txt','r')
            s11_1=f11.readline()
            f11.close()"""

        s[0][0]=float(s1_1)
        s[1][0]=float(s2_1)
        s[2][0]=float(s3_1)
        s[3][0]=float(s4_1)
        s[4][0]=float(s5_1)
        s[5][0]=float(s6_1)
        s[6][0]=float(s7_1)
        s[7][0]=float(s8_1)
        s[8][0]=float(s9_1)

        pert[0][0]=float(s10_1)
        pert[1][0]=float(s11_1)

        y=np.dot(kij,s)+np.dot(kip,pert)
        
        for i in range(0,9):
            ematriz[i][0]=Ref[i][0]-y[i][0]
            upid[i][0]=Amatriz[i]*ematriz[i][0]+Bmatriz[i]*ematriz[i][1]+Cmatriz[i]*ematriz[i][2]+upid[i][1]
            upid[i][0]=round(upid[i][0],2)
        
        for q in range(9):
            if upid[q][0]>100:
                upid[q][0]=100
    
        for p in range(9):
            if upid[p][0]<0:
                upid[p][0]=0.01
        
        fa=open('exp1.txt','w')
        fa.write(str(upid[0][0]))
        fa.close()
        fa=open('exp2.txt','w')
        fa.write(str(upid[1][0]))
        fa.close()
        fa=open('exp3.txt','w')
        fa.write(str(upid[2][0]))
        fa.close()
        fa=open('exp4.txt','w')
        fa.write(str(upid[3][0]))
        fa.close()
        fa=open('exp5.txt','w')
        fa.write(str(upid[4][0]))
        fa.close()
        fa=open('exp6.txt','w')
        fa.write(str(upid[5][0]))
        fa.close()
        fa=open('exp7.txt','w')
        fa.write(str(upid[6][0]))
        fa.close()
        fa=open('exp8.txt','w')
        fa.write(str(upid[7][0]))
        fa.close()
        fa=open('exp9.txt','w')
        fa.write(str(upid[8][0]))
        fa.close()

        for k in range(0,9):
            for i in range(0,2):
                ematriz[k][i+1]=ematriz[k][i]
                upid[k][i+1]=upid[k][i]

        print(y)
        print(upid)

        fin=time()-start
        Pause=Ts-fin
        if Pause<0:
            Pause=0    
        sleep(Pause)

########################################################################################

###################################### Replicador #################################################

def REPLICADOR(stop):
    global y

    global Ref
    global alpha
    global beta

    while True:
        start=time()
        if stop():
            break
        
        s1_1=100
        s2_1=100
        s3_1=100
        s4_1=100
        s5_1=100
        s6_1=100
        s7_1=100
        s8_1=100
        s9_1=100
        s10_1=100
        s11_1=100

        """arduino1=serial.Serial('/dev/rfcomm0',9600)
        arduino2=serial.Serial('/dev/rfcomm1',9600)
        arduino3=serial.Serial('/dev/rfcomm2',9600)
        arduino4=serial.Serial('/dev/rfcomm3',9600)
        arduino5=serial.Serial('/dev/rfcomm4',9600)
        arduino6=serial.Serial('/dev/rfcomm5',9600)
        arduino7=serial.Serial('/dev/rfcomm6',9600)

        s1_1=arduino1.readline()
        s2_1=arduino2.readline()
        s3_1=arduino3.readline()
        s4_1=arduino4.readline()
        s5_1=arduino5.readline()
        s6_1=arduino6.readline()
        s7_1=arduino7.readline()
        
        while s1_1.find('.')==-1:
            s1_1=arduino1.readline()
        
        while s2_1.find('.')==-1:
            s2_1=arduino2.readline()
        
        while s3_1.find('.')==-1:
            s3_1=arduino3.readline()
        
        while s4_1.find('.')==-1:
            s4_1=arduino4.readline()

        while s5_1.find('.')==-1:
            s5_1=arduino5.readline()
        
        while s6_1.find('.')==-1:
            s6_1=arduino6.readline()
        
        while s7_1.find('.')==-1:
            s7_1=arduino7.readline()

        arduino1.close()
        arduino2.close()
        arduino3.close()
        arduino4.close()
        arduino5.close()
        arduino6.close()
        arduino7.close()

        f8=open ('med8.txt','r')
        s8_1=f8.readline()
        f8.close()
        while s8_1.find('.')==-1:
            f8=open ('med8.txt','r')
            s8_1=f8.readline()
            f8.close()

        f9=open ('med9.txt','r')
        s9_1=f9.readline()
        f9.close()
        while s9_1.find('.')==-1:
            f9=open ('med9.txt','r')
            s9_1=f9.readline()
            f9.close()    

        f10=open ('med10.txt','r')
        s10_1=f10.readline()
        f10.close()
        while s10_1.find('.')==-1:
            f10=open ('med10.txt','r')
            s10_1=f10.readline()
            f10.close()

        f11=open ('med11.txt','r')
        s11_1=f11.readline()
        f11.close()
        while s11_1.find('.')==-1:
            f11=open ('med11.txt','r')
            s11_1=f11.readline()
            f11.close()"""

        s[0][0]=float(s1_1)
        s[1][0]=float(s2_1)
        s[2][0]=float(s3_1)
        s[3][0]=float(s4_1)
        s[4][0]=float(s5_1)
        s[5][0]=float(s6_1)
        s[6][0]=float(s7_1)
        s[7][0]=float(s8_1)
        s[8][0]=float(s9_1)

        pert[0][0]=float(s10_1)
        pert[1][0]=float(s11_1)

        y=np.dot(kij,s)+np.dot(kip,pert)

        # Replicador
        for i in range(0,9):
            f[i][0]=x[i][0]*(B+Ref[i][0]-y[i][0])

        suma=0
        for l in range(0,9):
            suma=f[l][0]+suma

        f[9][0]=x[9][0]*B
        suma=suma+f[9][0]
        Fb=suma/P

        for p in range(0,9):
            x[p][0]=round(xa[p][0]*(beta*(alpha+(B+Ref[p][0]-ya[p][0]))/(alpha+Fb)),2)

        x[9][0]=round(xa[9][0]*(beta*(alpha+B)/(alpha+Fb)),2)
        
        xa[:,:]=x
        ya[:,:]=y
        
        for q in range(0,9):
            uarep[q][0]=x[q][0]+10

            if uarep[q][0]>100:
                uarep[q][0]=100
        uarep[9][0]=x[9][0]    

        for q in range(0,9):
            if (x[q][0]<10):
                aux2=10-x[q][0]
                x[q][0]=10
                x[9][0]=x[9][0]-aux2

        uarep[9][0]=x[9][0]
        
        fa=open('exp1.txt','w')
        fa.write(str(uarep[0][0]))
        fa.close()
        fa=open('exp2.txt','w')
        fa.write(str(uarep[1][0]))
        fa.close()
        fa=open('exp3.txt','w')
        fa.write(str(uarep[2][0]))
        fa.close()
        fa=open('exp4.txt','w')
        fa.write(str(uarep[3][0]))
        fa.close()
        fa=open('exp5.txt','w')
        fa.write(str(uarep[4][0]))
        fa.close()
        fa=open('exp6.txt','w')
        fa.write(str(uarep[5][0]))
        fa.close()
        fa=open('exp7.txt','w')
        fa.write(str(uarep[6][0]))
        fa.close()
        fa=open('exp8.txt','w')
        fa.write(str(uarep[7][0]))
        fa.close()
        fa=open('exp9.txt','w')
        fa.write(str(uarep[8][0]))
        fa.close()
        print(y)
        print(x)

        fin=time()-start
        Pause=Ts-fin
        if Pause<0:
            Pause=0
        sleep(Pause)

######################################## MPC ##################################################

def MPC(stop):
    global y

    global umpc
    global Ref1
    global Ref
    global Np
    global Nc
    global n
    global n_in

    m1=len(Cd)
    n1=len(Bd)
    n=len(Cn)

    F=np.zeros((Np*n,2*n_in))
    F[0:n,:]=np.dot(Cn,An)
    for i in range(0,Np-1):
        A1=An
        j=1
        while j<=i+1:
            A1=np.dot(A1,An)
            j=j+1
        F[n:n+n1,:]=np.dot(Cn,A1)
        n=n+n1
    n=len(Cn)
    m=len(Cn[0])
    n2=n
    Phi=np.zeros((Np*n,Nc*n))
    Phi[0:n,0:n]=np.dot(Cn,Bn)
    for i in range(0,Np-1):
        j=1
        A1=An
        while j<i+1:
            A1=np.dot(A1,An)
            j=j+1
        Prod=np.dot(Cn,A1)
        Phi[n:n+n1,0:n2]=np.dot(Prod,Bn)
        Phi[n:n+n1,n2:len(Phi[0])]=Phi[n-n1:n,0:n2*(Nc-1)]
        n=n+n1
    Phi_Phi=np.dot(Phi.transpose(),Phi)
    Phi_F=np.dot(Phi.transpose(),F)
    Phi_R=Phi_F[:,len(Phi_F[0])-n2:len(Phi_F[0])]
    xm=np.zeros((9,1))
    xm_old=np.zeros((9,1))
    Xf=np.zeros((2*n2,1))

    u=np.zeros((9,1))
    ua=np.zeros((9,1))

    R=2100*np.eye(len(Phi_Phi))

    #########Restricciones####################
    n_in=9
    aux=9

    A_cons=np.zeros((2*Nc*n_in,n_in*Nc))

    A_cons[0:n_in,0:n_in]=np.eye(n_in)

    for i in range(1,Nc):
        A_cons[aux:aux+n_in,n_in:len(A_cons[0])]=A_cons[aux-n_in:aux,0:n_in*(Nc-1)]
        A_cons[aux:aux+n_in,0:n_in]=np.eye(n_in)
        aux=aux+n_in

    A_cons[aux:aux+n_in,0:n_in]=-1*np.eye(n_in)

    A_cons[aux:aux+n_in,n_in:len(A_cons[0])]=np.zeros((n_in,n_in*(Nc-1)))
    for i in range(0,Nc):
        A_cons[aux:aux+n_in,n_in:len(A_cons[0])]=A_cons[aux-n_in:aux,0:n_in*(Nc-1)]
        A_cons[aux:aux+n_in,0:n_in]=-1*np.eye(n_in)
        aux=aux+n_in

    DeltaU=np.zeros((Nc*9,1))
    b=np.zeros((2*n_in*Nc,1))

    while True:
        start=time()
        if stop():
            break

        s1_1=100
        s2_1=100
        s3_1=100
        s4_1=100
        s5_1=100
        s6_1=100
        s7_1=100
        s8_1=100
        s9_1=100
        s10_1=100
        s11_1=100

        """arduino1=serial.Serial('/dev/rfcomm0',9600)
        arduino2=serial.Serial('/dev/rfcomm1',9600)
        arduino3=serial.Serial('/dev/rfcomm2',9600)
        arduino4=serial.Serial('/dev/rfcomm3',9600)
        arduino5=serial.Serial('/dev/rfcomm4',9600)
        arduino6=serial.Serial('/dev/rfcomm5',9600)
        arduino7=serial.Serial('/dev/rfcomm6',9600)

        s1_1=arduino1.readline()
        s2_1=arduino2.readline()
        s3_1=arduino3.readline()
        s4_1=arduino4.readline()
        s5_1=arduino5.readline()
        s6_1=arduino6.readline()
        s7_1=arduino7.readline()

        while s1_1.find('.')==-1:
            s1_1=arduino1.readline()
        
        while s2_1.find('.')==-1:
            s2_1=arduino2.readline()
        
        while s3_1.find('.')==-1:
            s3_1=arduino3.readline()
        
        while s4_1.find('.')==-1:
            s4_1=arduino4.readline()

        while s5_1.find('.')==-1:
            s5_1=arduino5.readline()
        
        while s6_1.find('.')==-1:
            s6_1=arduino6.readline()
        
        while s7_1.find('.')==-1:
            s7_1=arduino7.readline()

        arduino1.close()
        arduino2.close()
        arduino3.close()
        arduino4.close()
        arduino5.close()
        arduino6.close()
        arduino7.close()

        
        f8=open ('med8.txt','r')
        s8_1=f8.readline()
        f8.close()
        while s8_1.find('.')==-1:
            f8=open ('med8.txt','r')
            s8_1=f8.readline()
            f8.close()

        f9=open ('med9.txt','r')
        s9_1=f9.readline()
        f9.close()
        while s9_1.find('.')==-1:
            f9=open ('med9.txt','r')
            s9_1=f9.readline()
            f9.close()         

        f10=open ('med10.txt','r')
        s10_1=f10.readline()
        f10.close()
        while s10_1.find('.')==-1:
            f10=open ('med10.txt','r')
            s10_1=f10.readline()
            f10.close()

        f11=open ('med11.txt','r')
        s11_1=f11.readline()
        f11.close()
        while s11_1.find('.')==-1:
            f11=open ('med11.txt','r')
            s11_1=f11.readline()
            f11.close()""" 
        
        
        #DeltaU=np.dot(inv(Phi_Phi+R),np.dot(Phi_R,Ref)-np.dot(Phi_F,Xf))
        H=2*Phi_Phi+R
        f=-2*(np.dot(Phi_R,Ref)-np.dot(Phi_F,Xf))
        aux=0
        for i in range(0,Nc):
            b[aux:aux+n_in,:]=umax*np.ones((9,1))-umpc
            aux=aux+n_in
        for i in range(0,Nc):
            b[aux:aux+n_in,:]=umin*np.ones((9,1))+umpc
            aux=aux+n_in
        #print(A_cons)
        DeltaU=QPhild(H,f,A_cons,b)
        deltau=DeltaU[0:n2,:]
        
        umpc=umpc+deltau
        xm_old[:,:]=xm

        xm[0][0]=float(s1_1)
        xm[1][0]=float(s2_1)
        xm[2][0]=float(s3_1)
        xm[3][0]=float(s4_1)
        xm[4][0]=float(s5_1)
        xm[5][0]=float(s6_1)
        xm[6][0]=float(s7_1)
        xm[7][0]=float(s8_1)
        xm[8][0]=float(s9_1)

        pert[0][0]=float(s10_1)
        pert[1][0]=float(s11_1)

        yr=np.dot(kij,xm)

        y=yr+np.dot(kip,pert)

        Ref=Ref1-np.dot(kip,pert)

        Xf[0:n2,:]=xm-xm_old
        Xf[n2:2*n2,:]=yr
        
        ua[0][0]=-6.5559/(2*0.1215)+(((6.5559)**2-4*0.1215*(-1.6376-d[0][0]*umpc[0][0]))**0.5)/(2*0.1215)
        ua[1][0]=-7.1012/(2*0.3183)+(((7.1012)**2-4*0.3183*(-85.205-d[1][0]*umpc[1][0]))**0.5)/(2*0.3183)
        ua[2][0]=-4.7038/(2*0.1976)+(((4.7038)**2-4*0.1976*(-54.036-d[2][0]*umpc[2][0]))**0.5)/(2*0.1976)
        ua[3][0]=-7.1792/(2*0.3281)+(((7.1792)**2-4*0.3281*(-83.856-d[3][0]*umpc[3][0]))**0.5)/(2*0.3281)
        ua[4][0]=-4.2495/(2*0.1408)+(((4.2495)**2-4*0.1408*(-43.973-d[4][0]*umpc[4][0]))**0.5)/(2*0.1408)
        ua[5][0]=-3.8355/(2*0.1094)+(((3.8355)**2-4*0.1094*(-34.462-d[5][0]*umpc[5][0]))**0.5)/(2*0.1094)
        ua[6][0]=-3.7761/(2*0.1879)+(((3.7761)**2-4*0.1879*(-49.251-d[6][0]*umpc[6][0]))**0.5)/(2*0.1879)
        ua[7][0]=-5.3571/(2*0.1933)+(((5.3571)**2-4*0.1933*(-54.307-d[7][0]*umpc[7][0]))**0.5)/(2*0.1933)
        ua[8][0]=-5.8920/(2*0.2418)+(((5.8920)**2-4*0.2418*(-67.466-d[8][0]*umpc[8][0]))**0.5)/(2*0.2418)

        
        for q in range(9):
            ua[q][0]=round(ua[q][0],2)

            if ua[q][0]>100:
                ua[q][0]=100
            
            if ua[q][0]<0:
                ua[q][0]=0.01

        fa=open('exp1.txt','w')
        fa.write(str(ua[0][0]))
        fa.close()
        fa=open('exp2.txt','w')
        fa.write(str(ua[1][0]))
        fa.close()
        fa=open('exp3.txt','w')
        fa.write(str(ua[2][0]))
        fa.close()
        fa=open('exp4.txt','w')
        fa.write(str(ua[3][0]))
        fa.close()
        fa=open('exp5.txt','w')
        fa.write(str(ua[4][0]))
        fa.close()
        fa=open('exp6.txt','w')
        fa.write(str(ua[5][0]))
        fa.close()
        fa=open('exp7.txt','w')
        fa.write(str(ua[6][0]))
        fa.close()
        fa=open('exp8.txt','w')
        fa.write(str(ua[7][0]))
        fa.close()
        fa=open('exp9.txt','w')
        fa.write(str(ua[8][0]))
        fa.close()  
        
        print(y)
        print(umpc)

        fin=time()-start
        Pause=Ts-fin
        if Pause<0:
            Pause=0    
        sleep(Pause) 


def activarPID():
    global Refg
    global Ref
    global Ts
    global Amatriz
    global Bmatriz
    global Cmatriz
    global kp 
    global ki
    global kd
    global Amatriz
    global Bmatriz
    global Cmatriz
    global tmp1
    global tmp2
    global tmp3
    if tmp1.isAlive()==False:
        if tmp2.isAlive()==False:
            if tmp3.isAlive()==False:
                Refg=float(vo.get())
                kp=float(vo1.get())
                ki=float(vo2.get())
                kd=float(vo3.get())
                Ref=Refg*np.ones((9,1))
                
                Ti=kp/ki
                Td=kd/kp
                Kpd=kp-kp*Ts/(2*Ti)
                Kid=kp*Ts/Ti
                Kdd=Td*kp/Ts
                a=Kid+Kpd+Kdd
                b=-Kpd-2*Kdd
                c=Kdd

                Amatriz=[a,a,a,a,a,a,a,a,a]
                Bmatriz=[b,b,b,b,b,b,b,b,b]
                Cmatriz=[c,c,c,c,c,c,c,c,c]

                tmp1.start()

def activarReplicador():
    global Refg
    global Ref
    global alpha
    global beta
    global tmp1
    global tmp2
    global tmp3

    if tmp1.isAlive()==False:
        if tmp2.isAlive()==False:
            if tmp3.isAlive()==False:
                Refg=float(vo.get())
                alpha=float(vo4.get())
                beta=float(vo5.get())
                Ref=Refg*np.ones((9,1))
                tmp2.start()

def activarMPC():
    global Refg
    global Ref
    global Ref1
    global Np
    global Nc
    global tmp1
    global tmp2
    global tmp3

    if tmp1.isAlive()==False:
        if tmp2.isAlive()==False:
            if tmp3.isAlive()==False:
                Refg=float(vo.get())
                Np=int(vo6.get())
                Nc=int(vo7.get())
                Ref=Refg*np.ones((9,1))
                Ref1=Refg*np.ones((9,1))
                tmp3.start()

def desactivar():
    global tmp1
    global tmp2
    global tmp3
    global stop_threads

    if tmp1.isAlive()==True:
        stop_threads = True
        tmp1.join()
        stop_threads = False
        tmp1 = threading.Thread(target=PID, args=(lambda: stop_threads,))

    if tmp2.isAlive()==True:
        stop_threads = True
        tmp2.join()
        stop_threads = False
        tmp2 = threading.Thread(target=REPLICADOR, args=(lambda: stop_threads,))

    if tmp3.isAlive()==True:
        stop_threads = True
        tmp3.join()
        stop_threads = False
        tmp3 = threading.Thread(target=MPC, args=(lambda: stop_threads,))

COUNT = 100

def next():
    global data_time
    global data_y1
    global data_y2
    global data_y3
    global data_y4
    global data_y5
    global data_y6
    global data_y7
    global data_y8
    global data_y9

    it = 0
    while it <= COUNT:
        it += 1
        if it==100:
            it=0
            
        yield it

def update(it):

    data_time[tiempo]=tiempo

    data_y1[tiempo]=float(y[0])
    data_y2[tiempo]=float(y[1])
    data_y3[tiempo]=float(y[2])
    data_y4[tiempo]=float(y[3])
    data_y5[tiempo]=float(y[4])
    data_y6[tiempo]=float(y[5])
    data_y7[tiempo]=float(y[6])
    data_y8[tiempo]=float(y[7])
    data_y9[tiempo]=float(y[8])
 

    line1.set_data(data_time, data_y1)
    line2.set_data(data_time, data_y2)
    line3.set_data(data_time, data_y3)
    line4.set_data(data_time, data_y4)
    line5.set_data(data_time, data_y5)
    line6.set_data(data_time, data_y6)
    line7.set_data(data_time, data_y7)
    line8.set_data(data_time, data_y8)
    line9.set_data(data_time, data_y9)

    return line1,line2,line3,line4,line5,line6,line7,line8,line9,

def graficas():
    global xdata
    global ydata
    global line1,line2,line3,line4,line5,line6,line7,line8,line9
    global data_time
    global data_y1
    global data_y2
    global data_y3
    global data_y4
    global data_y5
    global data_y6
    global data_y7
    global data_y8
    global data_y9


    root.quit()     
    root.destroy()

    fig, ax = plt.subplots(3,3)

    line1, = ax[0][0].plot([], [])
    line2, = ax[0][1].plot([], [])
    line3, = ax[0][2].plot([], [])
    line4, = ax[1][0].plot([], [])
    line5, = ax[1][1].plot([], [])
    line6, = ax[1][2].plot([], [])
    line7, = ax[2][0].plot([], [])
    line8, = ax[2][1].plot([], [])
    line9, = ax[2][2].plot([], [])

    ax[0][0].set_ylim([0, 300])
    ax[0][0].set_xlim(0, COUNT)

    ax[0][1].set_ylim([0, 300])
    ax[0][1].set_xlim(0, COUNT)

    ax[0][2].set_ylim([0, 300])
    ax[0][2].set_xlim(0, COUNT)

    ax[1][0].set_ylim([0, 300])
    ax[1][0].set_xlim(0, COUNT)

    ax[1][1].set_ylim([0, 300])
    ax[1][1].set_xlim(0, COUNT)

    ax[1][2].set_ylim([0, 300])
    ax[1][2].set_xlim(0, COUNT)

    ax[2][0].set_ylim([0, 300])
    ax[2][0].set_xlim(0, COUNT)

    ax[2][1].set_ylim([0, 300])
    ax[2][1].set_xlim(0, COUNT)

    ax[2][2].set_ylim([0, 300])
    ax[2][2].set_xlim(0, COUNT)


    data_time = list(np.linspace(0,99,100))
    data_y1 = [0]*COUNT
    data_y2 = [0]*COUNT 
    data_y3 = [0]*COUNT
    data_y4 = [0]*COUNT 
    data_y5 = [0]*COUNT
    data_y6 = [0]*COUNT 
    data_y7 = [0]*COUNT
    data_y8 = [0]*COUNT 
    data_y9 = [0]*COUNT

    ao = animation.FuncAnimation(fig, update, next, blit = True, interval = 0,repeat = False)
    plt.show()
    GUI()

def GUI():
    global root, vo, vo1, vo2, vo3, vo4, vo5, vo6, vo7
    root = Tk()

    root.title('Control de iluminacion')
    miFrame=Frame(root, width=1000, height= 1000,background='#264362',bd=25)
    miFrame.grid(padx=20,pady=10)
    
    Title=Label(miFrame, text='Control de Iluminacion', fg='#4190E7',font=('Times',20),background='#1E354E').grid(row = 0, column = 2, pady=5, columnspan=3)

    Set_Point_Label=Label(miFrame, text='Set-Point:', fg='white', font=('Times',10),background='#1E354E').grid(row = 1, column = 0)

    vo = StringVar()

    Set_Point=Entry(miFrame,textvariable=vo,width=10).grid(row = 1, column = 1)

    vo.set('100')
    
    Unity_Label=Label(miFrame, text='(Luxes)', fg='white',background='#1E354E').grid(row = 1, column = 2, pady = 10)


    IniciarPID=Button(miFrame,text='Iniciar PID',command=activarPID,background='#1E354E').grid(row = 4, column = 0,pady=1)

    IniciarRep=Button(miFrame,text='Iniciar Replicador',command=activarReplicador,background='#1E354E').grid(row = 5, column = 0,pady=5)

    IniciarMPC=Button(miFrame,text='Iniciar MPC',command=activarMPC,background='#1E354E').grid(row = 6, column = 0,pady=5)


    Parar=Button(miFrame,text='Parar',command=desactivar,background='#1E354E', width=10, height=1).grid(row = 4, column = 2)

    IniciarGRAFICAS=Button(miFrame,text='Graficas',command=graficas,background='#1E354E').grid(row = 6, column = 2)

    conf_PID=Label(miFrame, text='Control PID', fg='white',background='#1E354E').grid(row = 1, column = 4, pady = 10, padx=100, columnspan = 2)

    Label_Kp=Label(miFrame, text='Constante kp', fg='black',background='#30547B').grid(row = 2, column = 4)

    vo1 = StringVar()
    Entry_Kp=Entry(miFrame,textvariable=vo1,width=5).grid(row = 2, column = 5)
    vo1.set('0.15')

    Label_Ki=Label(miFrame, text='Constante ki', fg='black',background='#30547B').grid(row = 3, column = 4)


    vo2 = StringVar()
    Entry_Ki=Entry(miFrame,textvariable=vo2,width=5).grid(row = 3, column = 5)
    vo2.set('0.01')


    Label_Kd=Label(miFrame, text='Constante kd', fg='black',background='#30547B').grid(row = 4, column = 4)

    vo3 = StringVar()
    Entry_Kd=Entry(miFrame,textvariable=vo3,width=5).grid(row = 4, column = 5)
    vo3.set('0')


    conf_REP=Label(miFrame, text='Dinamicas de Replicador', fg='white',background='#1E354E').grid(row = 5, column = 4, pady = 10, columnspan = 2)

    Label_alfa=Label(miFrame, text='Parametro alfa', fg='black',background='#30547B').grid(row = 6, column = 4)

    vo4 = StringVar()
    Entry_alfa=Entry(miFrame,textvariable=vo4,width=5).grid(row = 6, column = 5)
    vo4.set('0.1')

    Label_beta=Label(miFrame, text='Parametro beta', fg='black',background='#30547B').grid(row = 7, column = 4)

    vo5 = StringVar()
    Entry_beta=Entry(miFrame,textvariable=vo5,width=5).grid(row = 7, column = 5)
    vo5.set('0.44')

    conf_MPC=Label(miFrame, text='Control MPC', fg='white',background='#1E354E').grid(row = 1, column = 7, pady = 10, padx=70, columnspan = 2)

    Label_Np=Label(miFrame, text='Horizonte de prediccion', fg='black',background='#30547B').grid(row = 2, column = 7)

    vo6 = StringVar()
    Entry_Np=Entry(miFrame,textvariable=vo6,width=5).grid(row = 2, column = 8)
    vo6.set('10')

    Label_Nc=Label(miFrame, text='Horizonte de control', fg='black',background='#30547B').grid(row = 3, column = 7)

    vo7 = StringVar()
    Entry_Nc=Entry(miFrame,textvariable=vo7,width=5).grid(row = 3, column = 8)
    vo7.set('4')

    root.resizable(width=False, height=False)
    root.mainloop()

tiempo=0

def segundos(stop):
    global tiempo
    while True:
        if stop():
            break

        if tiempo==(COUNT-1):
            sleep(1)
            tiempo=0
        sleep(1)
        tiempo=tiempo+1



if __name__ == '__main__':

    stop_threads = False
    tmp1 = threading.Thread(target=PID, args=(lambda: stop_threads,))
    tmp2 = threading.Thread(target=REPLICADOR, args=(lambda: stop_threads,))
    tmp3 = threading.Thread(target=MPC, args=(lambda: stop_threads,))

    stop_threads_4 = False
    tmp4 = threading.Thread(target=segundos, args=(lambda: stop_threads_4,))

    tmp4.start()
    GUI()

    if tmp1.isAlive()==True:
        stop_threads = True
        tmp1.join()
        print('Fin PID.')
    if tmp2.isAlive()==True:
        stop_threads = True
        tmp2.join()
        print('Fin Replicador.')

    if tmp3.isAlive()==True:
        stop_threads = True
        tmp3.join()
        print('Fin MPC.')

    if tmp4.isAlive()==True:
        stop_threads_4 = True
        tmp4.join()
        print('Fin.')