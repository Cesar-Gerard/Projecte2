import pymongo
import pandas as pd
import os
from dotenv.main import load_dotenv
from datetime import * 
from datetime import timedelta as td

#CARREGUEM EL DOCUMENT DE PROPIETATS
load_dotenv()

user=os.environ['USER']
password=os.environ['PASSWORD']
cluster=os.environ['CLUSTER']
#//////////////////////////////////////////////////////////////////////////////////////////////////////////////////


#CONNECTEM AMB MONGODB
client = pymongo.MongoClient("mongodb+srv://"+user+":"+password+"@"+cluster+".swmytsn.mongodb.net/?retryWrites=true&w=majority")
db = client.gcesar

crear_colection_USUARIS= db.USUARIS
crear_colection_PACIENTS= db.PACIENTS
crear_colection_METGES= db.METGES
#//////////////////////////////////////////////////////////////////////////////////////////////////////


#LLEGIM EL DOCUMENT I PREPAREM LES COLECCIONS
p="Tasca3.xlsx"

#LLista de valors de el document
excel_data = pd.read_excel(p)
excel_horari=pd.read_excel(p,sheet_name="HORARIS")
excel_visites=pd.read_excel(p,sheet_name="VISITES")

#///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

#CREACIÓ DE LES CLASSES NECESSARIES Y LA SEVA CORRESPONENT LECTURA EN EL EXCEL

# Creem la classe HORARI per guardar els seus valors

class HORARI():
    def __init__(self,id_temporal,dilluns,dimarts,dimecres,dijous,divendres,dissabte,inici,fi):
        self.id_temporal=id_temporal
        self.dilluns=dilluns
        self.dimarts=dimarts
        self.dimecres=dimecres
        self.dijous=dijous
        self.divendres=divendres
        self.dissabte=dissabte
        self.inci=inici
        self.fi= fi


#Guardem les variables de HORARI en una llista
llista_horari= [(HORARI(row[0],row[1],row[2],row[3],row[4],row[5],row[6],row[7],row[8])) for index, row in excel_horari.iterrows() ] 
  


#Creem la classe persona per guardar els seus valors
class PERSONA ():
    def __init__(self,id_temporal,DNI,login,Sexe,Cognoms_i_Nom,Data_Naixement,Adreça,Poblacio,CP,Provincia,Pais,Mutua,Num_mutualista,Especialitat,Num_colegiat):
        self.id_temporal=id_temporal
        self.DNI= DNI
        self.login= login
        self.Sexe=Sexe
        self.Cognoms_i_Nom=Cognoms_i_Nom
        self.Data_Naixement=Data_Naixement
        self.Adreça=Adreça
        self.Poblacio=Poblacio
        self.CP=CP
        self.Provincia=Provincia
        self.Pais=Pais
        self.Mutua=Mutua
        self.Num_mutualista=Num_mutualista
        self.Especialitat=Especialitat
        self.Num_colegiat=Num_colegiat
        self.id = id
        self.rang=[]
        self.inici=0
        self.final=0
        
        


       
#Guardem les variables de PERSONA en una llista
llista_personas= [(PERSONA(row[0],row[1],row[2],row[3],row[4],row[5],row[6],row[7],row[8],row[9],row[10],row[11],row[12],row[13],row[14])) for index, row in excel_data.iterrows() ] 
    




#Creem la classe VISITA per guardar les vistes registrades

class VISITES():
    def __init__(self,id_metge,Moment_visita,id_pacient,Realitzada,Informe):
        self.id_metge=id_metge
        self.Moment_visita=Moment_visita
        self.id_pacient=id_pacient
        self.Realitzada=Realitzada
        self.Informe=Informe

#Guardem les variables de VISITA en una llista
llista_visites= [(VISITES(row[0],row[1],row[2],row[3],row[4])) for index, row in excel_visites.iterrows() ] 
  
#///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

#MÉTODES

def creacio_calendari():
   #Rang de dates que ens demanen
    start = datetime.strptime("2023-01-01T00:00:00.00Z", "%Y-%m-%dT%H:%M:%S.%fZ")
    date_generated = pd.date_range(start, periods=181)

    #Treiem els diumenges
    resultat = [fecha for fecha in date_generated if fecha.weekday() < 6]

    

    #Eliminem les dates que no entren en la agenda per no treballar
    for x in resultat:
        if x == datetime.strptime("2023-05-01T00:00:00.00Z", "%Y-%m-%dT%H:%M:%S.%fZ"):
            resultat.remove(x)

        if x == datetime.strptime("2023-06-24T00:00:00.00Z", "%Y-%m-%dT%H:%M:%S.%fZ"):
           resultat.remove(x)

        if x == datetime.strptime("2023-04-02T00:00:00.00Z", "%Y-%m-%dT%H:%M:%S.%fZ"):
           resultat.remove(x)

        if x == datetime.strptime("2023-04-03T00:00:00.00Z", "%Y-%m-%dT%H:%M:%S.%fZ"):
            resultat.remove(x)

        if x == datetime.strptime("2023-04-04T00:00:00.00Z", "%Y-%m-%dT%H:%M:%S.%fZ"):
            resultat.remove(x)

        if x == datetime.strptime("2023-04-05T00:00:00.00Z", "%Y-%m-%dT%H:%M:%S.%fZ"):
            resultat.remove(x)

        if x == datetime.strptime("2023-04-06T00:00:00.00Z", "%Y-%m-%dT%H:%M:%S.%fZ"):
            resultat.remove(x)

        if x == datetime.strptime("2023-04-07T00:00:00.00Z", "%Y-%m-%dT%H:%M:%S.%fZ"):
            resultat.remove(x)

        if x == datetime.strptime("2023-04-08T00:00:00.00Z", "%Y-%m-%dT%H:%M:%S.%fZ"):
            resultat.remove(x)

        if x == datetime.strptime("2023-04-09T00:00:00.00Z", "%Y-%m-%dT%H:%M:%S.%fZ"):
            resultat.remove(x)

        if x == datetime.strptime("2023-04-10T00:00:00.00Z", "%Y-%m-%dT%H:%M:%S.%fZ"):
            resultat.remove(x)

        if x == datetime.strptime("2023-04-11T00:00:00.00Z", "%Y-%m-%dT%H:%M:%S.%fZ"):
            resultat.remove(x)

    date=[]
    for r in resultat:
        for d in range(24):
            date.append(r+timedelta(hours=d,minutes=0))
            date.append(r+timedelta(hours=d,minutes=30))
            



    
    return date



def inserirdatacita(id,MV,pacient,realitzada,informe):
    crear_colection_METGES.update_one({"_id": id}, {
                    "$push": {
                    "Agenda": {
                        "$each": [{
                            "Moment_Temporal":MV,
                            "Pacient":pacient,
                            "Realitzada":realitzada,
                            "Informe":informe
                           
                    }],
                        "$position": 0
                    }      
                    }
                    })


def inserirdataNOcita(id):
    crear_colection_METGES.update_one({"_id": id}, {
                    "$push": {
                    "Agenda": {
                        "$each": [{
                           "Moment_Temporal":"",
                            "Pacient":"",
                            "Realitzada":"",
                            "Informe":""
                           
                    }],
                        
                    }      
                    }
                    })


#Transformo el tipus de dada de la llista_visites
def transformar_visites():
    for trans in llista_visites:
        trans.Moment_visita=datetime.strptime(trans.Moment_visita, "%Y-%m-%dT%H:%M:%S.%fZ")


#TReiem els dies dels metges que no treballen
def treure_dies_no_treballats():
    for p in llista_personas:

        if(p.inici!=0):
            for dias in llista_dies_laborals:
                if datetime.time(dias) >= p.inici and datetime.time(dias)<p.final:
                    for h in llista_horari:
                        if h.id_temporal == p.id_temporal:
                            if datetime.weekday(dias)==0 and h.dilluns=="s" :
                                p.rang.append(dias)

                            if datetime.weekday(dias)==1 and h.dimarts=="s" :
                                p.rang.append(dias)

                            if datetime.weekday(dias)==2 and h.dimecres=="s" :
                                   p.rang.append(dias)

                            if datetime.weekday(dias)==3 and h.dijous=="s" :
                                p.rang.append(dias)

                            if datetime.weekday(dias)==4 and h.divendres=="s" :
                              p.rang.append(dias)

                            if datetime.weekday(dias)==5 and h.dissabte=="s" :
                              p.rang.append(dias)

#Assignem a cada persona que treballa el seu horari de entrada y sortida
def donar_horari():
    for y in llista_horari:
       for m in llista_personas:
        if m.id_temporal==y.id_temporal:
            m.inici=y.inci
            m.final=y.fi

   
#///////////////////////////////////////////////////////////////////////////////////////




llista_dies_laborals=creacio_calendari()


         
donar_horari();        
        
    
treure_dies_no_treballats(); 


transformar_visites()  





#Inserció Colecció Usuaris amb les dades primordials
for x in llista_personas:

    llista_nom_cognom=(str) (x.Cognoms_i_Nom).split(',')
    name=llista_nom_cognom[1]
    last_name=llista_nom_cognom[0]

    mydict={
                "DNI":x.DNI,
                "Login":x.login,
                "Sexe":x.Sexe,
                "Cognom":last_name,
                "Nom":name,
                "Data_Naixement":x.Data_Naixement,
                "Adreça":{
                    "Direcció": x.Adreça, 
                    "Població":x.Poblacio,
                    "CP":x.CP,
                    "Província":x.Provincia,
                    "País":x.Pais
                }
               
            }
    id_object=crear_colection_USUARIS.insert_one(mydict)
    x.id=id_object.inserted_id
    
    #Inserció a la colecció Pacient amb les dades característiques

    if(pd.notna(x.Mutua)):
        mydict2={
            "_id":x.id,
            "Mutualitat":x.Mutua,
            "Num_mutualista":x.Num_mutualista,
            
        }

        crear_colection_PACIENTS.insert_one(mydict2)
    

    #Inserció en la colecció METGE amb les seves dades



    if (pd.notna(x.Num_colegiat)):
        
        
        mydict3={
            "_id":x.id,
            "Especialitat":x.Especialitat,
            "Num_colegiat":x.Num_colegiat,
            
        }

        crear_colection_METGES.insert_one(mydict3)

        
        for dias in x.rang:
            for v in llista_visites:
                if x.id_temporal == v.id_metge:
                    
                    if dias == v.Moment_visita:
                        inserirdatacita(x.id,v.Moment_visita,v.id_pacient,v.Realitzada,v.Informe)
                        
                    else:
                        inserirdataNOcita(x.id)
                        
        
        
        
            
                              
       