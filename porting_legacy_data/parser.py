import xlsxwriter

option = "essais"

typeCodeNatureLength = {"A1":[100],"B1":[100], "B3":[2," cubes en beton, dimensions declares: ",27],"B4":[2," cylindres en beton, dimensions nominales: ",27],"B5":[2," carottes en beton ",50],"B6":[2," blocs en beton, ",30," dimensions nominales: ",2,"x",2,"x",2,"cm"], "B7":[2," tuyaux en beton ",50],"B8":[2," tuyaux en beton arm, ( Max. 10 par demande ) ",50],"B9":[2," briques ",39," dimensions nominales: ",3,"x",3,"x",3,"mm"]}
essaisColumns=[1,0,2,4,5,8,10,12,11,13,14,15,16,20,20,20,20,20,20]
essaistableheaders = ["ID","Type","Version","sortiLe","Demandeur","Payeur","EDemandeur","EPayeur","References","Quantity","Nature Du Produit", "Date De Reception", "Provenance", "Essais demandes", "Norme", "Remarques", "Technicien","Interlocuteur"]
clientstableheaders=["ID","Nom","Adresse","Autre","Remarques"]

if option == "clients":
	wb = xlsxwriter.Workbook('LegacyClients250918.xls')
	ws = wb.add_worksheet("Clients")
	myFile = open("CLIENT.DAT", "rb")
	tableheaders = clientstableheaders
	tablename = "clientsTable"
elif option == "essais":
	wb = xlsxwriter.Workbook('LegacyEssais250918.xls')
	myFile = open("DEMESS.DAT", "rb")
	ws = wb.add_worksheet("Essais")
	tableheaders = essaistableheaders
	tablename = "essaisTable"

myData = myFile.read().replace("\n","")
print len(myData)



cc=0
col=0
recordcount = 1
clientcount = 1

def nextSlice(chars):
	global cc
	oldcc = cc
	cc+=chars

	slicetext = myData[oldcc:cc]
	return removeProblems(slicetext).strip()
	
def removeProblems(string):
	returnstring = ""
	for c in string:
		try:
			returnstring += c.decode("cp850")#c.encode("utf-8")
		except UnicodeDecodeError:
			print "UnicodeError"
			returnstring += "?"
	return returnstring

def iterateFieldLengths(fieldLengths):
	global recordcount; global col
	for fieldLength in fieldLengths:
		ws.write(recordcount,essaisColumns[col],nextSlice(fieldLength))
		col+=1

def nextRecord():
	global cc
	global recordcount
	global col
	col=0
	recordType = nextSlice(2)
	cc+=50 #skip some nonsense
	#find start
	while myData[cc] == " ":
		cc+=1
		
	#start of record
	iterateFieldLengths([4,7,1,4,4,100])
	
	#nature du produit
	natureduproduit=""
	for field in typeCodeNatureLength[recordType]:
		try:
			field+=1; field-=1
			#is a fieldlength
			natureduproduit += nextSlice(field)
		except TypeError:
			#is some text
			natureduproduit += field
	
	ws.write(recordcount,essaisColumns[col],natureduproduit)
	col+=1

	#carry on
	iterateFieldLengths([100])
	
	#datedereception
	datedereception = nextSlice(2) + "/" + nextSlice(2) + "/" + nextSlice(2)
	ws.write(recordcount,essaisColumns[col],datedereception)
	col+=1
	
	#carry on
	iterateFieldLengths([100])
	
	#skip norme
	col+=1
	
	#remarques
	remarques = nextSlice(150) + " | " + nextSlice(150)
	ws.write(recordcount,essaisColumns[col],remarques)
	col+=1
	
	#carry on
	iterateFieldLengths([26])
			
	#done
	recordcount += 1
	#find next one
	while (myData[cc] != "A" and myData[cc] != "B"):
		cc+=1	

def nextClient():
	global cc
	global clientcount	
	#find start
	while myData[cc] == " ":
		cc+=1
	fieldLengths = [4,35,105,35,7,35,14,14,14,1,9,12,1,1,1,30,2,30]
	
	#start of record	
#	for (ind, fieldLength) in enumerate(fieldLengths):
#		ws.write(clientcount,clientColumns[ind],nextSlice(fieldLength))
	ws.write(clientcount,0, nextSlice(4)) #ID
	ws.write(clientcount,1, nextSlice(35) + ", " + nextSlice(105)) #cle and nom
	ws.write(clientcount,2, nextSlice(35) + " " + nextSlice(7) + " " + nextSlice(35)) #adresse, cpost, local
	ws.write(clientcount,3, "Tel: " + nextSlice(14) + nextSlice(14)[:0] + " ext: " + nextSlice(14))
	ws.write(clientcount,4, nextSlice(87))
		
	#done
	clientcount +=1
		

		
for i in range(10000):
	try:
		if option == "essais": 
			nextRecord()
		elif option == "clients":
			nextClient()
	except IndexError:
		#done
		break

headerlist = []
for i in tableheaders:
	headerlist += [{"header":i}]

if option == "essais": tablerows = recordcount
if option == "clients": tablerows = clientcount

ws.add_table(0,0,tablerows-1,len(tableheaders)-1,{"columns":headerlist,"name":tablename})
wb.close()
