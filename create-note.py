import random, os

people = int(input("Please enter total of employee: "))
cf = open("sales-report.txt", "w")
name=[]
state=[]
sales=[]
country = ["Johor","Kuala Lumpur","Melaka","Selangor","Kedah","Pahang","Perlis","Sarawak","Sabah","Perak","Penang","Terrengganu"]
for x in range (0,people):
	name.append("emp-name" + str(x))
	state.append(random.choice(country))
	sales.append(random.randint(1000,10000))
	cf.write(name[x]+"|"+state[x]+"|"+str(sales[x])+"|\n");	
cf.close()

of = open("sales-report.txt", "r")
print("\n\n\n************************************************************")
print("*************   Creating txt file .....  *******************")
print("************************************************************")
of.close()
print("\n\n\nSales-report.txt Successfull Created")
os.system("python createXcel.py")
