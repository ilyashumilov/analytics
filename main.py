import pandas as pd
#Shipments
d1 = pd.read_excel('1.xlsx',sheet_name='CH')
d2 = pd.read_excel('1.xlsx',sheet_name='GR')
d3 = pd.read_excel('1.xlsx',sheet_name='US')

#SO
d4 = pd.read_excel('2.xlsx')
# Read the values of the file in the dataframe
df1 = pd.DataFrame(d1, columns=['DELIVERY DISPLAY NR','Quant'])
df2 = pd.DataFrame(d2, columns=['DELIVERY DISPLAY NR','Quant'])
df3 = pd.DataFrame(d3, columns=['DELIVERY DISPLAY NR','Quant'])

df4 = pd.DataFrame(d4, columns=['Allocated Quantity','Contract','Price','Currency','Purchased','POD/POL','Third Name'])
print(df4)
frames = [df1, df2, df3]
result = pd.concat(frames).reset_index(drop=True)
# print(result)
# Print the contents
rt = {}
cnt = 2
so=None
po=None
# fin = ['Number','MT','Price']
fin = []

for i in df4['Contract']:
    try:
        if len(i) == 7:
            if so != None:
                rt[so]=po
            # print(i, ' ', df4['Allocated Quantity'][cnt])
            so = i
            fin.append([i,df4['Allocated Quantity'][cnt],df4['Price'][cnt-2],df4['Currency'][cnt-2],df4['POD/POL'][cnt-2],df4['Third Name'][cnt-2]])
            po = []
            # print(i)
        elif (len(i) == 11):
            fin.append([i,df4['Purchased'][cnt-1],df4['Price'][cnt-1],df4['Currency'][cnt-1],df4['POD/POL'][cnt-2],df4['Third Name'][cnt-2]])
            l = []
            c = 0
            for n in result['DELIVERY DISPLAY NR']:
            #     # print(n)
                if n[:11] == i:
                    fin.append([n,str(result['Quant'][c])[:3]+','+str(result['Quant'][c])[3:],'',''])
                c += 1
            # po.append({i:l})
    except Exception as e:
        pass
        # print(e)
    cnt += 1
final = pd.DataFrame(fin,columns = ['Number', 'MT', 'Price', 'Currency','POL/POD','Third Name'])
final.to_excel("output.xlsx")
print(final)
# for i in fin:
#     print(i)

 # and i[:7] == so and so != None