import pandas as pd
df = pd.read_excel(r"C:\Users\USER\Documents\github\2001ME66_2022\tut05\octant_input.xlsx")
x = df.shape[0]
df.head(15)
#finding mean for every value
U_avg = df['U'].mean()
V_avg = df['V'].mean()
W_avg = df['W'].mean()

#making columns to store average value of U,V,W
df['U Avg']=U_avg
#and here I have given only 1st place of line to U avg otherwise average value will print in whole column
df['U Avg']=df['U Avg'].head(1)

#similarly for V and W,doing the same thing
df['V Avg']=V_avg
df['V Avg']=df['V Avg'].head(1)

df['W Avg']=V_avg
df['W Avg']=df['W Avg'].head(1)

df.head()
X = df['U'] - U_avg
Y = df['V'] - V_avg
Z = df['W'] - W_avg

#made column for storing X and named the column as U'=U - U_avg,similarily Y for V'=V - V_avg and Z for W'=W - W_avg.
df["U'=U - U_avg"] = X
df["V'=V - V_avg"] = Y
df["W'=W - W_avg"] = Z
df.head()
df.insert(10, column="Octant", value="")

#using loop
for i in range(0,x):
    M= df["U'=U - U_avg"][i]
    N= df["V'=V - V_avg"][i]
    O= df["W'=W - W_avg"][i]

#using loop
for i in range(0,x):
    M= df["U'=U - U_avg"][i]
    N= df["V'=V - V_avg"][i]
    O= df["W'=W - W_avg"][i]
    
    
    if M>0 and N>0 and O>0:
        print(1)
        df["Octant"][i] = 1
    elif M>0 and N>0 and O<0:
        print(-1)
        df["Octant"][i] =-1
    elif M<0 and N>0 and O>0:
        print(2)
        df["Octant"][i] =2
    elif M<0 and N>0 and O<0:
        print(-2)
        df["Octant"][i] =-2
    elif M<0 and N<0 and O>0:
        print(3)
        df["Octant"][i] =3
    elif M<0 and N<0 and O<0:
        print(-3)
        df["Octant"][i] =-3
    elif M>0 and N<0 and O>0:
        print(4)
        df["Octant"][i] =4
    elif M>0 and N<0 and O<0:
        print(-4)
        df["Octant"][i] =-4
df.at[1,''] = 'User Input'

df.at[0,'Octant ID'] = 'Overall Count'
df.at[1,''] = 'User Input'
mod = 5000
mod_max_value=30000
n=mod_max_value//mod
q_list=[]
for k in range(0,n+2):
    if(k==0):
        df.at[k,'Octant ID'] = 'Overall Count'
    elif(k==1):
        df.at[k,'Octant ID'] ="mod"+ " " +str(mod)
    elif(k==2):
        df.at[k,'Octant ID'] = str((k-2)*mod) +"-"+str((k-1)*mod-1)
    else:
        df.at[k,'Octant ID'] = str((k-2)*mod+1) +"-"+str((k-1)*mod-1)
range_value = int(2 + (mod_max_value/mod))
q_list = [1,-1,2,-2,3,-3,4,-4]
for j in q_list:
    df.at[0,j] = list(df['Octant']).count(j)
    for i in range(0,n):
        if(i==0):
            df.at[i+2,j] = list(df['Octant'][i*mod:(i+1)*mod-1]).count(j) 
        else :
            df.at[i+2,j] = list(df['Octant'][i*mod-1:(i+1)*mod-1]).count(j) 
for j in range(0,range_value):
    octant_name_id_mapping = {"1":"Internal outward interaction", "-1":"External outward interaction", "2":"External Ejection", "-2":"Internal Ejection", "3":"External inward interaction", "-3":"Internal inward interaction", "4":"Internal sweep", "-4":"External sweep"}
    octant_count = []
    for i in q_list:
        octant_count.append(df.loc[j,i]) 
    octant_count.sort(reverse=True)

    for i in q_list:
        for x in range(0,8):
            if(octant_count[x]==df.loc[j,i]):
                df.loc[j,"Rank("+str(i)+")"] = x+1
    
    for k in q_list:
        if(df.loc[j,"Rank("+str(k)+")"]==1):
            df.loc[j,"Rank1 Octant ID"] = k
            df.loc[j,"Rank1 Octant Name"]=octant_name_id_mapping[str(k)]
df = pd.concat([df.columns.to_frame().T, df], ignore_index=True)
df.to_excel(r"C:\Users\USER\Documents\github\2001ME66_2022\tut05\octant_output_ranking_excel.xlsx",index=False,header=None)

