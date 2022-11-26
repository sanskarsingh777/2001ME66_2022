#Help https://youtu.be/H37f_x4wAC0
from operator import length_hint
import pandas as pd
import math

def octant_longest_subsequence_count_with_range():

    df = pd.read_excel(r'C:\Users\USER\Documents\github\2001ME66_2022\tut04\input_octant_longest_subsequence_with_range.xlsx')

    df["v_"] = df["V"]-df["V"].mean()
    df["w_"] = df["W"]-df["W"].mean()
    df["u_"] = df["U"]-df["U"].mean()
    
    df.loc[0, "U_Avg"] = df["U"].mean()
    df.loc[0, "W_Avg"] = df["W"].mean()
    df.loc[0, "V_Avg"] = df["V"].mean()


    df.loc[((df.u_ > 0) & (df.v_ > 0) & (df.w_ > 0)), "Octant"] = "+1"
    df.loc[((df.u_ < 0) & (df.v_ > 0) & (df.w_ < 0)), "Octant"] = "-2"
    df.loc[((df.u_ < 0) & (df.v_ < 0) & (df.w_ > 0)), "Octant"] = "+3"
    df.loc[((df.u_ < 0) & (df.v_ < 0) & (df.w_ < 0)), "Octant"] = "-3"
    df.loc[((df.u_ > 0) & (df.v_ < 0) & (df.w_ > 0)), "Octant"] = "+4"
    df.loc[((df.u_ > 0) & (df.v_ > 0) & (df.w_ < 0)), "Octant"] = "-1"
    df.loc[((df.u_ < 0) & (df.v_ > 0) & (df.w_ > 0)), "Octant"] = "+2"
   

    df.loc[((df.u_ > 0) & (df.v_ < 0) & (df.w_ < 0)), "Octant"] = "-4"


    max_len = len(df)

    OctantCount = [0, 0, 0, 0, 0, 0, 0, 0]
    for i in range(max_len):
        str = df["Octant"][i]
        if str == "+1":
            OctantCount[0] += 1
        elif str == "+4":
            OctantCount[6] += 1
        elif str == "-4":
            OctantCount[7] += 1
        elif str == "+2":
            OctantCount[2] += 1
        elif str == "-2":
            OctantCount[3] += 1
        elif str == "-3":
            OctantCount[5] += 1
        elif str == "+3":
            OctantCount[4] += 1
        elif str == "-1":
            OctantCount[1] += 1
        



    Random = [1, 1, 1, 1, 1, 1, 1, 1]
    length_OfLongestSubsequ = [0, 0, 0, 0, 0, 0, 0, 0]

    for i in range(1, max_len):
        str = df["Octant"][i-1]
        if df["Octant"][i] == df["Octant"][i-1]:
            if str == "+1":
                Random[0] += 1
            elif str == "+3":
                Random[4] += 1
            elif str == "-1":
                Random[1] += 1
            elif str == "+4":
                Random[6] += 1
            elif str == "-4":
                Random[7] += 1
            elif str == "+2":
                Random[2] += 1
            elif str == "-2":
                Random[3] += 1
            elif str == "-3":
                Random[5] += 1
            
        else:
            for j in range(8):
                length_OfLongestSubsequ[j] = max(length_OfLongestSubsequ[j], Random[j])

            Random = [1, 1, 1, 1, 1, 1, 1, 1]


    count_length_OfLongestSubsequ = [0, 0, 0, 0, 0, 0, 0, 0]
    Random = [1, 1, 1, 1, 1, 1, 1, 1]


    for i in range(1, max_len):
        str = df["Octant"][i-1]
        if df["Octant"][i] == df["Octant"][i-1]:
            if str == "+1":
                Random[0] += 1
            elif str == "-1":
                Random[1] += 1
            elif str == "+2":
                Random[2] += 1
            elif str == "-2":
                Random[3] += 1
            elif str == "+3":
                Random[4] += 1
            elif str == "-3":
                Random[5] += 1
            elif str == "+4":
                Random[6] += 1
            elif str == "-4":
                Random[7] += 1
        else:
            for j in range(8):
                if Random[j] == length_OfLongestSubsequ[j] and Random[j] == 1:
                    count_length_OfLongestSubsequ[j] = OctantCount[j]
                elif Random[j] == length_OfLongestSubsequ[j] and Random[j] != 1:
                    count_length_OfLongestSubsequ[j] += 1
                
            Random = [1, 1, 1, 1, 1, 1, 1, 1]


    df["Count"] = ""
    df["Longest Subsequence Length"] = ""
    df["Count2"] = ""


    df.loc[4, "Count"] = "+3"
    df.loc[0, "Count"] = "+1"
    df.loc[1, "Count"] = "-1"
    df.loc[5, "Count"] = "-3"
    df.loc[6, "Count"] = "+4"
    df.loc[7, "Count"] = "-4"
    df.loc[2, "Count"] = "+2"
    df.loc[3, "Count"] = "-2"
    
    

    for i in range(8):
        df.loc[i, "Count2"] = count_length_OfLongestSubsequ[i]
        df.loc[i, "Longest Subsequence Length"] = length_OfLongestSubsequ[i]
        


    df["Count3"] = ""
    df["Longest Subsequence Length2"] = ""
    df["Count4"] = ""

    n = 0
    for i in range(8):
        df.loc[i+n, "Count3"] = df.loc[i, "Count"]
        df.loc[i+n, "Longest Subsequence Length2"] = df.loc[i,
                                                                "Longest Subsequence Length"]
        df.loc[i+n, "Count4"] = df.loc[i, "Count2"]
        df.loc[i+n+1, "Count3"] = "Time"
        df.loc[i+n+1, "Longest Subsequence Length2"] = "From"
        df.loc[i+n+1, "Count4"] = "To"
        n = n+count_length_OfLongestSubsequ[i]+1


    Random = [1, 1, 1, 1, 1, 1, 1, 1]
    startlist = []
    endlist = []

    elist = [0, 0, 0, 0, 0, 0, 0, 0]
    slist = [0, 0, 0, 0, 0, 0, 0, 0]
    

    for i in range(1, max_len):
        str = df["Octant"][i-1]
        if df["Octant"][i] == df["Octant"][i-1]:
            if str == "+1":
                Random[0] += 1
                elist[0] = i
                slist[0] = i-length_OfLongestSubsequ[0]+1
            elif str == "-1":
                Random[1] += 1
                elist[1] = i
                slist[1] = i-length_OfLongestSubsequ[1]+1
            elif str == "+3":
                Random[4] += 1
                elist[4] = i
                slist[4] = i-length_OfLongestSubsequ[4]+1
            
            elif str == "-4":
                Random[7] += 1
                elist[7] = i
                slist[7] = i-length_OfLongestSubsequ[7]+1
            elif str == "+2":
                Random[2] += 1
                elist[2] = i
                slist[2] = i-length_OfLongestSubsequ[2]+1
            elif str == "-3":
                Random[5] += 1
                elist[5] = i
                slist[5] = i-length_OfLongestSubsequ[5]+1
            elif str == "+4":
                Random[6] += 1
                elist[6] = i
                slist[6] = i-length_OfLongestSubsequ[6]+1
            elif str == "-2":
                Random[3] += 1
                elist[3] = i
                slist[3] = i-length_OfLongestSubsequ[3]+1
            
        else:
            for j in range(8):
                if Random[j] == length_OfLongestSubsequ[j] and Random[j] != 1:
                    endlist.append(df.loc[elist[j], "Time"])
                    startlist.append(df.loc[slist[j], "Time"])
            Random = [1, 1, 1, 1, 1, 1, 1, 1]


    n = 2
    x = 0
    for i in range(8):
        l = count_length_OfLongestSubsequ[i]
        for j in range(l):
            df.loc[n+j, "Count4"] = endlist[x+j]
            df.loc[n+j, "Longest Subsequence Length2"] = startlist[x+j]
            
        x = x+l
        n = n + count_length_OfLongestSubsequ[i]+2


    df.to_excel(r"C:\Users\USER\Documents\github\2001ME66_2022\tut04\output_octant_longest_subsequence_with_range.xlsx")


from platform import python_version
ver = python_version()

if ver == "3.8.10":
    print("Correct Version Installed")
else:
    print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")


octant_longest_subsequence_count_with_range()
