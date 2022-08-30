# pip3 install pandas
# pip3 install openpyxl
# pip3 install Jinja2
import pandas as pd
import sys

def color(changed): 
    if len(changed) > 1:
        color = 'yellow'
        return 'background-color: %s' % color 

def stat(file):
    print("stat file: ", file)

    staticDict = {}

    df = pd.read_excel(file,sheet_name=0)

    for index, row in df.iterrows():
        if index == 0:
            continue

        storeId = str(df.loc[index, 'CREATE STORE'])
        group = df.loc[index, 'Group']
        # print(">>>", index, group, str(storeId))
        if not storeId or not group:
            print("no store:", index)
            continue

        if storeId not in staticDict:
            staticDict[storeId] = {}
        if group not in staticDict[storeId]:
            staticDict[storeId][group] = 0
        staticDict[storeId][group] += 1

    # write to new xlsx
    data = {
        'STORE': [],
        'Group 1': [],
        'Group 2': [],
        'Group 3': [],
        'Group Total': [],
        }

    for key in staticDict:
        # print(key, ":", staticDict[key])
        data['STORE'].append(key)
        total = 0
        for i in range (1, 4):
            if i in staticDict[key]:
                data['Group {}'.format(i)].append(staticDict[key][i])
                total += int(staticDict[key][i])
            else:
                data['Group {}'.format(i)].append(0)
        data['Group Total'].append(total)
    newDf = pd.DataFrame(data)
    newDf.to_excel('stat.xlsx', "Sheet1", index=False)
    print("OK")

def change(file):
    df = pd.read_excel(file)
    statDf = pd.read_excel("stat.xlsx")

    df.insert(df.shape[1], "oldSTORE", '')

    for index, statRow in statDf.iterrows():
        storeId = str(statRow['STORE'])
        group1 = str(statRow['Group 1'])
        group2 = str(statRow['Group 2'])
        group3 = str(statRow['Group 3'])
        toStoreId = str(statRow['toSTORE'])
        toGroup1 = str(statRow['toGroup1'])
        toGroup2 = str(statRow['toGroup2'])
        toGroup3 = str(statRow['toGroup3'])
        if toStoreId == 'nan':
            continue
        if toGroup1 == 'nan' and toGroup2 =='nan' and toGroup3 == 'nan':
            continue

        if toGroup1 != 'nan':
            toGroup1 = int(float(toGroup1))
        else:
            toGroup1 = 0

        if toGroup2 == 'nan':
            toGroup2 = 0
        else:
            toGroup2 = int(float(toGroup2))

        if toGroup3 == 'nan':
            toGroup3 = 0
        else:
            toGroup3 = int(float(toGroup3))

        for index, row in df.iterrows():
            nowStoreId = str(row['CREATE STORE'])
            group = row['Group']
            
            if nowStoreId == storeId:
                if int(group) == 1:
                    if toGroup1 > 0:
                        row['oldSTORE'] = row['CREATE STORE']
                        row['CREATE STORE'] = toStoreId
                        toGroup1 -= 1
                elif int(group) == 2:
                    if toGroup2 > 0:
                        row['oldSTORE'] = row['CREATE STORE']
                        row['CREATE STORE'] = toStoreId

                        toGroup2 -= 1
                elif int(group) == 3:
                    if toGroup3 > 0:
                        row['oldSTORE'] = row['CREATE STORE']
                        row['CREATE STORE'] = toStoreId
                        toGroup3 -= 1
            df.loc[index] = row
            
        writer = pd.ExcelWriter('output.xlsx')
        df.style.applymap(color, subset=['oldSTORE']).to_excel(writer, "Sheet1", index=False)
        writer.save()

        print("OK")

if __name__ == "__main__":
    # print(sys.argv)
    if len(sys.argv) <= 2:
        print("Usage: python main.py [stat|change] xxx.xlsx")
        sys.exit()

    if sys.argv[1] == 'stat':
        stat(sys.argv[2])
    if sys.argv[1] == 'change':
        change(sys.argv[2])


