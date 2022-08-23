import pandas as pd
import math

partDf = pd.read_excel('partData.xlsx')
newpartDf = partDf[['Maximum_stock_level', 'Total_Volume', 'Max', 'Gross_Weight(LB)']]
tot_vol = tot_vol1 = 0
max_inv = max_inv1 = 0
tempDf = []
tempDf1 = []
i = 0
storage_type = []
No_of_bins = []
max_dim = max_dim1 = 0
part_weight1 = part_weight = 0
total_weight: int = 0

for row in newpartDf.iterrows():
    max_inv1 = newpartDf['Maximum_stock_level']
    max_inv = max_inv1[i]

    tot_vol1 = newpartDf['Total_Volume']
    tot_vol = tot_vol1[i]
    binDf = pd.read_excel('binCategory.xlsx')
    newBinDf = pd.DataFrame()
    tot_Bin_Ser =[]
    parts_per_bin_Ser =[]
    pc_volume_Ser =[]
    Utilization_Ser =[]
    No_of_bin_Ser =[]
    factor_Ser =[]
    Bin_Weight_Ser = []
    total_bin_vol = pd.DataFrame()
    max_dim1 = newpartDf['Max']
    max_dim = max_dim1[i]
    part_weight1 = newpartDf['Gross_Weight(LB)']
    part_weight = part_weight1[i]
    j = 0
    x = y = 0
    newBinDf1 = []
    for index, row1 in binDf.iterrows():
        row1['No_of_Bin'] = tot_vol / row1['Bin_Volume']
        row1['No_of_Bin'] = math.ceil(row1['No_of_Bin'])
        No_of_bin_Ser.append(row1['No_of_Bin'])
        Bin_Weight_Ser.append(part_weight * (max_inv / row1['No_of_Bin']))
        tot_Bin_Ser.append(row1['No_of_Bin'] * row1['Bin_Volume'])  # Total bin volume
        parts_per_bin_Ser.append(max_inv / row1['No_of_Bin'])  # parts per bin
        pc_volume_Ser.append(tot_vol / max_inv)  # pc volume
        Utilization_Ser.append((max_inv / row1['No_of_Bin']) * (tot_vol / max_inv) / (row1['Bin_Volume']))  # utilization
        factor_Ser.append((row1['No_of_Bin']) * row1['No_of_Bin'] * row1['Bin_Volume'])  # no of bin* total bin volume

        # while Bin_Weight_Ser[0] > 25:
        #     No_of_bin_Ser[0]= row1['No_of_bin'].iloc[0] + 1
        #     parts_per_bin_Ser[0] = (max_inv / row1['No_of_bin'])
        #     tot_Bin_Ser[0] = (row1['No_of_bin'] * 232.9)
        #     pc_volume_Ser[0] = (tot_vol / max_inv)
        #     Utilization_Ser[0] = ((row1['parts_per_bin'] * row1['pc_volume'])/232.9)
            #newBinDf = newBinDf.append(row1)

        newBinDf = newBinDf.append(row1)
    newBinDf['No_of_bin'] = No_of_bin_Ser
    newBinDf['total_Bin_Volume'] = tot_Bin_Ser
    newBinDf['parts_per_bin'] = parts_per_bin_Ser
    newBinDf['pc_volume'] = pc_volume_Ser
    newBinDf['Utilization'] = Utilization_Ser
    newBinDf['Factor'] = factor_Ser
    newBinDf['Bin_Weight'] = Bin_Weight_Ser

    # Red Bin weight criteria to be checked

    # for index, row3 in newBinDf.iterrows():
    #     if (row3['Storage_Type']== "Plastic Red Bin"):
    #     #newBinDf['Bin Weight'] = part_weight * newBinDf['parts_per_bin']
    #         while row3['Bin_Weight'] > 25:
    #             No_of_bin_Ser.append(row3['No_of_bin'] +1)
    #             parts_per_bin_Ser.append(max_inv / row3['No_of_bin'])
    #             tot_Bin_Ser.append(row3['No_of_bin'] * 232.9)
    #             pc_volume_Ser.append(tot_vol / max_inv)
    #             Utilization_Ser.append((row3['parts_per_bin'] * row3['pc_volume'] )/232.9)
    #             newBinDf = newBinDf.append(row3)
    #
    # newBinDf['No_of_bin'] = No_of_bin_Ser
    # newBinDf['total_Bin_Volume'] = tot_Bin_Ser
    # newBinDf['parts_per_bin'] = parts_per_bin_Ser
    # newBinDf['pc_volume'] = pc_volume_Ser
    # newBinDf['Utilization'] = Utilization_Ser

    # sorting the new dataframe
    newBinDf.sort_values(by=['Factor', 'Utilization'], inplace=True, ascending=[True, False])

    # casehandling - getting lowest bin volume and then picking the storage type for same
    x = newBinDf['total_Bin_Volume'].iloc[0]
    y = newBinDf['No_of_bin'].iloc[0]
    storage_type = newBinDf['Storage_Type'].iloc[0]

    for index, row2 in newBinDf.iterrows():
        if row2['total_Bin_Volume'] < x and row2['No_of_bin'] <= 3:
            storage_type = row2['Storage_Type']


    # max dimension check
    if (storage_type == "Plastic Red Bin") and (max_dim >= 11.5):
        j = j + 1
        storage_type = newBinDf['Storage_Type'].iloc[j]
    if (storage_type == "Metal Green Bin") and (max_dim >= 15.0):
        j = j + 1
        storage_type = newBinDf['Storage_Type'].iloc[j + 1]
    if (storage_type == "Small Grey Tub") and (max_dim >= 31.0):
        j = j + 1
        storage_type = newBinDf['Storage_Type'].iloc[j + 1]
    if (storage_type == "Large Tub") and (max_dim >= 36):
        j = j + 1
        storage_type = newBinDf['Storage_Type'].iloc[j + 1]

    tempDf.append(storage_type)
    No_of_bins = newBinDf['No_of_bin'].iloc[0]
    tempDf1.append(No_of_bins)
    i = i + 1

partDf['bin'] = tempDf
partDf['No_of _bins'] = tempDf1
# partDf['No of Bin'] =
writer = pd.ExcelWriter("result.xlsx", engine='xlsxwriter')
partDf.to_excel(writer, index=False)
writer.save()
