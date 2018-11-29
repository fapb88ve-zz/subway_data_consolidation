import pandas as pd
import numpy as np

def cat_reader():
    all_data = pd.DataFrame()
    for sheet in ['Food Items - 672', 'Packaging Items - 213',
                  'Beverages and Chips - 109', 'Cleaning Items - 83',
                 'Other']:
        df = pd.read_excel('sup_cat.xlsx', sheet_name = sheet)[['Super Category', 'Category']]
        df = df.fillna(method = 'ffill').drop_duplicates(subset = 'Category')
        all_data = all_data.append(df,ignore_index = True)

    return all_data

def cat_formatter(df):
    supCat = ['Food',
         'Beverages and Chips',
         'Packaging',
         'Cleaning']
    concept = ['MDP', 'SC', 'Auntie Annes', 'NonTrad', 'Walmart']
    cats = {}
    subCat = {}

    cat_read = cat_reader()
    for row in cat_read.iterrows():
        cats[row[1]['Category'].strip()] = row[1]['Super Category']

    sup_cat = []
    cat = []
    mm_list = []
    con = []

    for row in df.ItemName:
        name = [i.strip() for i in row.split(",")]
        if name[0] in supCat:
            sup_cat.append(name[0])
            con.append('Subway')
            if name[0] == 'Cleaning':
                cat.append(name[0])
                if name[1] == 'Cleaning': mm_list.append(True)
                else: mm_list.append(False)

            elif name[1] in cats:
                mm_list.append(True)
                cat.append(name[1])
            else:
                cat.append('NULL')
                mm_list.append(True)
        elif name[0] in cats:
            con.append('Subway')
            cat.append(name[0])
            sup_cat.append(cats[name[0]])
            mm_list.append(False)
        elif name[0] in concept:
            con.append(name[0])
            sup_cat.append('NULL')
            cat.append(name[1])
            mm_list.append(False)

        #Request from the boss
        elif "Mama Delucas" in name:
            con.append('MDP')
            sup_cat.append('NULL')
            cat.append("NULL")
            mm_list.append(False)

        else:
            sup_cat.append('NULL')
            cat.append('NULL')
            mm_list.append(False)
            con.append('NULL')

    df['SupCategory'] = sup_cat
    df['Category'] = cat
    df['Modified in MM 2.0'] = mm_list
    df['Concept'] = con
    return df

def col_format(df):
    fname = 'col_format.xlsx'

    ing = pd.read_excel(fname, sheet_name = 'Ingredients')[['IngredientId', 'IngredientName']]
    ing_dict = {}
    for row in ing.iterrows():
        ing_dict[row[1]['IngredientId']] = row[1]['IngredientName']
    df['IngredientName'] = df.IngredientId.map(ing_dict)

    status = pd.read_excel(fname, sheet_name = 'StatusTypes')[['StatusTypeId', 'Description']]
    status_dict = {}
    for row in status.iterrows():
        status_dict[row[1]['StatusTypeId']] = row[1]['Description']
    df['StatusTypeId'] = df.StatusTypeId.map(status_dict)

    delivery = pd.read_excel(fname, sheet_name = 'DeliveryUnits')[['DeliveryUnitTypeId', 'Description']]
    delivery_dict = {}
    for row in delivery.iterrows():
        delivery_dict[row[1]['DeliveryUnitTypeId']] = row[1]['Description']
    df['DeliveryUnitTypeId'] = df.DeliveryUnitTypeId.map(delivery_dict)

    pack = pd.read_excel(fname, sheet_name = 'PackDescTypes')[['PackDescriptionTypeId', 'Description']]
    pack_dict = {}
    for row in pack.iterrows():
        pack_dict[row[1]['PackDescriptionTypeId']] = row[1]['Description']
    df['PackDescriptionTypeId'] = df.PackDescriptionTypeId.map(pack_dict)

    packUOM = pd.read_excel(fname, sheet_name = 'PackUOMTypes')[['PackUOMTypeId', 'Description']]
    packUOM_dict = {}
    for row in packUOM.iterrows():
        packUOM_dict[row[1]['PackUOMTypeId']] = row[1]['Description']
    df['PackUOMTypeId'] = df.PackUOMTypeId.map(packUOM_dict)

    #MIGHT BE PROBLEMATIC
    portionUOM = pd.read_excel(fname, sheet_name = 'PortionUOMTypes')[['PortionUOMTypeId', 'Description']]
    portionUOM_dict = {}
    for row in portionUOM.iterrows():
        portionUOM_dict[row[1]['PortionUOMTypeId']] = row[1]['Description']
    df['PortionUOMOverrideId'] = df.PortionUOMOverrideId.map(portionUOM_dict)

    standardUOM = pd.read_excel(fname, sheet_name = 'StandardUOMTypes')[['StandardUOMTypeId', 'Description']]
    standardUOM_dict = {}
    for row in standardUOM.iterrows():
        standardUOM_dict[row[1]['StandardUOMTypeId']] = row[1]['Description']
    df['StandardUOMTypeId'] = df.StandardUOMTypeId.map(packUOM_dict)
    df['StandardUOMOverrideId'] = df.StandardUOMOverrideId.map(standardUOM_dict)

    portUOM_dict = {'Each': 'Each',
                    'Liter': 'Milliliter',
                    'Kilogram': 'Gram',
                    'Pound': 'Ounce',
                    'Gallon': 'Fluid Ounce',
                    'Milliliter': 'Milliliter',
                    'Gram': 'Gram',
                    'Ounce': 'Ounce',
                    'Fluid Ounce': 'Fluid Ounce'}
    df['PortionUOMTypeId'] = df.StandardUOMTypeId.map(portUOM_dict)

    df['StandardFinalUOM'] = df.StandardUOMOverrideId.combine_first(df.StandardUOMTypeId)

    df['PortionFinalUOM'] = df.PortionUOMOverrideId.combine_first(df.PortionUOMTypeId)


    return df

def region_describer():
    inv_hir = pd.read_excel('Inventory Hierarchy.xlsx', sheet_name = 'InventoryHierarchyAssignments')
    country =  pd.read_excel('Inventory Hierarchy.xlsx', sheet_name = 'Country')
    market = pd.read_excel('Inventory Hierarchy.xlsx', sheet_name = 'Market')
    region = pd.read_excel('Inventory Hierarchy.xlsx', sheet_name = 'Global Region')

    region_desc = {}
    country_desc = {}
    market_desc = {}

    for row in region.iterrows():
        region_desc[row[1][0]] = row[1]['Description']

    for row in country.iterrows():
        country_desc[row[1]['Id']] = row[1]['Description'].lower().title()

    for row in market.iterrows():
        market_desc[row[1]['Id']] = row[1]['Description'].lower().title()

    market_country = {}
    market_region = {}
    country_region = {}

    for row in market.iterrows():
        market_country[row[1]['Id']] = row[1]['CountryCodeId']

    for row in market.iterrows():
        market_region[row[1]['Id']] = row[1]['GlobalRegionId']

    for row in country.iterrows():
        country_region[row[1]['Id']] = row[1]['GlobalRegionId']


    geo = []

    for row in inv_hir.iterrows():
        typeid = row[1]['TypeId']
        itemid = row[1]['InventoryItemId']
        memberid = row[1]['MemberId']

        temp = [itemid]

        if typeid == 1:
            temp.extend([memberid, np.nan, np.nan])

        elif typeid == 2:
            temp.extend([country_region[memberid], memberid, np.nan])

        else:
            temp.extend([market_region[memberid], market_country[memberid], memberid])

        geo.append(temp)

    places = pd.DataFrame(data = [geo[0]])

    for row in geo[1:]:
        places = places.append(pd.Series(row, index = places.columns), ignore_index = True)

    places.columns = ['InventoryItemId', 'GlobalRegion', 'Country', 'Market']

    places['AccessLevel'] = inv_hir.TypeId

    places.AccessLevel = places.AccessLevel.map({1: "Region", 2:'Country', 3:'Market'})

    places.InventoryItemId = places.InventoryItemId.astype('int')

    places.GlobalRegion = places.GlobalRegion.map(region_desc)

    places.Country = places.Country.map(country_desc)

    places.Market = places.Market.map(market_desc)

    return places

def region_splitter(df):
    geo = []

    for row in df.InventoryItemId.unique():
        temp = [row]
        table = df[df.InventoryItemId == row]
        usa = []
        canada = []
        intl = []
        australia_nz = []
        europe = []
        latin_america = []
        middle_east = []
        asia = []
        for index, i in enumerate(table.iterrows()):
            if i[1]['AccessLevel'] == 'Region':
                if 'USA' in i[1]['GlobalRegion']:
                    usa.extend(['United States', 'Bahamas'])
                    canada.append('Canada')
                elif 'Asia' in i[1]['GlobalRegion']:
                    asia.append(i[1]['GlobalRegion'])
                elif 'Latin' in i[1]['GlobalRegion']:
                    latin_america.append(i[1]['GlobalRegion'])
                elif 'Middle' in i[1]['GlobalRegion']:
                    middle_east.append(i[1]['GlobalRegion'])
                elif 'Europe' in i[1]['GlobalRegion']:
                    europe.append(i[1]['GlobalRegion'])
                elif 'Australia' in i[1]['GlobalRegion']:
                    australia_nz.append(i[1]['GlobalRegion'])

            elif i[1]['AccessLevel'] == 'Country':
                if 'United States' in i[1]['Country']:
                    usa.append(i[1]['Country'])
                elif 'Canada' in i[1]['Country']:
                    canada.append(i[1]['Country'])
                elif 'Bahamas' in i[1]['Country']:
                    usa.append('Bahamas')
                elif 'Asia' in i[1]['GlobalRegion']:
                    asia.append(i[1]['Country'])
                elif 'Latin' in i[1]['GlobalRegion']:
                    latin_america.append(i[1]['Country'])
                elif 'Middle' in i[1]['GlobalRegion']:
                    middle_east.append(i[1]['Country'])
                elif 'Europe' in i[1]['GlobalRegion']:
                    europe.append(i[1]['Country'])
                elif 'Australia' in i[1]['GlobalRegion']:
                    australia_nz.append(i[1]['Country'])
            else:
                if 'United States' in i[1]['Country']:
                    usa.append(i[1]['Market'])
                elif 'Canada' in i[1]['Country']:
                    canada.append(i[1]['Market'])
                else:
                    if not i[1]['Market']:
                        if 'Bahamas' in i[1]['Country']:
                            usa.append('Bahamas')
                        elif 'Asia' in i[1]['GlobalRegion']:
                            asia.append(i[1]['Country'])
                        elif 'Latin' in i[1]['GlobalRegion']:
                            latin_america.append(i[1]['Country'])
                        elif 'Middle' in i[1]['GlobalRegion']:
                            middle_east.append(i[1]['Country'])
                        elif 'Europe' in i[1]['GlobalRegion']:
                            europe.append(i[1]['Country'])
                        elif 'Australia' in i[1]['GlobalRegion']:
                            australia_nz.append(i[1]['Country'])
                    else:
                        if 'Bahamas' in i[1]['Country']:
                            usa.append('Bahamas*')
                        elif 'Asia' in i[1]['GlobalRegion']:
                            asia.append(i[1]['Country']+"*")
                        elif 'Latin' in i[1]['GlobalRegion']:
                            latin_america.append(i[1]['Country']+"*")
                        elif 'Middle' in i[1]['GlobalRegion']:
                            middle_east.append(i[1]['Country']+"*")
                        elif 'Europe' in i[1]['GlobalRegion']:
                            europe.append(i[1]['Country']+"*")
                        elif 'Australia' in i[1]['GlobalRegion']:
                            australia_nz.append(i[1]['Country']+"*")

            if index == len(table)-1:
                temp.append(", ".join(usa))
                temp.append(", ".join(canada))
                temp.append(", ".join(asia))
                temp.append(", ".join(latin_america))
                temp.append(", ".join(middle_east))
                temp.append(", ".join(europe))
                temp.append(", ".join(australia_nz))
                geo.append(temp)

    final = pd.DataFrame(data = [geo[0]])
    for row in geo[1:]:
        final = final.append(pd.Series(row, index = final.columns), ignore_index = True)
    final.columns = ['InventoryItemId', 'MarketsUSARegion', 'MarketsCanadaRegion', 'AsiaRegion', 'LatinAmericaRegion', 'MiddleEastRegion', 'EuropeRegion', 'AustraliaNZRegion']
    final = pd.merge(df.iloc[:,0:7], final, on = "InventoryItemId")
    return final

def countByRegion(df):
    inv_hir = pd.read_excel('Inventory Hierarchy.xlsx', sheet_name = 'InventoryHierarchyAssignments')
    country =  pd.read_excel('Inventory Hierarchy.xlsx', sheet_name = 'Country')[['Id', 'GlobalRegionId']]
    market = pd.read_excel('Inventory Hierarchy.xlsx', sheet_name = 'Market')[['Id', 'GlobalRegionId']]

    countryid = {}
    for i in country.iterrows():
        countryid[i[1]['Id']] = i[1]['GlobalRegionId']

    marketid = {}
    for i in market.iterrows():
        marketid[i[1]['Id']] = i[1]['GlobalRegionId']

    regionid = []

    for row in inv_hir.iterrows():
        level = row[1]['TypeId']
        member = row[1]['MemberId']
        if level == 1:
            regionid.append(member)
        elif level == 2:
            regionid.append(countryid[member])
        else: regionid.append(marketid[member])

    inv_hir['GlobalRegionId'] = regionid
    countByRegion = inv_hir.groupby(['InventoryItemId', 'GlobalRegionId']).size().unstack()

    region =  pd.read_excel('Inventory Hierarchy.xlsx', sheet_name = 'Global Region')[['Id', 'Description']]
    describer = {}
    for row in region.iterrows():
        describer[row[1]['Id']] = row[1]['Description']

    countByRegion.columns = [describer[i] + " Count" for i in countByRegion.columns]

    return df.merge(countByRegion, on = 'InventoryItemId')

def main(file_name, df_output = True, file_output = False):

    df = pd.read_excel(file_name)

    colsToDrop = ['CreatedBy', 'DeletedUserName',
                  'TrackWaste', 'CreatedDT', 'CreatedBy',
                  'LastUpdateBy', 'DeletedBy', 'DeletedDT']
    df = df.drop(colsToDrop, axis = 1)
    df = cat_formatter(df)
    df = col_format(df)
    cols = ['InventoryItemId',
            'Concept',
            'SupCategory',
            'Category',
            'ItemName',
            'ItemShortDescription',
            'IngredientId',
            'IngredientName',
            'DeliveryUnitTypeId',
            'PackPerCase',
            'PackDescriptionTypeId',
            'PackSize',
            'PackUOMTypeId',
            'StandardUOMTypeId',
            'StandardUOMOverrideId',
            'StandardConversionFactor',
            'StandardFinalUOM',
            'PortionUOMTypeId',
            'PortionsPerCase',
            'PortionFinalUOM',
            'PortionUOMOverrideId',
            'PortionConversionFactor',
            'CaseCost',
            'PortionCost',
            'StatusTypeId',
            'Deleted',
            'CreatedUserName',
            'UpdatedUserName',
            'Modified in MM 2.0']
    df = df[cols]

    countRegion = countByRegion(df)

    df2 = df[['InventoryItemId',
            'Concept',
            'SupCategory',
            'Category',
            'ItemName',
            'ItemShortDescription']]
    places = region_describer()[['InventoryItemId', 'AccessLevel','GlobalRegion', 'Country', 'Market']]
    places = pd.merge(df2, places, on = 'InventoryItemId')

    final = region_splitter(places)

    if df_output and file_output:
        try:
            writer = pd.ExcelWriter('[Cleaned] ' + file_name)
            countRegion.to_excel(writer, sheet_name = 'InventoryItems', index = False)
            places.to_excel(writer, sheet_name = 'InventoryHierarchy', index = False)
            final.to_excel(writer, sheet_name = "ListCountries", index = False)
            print("Success!")
        except:
            print("File Already Exist in Folder")
        return df
    elif df_output and not file_output:
        return df
    elif not df_output and file_output:
        try:
            writer = pd.ExcelWriter('[Cleaned] ' + file_name)
            countRegion.to_excel(writer, sheet_name = 'InventoryItems', index = False)
            places.to_excel(writer, sheet_name = 'InventoryHierarchy', index = False)
            final.to_excel(writer, sheet_name = "ListCountries", index = False)
            writer.save()
            print("Success!")
        except:
            print("File Already Exist in Folder")
    else:
        return df
