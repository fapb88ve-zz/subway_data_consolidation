import pandas as pd
import numpy as np

def cat_reader():
    all_data = pd.DataFrame()
    for sheet in ['Food Items - 672', 'Packaging Items - 213', 
                  'Beverages and Chips - 109', 'Cleaning Items - 83',
                 'Other']:
        df = pd.read_excel('sup_cat.xlsx', sheet_name = sheet)[['Super Category', 'Category', 'SubCategory']]
        df = df.fillna(method = 'ffill').drop_duplicates(subset = 'SubCategory')
        all_data = all_data.append(df,ignore_index = True)
        
    return all_data

def cat_reader():
    all_data = pd.DataFrame()
    for sheet in ['Food Items - 672', 'Packaging Items - 213', 
                  'Beverages and Chips - 109', 'Cleaning Items - 83',
                 'Other']:
        df = pd.read_excel('sup_cat.xlsx', sheet_name = sheet)[['Super Category', 'Category', 'SubCategory']]
        df = df.fillna(method = 'ffill').drop_duplicates(subset = 'SubCategory')
        all_data = all_data.append(df,ignore_index = True)
        
    return all_data


def cat_formatter(df):
    supCat = ['Food', 
         'Beverages and Chips',  
         'Packaging',
         'Cleaning']
    cats = {}
    subCat = {}

    cat_read = cat_reader()
    for row in cat_read.iterrows():
        cats[row[1]['Category']] = row[1]['Super Category']
        

    sup_cat = []
    cat = []
    mm_list = []
    
    for row in df.ItemName:
        name = [i.strip() for i in row.split(",")]
        if name[0] in supCat:
            sup_cat.append(name[0])
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
            cat.append(name[0])
            sup_cat.append(cats[name[0]])
            mm_list.append(False)
        else:
            sup_cat.append('NULL')
            cat.append('NULL')
            mm_list.append(False)

    df['SupCategory'] = sup_cat
    df['Category'] = cat
    df['Modified in MM 2.0'] = mm_list
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
        
    countByRegion.columns = [describer[i] + " Count" for i in describer]
    
    return df.merge(countByRegion, on = 'InventoryItemId')



def main(file_name, indoor = False):
    
    df = pd.read_excel(file_name)
    
    colsToDrop = ['CreatedBy', 'DeletedUserName', 
                  'TrackWaste', 'CreatedDT', 'CreatedBy', 
                  'LastUpdateBy', 'DeletedBy', 'DeletedDT']
    df = df.drop(colsToDrop, axis = 1)
    df = cat_formatter(df)
    df = col_format(df)
    cols = ['InventoryItemId', 
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
    
    df = countByRegion(df)
    
    try:
        df.to_excel('[Cleaned] ' + file_name, index = False)
    except:
        print("File Already Exist in Folder")
    if not indoor:
        return 'Success'
    else: return df


    
    