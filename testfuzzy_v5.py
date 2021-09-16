import xlrd
import xlsxwriter
import os
from fuzzywuzzy import process
from fuzzywuzzy import fuzz
import math
import re


def reduce(dic):
    new_dic = {}
    for key,val in dic.items():
        if val!=1:
            new_dic.update({key:dic[key]})
    
    return new_dic


def get_size(code):
    reg = "\d+[\s\`\"]*x\s*\d+[\`\"]*|\d+[\s\`\"]*X\s*\d+[\`\"]*"
    lst = re.findall(reg,code)
    if len(lst) == 0:
        return ''
    else:
        lst = lst[0]
        lst = lst.replace('x', 'X')
        lst = lst.replace(' ', '')
        lst = lst.replace('`', '')
        return lst


def get_matches(element, list1, list2):
    matches = []
    if element in list1:
        for e in list1:
            if e==element:
                matches.append(list2[list1.index(e)])
        
        return matches
    elif element in list2:
        for e in list2:
            if e==element:
                matches.append(list1[list2.index(e)])
        
        return matches
    else:
        return matches


# filter_1 -> dimensions
def filter_1(elem,lst,size1,size2):
    elem_size = get_size(elem)
    matching_sizes = []
    
    if elem_size in size1:
        for s in size1:
            if s in size2:
                matching_sizes.append( s )
        
        matches = []
        for l in lst:
            l_size = get_size(l)
            if l_size in matching_sizes:
                matches.append( l )
        
        return matches
    
    else:
        #print('No matching sizes in config file.')
        return 0


# filter_2 -> matching words
def filter_2(elem,lst):
    elem_size = get_size(elem)
    elem_split = elem.replace(elem_size,'')
    elem_split = elem_split.split(' ')
    matches = []
    
    for e in elem_split:
        for l in lst:
            ratio = fuzz.partial_ratio(e, l)
            if (ratio>80) and (l not in matches):
                matches.append(l)
    
    if len(matches)==0:
        #print('No partial matches found.')
        return 0
    else:
        return matches


# Logic
def match(source, target, size1, size2):
    sources = []
    matches = []
    del source['NA']
    for key,val in source.items():
        targ = list( target.keys() )
        #print('Working on '+key)
        options_1 = filter_1(key,targ,size1,size2)
        options_2 = []
        match = []
        
        if options_1!=0:
            options_2 = filter_2(key,options_1)
            if options_2!=0:
                match = process.extract(key, options_2, limit=1)
            else:
                match = process.extract(key, options_1, limit=1)
        else:
            options_2 = filter_2(key,targ)
            if options_2!=0:
                match = process.extract(key, options_2, limit=1)
            else:
                match = process.extract(key, targ, limit=1)
        
        if len(match)==0:
            matches.append( 'No matches found.' )
            sources.append( key )
        else:
            matches.append( match[0][0] )
            sources.append( key )
            
            target = reduce(target)
        
        #print('Match is '+str(matches[-1]))
    print('Matching done.')    
    
    return sources,matches


# Read configuration file
def read_conf(filename):
    size1 = []
    size2 = []
    path = os.getcwd()+'/'+filename
    wb = xlrd.open_workbook(path)
    sheet = wb.sheet_by_index(0)
    sheet.cell_value(0, 0)
    for i in range(1, sheet.nrows):
        size1.append( sheet.cell_value(i, 0) )
        size2.append( sheet.cell_value(i, 1) )
    print('Conf read complete.')
    
    return size1,size2


# write file
def write(source, target, sku):
    workbook = xlsxwriter.Workbook('./Result_v5.xlsx')
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True})
    worksheet.write(0, 0, "SKU code", bold)
    worksheet.write(0, 1, "Distributed Item Name", bold)
    worksheet.write(0, 2, "Company Description", bold)
    for i in range(0,len(target)):
        #if source[i] in sku.keys():
        worksheet.write(i+1, 0, sku[ target[i] ])
        worksheet.write(i+1, 1, source[i])
        worksheet.write(i+1, 2, target[i])
    workbook.close()
    print('Result file written.')


# Read SKU file
comp_desc = {}
dist_desc = {}
sku = {}
path = os.getcwd()+'/SKU Mapping_for_v5.xlsx'
wb = xlrd.open_workbook(path)
sheet = wb.sheet_by_index(0)
sheet.cell_value(0, 0)
for i in range(1, sheet.nrows):
    com = sheet.cell_value(i, 2)
    comp_desc.update( {com : 0} )
    sku.update( {com : sheet.cell_value(i, 1)} )
    com = sheet.cell_value(i, 0)
    dist_desc.update( {com : 0} )

print('File read complete.')
print('No. of source(company codes) : '+str(len(comp_desc)))
print('No. of targets(distributor codes) : '+str(len(dist_desc)))
print('No. of SKUs : '+str(len(sku)))

# main program
size1,size2 = read_conf('Size Configurations.xlsx')
sources,matches = match(comp_desc, dist_desc, size1, size2)
write(matches,sources,sku)
