# -*- coding: utf-8 -*-

#Import the necessary libraries for reading excel documents and ploting graphs
import matplotlib.pyplot as plt
import pandas as pd
import os

my_path = os.path.dirname(__file__)
patients = {}
#Function to separate patient information using the visit as split factor
def check_visit(visit,sample_id ,patients, row, test_col, table):
    sample_id = table[test_col[0]][row].split(visit)
    patients['ID'].append(sample_id[0])
    patients['VISIT'].append(visit)
    patients['DILUTION'].append(sample_id[1]) 
    for col in range(1,len(test_col)):
            patients[table[test_col[col]][0]].append(str(table[test_col[col]][row]))
            
            
##############################################################################################            
#Function to read excel files and plot them
#then it returns a table ready to be stored in an xlsl file            
def read_excel(sheet_num):
    print "Generating graphs for Plate number " ,sheet_num+1, "..."
    table = pd.read_excel('06222016-Staph-Array-Data.xlsx',sheetname=sheet_num)
    patients = {}
    patients['ID'] = []
    patients['VISIT'] = []
    patients['DILUTION'] = []
    test_col = table.columns.values
    for column in range(1,len(test_col)):
        if table[test_col[column]][0] in patients:
            table[test_col[column]][0] = " " + table[test_col[column]][0] 
        patients[table[test_col[column]][0]] = []
    #sample_id2 = table[test_col[0]][23].split()
    for row in range(1,len(table[test_col[0]])):
        #print "Check row number :", row
        sample_id = table[test_col[0]][row].split()
        test_length = len(sample_id)
        #simple if-else statements that separate IDs, Visits and Dilutions
        #we feel the unknown data with NaN
        if test_length > 1: 
            if ' V1 ' in table[test_col[0]][row]:
                check_visit(' V1 ',sample_id ,patients, row, test_col, table)
            elif ' v1 ' in table[test_col[0]][row]:
                check_visit(' v1 ',sample_id ,patients, row, test_col, table)
            elif ' V2' in table[test_col[0]][row]:
                check_visit(' V2',sample_id ,patients, row, test_col, table)
            elif ' v2 ' in table[test_col[0]][row]:
                check_visit(' v2 ',sample_id ,patients, row, test_col, table)
            elif ' V3 ' in table[test_col[0]][row]:
                check_visit(' V3 ',sample_id ,patients, row, test_col, table)
            elif ' v3 ' in table[test_col[0]][row]:
                check_visit(' v3 ',sample_id ,patients, row, test_col, table)
            else:
                patients['ID'].append(" ".join(sample_id[:(len(sample_id) - 1)]))
                patients['VISIT'].append('NaN')
                patients['DILUTION'].append(sample_id[len(sample_id) - 1])
                for col in range(1,len(test_col)):
                    #if len(patients[table[test_col[col]][0]]) != 47:
                        patients[table[test_col[col]][0]].append(str(table[test_col[col]][row]))     
        else:
                patients['ID'].append(sample_id[0])
                patients['VISIT'].append('NaN')
                patients['DILUTION'].append('NaN')
                for col in range(1,len(test_col)):
                    #if len(patients[table[test_col[col]][0]]) != 47:
                        patients[table[test_col[col]][0]].append(str(table[test_col[col]][row])) 

    #from this point onwards are calculations for plotting and saving the graphs
    pat_id = patients['ID'][0]
    id_ind = 0
    #loop for all the patients IDs
    for x in range(1,len(patients['ID'])):
        #separate patients by their IDs
        if pat_id != patients['ID'][x] or x == (len(patients['ID'])-1):
            #print id_ind,x
            if patients['DILUTION'][id_ind] != 'NaN':
                for prot in range(0,len(patients.keys())):
                    plt.clf()
                    if patients.keys()[prot] != 'ID' and\
                        patients.keys()[prot] != 'VISIT' and\
                        patients.keys()[prot] != 'DILUTION' and\
                        patients.keys()[prot] != 'Age' and\
                        patients.keys()[prot] != 'Gender' and\
                        patients.keys()[prot] != 'Hospital ' and\
                        patients[patients.keys()[prot]][0] != 'nan':
                            gender = 'nan'
                            age = 'nan'
                            hospital = 'nan'
                            if 'Gender' in patients.keys():
                                gender = patients['Gender'][id_ind]
                            if 'Age' in patients.keys():
                                age = patients['Age'][id_ind]
                            if 'Hospital ' in patients.keys():
                                hospital = patients['Hospital '][id_ind]
                            #print 'We got protein :',patients.keys()[prot]
                            vis_id = id_ind
                            pat_vis = patients['VISIT'][id_ind]
                            for_legend = [pat_vis]
                            
                            for y in range(id_ind,x):
                                #print 'We got visit :',pat_vis
                                if pat_vis != patients['VISIT'][y]:
                                    #TO DO FOR LOOP
                                    #print 'INSIDE IF STATEMENT!!!'
                                    plt.plot(patients['DILUTION'][vis_id:y],patients[patients.keys()[prot]][vis_id:y],'o',linestyle='-')
                                    #print patients[patients.keys()[prot]][vis_id:y]
                                    #print patients['DILUTION'][vis_id:y]
                                    vis_id = y
                                    pat_vis = patients['VISIT'][y]
                                    for_legend.append(pat_vis)
                                if y == (x-1):
                                    plt.plot(patients['DILUTION'][vis_id:y+1],patients[patients.keys()[prot]][vis_id:y+1],'o',linestyle='-')

                                    vis_id = y
                                    pat_vis = patients['VISIT'][y]
                                    for_legend.append(pat_vis)        
                                    
                            plt.yscale('log')
                            plt.xscale('log')
                            plt.ylabel('Intensity')
                            plt.xlabel('Dilution')
                            plt.title(str(pat_id)+'('+gender+" "+age+" yr "+hospital+")"+patients.keys()[prot])
                            plt.legend(for_legend)
                            output_dir1 = "Plate_"+str(sheet_num+1)
                            if not os.path.isdir(output_dir1): os.makedirs(output_dir1)
                            output_dir2 = str(pat_id)
                            final_out = os.path.join(output_dir1, output_dir2)
                            if not os.path.isdir(final_out): os.makedirs(final_out)
                            path_to_save = os.path.join(my_path,final_out)
                            my_file = str(patients.keys()[prot])+'.png'
                            plt.savefig(os.path.join(path_to_save, my_file))
                            
                    
                
                #move to the nexct id number
            pat_id = patients['ID'][x]
            id_ind = x
                #if patients['DILUTION'][x] != 'NaN':
                
            
    return patients
   # return patients
#################################################################################    
    
    
                        
writer = pd.ExcelWriter('Scores.xlsx')   
#read excel files, 'i' represents the number of the plate
#in the given xlsl    
for i in range(0,11):
    patients = read_excel(i)
    f_table =  pd.DataFrame(patients)
    plate_num = 'Plate '+str(i+1)
    f_table.to_excel(writer, plate_num)
    patients
#save our new structure of xlsl    
writer.save()




