##########################
# Notes for helping to read and understand logic
##########################
# variables will be prefixed accordingly with the following method:
#   i = integer
#   s = string
#   b = boolean
#   d = dictionary
#   df = dataframe
#   a = array
#   l = lists
#   v = values (for loops)
##########################
# card Creation
# two while loops
# first to loop over the number of cards to create
# second is to loop over the number of rows required for each card
##########################
# Simulation Creation
# single while loop that creates a dictionary to hold x number of simulations desired
# and the proper range assigned
##########################
# Simulating over the cards
# two while loops, 5 for loops
# two while loops go over each bingo sheet and then loops over each simulation assessing results
# 5 for loops dynamically assign dictionary of column and row combinations
# assigns results to a dataframe for analysis
#########################
# Importation of packages
# dataframe is datatype for pandas. data that pandas can manipulate. python arrays are converted to dataframe in pandas.
# html2image converts dataframes to images. For our program, cards to displayed images.
#########################
import numpy as np
import random
import time
import copy
import pandas as pd
import statistics
import sys
import math
import matplotlib.pyplot as plt
import scipy.stats as stats
from scipy.stats import kurtosis
import dataframe_image as dfi
from html2image import Html2Image
import glob
from fpdf import FPDF
import seaborn as sns
import os


#########################
# Define functions
#########################
def split_into_parts(number, n_parts):
    return np.linspace(0, number, n_parts + 1)[1:]


def run_game(sUSERDECISION, iCARDS, iSIMULATION, iROWS, iCOLUMNS, iZEROS, iBEGINNINGRANGE, iENDINGRANGE):
    #################
    # User decides which program to run
    #################

    #########################
    # Generate bingo cards depending on user selection
    #########################

    if sUSERDECISION == "Standard":
        #########################
        # Declaring Variables
        #########################
        #### Number of Cards input

        iCARD_SIZE = 5
        iCARDS_COUNTER = 1
        start_time = time.time()
        dBINGO_SHEETS = {}

        #########################
        # While Loop - Bingo Board Creations
        #########################
        while iCARDS_COUNTER <= iCARDS:
            #########################
            # Variables, dictionary, arrays created within loop
            #########################
            sCARDS = str(iCARDS_COUNTER)
            dBINGO_SHEETS["Bingo_Sheet_{0}".format(iCARDS_COUNTER)] = sCARDS
            first_row = random.sample(range(1, 15, 1), k=iCARD_SIZE)
            second_row = random.sample(range(16, 30, 1), k=iCARD_SIZE)
            third_row = random.sample(range(31, 45, 1), k=iCARD_SIZE)
            fourth_row = random.sample(range(46, 60, 1), k=iCARD_SIZE)
            fifth_row = random.sample(range(61, 75, 1), k=iCARD_SIZE)

            #########################
            # Concatenate,transpose and reshape the arrays to fit the 5x5 requirement
            # Replace the middle number with 0 to represent a free value
            #########################
            dBINGO_SHEETS["Bingo_Sheet_{0}".format(iCARDS_COUNTER)] = np.concatenate(
                (first_row, second_row, third_row, fourth_row, fifth_row))
            dBINGO_SHEETS["Bingo_Sheet_{0}".format(iCARDS_COUNTER)] = np.transpose(
                dBINGO_SHEETS["Bingo_Sheet_{0}".format(iCARDS_COUNTER)])
            dBINGO_SHEETS["Bingo_Sheet_{0}".format(iCARDS_COUNTER)] = dBINGO_SHEETS[
                "Bingo_Sheet_{0}".format(iCARDS_COUNTER)].reshape(iCARD_SIZE, iCARD_SIZE)
            dBINGO_SHEETS["Bingo_Sheet_{0}".format(iCARDS_COUNTER)][2][2] = 0

            #########################
            # Print statement printing each card
            #########################
            # print("Bingo_Sheet_"+sCARDS)
            print(dBINGO_SHEETS["Bingo_Sheet_{0}".format(iCARDS_COUNTER)])

            #########################
            # Increment
            #########################
            iCARDS_COUNTER += 1

            #########################
        # Ending Timestamp
        #########################
        print("--- Run time is %s seconds ---" % (time.time() - start_time))
    elif sUSERDECISION == "Dynamic":
        ##########################
        # Print statements introducing user to the game
        #### These two variables are backups - place holders for when variables need to be reset
        iBEGINNINGRANGE_BACKUP = iBEGINNINGRANGE
        iENDINGRANGE_BACKUP = iENDINGRANGE
        #### Finds the numerical difference between starting and ending range
        iRANGE = (iENDINGRANGE - iBEGINNINGRANGE) + 1

        #### error handling for range inputs

        ####################################################
        #### Ending of input variables and error handling
        ####################################################
        ##########################
        # Declaring non-inputted variables
        # dictionaries, counters for loops, timestamps, etc.
        ##########################
        dCOORDINATES = {}
        dROWS = {}
        dBINGO_SHEETS = {}
        iCARDS_Counter = 1
        lCOORDINATES = [*range(iBEGINNINGRANGE, iROWS * iCOLUMNS + 1, 1)]
        iCOLUMNCOUNTER = 0
        iNAMINGCOUNTER = 1

        #########################
        # START THE CLOCK !!!!!!!
        #########################
        start_time = time.time()

        #########################
        # While Loops for random coordinates
        #########################
        while iCOLUMNCOUNTER < iCOLUMNS:
            iROWCOUNTER = 0
            while iROWCOUNTER < iROWS:
                dCOORDINATES["{0}_Coordinates".format(iNAMINGCOUNTER)] = (iCOLUMNCOUNTER, iROWCOUNTER)
                iNAMINGCOUNTER += 1
                iROWCOUNTER += 1
            iCOLUMNCOUNTER += 1
        lSELECTED = random.sample(lCOORDINATES, iZEROS)
        #########################
        # While Loops for Bingo Board Creations
        #########################
        # First loop determines how many cards to create
        while iCARDS_Counter <= iCARDS:

            #########################
            # Variables created within loop
            # Name our bingo sheet_1,2,3 ...  and assign to empty dictionary
            # resetting beginning of range for each bingo sheet
            #########################
            sCARDS = str(iCARDS_Counter)
            dBINGO_SHEETS["Bingo_Sheet_{0}".format(iCARDS_Counter)] = sCARDS
            iBEGINNINGRANGE = iBEGINNINGRANGE_BACKUP
            # end of variable declaration in loops
            #########################

            #########################
            # Dynamically creates arrays for each row
            # 1) if statement - Determines how to divide our range whether by the column or row number, whichever is greater
            # 2) if statement - handles for first row and then all remaining rows
            #########################
            iROWSCOUNTER = 0
            while iROWSCOUNTER < iROWS:
                if iCOLUMNS >= iROWS:
                    aSPLITRANGE = split_into_parts(iRANGE, iCOLUMNS)
                else:
                    aSPLITRANGE = split_into_parts(iRANGE, iROWS)
                # finds upper limit of row range
                iUPPERRANGE = int(math.floor(aSPLITRANGE[iROWSCOUNTER])) + 1
                # handles first row
                if iROWSCOUNTER == 0:#replaced iBEGINNINGRANGE with 0
                    dROWS["Row_{0}".format(iROWSCOUNTER)] = random.sample(range(1, iUPPERRANGE, 1),
                                                                          k=iCOLUMNS)

                # handles all other rows
                else:
                    iBEGINNINGRANGE = int(math.floor(aSPLITRANGE[iROWSCOUNTER - 1])) + 1
                    dROWS["Row_{0}".format(iROWSCOUNTER)] = random.sample(range(iBEGINNINGRANGE, iUPPERRANGE, 1),
                                                                          k=iCOLUMNS)
                # increments to the next row creation
                iROWSCOUNTER += 1
                # Exit row creation loop
            #########################

            #########################
            # concatenate,transpose and reshape the arrays to fit the Column x Row input requirement
            # traverse the keys and values in our arrays' dictionary (test)
            #########################
            dBINGO_SHEETS["Bingo_Sheet_{0}".format(iCARDS_Counter)] = np.concatenate(
                [v for k, v in sorted(dROWS.items())],
                0)
            dBINGO_SHEETS["Bingo_Sheet_{0}".format(iCARDS_Counter)] = np.transpose(
                dBINGO_SHEETS["Bingo_Sheet_{0}".format(iCARDS_Counter)])
            dBINGO_SHEETS["Bingo_Sheet_{0}".format(iCARDS_Counter)] = dBINGO_SHEETS[
                "Bingo_Sheet_{0}".format(iCARDS_Counter)].reshape(iROWS, iCOLUMNS)

            #########################
            # If statements applied to each bingo sheet to determine 0 placements
            # If it is a 5x5 bingo sheet and 1 zero requested put it in the middle
            # Otherwise randomly add the zeros throughout the bingo sheet
            #########################
            # 1 zero and standard 5x5 matrix
            if iZEROS == 1 and iCOLUMNS == 5 and iROWS == 5:
                dBINGO_SHEETS["Bingo_Sheet_{0}".format(iCARDS_Counter)][2][2] = 0
            # standard 5x5 matrix but with multiple zeros
            # will still add multiple zeros but makes sure there is one in the center
            elif iZEROS > 1 and iCOLUMNS == 5 and iROWS == 5:
                dBINGO_SHEETS["Bingo_Sheet_{0}".format(iCARDS_Counter)][2][2] = 0
                for i in lSELECTED[:-1]:
                    x_cord = dCOORDINATES[str(i) + "_Coordinates"][1]
                    y_cord = dCOORDINATES[str(i) + "_Coordinates"][0]
                    dBINGO_SHEETS["Bingo_Sheet_{0}".format(iCARDS_Counter)][x_cord][y_cord] = 0
                    # all other cases add zeros randomly throughout the matrix
            else:
                for i in lSELECTED:
                    x_cord = dCOORDINATES[str(i) + "_Coordinates"][1]
                    y_cord = dCOORDINATES[str(i) + "_Coordinates"][0]
                    dBINGO_SHEETS["Bingo_Sheet_{0}".format(iCARDS_Counter)][x_cord][y_cord] = 0
            # End of If-statements
            #########################

            #########################
            # Print bingo cards
            #     if bCARDS == "Yes":
            #########################
            print("Bingo_Sheet_" + sCARDS)
            print(dBINGO_SHEETS["Bingo_Sheet_{0}".format(iCARDS_Counter)])

            #########################
            # Increment to the next bingo sheet
            #########################
            iCARDS_Counter += 1

            #########################
        # Print total run time in seconds
        #########################
        print("--- %s seconds ---" % (time.time() - start_time))

    #########################
    # Create simulations depending on the user choice
    #########################

    if sUSERDECISION == "Standard":
        #########################
        # Declaring Variables
        #########################

        iSIMULATION_RANGE = range(1, 76, 1)
        dSIMULATION_SEQUENCES = {}
        iSIMULATION_COUNTER = 1
        start_time = time.time()

        #########################
        # While loop
        #########################
        while iSIMULATION_COUNTER <= iSIMULATION:
            #########################
            # Variables and Dictionaries created within loop
            #########################
            sSIMULATION = str(iSIMULATION_COUNTER)
            dSIMULATION_SEQUENCES["Simulation_{0}".format(iSIMULATION_COUNTER)] = sSIMULATION
            dSIMULATION_SEQUENCES["Simulation_{0}".format(iSIMULATION_COUNTER)] = random.sample(iSIMULATION_RANGE, 75)

            #########################
            # Print statement of each simulation
            #########################
            print("Simulation_" + sSIMULATION)
            print(dSIMULATION_SEQUENCES["Simulation_{0}".format(iSIMULATION_COUNTER)])

            #########################
            # Increment
            #########################
            iSIMULATION_COUNTER += 1

            #########################
        # Ending Timestamp
        #########################
        print("--- Run time is %s seconds ---" % (time.time() - start_time))

    elif sUSERDECISION == "Dynamic":
        #########################
        # Declaring Variables and Inputs
        #########################

        iSIMULATION_RANGE = range(iBEGINNINGRANGE_BACKUP, iENDINGRANGE_BACKUP + 1, 1)
        dSIMULATION_SEQUENCES = {}
        iSIMULATIONCOUNTER = 1

        #########################
        # START THE CLOCK !!!!!!!
        #########################
        start_time = time.time()

        #########################
        # while loop
        #########################
        while iSIMULATIONCOUNTER <= iSIMULATION:
            #########################
            # Variables created within loop
            #########################
            sSIMULATION = str(iSIMULATIONCOUNTER)

            #########################
            # Create dictionary of all the simulations and randomize the order
            #########################
            dSIMULATION_SEQUENCES["Simulation_{0}".format(iSIMULATIONCOUNTER)] = sSIMULATION
            dSIMULATION_SEQUENCES["Simulation_{0}".format(iSIMULATIONCOUNTER)] = random.sample(iSIMULATION_RANGE,
                                                                                               iRANGE)

            #########################
            # View the simulation in
            #########################
            print("Simulation_" + sSIMULATION)
            print(dSIMULATION_SEQUENCES["Simulation_{0}".format(iSIMULATIONCOUNTER)])

            #########################
            # Increment to the next simulation
            #########################
            iSIMULATIONCOUNTER += 1
            # end of loop

        #########################
        # Print total run time in seconds
        #########################
        print("--- %s seconds ---" % (time.time() - start_time))

    #########################
    # Run simulations depending on the user choice
    #########################

    if sUSERDECISION == "Standard":
        #########################
        # Declaring Variables
        #########################
        start_time = time.time()
        iSHEET_COUNTER = 1
        df_BINGO_RESULTS = pd.DataFrame(
            columns=['Bingo Sheet', 'Simulation', 'Primary Diagonal', 'Secondary Diagonal', 'First Horizontal',
                     'Second Horizontal', 'Third Horizontal'
                , 'Fourth Horizontal', 'Fifth Horizontal', 'First Vertical', 'Second Vertical', 'Third Vertical'
                , 'Fourth Vertical', 'Fifth Vertical'])
        #########################
        # While loop - Loop over cards
        #########################
        while iSHEET_COUNTER <= iCARDS:

            #########################
            # Inner while loop - Loop over each simulation for each bingo card
            # Variables created within loop
            #########################
            iSIMULATION_COUNTER = 1
            while iSIMULATION_COUNTER <= iSIMULATION:

                #########################
                # For each simulation reset the success flags
                # deep copy the current bingo sheet, so we have an original as we modify the copy
                #########################
                # pd is Primary Diagonal, od is Opposite Diagonal
                pdsuccess_flag = 0
                odsuccess_flag = 0
                # first, second, ...h indicates flags for horizontal combinations
                firsthsuccess_flag = 0
                secondhsuccess_flag = 0
                thirdhsuccess_flag = 0
                fourthhsuccess_flag = 0
                fifthhsuccess_flag = 0
                # firstv, secondv, ...v indicates flags for vertical combinations
                firstvsuccess_flag = 0
                secondvsuccess_flag = 0
                thirdvsuccess_flag = 0
                fourthvsuccess_flag = 0
                fifthvsuccess_flag = 0
                # Deepcopy the current bingo sheet
                aCURRENT_ARRAY = copy.deepcopy(dBINGO_SHEETS["Bingo_Sheet_{0}".format(iSHEET_COUNTER)])

                #########################
                # for loop through each simulation's numbers
                # checking if the number exists in the bingo sheet and logging results in a dataframe
                #########################
                for index, x in enumerate(dSIMULATION_SEQUENCES["Simulation_{0}".format(iSIMULATION_COUNTER)]):
                    if x in dBINGO_SHEETS["Bingo_Sheet_{0}".format(iSHEET_COUNTER)]:
                        # print("yes!!! ", x, " was found in Bingo Sheet: ",str(iSHEET_COUNTER), "in the", index, "The simulation sheet is: ",str(iSIMULATION_COUNTER))
                        # print(aCURRENT_ARRAY)

                        #########################
                        # find coordinates of number and replace with 0
                        #########################
                        result = np.where(aCURRENT_ARRAY == x)
                        listOfCoordinates = list(zip(result[0], result[1]))
                        for cord in listOfCoordinates:
                            x_cord = cord[0]
                            y_cord = cord[1]
                        aCURRENT_ARRAY[x_cord][y_cord] = 0

                        #########################
                        # sum each winning combination
                        #########################
                        primary_diagonal = sum(aCURRENT_ARRAY.diagonal())
                        opposite_diagonal = sum(np.fliplr(aCURRENT_ARRAY).diagonal())
                        first_horizontal = sum(aCURRENT_ARRAY[0])
                        second_horizontal = sum(aCURRENT_ARRAY[1])
                        third_horizontal = sum(aCURRENT_ARRAY[2])
                        fourth_horizontal = sum(aCURRENT_ARRAY[3])
                        fifth_horizontal = sum(aCURRENT_ARRAY[4])
                        first_column = np.sum(aCURRENT_ARRAY, axis=0)[0]
                        second_column = np.sum(aCURRENT_ARRAY, axis=0)[1]
                        third_column = np.sum(aCURRENT_ARRAY, axis=0)[2]
                        fourth_column = np.sum(aCURRENT_ARRAY, axis=0)[3]
                        fifth_column = np.sum(aCURRENT_ARRAY, axis=0)[4]

                        #########################
                        # if checks on each combination and registering the win
                        #########################
                        if primary_diagonal == 0 and pdsuccess_flag == 0:
                            print("BINGO!!! From the Primary Diagonal", index)
                            pd_index = index
                            pdsuccess_flag = 1
                        if opposite_diagonal == 0 and odsuccess_flag == 0:
                            print("BINGO!!! From the Opposite Diagonal", index)
                            sd_index = index
                            odsuccess_flag = 1
                        if first_horizontal == 0 and firsthsuccess_flag == 0:
                            print("BINGO!!! From the First Horizontal", index)
                            firsth_index = index
                            firsthsuccess_flag = 1
                        if second_horizontal == 0 and secondhsuccess_flag == 0:
                            print("BINGO!!! From the Second Horizontal", index)
                            secondh_index = index
                            secondhsuccess_flag = 1
                        if third_horizontal == 0 and thirdhsuccess_flag == 0:
                            print("BINGO!!! From the Third Horizontal", index)
                            thirdh_index = index
                            thirdhsuccess_flag = 1
                        if fourth_horizontal == 0 and fourthhsuccess_flag == 0:
                            print("BINGO!!! From the Fourth Horizontal", index)
                            fourthh_index = index
                            fourthhsuccess_flag = 1
                        if fifth_horizontal == 0 and fifthhsuccess_flag == 0:
                            print("BINGO!!! From the Fifth Horizontal", index)
                            fifthh_index = index
                            fifthhsuccess_flag = 1
                        if first_column == 0 and firstvsuccess_flag == 0:
                            print("BINGO!!! From the First Column", index)
                            firstv_index = index
                            firstvsuccess_flag = 1
                        if second_column == 0 and secondvsuccess_flag == 0:
                            print("BINGO!!! From the Second Column", index)
                            secondv_index = index
                            secondvsuccess_flag = 1
                        if third_column == 0 and thirdvsuccess_flag == 0:
                            print("BINGO!!! From the Third Column", index)
                            thirdv_index = index
                            thirdvsuccess_flag = 1
                        if fourth_column == 0 and fourthvsuccess_flag == 0:
                            print("BINGO!!! From the Fourth Column", index)
                            fourthv_index = index
                            fourthvsuccess_flag = 1
                        if fifth_column == 0 and fifthvsuccess_flag == 0:
                            print("BINGO!!! From the Fifth Column", index)
                            fifthv_index = index
                            fifthvsuccess_flag = 1

                        #########################
                        # break the cycle if every winning combination has already been hit
                        #########################
                        if pdsuccess_flag == 1 and odsuccess_flag == 1 and firsthsuccess_flag == 1 and secondhsuccess_flag == 1 and thirdhsuccess_flag == 1 and fourthhsuccess_flag == 1 and fifthhsuccess_flag == 1 and firstvsuccess_flag == 1 and secondvsuccess_flag == 1 and thirdvsuccess_flag == 1 and fourthvsuccess_flag == 1 and fifthvsuccess_flag == 1:
                            break
                #########################
                # Create dictionary of results
                # joins that dictionary to a dataframe after each simulation
                #########################
                new_row = {'Bingo Sheet': iSHEET_COUNTER, 'Simulation': iSIMULATION_COUNTER,
                           'Primary Diagonal': pd_index, 'Secondary Diagonal': sd_index,
                           'First Horizontal': firsth_index, 'Second Horizontal': secondh_index,
                           'Third Horizontal': thirdh_index, 'Fourth Horizontal': fourthh_index,
                           'Fifth Horizontal': fifthh_index, 'First Vertical': firstv_index,
                           'Second Vertical': secondv_index, 'Third Vertical': thirdv_index,
                           'Fourth Vertical': fourthv_index, 'Fifth Vertical': fifthv_index}
                df_BINGO_RESULTS = df_BINGO_RESULTS.append(new_row, ignore_index=True)

                #########################
                # Increments for each while loop
                #########################
                iSIMULATION_COUNTER += 1
            iSHEET_COUNTER += 1

        #########################
        # Modify Dataframe
        #########################
        df_BINGO_RESULTS['first win'] = df_BINGO_RESULTS.iloc[:, 2:].min(axis=1)

        #########################
        # Ending Timestamp
        #########################
        print("--- %s seconds ---" % (time.time() - start_time))
    elif sUSERDECISION == "Dynamic":
        #########################
        # Declaring Variables, dictionaries and dataframes
        #########################
        start_time = time.time()
        iSHEET_COUNTER = 1

        nCounterCOLUMNS = 1
        nCounterROWS = 1
        columnnames = {}
        rownames = {}
        indexs = {}
        df_BINGO_RESULTS = pd.DataFrame()

        #########################
        # Dynamically create the dataframe of results
        #########################
        while nCounterCOLUMNS <= iCOLUMNS:
            strCOLUMNS = str(nCounterCOLUMNS)
            columnnames["{0}_Column".format(nCounterCOLUMNS)] = 0, 0
            nCounterCOLUMNS += 1
        while nCounterROWS <= iROWS:
            strROWS = str(nCounterROWS)
            rownames["{0}_Row".format(nCounterROWS)] = 0, 0
            nCounterROWS += 1

        columnnames2 = {'Bingo Sheet': (0, 0),
                        'Simulation': (0, 0),
                        'Primary Diagonal': (0, 0),
                        'Secondary Diagonal': (0, 0)}
        colnames = {**columnnames, **rownames, **columnnames2}

        #########################
        # while loop - Loop over the number of cards we've created
        #########################
        while iSHEET_COUNTER <= iCARDS:
            #########################
            # while loop - Loop over for each card every simulation
            #########################
            iSIMULATION_COUNTER = 1
            while iSIMULATION_COUNTER <= iSIMULATION:

                ##########################
                # Deep copy the bingo sheets to a new variables that will be used for each iteration
                # This way we can gather our simulation results, and we preserve our original bingo sheets
                ##########################
                aCURRENT_ARRAY = copy.deepcopy(dBINGO_SHEETS["Bingo_Sheet_{0}".format(iSHEET_COUNTER)])

                ##########################
                # add to our columns dictionary
                # first value is our column sum
                # second value is our success-flag once flipped to 1
                # we register that winning combination and don't hit it again for that bingo sheet
                ##########################
                for k, v in columnnames.items():
                    columnnames[k] = 0, 0

                ##########################
                # Same process for rows as columns
                ##########################
                for k, v in rownames.items():
                    rownames[k] = 0, 0

                ##########################
                # hard coded our success and sum
                # for Primary Diagonal and Opposite Diagonal
                ##########################
                pdsuccess_counter = 0
                odsuccess_counter = 0

                ##########################
                # Updated dictionary that will be joined to a dataframe for each iteration
                ##########################
                indexs.update({'Bingo Sheet': iSHEET_COUNTER})
                indexs.update({'Simulation': iSIMULATION_COUNTER})

                ##########################
                # For loop of each number in our simulation array for x simulation
                ##########################
                for index, x in enumerate(dSIMULATION_SEQUENCES["Simulation_{0}".format(iSIMULATION_COUNTER)]):
                    # if statement to see check that x number in the array is selected
                    if x in dBINGO_SHEETS["Bingo_Sheet_{0}".format(iSHEET_COUNTER)]:

                        ##########################
                        # print the array, find the number called, replace it with a 0
                        ##########################
                        # print(aCURRENT_ARRAY)
                        result = np.where(aCURRENT_ARRAY == x)

                        listOfCoordinates = np.array(result)
                        x_cord = int(listOfCoordinates[0])
                        y_cord = int(listOfCoordinates[1])
                        aCURRENT_ARRAY[x_cord][y_cord] = 0

                        ##########################
                        # Checking for winning formulas dynamically
                        #  checking rows and columns using dictionaries established before
                        ##########################
                        for k, v in columnnames.items():
                            col_totals = [sum(x) for x in zip(*aCURRENT_ARRAY)]
                            colsums = col_totals[int(k[0]) - 1]
                            counter = v[1]
                            if colsums == 0 and counter == 0:
                                columnnames[k] = colsums, 1
                                indexs.update({k: index + 1})
                        for k, v in rownames.items():
                            rowsums = sum(aCURRENT_ARRAY[int(k[0]) - 1])
                            rowcounter = v[1]
                            if rowsums == 0 and rowcounter == 0:
                                rownames[k] = rowsums, 1
                                indexs.update({k: index + 1})

                                ##########################
                        # Checking for winning formulas of primary and opposite diagonal
                        ##########################
                        primary_diagonal = sum(aCURRENT_ARRAY.diagonal())
                        opposite_diagonal = sum(np.fliplr(aCURRENT_ARRAY).diagonal())
                        if primary_diagonal == 0 and pdsuccess_counter == 0:
                            indexs.update({'primary_diagonal': index + 1})
                            pdsuccess_counter = 1
                        if opposite_diagonal == 0 and odsuccess_counter == 0:
                            indexs.update({'opposite_diagonal': index + 1})
                            odsuccess_counter = 1

                ##########################
                # write to our dataframe after each simulation is complete
                # which turn each winning combination won on
                ##########################
                df_TEMP = pd.DataFrame([indexs], columns=indexs.keys())
                df_BINGO_RESULTS = pd.concat([df_BINGO_RESULTS, df_TEMP])

                ##########################
                # Increment to the next simulation and iterate over the same bingo card
                ##########################
                iSIMULATION_COUNTER += 1

                ##########################
            # Move to the next bingo sheet
            ##########################
            iSHEET_COUNTER += 1

        ##########################
        # Clean and Modify DataFrame
        ##########################
        df_BINGO_RESULTS = df_BINGO_RESULTS.reset_index()
        del df_BINGO_RESULTS["index"]
        df_BINGO_RESULTS['first win'] = df_BINGO_RESULTS.iloc[:, 2:].min(axis=1)

        #########################
        # Print total run time in seconds
        #########################
        print("--- %s seconds ---" % (time.time() - start_time))

    #########################
    # Print Descriptive Analysis depending on the user choice
    #########################
    if sUSERDECISION == "Standard":
        #########################
        # For loop over columns
        # Descriptive Analytics of each column
        #########################
        for (columnName, columnData) in df_BINGO_RESULTS.items():  # removed iteritems()
            if columnName == 'Bingo Sheet' or columnName == 'Simulation':
                nothing = 'nothing'
            else:
                print('Column Name:', columnName)
                print('Median:', statistics.median(columnData.values))
                print('Min:', min(columnData.values))
                print('First Quartile:', np.quantile(columnData.values, 0.25))
                print('Mean:', statistics.mean(columnData.values))
                print('Third Quartile:', np.quantile(columnData.values, 0.75))
                print('Max:', max(columnData.values))
                print('Kurtosis:', kurtosis(columnData.values.astype(int).astype(float)))
                print('Standard Deviation:', statistics.stdev(columnData.values))
                print("----------------------------")

        print("Number of Bingo-Simulation Iterations:", iCARDS * iSIMULATION * 75, "Cards: ", iCARDS, ", Simulations: ",
              iSIMULATION, ", Iterations per simulation: ", 75)
    elif sUSERDECISION == "Dynamic":
        #########################
        # For loop over columns
        # Descriptive Analytics of each column
        #########################
        for (columnName, columnData) in df_BINGO_RESULTS.items():  # removed iteritems()
            if columnName == 'Bingo Sheet' or columnName == 'Simulation':
                nothing = 'nothing'
            else:
                try:#some values might raise error like single simulation
                    print('Column Name:', columnName)
                    print('Median:', statistics.median(columnData.values))
                    print('Min:', min(columnData.values))
                    print('First Quartile:', np.quantile(columnData.values, 0.25))
                    print('Mean:', statistics.mean(columnData.values))
                    print('Third Quartile:', np.quantile(columnData.values, 0.75))
                    print('Max:', max(columnData.values))
                    print('Kurtosis:', kurtosis(columnData.values.astype(int).astype(float)))
                    print('Standard Deviation:', statistics.stdev(columnData.values))
                    print("----------------------------")
                except ValueError:
                    continue

        print("Number of Bingo-Simulation Iterations:", iCARDS * iSIMULATION * iRANGE, "Cards: ", iCARDS,
              "Simulations: ",
              iSIMULATION, " Iterations per simulation: ", iRANGE)

    ###################
    # View the whole dataframe
    ###################
    print(df_BINGO_RESULTS)

    #############
    # could be one graph
    ##############
    # libraries & dataset

    # set a grey background (use sns.set_theme() if seaborn version 0.11.0 or above)
    sns.set(style="darkgrid")
    sns.histplot(df_BINGO_RESULTS["first win"], kde=True)
    # in the next version of the distplot function, one would have to write:
    # sns.distplot(data=df, x="sepal_length", kde=True, rug=True) # note that 'kind' is 'hist' by default
    # plt.show()

    print(dBINGO_SHEETS)

    l1 = [(v) for k, v in dBINGO_SHEETS.items()]
    l1
    # added absolute wdirectory
    path = os.getcwd()
    bingo_path = os.path.join(path, 'bingo.png')

    # for standard Bingo
    def standard_bingo_image(cards):
        for i in range(len(cards)):
            print(i, cards[i])
            card = pd.DataFrame(cards[i])
            s = 'BINGO'
            for col in card.columns:
                for c in s[col]:
                    card.rename(columns={col: c}, inplace=True)
            card['N'].replace(to_replace=0, value='<img src= "' + bingo_path + '" width="15" >', inplace=True)

            # .background_gradient()`Styler.hide(axis="index")`
            dfi.export(card.style.hide(axis="index").set_caption("Player {x}".format(x=str(int(i) + 1))),
                       'df_image{x}_new.png'.format(x=str(i)))

    # for dynamic Bingo
    def dynamic_bingo_image(cards):
        for i in range(len(cards)):
            print(i, cards[i])
            card = pd.DataFrame(cards[i])
            card_multplier,remainder = len(card.columns) // 5,len(card.columns) % 5#generate repeated bingo string for longer columns

            s = ['B','I','N','G','O'] * card_multplier
            if card_multplier > 1:  # prevent 0 multiplication
                for l in range(1,card_multplier):
                    for m in "BINGO":
                        s.append(m+str(l))
            if remainder != 0:
                for m in "BINGO"[0:remainder+1]:
                    s.append(m + str(card_multplier+1))

            card_cols = len(card.columns)
            for col in card.columns:
                for c in s[col]:
                    card.rename(columns={col: c}, inplace=True)
            try:  # this code prevents error for cards with columns less than 5 or more than 5
                card.replace(to_replace=0, value='<img src= "' + bingo_path + '" width="15" >',
                                  inplace=True)
            except KeyError:
                pass
            # .background_gradient()`Styler.hide(axis="index")`
            dfi.export(card.style.hide(axis="index").set_caption("Player {x}".format(x=str(int(i) + 1))),
                       'df_image{x}_new.png'.format(x=str(i)))
    if sUSERDECISION == "Standard":
        standard_bingo_image(l1)
    elif sUSERDECISION == "Dynamic":
        dynamic_bingo_image(l1)

    # pip install glob

    def write_pdf():
        imagelist = glob.glob('*_new.png')
        # print('imagelist:',imagelist)
        pdf = FPDF()
        # imagelist is the list with all image filenames
        for image in imagelist:
            pdf.add_page()
            pdf.image(image)
            timestr = str(time.strftime("%Y%m%d-%H%M%S"))
        pdf.output("generated_" + timestr + "_file.pdf", "F")
        print('PDF saved.')

    write_pdf()

    return df_BINGO_RESULTS, dBINGO_SHEETS, dSIMULATION_SEQUENCES
