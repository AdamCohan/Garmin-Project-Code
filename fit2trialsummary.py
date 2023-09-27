from garmin_fit_sdk import Decoder, Stream
import pandas as pd
import sys

def summarize(fitFile, intervalEndTimeList):
    defaultEndTimeList = [15,25,35,45]
    if intervalEndTimeList == []:
        intervalEndTimeList = defaultEndTimeList

    stream = Stream.from_file(fitFile)
    decoder = Decoder(stream)
    
    if(not decoder.is_fit()):
        raise RuntimeError("given file is not a valid .fit file")

    if(not decoder.check_integrity()):
       raise RuntimeError("given .fit file has bad integrity")

    messages, errors = decoder.read()

    if(errors != []):
        print("ERRORS: " + errors)

    record_mesgs_df = pd.DataFrame.from_records(messages['record_mesgs'])

    # given in the order that I expect them to show up in goldencheetah
    # lowercase and '_' is from record_mesgs
    # uppercase and ' ' is from developer_fields
    # columnsOfInterestSDK = ['timestamp', 'distance', 'RP_Power', 'heart_rate', 'speed', 'stance_time_balance', 'cadence', 'vertical_oscillation', 'stance_time',       # chart 1
    #                       'Power', 'Cadence', 'Ground Time', 'Vertical Oscillation', 'Form Power', 'Leg Spring Stiffness',                                             # chart 2
    #                       'stance_time_percent', 'vertical_ratio', 'step_length']                                                                                      # chart 3
    
    columnsOfInterestGC = ['timestamp', 'Distance', 'Power', 'Heartrate', 'Speed', 'Left/Right Balance', 'Run Cadence', 'Vertical Oscillation', 'GCT',                # chart 1
                          'POWER-2', 'CADENCE-2', 'RUNCONTACT-2', 'RUNVERT-2', 'Form Power', 'Leg Spring Stiffness',                                                # chart 2
                          'STANCETIMEPERCENT', 'VERTICALRATIO', 'STEPLENGTH']                                                                                       # chart 3


    # putting everything together and renaming to fit goldencheetah names
    dfOfInterest = getColsOfInterestRenamed(record_mesgs_df, columnsOfInterestGC)

    # adding new time series for seconds since start of trial
    dfWithTime = addSecondsFromStart(dfOfInterest)
    dfWithTime['timestamp'] = dfWithTime['timestamp'].dt.strftime("%Y-%m-%d %H:%M:%S")
    # dfWithTime['timestamp'] = dfWithTime['timestamp'].apply(lambda x: x.replace(tzinfo=None)) # not using anymore since i just convert everything to a string

    # set up final file (same name as the given fit file)
    xlsxName = fitFile[:-3] + 'xlsx'
    aggregateOverIntervals(dfWithTime, intervalEndTimeList, [2,4.5], columnsOfInterestGC,xlsxName)

def getColsOfInterestRenamed(record_mesgs, colsOfInterest):
    dev_fields = pd.DataFrame.from_records(list(record_mesgs['developer_fields']))
    dev_fields.rename(columns={0:'RP_Power', 1:'Power', 2:'Cadence', 3:'Ground Time', 4:'Vertical Oscillation', 5:'Form Power', 6:'Leg Spring Stiffness', 7:'Air Power'}, inplace=True, errors='raise')

    if record_mesgs.shape[0] == dev_fields.shape[0]:
        full_mesgs_df = pd.concat([record_mesgs, dev_fields], axis=1, join='inner')
    else:
        print("not the same size")

    full_mesgs_df.rename(columns={'stance_time_percent':'STANCETIMEPERCENT', 'vertical_ratio':'VERTICALRATIO', 'step_length':'STEPLENGTH',                              # chart 3
                                  'Power':'POWER-2', 'Cadence':'CADENCE-2', 'Ground Time':'RUNCONTACT-2', 'Vertical Oscillation':'RUNVERT-2',                           # chart 2
                                        'Form Power':'Form Power', 'Leg Spring Stiffness':'Leg Spring Stiffness',
                                  'distance':'Distance', 'RP_Power':'Power', 'heart_rate':'Heartrate', 'speed':'Speed', 'stance_time_balance':'Left/Right Balance',     # chart 1
                                        'cadence':'Run Cadence', 'vertical_oscillation':'Vertical Oscillation', 'stance_time':'GCT'}, inplace=True, errors='raise')

    dfOfInterest = full_mesgs_df[colsOfInterest]

    return dfOfInterest

def addSecondsFromStart(dfOfInterest):
    baseDate = dfOfInterest['timestamp'][0]
    secondsSinceStart = dfOfInterest['timestamp'].apply(lambda x: (x - baseDate).seconds)
    secondsSinceStart.name = 'seconds_since_start'

    dfWithTime = pd.concat([secondsSinceStart, dfOfInterest], axis=1, join='inner')

    return dfWithTime

def aggregateOverIntervals(ogDF, intervalEndTimeList, intervalDurationList, colsOfInterest, xlsxName):
    '''
    All times are given in minutes
    intervalEndTimeList --> the end times for each interval (measured backwards from 15min, 25min, 35min, 45min)
    intervalDurationList --> the durations of intervals to be measuring from the end of each interval [2min, 5min]
    '''

    writer = pd.ExcelWriter(xlsxName, engine='xlsxwriter')

    if 'timestamp' in colsOfInterest:
        colsOfInterest.remove('timestamp')
        df = ogDF.drop(columns=['timestamp'])
    else:
        df = ogDF
    
    numIntervals = len(intervalEndTimeList)

    # for 2min, 5min, make the summary df
    for intervalLen in intervalDurationList:

        # Distance_1, Distance_2, Distance_3, Distance_4, GCT_1, GCT_2, ... (initializing with value of intervalLen)
        curListOfSeries = [intervalLen] * len(colsOfInterest) * numIntervals

        dfName = str(intervalLen) + 'm Interval Summary'

        # for 15, 25, 35, 45, get the summary stats and put them in the larger list of series
        for j in range(len(intervalEndTimeList)):
            eoiMin = intervalEndTimeList[j]
            duration = intervalLen * 60
            eoi = eoiMin * 60
            soi = eoi - duration

            rowsSinceStart = df[df['seconds_since_start'] >= soi]
            rowsInInterval = rowsSinceStart[rowsSinceStart['seconds_since_start'] <= eoi]

            aggDF = rowsInInterval.agg(['mean', 'std', 'max', 'min'])
            aggDF.loc['range'] = aggDF.apply(lambda x: x['max']-x['min'])
            aggDF.loc['interval'] = str(soi) + ' - ' + str(eoi)


            for k in range(len(colsOfInterest)):
                colNameOld = colsOfInterest[k]
                colNameNew = colNameOld + '_' + str(j + 1)
                aggDF.rename(columns={colNameOld:colNameNew}, inplace=True, errors='raise')
                curListOfSeries[numIntervals*k + j] = aggDF[colNameNew]  

        curDF = pd.DataFrame.from_dict(curListOfSeries).transpose()

        curDF.to_excel(writer, sheet_name=dfName)

    ogDF.to_excel(writer, sheet_name='Continuous', index=False)

    writer._save()

summarize(sys.argv[1], intervalEndTimeList=[(lambda x: int(x))(x) for x in sys.argv[2:]])
