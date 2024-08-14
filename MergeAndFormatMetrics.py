import os
import glob 
import pandas as pd 
import datetime
from datetime import datetime
import shutil


def mergeMetricsAndFormat(outputFileName):

    #date time for file extenstion 
    month, year, day = datetime.now().month, datetime.now().year, datetime.now().day

    # Get current directory and create a sessions_date folder if it does not exist
    cwd = os.getcwd()
    dir = "sessions_{}_{}_{}".format(month,day,year)
    path = os.path.join(cwd, dir)

    if not os.path.exists(path):
        os.makedirs(path)


    # dir_list = os.listdir(cwd)
    # print(dir_list)

    #get all files ending in csv and put them in the sessions folder while renaming for easier access
    files = glob.glob(cwd + "/*.csv") 
    i = 0
    for file in files:
        filename = os.path.basename(file)
        newfilename = file.replace(filename, "") +"{}\\".format(dir) + "session{}".format(str(i)) + ".csv"
        os.rename(file, newfilename)
        i +=1

    # print(path)

    #gets all our renamed session file names
    filesinsessionsdir = os.listdir(path)
    # print(filesinsessionsdir)

    #append root path to our filenames for access
    file_lists =[]
    for file in filesinsessionsdir:
        file_lists.append(path+"\\"+file)

    # print(file_lists)
    #merge all csv by creating main dataframe with our first file and appending the rest to it
    main_dataframe = pd.DataFrame(pd.read_csv(file_lists[0]))
    for i in range(1,len(file_lists)): 
        data = pd.read_csv(file_lists[i]) 
        df = pd.DataFrame(data) 
        main_dataframe = pd.concat([main_dataframe,df],axis=0) 
    # print(main_dataframe.head())
    df = main_dataframe

    #remove all rows with no initial user message
    df = df[df['InitialUserMessage'].notna()]
    # print(df['InitialUserMessage'])


    #neeeded for operations below
    pd.options.mode.chained_assignment = None


    #move topicId and SessionOutcome coulmns to the front
    TopicId = df['TopicId']
    SessionOutcome = df['SessionOutcome']
    df.drop(labels=['TopicId'], axis=1,inplace = True)
    df.drop(labels=['SessionOutcome'], axis=1,inplace = True)

    df.insert(0, 'TopicId', TopicId)
    df.insert(0, 'SessionOutcome', SessionOutcome)

    df.head()

    #create as many columns as the longest interaction and expand the interactions into their own columns 
    max_semicolons = df['ChatTranscript'].str.split(';').transform(len).max()
    df[[f'interaction {x}' for x in range(max_semicolons)]] = df['ChatTranscript'].str.split(';', expand=True)

    #chattrascript and the first interactions with the initial message of the bot can be removed
    df.drop(labels=['ChatTranscript'],  axis=1,inplace = True)
    df.drop(labels=['interaction 0'], axis=1, inplace=True)


    #filename provided in function call for differnt bots
    filename = "{}-{}_{}_{}.xlsx".format(outputFileName, month, day, year)

    #convert data frame to excell file
    writer = pd.ExcelWriter(filename) 
    df.to_excel(writer, sheet_name='sheetName', index=False)
    workbook  = writer.book
    wrap_format = workbook.add_format({'text_wrap': True}) #text wrap formatting

    #formatting for readability
    for column in df:
        # column_length = max(df[column].astype(str).map(len).max(), len(column))
        col_idx = df.columns.get_loc(column)
        if column == "TopicName":
            writer.sheets['sheetName'].set_column(col_idx, col_idx, 30)
        elif column == "StartDateTime(UTC)":
            writer.sheets['sheetName'].set_column(col_idx, col_idx, 20)
        elif column == "TopicId":
            writer.sheets['sheetName'].set_column(col_idx, col_idx, 38)
        elif column == "SessionOutcome":
            writer.sheets['sheetName'].set_column(col_idx, col_idx, 17)
        else:
            writer.sheets['sheetName'].set_column(col_idx, col_idx, 50, wrap_format) #wrap text for rest of columns
    #save excel file
    writer.close()

    #cleanup created directory along with sessions
    if  os.path.exists(path):
        shutil.rmtree(path)


if __name__ == "__main__":
    filename = "ITSM_test" #change for each bots.
    mergeMetricsAndFormat(filename)