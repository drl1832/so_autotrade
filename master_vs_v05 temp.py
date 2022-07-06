import pandas as pd
import numpy as np#not sure this is needed
import xlwings as xw
import time
import tweepy

#Set global variables.  Variables will not be changed     
# genaric lists
LST_WK=['SUN','MON','TUE','WED','THU','FRI','SAT']
LST_WK_FULL=['SUNDAY','MONDAY','TUESDAY','WEDNESDAY','THURSDAY','FRIDAY','SATURDAY']
LST_MTH=['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC']
LST_MTH_FULL=['JANUARY','FEBRUARY','MARCH','APRIL','MAY','JUNE','JULY','AUGUST','SEPTEMBER','OCTOBER','NOVEMBER','DECEMBER']
LST_PUNC=[',','.',':',';','(',')','?','!','$','%','+','-','=','*','"',"'"]

#Initialize Excel 
xw_path = r"C:\Users\drlai\OneDrive\SO_auto_trade 6_22\test_art_process_5.xlsm"#
# name sheets
# need to check if workbook is open. Right now won't work if wb is open
WB=xw.Book(xw_path)

SH_TWT=WB.sheets("sh_twt")
RG_TWT=SH_TWT.range('rg_twt')

SH_SO=WB.sheets('sh_so')
RG_SO=SH_SO.range('rg_so')

SH_ART=WB.sheets('sh_art')

SH_DICT=WB.sheets('sh_dict')###might seperate Twt dict from so dict
RG_TWT_DICT=SH_DICT.range('rg_twt_dict')
#Copy dataframes from exce to dataframes
#********  load data from excel to dataframe DF_TWT_DICT  Might change and use one dict for both twt and so  **************
DF_TWT_DICT=RG_TWT_DICT.options(pd.DataFrame, expand ='table', index = True, header = True).value
 

======================================
def ini_twr():
    consumer_key = "46Ahaz5e5mjNGssW1iCQWZeKz"
    consumer_secret = "6tth4QwdEJwpgBSSnUu7WnZci12yEWykNZetvim5w0O64tFBHM"
    access_token = "1130983997700370434-5UKCuShjlMa5ROMq5TTyfizeTJ2m74"
    access_token_secret = "KA65riQHjOnvfFJNjR6NdcfJGNrdMUKc6ZdU6qUnXWa7r"
    auth = tweepy.OAuthHandler(consumer_key, consumer_secret)
    auth.set_access_token(access_token, access_token_secret)
    twr_api = tweepy.API(auth)
    return twr_api

#==================================================================================================
#
#==================================================================================================
def df_twt_from_excel():
    df_twt=RG_TWT.options(pd.DataFrame, expand ='table', index =False, header = True).value
    return df_twt

def df_twt_new_create():
    df_twt_new_col=['twt_so_thrd_nbr','twt_id','twt_date','twt_userid','twt_so_url','twt_text','twt_tick','twt_yr','twt_mth',
        'twt_earn','twt_strat','twt_post_nbr']
    df_twt_new=pd.DataFrame(columns = df_twt_new_col)
    df_twt_new.at[0,:]='none'# can use either "at" or "loc" 
    return df_twt_new

def df_so_from_excel():
    df_so=RG_SO.options(pd.DataFrame, expand ='table', index = False, header = True).value
    return df_so

def df_so_new_create():
    df_so_new_col=['thrd_nbr','thrd_title','post_nbr','post_date','art_intro_open_close','art_intro_tick','art_intro_pos_size',
        'art_intro_stock_price','art_intro_return','art_strat_sent','art_earn_tf','art_earn_date','art_earn_bo_ac',
        'art_ttl_opt_price','art_ttl_opt_dr_cr','leg_nbr','leg_buy_sell','leg_open_close','leg_opt_nbr','leg_tick',
        'leg_expire_date','leg_strike_price','leg_call_put','leg_opt_price','leg_opt_dr_cr']
    df_so_new=pd.DataFrame(columns = df_so_new_col)
    df_so_new.at[0,:]='none'# can use either "at" or "loc"
    return df_so_new

#==================================================================================================
#
#==================================================================================================
def twt_new_parce(df_twt_new,twt_new):
        twt_strat=' '           # initialize/reset the strategy. Strategy can be multiple strategy words
        df_twt_new.twt_id[0]=f'SO-{twt_new.id_str}'
        df_twt_new.twt_date[0]=twt_new.created_at
        df_twt_new.twt_userid[0]=twt_new.user.id
        so_url=twt_new.entities['urls'][0]['expanded_url']
        df_twt_new.twt_so_url[0]=so_url
            
        find_string='forums/forum/topic/'
        start=so_url.find(find_string)+19
        end=start+4
        df_twt_new.twt_so_thrd_nbr[0]=so_url[start:end]
            
        twt_text=twt_new.text
        df_twt_new.twt_text[0]=twt_text
    
        #Parce twt_text.
        #Remove the url and other misc text at beginning and end of text
        twt_text_len=len(twt_text)
        twt_text_loc=twt_text.find('Read more at', 20)
        twt_text=twt_text[8:twt_text_loc]#remove "(TRADES)" and "Read more at"
        twt_text=twt_text.replace('.',' ')#remove "."
        twt_text=twt_text.replace('#',' #')# Some tweets have no space in front of #2 so the word is not split correctly
        twt_text=twt_text.upper() #Convert to upper
        twt_text=' '.join(filter(None,twt_text.split(' ')))#Standarize text by removing all exess spaces
            
        twt_text=twt_text.split() #NEED TO STANDARDIZE THE NAMEING CONVENTION.  SHOULD BE lst_twt_text
        for word in twt_text:
            word=word.strip()
    ##           word_query=df_dict.query(f'{df_dict_search_col} == "{word}"')[f'{df_dict_return_col}']
    ##           word_type=word_query.iloc[0]#returns a list and we only want the value
            try:
                word_type=DF_TWT_DICT[DF_TWT_DICT['word']==word]['type'].values[0]# need to check if this returns the value  https://stackoverflow.com/questions/36684013/extract-column-value-based-on-another-column-pandas-dataframe                              
                if word_type == 'ticker':               
                    df_twt_new.twt_tick[0]=word
                elif word_type == 'year':
                    df_twt_new.twt_yr[0]=word
                elif word_type == 'month':
                    df_twt_new.twt_mth[0]=word
                elif word_type == 'non_earnings':
                    df_twt_new.twt_earn[0]=False
                elif word_type == 'strat_word':
                    df_twt_new.twt_strat[0]=df_twt_new.twt_strat[0]+' '+ word
                    df_twt_new.twt_strat[0].strip()
                elif word_type=='post_number':
                    df_twt_new.twt_post_nbr[0]=word
                elif word_type=='misc': #ignore
                    a=1
                else:
    #Not sure what the difference between errors 'not placed in field' and 'not in dictionary'
                    error_100=f'tweet id {df_twt_new.twt_id}: word: {word} not placed if field' #not sure this is needed.  Can't think of a scenario where the word is found in the dict and not placed
                    return error_100
            except:
                error_101=f'tweet id {df_twt_new.twt_id}: word: {word} not in dictionary' #word hasn't been found in the dictionary
                return error_101  
    #Non-Earnings and strategy number are not always in text.
    #Earnings/Non-Earnings strategy
        if df_twt_new.twt_earn[0] == '':
            df_twt_new.twt_earn[0]=True
        #Post/Strategy Number            
        if df_twt_new.twt_post_nbr[0] == ' ':
            df_twt_new.twt_post_nbr[0]='#1'
        # insure all values are filled 
        return df_twt_new     
#==================================================================================================
#
#==================================================================================================
def twt_new_process():
    twr_api=ini_twr()
    df_twt=df_twt_from_excel()
    df_twt_new=df_twt_new_create()
    
    #twt_download_type
    #  1=download twitter history based on user id.  I think its up to 200 in encrements of 100
    #  2=download twitter history based on home page timeline.  I think its last 7 days with max of 40
    #  3= check for new tweet every x seconds
    twt_download_type=2
    ttl_twt_to_download=1
    
    if twt_download_type==1:
        twt_timeline = twr_api.user_timeline(screen_name='SteadyOptions',count = ttl_twt_to_download)          
    elif twt_download_type==2:
        twt_timeline = twr_api.home_timeline(count=ttl_twt_to_download)  
    #parce and save 1 tweet at a time
    n=1
    for twt_new in twt_timeline:
        twt_id_new=f'SO-{twt_new.id_str}'
        df_twt_new.at[0,:]=' '  #reset the new tweet dataframe
        if twt_id_new not in df_twt.twt_id.values:
            df_twt_new=twt_new_parce(df_twt_new,twt_new)
            print(n)
            n+=1
            print(df_twt_new)
            df_twt=pd.concat([df_twt,df_twt_new])
            df_twt.reset_index(drop=True, inplace=True)
    #if twt_download_type==3:
    #needs to be completed
    print(df_twt_new)
    RG_TWT.expand().clear_contents()
    RG_TWT.options(pd.DataFrame, expand ='table', index = False, header = True).value=df_twt
    print("Finished twt_new_process")
###################################################################################################################################    
#def main_twt():
#    twr_api=ini_twr()
#    #twt_download_type
#    #  1=download twitter history based on user id.  I think its up to 200 in encrements of 100
#    #  2=download twitter history based on home page timeline.  I think its last 7 days with max of 40
#    #  3= check for new tweet every x seconds
#    twt_download_type=2
#    ttl_twt_to_download=1
#    twt_new_process(twr_api,twt_download_type,ttl_twt_to_download)
#    print("Finished Main")
###################################################################################################################################
twt_new_process()
