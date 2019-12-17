# -*- coding: utf-8 -*-
"""
Created on Fri Dec  6 

@author: FENG Lizheng The university of Hong Kong

@E-mail:fenglz@hku.hk

"""

# coding=utf-8
import pandas as pd
from bs4 import BeautifulSoup
import re
import xlrd
from selenium import webdriver

def openChrome():
    option = webdriver.ChromeOptions()
    option.add_argument('disable-infobars')
    driver = webdriver.Chrome(chrome_options=option)
    return driver

def websurf(url):
       
    driver.get(url)   
    page = driver.page_source
    return page
    




if __name__=="__main__":
    driver = openChrome()
    
    df={'Movie':[],'ASIN':[],'Author':[],
        'Date':[],'Rating':[],'Find Helpful':[],
        'Format':[],'Verified Purchase':[],
        'Subject':[],'ReviewID':[],'Review':[]}
    
    
    excel_source=xlrd.open_workbook('title2.xlsx')
    sheet_source = excel_source.sheets()[0]
    nrows = sheet_source.nrows
    print('The total movie is:',nrows)
    j=1
    for i in range(0, nrows):
     MovieName=sheet_source.cell_value(i,1)
     ASIN=sheet_source.cell_value(i,5)
     print('Under Processing:',i,MovieName)
     
     try:
    
      if ASIN!='no_result':
        url = 'https://www.amazon.com/s?k='+ASIN+'&ref=nb_sb_noss'  
        page=websurf(url)   
        soup1 = BeautifulSoup(page, "lxml")    
        part1='https://www.amazon.com/'
        part2=soup1.select(".a-link-normal")[7].get('href').split('/',3)[1]
        part3='/product-reviews/'
        part4=ASIN  
        part5='/ref=cm_cr_dp_d_show_all_btm?ie=UTF8&reviewerType=all_reviews'     
        reviewwebsite=part1+part2+part3+part4+part5
        review_page=websurf(reviewwebsite)  
        html_source = driver.page_source
        soup2 = BeautifulSoup(review_page, "lxml")
        
        
        for item in soup2.select(".review"):
            author = item.select(".a-profile-name")
            author=str(author)
            author=re.findall(pattern='a-profile-name">(.*?)</span>',string=author,flags=re.S)
            if len(author)!=0:
                author=author[0]
            else:
                continue
                
            rating=item.select(".a-icon-alt")
            if len(rating)!=0:
                rating=rating[0].text
            else:
                continue
            
            title=item.select(".review-title")
            if len(title)!=0:
                title=title[0].text
            else:
                continue
            
            date=item.select(".review-date")
            if len(date)!=0:
                date=date[0].text
            else:
                continue
            
            THEformat=item.select(".review-format-strip")
            if len(THEformat)!=0:
                THEformat=THEformat[0].text
            else:
                continue
            
            VerfiedPurchase=item.select(".a-color-state")
            if len(VerfiedPurchase)!=0:
                VerfiedPurchase=VerfiedPurchase[0].text
            else:
                continue
            ReviewContext=item.select(".review-text-content")
            if len(ReviewContext)!=0:
                ReviewContext=ReviewContext[0].text
            else:
                continue
            
            Find_helpful=item.select(".cr-vote-text")
            if len(Find_helpful)!=0:
                Find_helpful=Find_helpful[0].text
            else:
                continue
            
            
            ReviewID=item.select(".review-title")
            ReviewID=str(ReviewID)
            ReviewID=re.findall(pattern='reviews/(.*?)/ref',string=ReviewID,flags=re.S)
            if len(ReviewID)!=0:
                ReviewID=ReviewID[0]
            else:
                continue
            
            
            df['Movie'].append(MovieName) 
            df['ASIN'].append(ASIN) 
            df['Author'].append(author) 
            df['Date'].append(date)
            df['Rating'].append(rating)
            df['Find Helpful'].append(Find_helpful)
            df['Format'].append(THEformat)
            df['Verified Purchase'].append(VerfiedPurchase)
            df['Subject'].append(title)
            df['Review'].append(ReviewContext)   
            df['ReviewID'].append(ReviewID)  
            
        label=soup2.select(".a-last")[0]
        label=str(label)      
        label=re.findall(pattern='&amp;(.*?)&amp;',string=label,flags=re.S)[0]
        #print(label)
        if len(label)!=0:
            
                    Flag=True
        else:
                    Flag=False
                
                
        
        
        
        while Flag:
         
            
            reviewwebsite=reviewwebsite+'&'+label
            review_page=websurf(reviewwebsite) 
            html_source = driver.page_source
            soup3 = BeautifulSoup(review_page, "lxml")
            for item in soup3.select(".review"):
                author = item.select(".a-profile-name")
                author=str(author)
                author=re.findall(pattern='a-profile-name">(.*?)</span>',string=author,flags=re.S)
                if len(author)!=0:
                    author=author[0]
                else:
                    continue
                #print(author)
                
                rating=item.select(".a-icon-alt")
                if len(rating)!=0:
                    rating=rating[0].text
                else:
                    continue
            
                title=item.select(".review-title")
                if len(title)!=0:
                    title=title[0].text
                else:
                    continue
            
                date=item.select(".review-date")
                if len(date)!=0:
                    date=date[0].text
                else:
                    continue
            
                THEformat=item.select(".review-format-strip")
                if len(THEformat)!=0:
                    THEformat=THEformat[0].text
                else:
                    continue
            
                VerfiedPurchase=item.select(".a-color-state")
                if len(VerfiedPurchase)!=0:
                    VerfiedPurchase=VerfiedPurchase[0].text
                else:
                    continue
                
                ReviewContext=item.select(".review-text-content")
                if len(ReviewContext)!=0:
                    ReviewContext=ReviewContext[0].text
                else:
                    continue
                Find_helpful=item.select(".cr-vote-text")
                if len(Find_helpful)!=0:
                    Find_helpful=Find_helpful[0].text
                else:
                    continue
                
                ReviewID=item.select(".review-title")
                ReviewID=str(ReviewID)
                ReviewID=re.findall(pattern='reviews/(.*?)/ref',string=ReviewID,flags=re.S)
                if len(ReviewID)!=0:
                    ReviewID=ReviewID[0]
                else:
                    continue
                
                df['Movie'].append(MovieName) 
                df['ASIN'].append(ASIN) 
                df['Author'].append(author) 
                df['Date'].append(date)
                df['Rating'].append(rating)
                df['Find Helpful'].append(Find_helpful)
                df['Format'].append(THEformat)
                df['Verified Purchase'].append(VerfiedPurchase)
                df['Subject'].append(title)
                df['Review'].append(ReviewContext)
                df['ReviewID'].append(ReviewID)  
               
                
                
            label=soup3.select(".a-last")
            label=str(label)            
            label=re.findall(pattern='&amp;(.*?)&amp;',string=label,flags=re.S)
                                   
            if len(label)!=0:
                    Flag=True
                    label=label[0]
                    reviewwebsite=part1+part2+part3+part4+part5
            else:
                    Flag=False
        j+=1        
        if j%5==0:
                print('Under Saving segment') 
                destinationFileName='ReviewAmazon_segment.csv'    
                pd.DataFrame(df).to_csv(destinationFileName, index=False)
                print('Already Saved segment')  
                
     except:
            print('no review')               
                    

print('Under Saving Process') 
                
destinationFileName='ReviewAmazon.csv'                
pd.DataFrame(df).to_csv(destinationFileName, index=False)

print('Already Saved')   
        


