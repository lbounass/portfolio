
# coding: utf-8

# # Importing All Relevant Libraries

# In[2]:

from urllib2 import HTTPError
import os, shutil, glob
import xlrd
from xlrd.sheet import ctype_text   
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, InvalidElementStateException, StaleElementReferenceException 
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
from collections import defaultdict
import time
import math


# # Setting Up Excel File

# In[3]:

book = xlrd.open_workbook("/Users/apple/Downloads/Bestsellers_CN_7-16-17.xlsx")
sheet_names = book.sheet_names()


# # Loading Books & Getting Authors Array

# In[4]:

Books = book.sheet_by_index(0)
print ('Sheet name: %s' % Books.name)

Books_Authors = []
print(Books.nrows)


# # Create Text to Keep Track of Authors

# In[5]:

for i in range (1, Books.nrows):
	if((Books.cell(i-1,5)).value != (Books.cell(i, 5)).value):
		Books_Authors.append((Books.cell(i,5)).value)
	else:
		i = i+1


# # Loading Authors & Getting Authors Array

# In[6]:

Authors = book.sheet_by_index(1)	
print ('Sheet name: %s' % Authors.name, Authors.nrows)


# # Create Arrays for Authors' IDs and Names

# In[7]:

Author_ID=[]
Author_Name=[]

for i in range (1, Authors.nrows):
	Author_ID.append((Authors.cell(i,1)).value)
    
for i in range (1, Authors.nrows):
	Author_Name.append((Authors.cell(i,0)).value)
    
print(Author_Name)


# # Create Dictionary for Authors Page (order dependent on hash value, not alphabet)

# In[8]:

Author_Aliases = {}

for author in range (0, Authors.nrows - 1):
    
	Author_Aliases[Author_Name[author]] = {}
	lc = 4
	lc_index = 0
    
	for lc in range(4, 37, 2):
        
		Author_Aliases[Author_Name[author]][lc_index] = Authors.cell(author+1,lc).value; 
		lc_index = lc_index + 1
        
print(Author_Aliases)


# # TO DO ONLY ONCE ( 2 THINGS ) :

# # 1. Check for Overlaps in Dictionary ( i.e.: no alias name is listed as an authentic name)

# In[ ]:

for author in range(0, Authors.nrows - 1):
    
    print("For author #", author)
    
    if (Author_Aliases.get(Author_Aliases.keys()[author])[0] == "N/A"):        
        print("Not applicable")
        continue
        
    else:      
        for author_2 in range(0,Authors.nrows - 1):
            
            if (author == author_2):                
                continue
                
            else:               
                for lc in range(0, 16):
                    
                    name = Author_Aliases.keys()[author]
                    name_2 = Author_Aliases.keys()[author_2]
                    
                    if (Author_Aliases.get(name)[0] == Author_Aliases.get(name_2)[lc]):
                        print("Found duplicate:", name, name_2)


# # 2. Create Folders in Directory 

# In[9]:

for i in range (0, len(Author_ID)):
	newpath = "/Users/apple/Desktop/RA Final/" + str(Author_ID[i])[0:(len(str(Author_ID[i])) - 2)] + "_" + Author_Name[i]  
	if not os.path.exists(newpath):
		os.makedirs(newpath)


# # Create Dictionary That Maps LC_Name to LC_ID

# In[9]:

#### FIRST CREATE DICTIONARY THAT MAPS THE LC NAME TO THE LC ID
Auth_LC = {}
Auth_LC_Name = []
Auth_LC_ID = []
Corrected_Unicode = []

for i in range (0, Authors.nrows-1):
    
    for j in range (4, 37, 2):
        
        if ((Authors.cell(i+1,j)).value):
            Auth_LC_Name.append((Authors.cell(i+1,j)).value)
            Auth_LC_ID .append((Authors.cell(i+1,j-1).value));
            pass
        
        else:
            break
            
for i in range (0,len(Auth_LC_Name)):
    
    Auth_LC [Auth_LC_Name[i] ] =Auth_LC_ID[i] 


# # Replace Name with Special Characters (this can also be bypassed by just running the main code Webscraping code and checking the array with unicode exceptions) [can be skipped]

# In[11]:

summary_test = open ("/Users/apple/Desktop/RA Final/SummaryOfAuthors_test.txt", "w")
summary_test.write("Checking with special characters:" + "\n")
Test_Unicode_Error_Author=[]

for author in range(0,len(Author_Aliases)):
        name=Author_Aliases.keys()[author]
        try:
            summary_test.write("\n" + "For #" + str(author) + " " + str(Author_Aliases.keys()[author]) + "\n")
        
        except UnicodeEncodeError:
            if Author_Aliases.keys()[author] not in  Test_Unicode_Error_Author:
                Test_Unicode_Error_Author.append(Author_Aliases.keys()[author])
                summary_test.write("\n" + "For author number " + str(author) + "\n")

            for lc in range (0 , 17 ):
                if(Author_Aliases.get(name)[lc]):
                    try:
                        summary_test.write("\n"+"          "+Author_Aliases.get(name)[lc]+ " -Good- " + "\n")
                    
                    except UnicodeEncodeError:
                        summary_test.write("\n"+"          "+"Alias # " + str(lc) + " -Unicode Error- " +"\n")
                        print (author, Author_Aliases.get(name)[lc])
                        
                        for i in range (0, len(Auth_LC)):
                            
                            if (Author_Aliases.get(name)[lc] == Auth_LC.keys()[i] ):
                                dummy = Auth_LC.keys()[i]
                                if (Auth_LC.get(dummy) != "N/A"):
                                    driver = webdriver.Firefox()
                                    driver.get("https://lccn.loc.gov/"+Auth_LC.get(dummy))             
                                    Author_Aliases.get(name)[lc]=(driver.find_element_by_xpath("//*[@id='title-top']/h1").text)        
                                    driver.quit() 
                                    
                else:
                    break


# # [can be skipped]

# In[93]:

summary_test_1 = open ("/Users/apple/Desktop/RA Final/SummaryOfAuthors_test_1.txt", "w")
summary_test_1.write("Checking with special characters:" + "\n")
#Test_1_Unicode_Error_Author = []
#list_author = [1153, 1387, 1630, 1936, 2159, 2276, 2294, 2359, 2563, 2804, 2991, 3040, 3186, 3252, 3296, 3378, 3473,
#3715, 3887, 4058, 4143, 4261, 4463, 4500, 4523, 4733, 4783, 4901, 4916, 4976, 5160, 6133, 5700, 6197 ]
#list_author = [3, 7, 34,39, 59, 60, 85, 91, 92, 93, 114, 115, 127, 133, 135, 141, 156, 158, 174, 180, 196, 202, 231, 233]
for author in list_author:
        name=Author_Aliases.keys()[author]
        try:
            summary_test_1.write("\n" + "For #" + str(author) + " " + str(Author_Aliases.keys()[author]) + "\n")
        
        except UnicodeEncodeError:
            if Author_Aliases.keys()[author] not in  Test_1_Unicode_Error_Author:
                Test_1_Unicode_Error_Author.append(Author_Aliases.keys()[author])
                summary_test_1.write("\n" + "For author number " + str(author) + "\n")

            for lc in range (0 , 17 ):
                if(Author_Aliases.get(name)[lc]):
                    try:
                        summary_test_1.write("\n"+"          "+Author_Aliases.get(name)[lc]+ " -Good- " + "\n")
                    
                    except UnicodeEncodeError:
                        summary_test_1.write("\n"+"          "+"Alias # " + str(lc) + " -Unicode Error- " +"\n")
                        print (author, Author_Aliases.get(name)[lc])
                        
                        for i in range (0, len(Auth_LC)):
                            
                            if (Author_Aliases.get(name)[lc] == Auth_LC.keys()[i] ):
                                dummy = Auth_LC.keys()[i]
                                if (Auth_LC.get(dummy) != "N/A"):
                                    driver = webdriver.Firefox()
                                    driver.get("https://lccn.loc.gov/"+Auth_LC.get(dummy))             
                                    Author_Aliases.get(name)[lc]=(driver.find_element_by_xpath("//*[@id='title-top']/h1").text)        
                                    driver.quit() 
                                    
                else:
                    break
                    
        for lc in range (0 , 17 ):
                if(Author_Aliases.get(name)[lc]):
                    try:
                        summary_test_1.write("\n"+"          "+Author_Aliases.get(name)[lc]+ " -Good- " + "\n")
                    
                    except UnicodeEncodeError:
                        summary_test_1.write("\n"+"          "+"Alias # " + str(lc) + " -Unicode Error- " +"\n")
                        print (author, Author_Aliases.get(name)[lc])
                        
                        for i in range (0, len(Auth_LC)):
                            
                            if (Author_Aliases.get(name)[lc] == Auth_LC.keys()[i] ):
                                dummy = Auth_LC.keys()[i]
                                print("Ok")
                                if (Auth_LC.get(dummy) != "N/A"):
                                    driver = webdriver.Firefox()
                                    driver.get("https://lccn.loc.gov/"+Auth_LC.get(dummy))             
                                    Author_Aliases.get(name)[lc]=(driver.find_element_by_xpath("//*[@id='title-top']/h1").text)        
                                    driver.quit() 
                                    
                else:
                    break


# # Initiating All The Exceptions and Needed Arrays & Creating Text File for Updates

# In[13]:

NoSuchElement_Errors = []
Unicode_Errors = []
Unicode_Auth_Names_Errors = []
Empty_Book_Set = [] #### For those with no book under their names
Large_Book_Set = [] #### For those with too many books under their name
Sample_Errors = [] #### For the future handling of NoSuchElement_Exceptions
NoSuchElement_Sample = []

summary = open ("/Users/apple/Desktop/RA Final/SummaryOfAuthors.txt", "w")
summary.write("Starting:" + "\n")

summary_nosuchelement = open ("/Users/apple/Desktop/RA Final/SummaryOfAuthors_nosuchelement.txt", "w")
summary_nosuchelement.write("No such element exceptions:" + "\n")

summary_empty = open ("/Users/apple/Desktop/RA Final/SummaryOfAuthors_empty.txt", "w")
summary_empty.write("Starting:" + "\n")

summary_large = open ("/Users/apple/Desktop/RA Final/SummaryOfAuthors_large.txt", "w")
summary_large.write("Starting:" + "\n")

summary_unicode = open ("/Users/apple/Desktop/RA Final/SummaryOfAuthors_unicode.txt", "w")
summary_unicode.write("Starting:" + "\n")


# # Firefox Script / Web Scraping Using Selenium

# In[58]:

#Author_Aliases= ToCheck
print((Author_Aliases))


# In[95]:

for author in range (0, len(Author_Aliases):
    
    name = Author_Aliases.keys()[author]
    print ("Author #", author, ": ", name)
    time.sleep(2)
    #author_unicode = 0
    
    try:
        summary.write("\n" + "For # "+str(author) + " " + str(Author_Aliases.keys()[author]) + ":" + "\n")
        
    except UnicodeEncodeError:
        
        print("Unicode Encode Error with ", Author_Aliases.keys()[author], "Continue with the alias names")
        summary.write("\n" + "Unicode Encode Error with for author at index " + str(author) + "\n")
        summary_unicode.write("\n" + "Unicode Error at Auth_Name at index " + str(author) + "\n")
        Unicode_Auth_Names_Errors.append(Author_Aliases.keys()[author])
        #author_unicode = 1
        
    for lc in range (0 , 17):
        
        if (Author_Aliases.get(name)[lc]):
            print("Author Alias # ", lc, ": ", Author_Aliases.get(name)[lc])  
            
            try:        
                #### OPEN BROWSER
                try:
                    driver = webdriver.Firefox()
                    driver.get("https://catalog.loc.gov")
                    driver.get("https://catalog.loc.gov/vwebv/searchBrowse")
                    

                #### BROWSE PAGE
                    search_code = Select(driver.find_element_by_id("search-code"))
                    search_code.select_by_visible_text("AUTHORS/CREATORS beginning with (enter last name first)")
                    driver.implicitly_wait(10) # seconds
                    
                    
                except WebDriverException:
                    
                    print ("WebDriver Exeption at ", Author_Aliases.get(name)[lc], "Let's wait 30 seconds.")
                    time.sleep (30)
                    driver.quit()
                    driver = webdriver.Firefox()
                    driver.get("https://catalog.loc.gov")
                    driver.get("https://catalog.loc.gov/vwebv/searchBrowse")
                    
                    search_code = Select(driver.find_element_by_id("search-code"))
                    search_code.select_by_visible_text("AUTHORS/CREATORS beginning with (enter last name first)")
                    driver.implicitly_wait(10) # seconds

                empty = 0
                
                if (Author_Aliases.get(name)[lc] == "N/A"):
                    summary.write("\n!!!!" + "          " + Author_Aliases.keys()[author] + " -Not in library- " + "\n")
                    driver.quit()
                    break
                    
                else:    
                    (driver.find_element_by_id("search-argument")).send_keys(Author_Aliases.get(name)[lc])
                    (driver.find_element_by_name("page.search.search.button")).click()
                    titles = (driver.find_element_by_class_name("search-results-browse-list-title-number")).text
                    max_titles = int(titles[1:(len(titles)-1)])
                    #print(max_titles)
                    
                    if (max_titles == 0):
                        Empty_Book_Set.append(Author_Aliases.get(name)[lc])
                        try:
                            summary.write("\n!!!!" + "          " + Author_Aliases.get(name)[lc] + " -Empty (Exception)- " + "\n")
                            
                        except UnicodeEncodeError:
                            
                            summary.write("\n!!!!" + "          " + "Author Alias # "+ str(lc) + " -Empty (Exception)- " + "\n")
                            summary_unicode.write("\n" + "Unicode Error for author index: " + str(author) + ", alias #: " + str(lc) + " -Empty (Exception)- " + "\n")
                        
                        driver.quit()
                        continue                        
                        
                    if (max_titles == 1):
                        nb_books = 0
                        (driver.find_element_by_css_selector("a[href*='search?searchType=7']")).click();
                        try:
                            summary.write("\n" + "          "+ Author_Aliases.get(name)[lc] + " -Good- " + titles +"\n")
                        
                        except UnicodeEncodeError:
                            
                            summary.write("\n" + "          " + "Author Alias # "+ str(lc) + " -Good- " +  titles + "\n")
                            summary_unicode.write("\n" + "Unicode Error for author index " + str(author) + ", alias #" + str(lc) + " -Good- " + "\n")
                        
                        time.sleep(3)
                        
                        if(driver.find_element_by_xpath("/html/body/main/article/div[2]/h1/small").text=="BOOK"):
                            nb_books = nb_books + 1
                            (driver.find_element_by_xpath("/html/body/main/article/div[2]/div/section/div/div[2]/div/a[2]")).click()
                            (driver.find_element_by_name("butExport")).click()
                            alert = driver.switch_to_alert()
                            time.sleep(5)
                            driver.quit()
                            
                        else:
                            print("No book for  author ", Author_Aliases.get(name)[lc])
                            Empty_Book_Set.append(Author_Aliases.get(name)[lc])
                            try:
                                summary.write("\n" + "          " + Author_Aliases.get(name)[lc] + " -No Book- " + titles + "\n")
                                summary_empty.write("\n" + Author_Aliases.get(name)[lc] + " -No Book- " + titles + "\n")
                                empty = 1 ## Because an IOException will be raised since there is no file to move
                            
                            except UnicodeEncodeError:
                            
                                summary.write("\n" + "          " + "Author Alias # "+ str(lc) + " -No Book- " + titles + "\n")  
                                summary_empty.write("\n" + "Author Index: " + str(author) + ", Alias # "+ str(lc) + " -No Book- " + titles + "\n")   
                                summary_unicode.write("\n" + "Unicode Error for author index: " + str(author) + ", alias #: " + str(lc) + " -No Book- " + "\n")
                                
                            driver.quit()                      
                        
                        if (nb_books == 0):
                            print ("No book to save for ", Author_Aliases.get(name)[lc])
                            
                        else:
                        #### MOVING FILE FOR CASE MAX_TITLES = 1
                            print("Moving files: ")	
                            source_dir="/Users/apple/Downloads"
                            dest_dir="/Users/apple/Desktop/RA Final"
                            files = glob.iglob(os.path.join(source_dir, "*.mrc.part"))
                            time.sleep(3)

                            for basename in os.listdir(source_dir):

                                if basename.endswith('.mrc.part'):
                                    pathname = os.path.join(source_dir, basename)

                                    if os.path.isfile(pathname):
                                        new_file = Author_Aliases.get(name)[lc] + ".mrc"
                                        print(new_file)
                                        shutil.copy2(pathname, dest_dir + "/" + new_file)
                                        os.unlink(pathname)

                            for i in range (0, len(Author_ID)):

                                if (Author_Aliases.keys()[author] == Author_Name[i]):
                                    print("Found it")
                                    shutil.move(dest_dir + "/" + new_file, "/Users/apple/Desktop/RA Final/" + str(Author_ID[i])[0:(len(str(Author_ID[i])) - 2)] + "_" + Author_Name[i] + "/" + str(Author_ID[i])[0:(len(str(Author_ID[i])) - 2)] + "_" + str(lc + 1) + "_" + new_file )
                                    summary.write("\n" + "          --> in Author_ID #: " + str(Author_ID[i])[0:(len(str(Author_ID[i])) - 2)] + "\n" )
                                    summary.write("\n" + "          --> Number of pages: 1" + "\n" )                                     
                                    summary.write("\n" + "          --> Number of books #: " + str(nb_books) + "\n" )
                                    nb_books = 0


                    #### IF MAX_TITLES > 2
                    else:	
                        (driver.find_element_by_css_selector("a[href*='search?searchType=7']")).click();
                        try:
                            summary.write("\n" + "          " + Author_Aliases.get(name)[lc] + " -Good- " + titles + "\n")
                        
                        except UnicodeEncodeError:              
                                summary.write("\n" + "          " + "Author Index: " + str(author) + ", Alias # "+ str(lc) + " -Good- " +  titles + "\n")                    
                                summary_unicode.write("\n" + "Unicode Error for author index: " + str(author) + ", alias #: " + str(lc) + " -Good- " + "\n")

                #### SHOULD MAXIMIZE WINDOW AND MAXIMIZE NUMBER OF RECORDS PER PAGE FOR CONVENIENCE
                        time.sleep(3)
                        driver.maximize_window()
                        time.sleep(3)
                        record = Select(driver.find_element_by_id("record-count"))
                        record.select_by_visible_text("100")
                        time.sleep(7)
                        
                #### 2 possibilities now: 
                #1) max of titles is less than 100, so we just check the books, select the books then click save.
                #2) max of titles is more than 100, so for the first to before last page, we check all 100 options, click Next until 
                # last page, where we take the maximum of options available (<100) then click save.
                
                ############## OPTION 1 
                        if (max_titles <= 100):
                            
                            #### KEEP TRACK OF NUMBER OF BOOKS
                            nb_books = 0
                            pages = 1

                            for i in range (1, max_titles + 1):
                                
                                if (driver.find_element_by_xpath("//div[@id='search-results']/ul/li[" + str(i) + "]/div[3]/div/div[1]").text=="Book"):
                                    nb_books = nb_books + 1
                                    (driver.find_element_by_xpath("//div[@id='search-results']/ul/li[" + str(i) + "]").find_element_by_name("titles")).click()
                                    time.sleep(0.2)
                                    
                                else:
                                    time.sleep(0.2)
                                    
                            print("Number of books by this author: ", nb_books)
                         
                            
                            #### NO NEED TO SCROLL OF MAX_TITLES < 2        
                            if (max_titles > 2):
                                driver.execute_script('window.scrollTo(0,document.body.scrollHeight)'); 

                            if (nb_books != 0):
                                (driver.find_element_by_name("ExportDialogServlet")).click()
                                time.sleep(2)
                                (driver.find_element_by_name("butExport")).click()
                                alert = driver.switch_to_alert()
                                time.sleep(7)	
                                driver.quit()
                                
                            else:
                                print("No book for  author ", Author_Aliases.get(name)[lc])
                                Empty_Book_Set.append(Author_Aliases.get(name)[lc])
                                
                                try:
                                    summary.write("\n" + "          " + Author_Aliases.get(name)[lc] + " -No Book- " + titles + "\n")
                                    summary_empty.write("\n" + Author_Aliases.get(name)[lc] + " -No Book- " + titles + "\n")
                                    
                                except UnicodeEncodeError:              
                                    summary.write("\n" + "          " + "Author Alias # " + str(lc) + " -No Book- " + titles + "\n") 
                                    summary_empty.write("\n" + "Author Index: " + str(author) + ", Alias # " + str(lc) + " -No Book- " + titles + "\n")
                                    summary_unicode.write("\n" + "Unicode Error for author index: " + str(author) + ", alias #: " + str(lc) + " -No Book- " + titles + "\n")
                                driver.quit()
                                
                ############## OPTION 2                            
                        else: 
                        
                            if (max_titles > 399):
                                
                                #### KEEP TRACK OF LARGE SETS FOR FUTURE VERIFICATION
                                if (Author_Aliases.get(name)[lc] not in Large_Book_Set):
                                    Large_Book_Set.append(Author_Aliases.get(name)[lc])
                                    
                                    try:
                                        summary_large.write("\n" + Author_Aliases.get(name)[lc] + " -Large Book Set- " + titles + "\n")
                                    
                                    except UnicodeEncodeError:
                                        summary_large.write("\n" + "Author Index: " + str(author) + ", Alias # " + str(lc) + " -Large Book Set- " + titles + "\n")
                                        summary_unicode.write("\n" + "Unicode Error for author index: " + str(author) + ", alias #: " + str(lc) + " -Large Book Set- " + titles + "\n")
                                
                                else:
                                    print("Already in Large_Book_Set")
                                    
                            pages = math.ceil(max_titles / 100) + 1
                            print("Nb of pages= " + str(pages))
                            k = 1
                            nb_books = 0
                            while (k < pages) :
                                
                                time.sleep(5)
                                print("Page: " , str(k))
                                
                                for i in range (1, 101):
                                    
                                    if (driver.find_element_by_xpath("//div[@id='search-results']/ul/li["+str(i)+"]/div[3]/div/div[1]").text=="Book"):
                                        nb_books = nb_books + 1; 
                                        (driver.find_element_by_xpath("//div[@id='search-results']/ul/li["+str(i)+"]").find_element_by_name("titles")).click()
                                        time.sleep(0.3)
                                        
                                    else:
                                        time.sleep(0.3)
                                      
                                if (k == 4) :
                                    time.sleep(0.5)
                                    (driver.find_element_by_xpath("//*[@id='results-form']/div[1]/ul/li[" + str(k + 1) + "]/a")).click();
                                    time.sleep(3)
                                    
                                elif (k == 6):
                                    time.sleep(0.5)
                                    (driver.find_element_by_xpath("//*[@id='results-form']/div[1]/ul/li[" + str(k + 1) + "]/a")).click();
                                    time.sleep(3)

                                elif (k == 7):
                                    time.sleep(0.5)
                                    (driver.find_element_by_xpath("//*[@id='results-form']/div[1]/ul/li[" + str(k) + "]/a")).click();
                                    time.sleep(3)

                                elif (k > 7):
                                    time.sleep(0.5)
                                    (driver.find_element_by_xpath("//*[@id='results-form']/div[1]/ul/li[7]/a")).click();
                                    time.sleep(3)

                                else:
                                    time.sleep(0.5)
                                    (driver.find_element_by_xpath("//*[@id='results-form']/div[1]/ul/li[" + str(k + 2) + "]/a")).click();
                                    time.sleep(3)
                              
                                k = k + 1
                                time.sleep(3)
                               
                            #### LAST PAGE
                            new = max_titles - (pages - 1) * 100
                            print("Last page:" , str(k), "Number left= ", str(new)) 
                            
                            for last_page in range (1, int(new + 1)):
                                
                                if (driver.find_element_by_xpath("//div[@id='search-results']/ul/li[" + str(last_page) + "]/div[3]/div/div[1]").text == "Book"):
                                    nb_books = nb_books + 1; 
                                    (driver.find_element_by_xpath("//div[@id='search-results']/ul/li[" + str(last_page) + "]").find_element_by_name("titles")).click()
                                    time.sleep(0.7)
                                else:
                                    time.sleep(0.7)
                                    
                            print("Number of books by this author: ", nb_books)
                            driver.execute_script('window.scrollTo(0,document.body.scrollHeight)'); 
                            
                            if (nb_books != 0):
                                (driver.find_element_by_name("ExportDialogServlet")).click()
                                (driver.find_element_by_name("butExport")).click()
                                alert = driver.switch_to_alert()
                            
                            #### THE MORE RECORDS, THE MORE TIME NEEDED
                                if (max_titles > 399):
                                    time.sleep(118)	
                                else:
                                    time.sleep(60)

                                driver.quit()
                                
                            else: 
                                print("No book for  author ", Author_Aliases.get(name)[lc])
                                Empty_Book_Set.append(Author_Aliases.get(name)[lc])
                                
                                try:
                                    summary.write("\n" + "          " + Author_Aliases.get(name)[lc] + " -No Book- " + titles + ", Pages: " + str(pages) + "\n" )
                                    summary_empty.write("\n" + Author_Aliases.get(name)[lc] + " -No Book- " + titles + ", Pages: " + str(pages) + "\n" )
                                    
                                except UnicodeEncodeError:              
                                    summary.write("\n" + "          " + "Author Alias # "+ str(lc) + " -No Book- " + titles + ", Pages: " + str(pages) +  "\n" )                    
                                    summary_empty.write("\n" + "          " + "Author Index: " + str(author) + ", Alias # " + str(lc) + " -No Book- " + titles + ", Pages: " + str(pages) +  "\n" )                    
                                    summary_unicode.write("\n" + "Unicode Error for author index: " + str(author) + ", alias #: " + str(lc) + " -No Book- " + titles  + ", Pages: " + str(pages) + "\n")

                                driver.quit()
                         
                        
                        if (nb_books == 0):
                            print ("No file to move")
                            
                        else:
                        #### MOVING FILE FOR CASE MAX_TITLES > 1
                            print("Moving files: ")	
                            source_dir="/Users/apple/Downloads"
                            dest_dir="/Users/apple/Desktop/RA Final"
                            files = glob.iglob(os.path.join(source_dir, "*.mrc.part"))
                            time.sleep(3)

                            for basename in os.listdir(source_dir):

                                if basename.endswith('.mrc.part'):
                                    pathname = os.path.join(source_dir, basename)

                                    if os.path.isfile(pathname):
                                        new_file=Author_Aliases.get(name)[lc] + ".mrc"
                                        print(new_file)
                                        shutil.copy2(pathname, dest_dir + "/" + new_file)
                                        os.unlink(pathname)

                            for i in range (0, len(Author_ID)):

                                if (Author_Aliases.keys()[author]==Author_Name[i]):

                                    print("Found it")
                                    shutil.move(dest_dir + "/" + new_file, "/Users/apple/Desktop/RA Final/" + str(Author_ID[i])[0:(len(str(Author_ID[i])) - 2)] + "_" + Author_Name[i] + "/" + str(Author_ID[i])[0:(len(str(Author_ID[i])) - 2)] + "_" + str(lc + 1) + "_" + new_file)       
                                    summary.write("\n" + "          --> in Author_ID #: " + str(Author_ID[i])[0:(len(str(Author_ID[i])) - 2)]  + "\n" )
                                    summary.write("\n" + "          --> Number of pages: " + str(pages) + "\n" )                                     
                                    summary.write("\n" + "          --> Number of books: " + str(nb_books) + "\n" ) 
                                    
                                    break

            #### EXCEPTIONS
            except StaleElementReferenceException:
                
                print("Stale Element Reference Exception at author: " + Author_Aliases.get(name)[lc] + ", let's count it as no such element exception")
                
                if (Author_Aliases.get(name)[lc] not in NoSuchElement_Errors):
                    NoSuchElement_Errors.append(Author_Aliases.get(name)[lc])
                    
                    try:   
                        summary.write("\n!!!!" + "          " + Author_Aliases.get(name)[lc] + " -Stale Element Reference Exception- " + "\n")
                        summary_nosuchelement.write("\n" + Author_Aliases.get(name)[lc] + " -Stale Element Reference Exception- " + "\n")
                 
                    except UnicodeEncodeError: 
                        summary.write("\n!!!!" + "          " + "Author Alias # "+ str(lc) + " -Stale Element Reference Exception- " + "\n")
                        summary_nosuchelement.write("\n" + "Author Index: " + str(author) + ", Alias # "+ str(lc) + " -Stale Element Reference Exception- " + "\n")
                        summary_unicode.write("\n" + "Unicode Error for author index: " + str(author) + ", alias #: " + str(lc) + " -Stale Element Reference Exception- " + "\n")
                else:
                    print(Author_Aliases.get(name)[lc] + " is a duplicate for No Such Element Error")
                    
                time.sleep(5) 
                driver.quit()
                time.sleep(60)
                continue

            except WebDriverException:
                
                print("WebDriver Exception at author: " + Author_Aliases.get(name)[lc] + ", let's count it as no such element exception")
                
                if (Author_Aliases.get(name)[lc] not in NoSuchElement_Errors):
                    NoSuchElement_Errors.append(Author_Aliases.get(name)[lc])
                    
                    try:   
                        summary.write("\n!!!!" + "          " + Author_Aliases.get(name)[lc] + " -WebDriver Exception- " + "\n")
                        summary_nosuchelement.write("\n" + Author_Aliases.get(name)[lc] + " -WebDriver Exception- " + "\n")
                        
                    except UnicodeEncodeError: 
                        summary.write("\n!!!!" + "          " + "Author Alias # "+ str(lc) + " -Web Driver Exception- " + "\n")
                        summary_unicode.write("\n" + "Unicode Error for author index: " + str(author) + ", alias #: " + str(lc) + " -Web Driver Exception- " + "\n")
                        summary_nosuchelement.write("\n" + "Author Index: " + str(author) + ", Alias # "+ str(lc) + " -Web Driver Exception- " + "\n")
                        
                else:
                    print(Author_Aliases.get(name)[lc] + " is a duplicate for No Such Element Error")
                
                time.sleep(5) 
                driver.quit()
                time.sleep(30)
                continue
                            
            except InvalidElementStateException:
                
                print("Invalid Element State Exception at author: " + Author_Aliases.get(name)[lc] + ", let's count it as no such element exception")
                
                if (Author_Aliases.get(name)[lc] not in NoSuchElement_Errors):
                    NoSuchElement_Errors.append(Author_Aliases.get(name)[lc])
                    
                    try:   
                        summary.write("\n!!!!" + "          " + Author_Aliases.get(name)[lc] + " -Invalid Element State Exception Exception- " + "\n")
                        summary_nosuchelement.write("\n" + Author_Aliases.get(name)[lc] + " -Invalid Element State Exception Exception- " + "\n")
                  
                    except UnicodeEncodeError: 
                        summary.write("\n!!!!" + "          " + "Author Alias # "+ str(lc) + " -Invalid Element State Exception Exception- " + "\n")
                        summary_unicode.write("\n" + "Unicode Error for author index: " + str(author) + ", alias #: " + str(lc) + " -Invalid Element State Exception Exception- " + "\n")
                        summary_nosuchelement.write("\n" + "Author Index: " + str(author) + ", Alias # " + str(lc) + " -Invalid Element State Exception Exception- " + "\n")
                        
                else:
                    print(Author_Aliases.get(name)[lc] + " is a duplicate for No Such Element Error")
                
                time.sleep(5) 
                driver.quit()
                time.sleep(60)
                continue

            except NoSuchElementException:
                
                print("No Such Element Exception at author: " + Author_Aliases.get(name)[lc] + ", let's continue")
                
                if (Author_Aliases.get(name)[lc] not in NoSuchElement_Errors):
                    NoSuchElement_Errors.append(Author_Aliases.get(name)[lc])
                    
                    try:   
                        summary.write("\n!!!!" + "          " + Author_Aliases.get(name)[lc] + " -No Such Element Exception Exception- " + "\n")
                        summary_nosuchelement.write("\n" + Author_Aliases.get(name)[lc] + " -No Such Element Exception Exception- " + "\n")
                   
                    except UnicodeEncodeError: 
                        summary.write("\n!!!!" + "          " + "Author Alias # "+ str(lc) + " -No Such Element Exception Exception- " + "\n")
                        summary_unicode.write("\n" + "Unicode Error for author index: " + str(author) + ", alias #: " + str(lc) + " -No Such Element Exception Exception- " + "\n")
                        summary_nosuchelement.write("\n" + "Author Index: " + str(author) + ", Alias # " + str(lc) + " -No Such Element Exception Exception- " + "\n")
                        
                else:
                    print(Author_Aliases.get(name)[lc] + " is a duplicate for No Such Element Error")
                
                time.sleep(5) 
                driver.quit()
                continue
                
            except UnicodeEncodeError:
                
                print("Unicode Exception at author: " + Author_Aliases.get(name)[lc] + ", let's continue")
                
                if (Author_Aliases.get(name)[lc] not in Unicode_Errors):
                    Unicode_Errors.append(Author_Aliases.get(name)[lc])
                    
                    try:   
                        summary.write("\n!!!!" + "          " + Author_Aliases.get(name)[lc] + " -Unicode Exception- " + "\n")
                        summary_unicode.write("\n" + Author_Aliases.get(name)[lc] + " -Unicode Exception- " + "\n")
                    
                    except UnicodeEncodeError: 
                        summary.write("\n!!!!" + "          " + "Author Alias # "+ str(lc) + " -Unicode Exception- " + "\n")
                        summary_unicode.write("\n" + "Unicode Error for author index: " + str(author) + ", alias #: " + str(lc) + " -Unicode Exception- " + "\n")
                        
                else:
                    print(Author_Aliases.get(name)[lc] + " is a duplicate for Unicode Error")
                
                time.sleep(5)   
                driver.quit()
                continue
                
            except IOError:
                
                print("IOError Exception at author: "+ Author_Aliases.get(name)[lc] + ", let's count it as no such element exception")
                
                if (empty == 1):
                    empty = 0
                    continue
                    
                else:
                    if (Author_Aliases.get(name)[lc] not in NoSuchElement_Errors):
                        NoSuchElement_Errors.append(Author_Aliases.get(name)[lc])
                
                    try:   
                        summary.write("\n!!!!" + "          " + Author_Aliases.get(name)[lc] + " -IOError Exception- " + "\n")
                        summary_nosuchelement.write("\n" + Author_Aliases.get(name)[lc] + " -IOError Exception- " + "\n")
                    
                    except UnicodeEncodeError: 
                        summary.write("\n!!!!" + "          " + "Author Alias # "+ str(lc) + " -IOError Exception- " + "\n")
                        summary_unicode.write("\n" + "Unicode Error for author index: " + str(author) + ", alias #: " + str(lc) + " -IOError Exception- " + "\n")
                        summary_nosuchelement.write("\n" + "Author Index: " + str(author) + ", Alias # " + str(lc) + " -IOError Exception- " + "\n")
                        
                    else:
                        print(Author_Aliases.get(name)[lc] + " is a duplicate for No Such Element Error")
                
                time.sleep(5)   
                driver.quit()
                
            except TimeoutException:
                
                print("Timeout Exception at author: " + Author_Aliases.get(name)[lc] + ", let's pause and count it as no such element exception")
                
                if (Author_Aliases.get(name)[lc] not in NoSuchElement_Errors):
                    NoSuchElement_Errors.append(Author_Aliases.get(name)[lc])

                    try:   
                        summary.write("\n!!!!" + "          " + Author_Aliases.get(name)[lc] + " -Timeout Exception- " + "\n")
                        summary_nosuchelement.write("\n" + Author_Aliases.get(name)[lc] + " -Timeout Exception- " + "\n")
                   
                    except UnicodeEncodeError: 
                        summary.write("\n!!!!" + "          " + "Author Alias # "+ str(lc) + " -Timeout Exception- " + "\n")
                        summary_unicode.write("\n" + "Unicode Error for author index: " + str(author) + ", alias #: " + str(lc) + " -Timeout Exception- " + "\n")
                        summary_nosuchelement.write("\n" + "Author Index: " + str(author) + ", Alias # " + str(lc) + " -Timeout Exception- " + "\n")
                    
                else:
                    print(Author_Aliases.get(name)[lc] + " is a duplicate for No Such Element Error")
                
                time.sleep(5)   
                driver.quit()
                time.sleep(120)
 
        #### IF NO OTHER ALIAS
        else:
            print("No other ALIAS names")
            summary.write("\n    " + "          " + "No other ALIAS names" + "\n" )


            break


# In[82]:

#Examples of output in NoSuchElement_Errors
NoSuchElement_Errors


# In[81]:

#To obtain the authors without any books under their name
Empty_Book_Set


# Because of server issues, we have to consider the instances in which the IP server gets blocked occasionally and an exception is raised. The NoSuchElement_Exceptions array takes care of that. It will take in any ALIAS name that was discarded because of the exception, and a new script adapted to that array will run. As long as NoSuchElement_Exceptions isn't empty, this script will go on and on.
# 
# As there might be special characters in the spreadhsheet, we will access the LC_Names through their LC_ID: first we will retrieve the appropriate alias names by searching their alias ID, then enter these names in the new list and run the script again.

# # Retrieve Correct LC_Names (Pre-Emptive Step)

# In[ ]:




# # Opening New Text File for No Such Element Exceptions

# In[14]:

summary_error = open ("/Users/apple/Desktop/RA Final/SummaryOfAuthors_error.txt", "w")
summary_error.write("Starting:" +"\n")
summary_nosuchelement_error = open ("/Users/apple/Desktop/RA Final/SummaryOfAuthors_nosuchelement_error.txt", "w")
summary_nosuchelement_error.write("No such element exceptions:" + "\n")
summary_empty_error = open ("/Users/apple/Desktop/RA Final/SummaryOfAuthors_empty_error.txt", "w")
summary_empty_error.write("Starting:" + "\n")
summary_unicode_error = open ("/Users/apple/Desktop/RA Final/SummaryOfAuthors_unicode_error.txt", "w")
summary_unicode_error.write("Starting:" + "\n")
summary_large_error = open ("/Users/apple/Desktop/RA Final/SummaryOfAuthors_large_error.txt", "w")
summary_large_error.write("Starting:" +"\n")


# In[87]:

while (NoSuchElement_Errors):

    Sample_Errors = NoSuchElement_Errors    
    print("Sample_Errors: ")
    print(Sample_Errors, len(Sample_Errors))
    
    for author in range(0 , len(Sample_Errors)):
        
        name = Sample_Errors[author]
        print("Author Alias #", author, ": ", name)
        time.sleep(2)
        #summary_nosuchelement.write("\n" + "For Author Alias #" + str(author) + " " + str(name) + ":" + "\n")
                    
        try:        
            #### OPEN BROWSER
            try:
                driver = webdriver.Firefox()
                driver.get("https://catalog.loc.gov")
                driver.get("https://catalog.loc.gov/vwebv/searchBrowse")
                

            #### BROWSE PAGE
                search_code = Select(driver.find_element_by_id("search-code"))
                search_code.select_by_visible_text("AUTHORS/CREATORS beginning with (enter last name first)")
                driver.implicitly_wait(10) # seconds
                
                
            except WebDriverException:
                
                print ("WebDriver Exeption at ", Sample_Errors[author], "Let's wait 30 seconds.")
                time.sleep (30)
                driver.quit()
                driver = webdriver.Firefox()
                driver.get("https://catalog.loc.gov")
                driver.get("https://catalog.loc.gov/vwebv/searchBrowse")
                
                search_code = Select(driver.find_element_by_id("search-code"))
                search_code.select_by_visible_text("AUTHORS/CREATORS beginning with (enter last name first)")
                driver.implicitly_wait(10) # seconds

            empty = 0
            if (Sample_Errors[author] == "N/A"):
                summary_error.write("\n!!!!" + "          " + Sample_Errors[author] + " -Not in library- " + "\n")
                driver.quit()
                break
                
            else:    
                (driver.find_element_by_id("search-argument")).send_keys(Sample_Errors[author])
                (driver.find_element_by_name("page.search.search.button")).click()
                titles = (driver.find_element_by_class_name("search-results-browse-list-title-number")).text
                max_titles = int(titles[1:(len(titles)-1)])
                #print(max_titles)
                
                if (max_titles == 0):
                    Empty_Book_Set.append(Sample_Errors[author])
                    try:
                        summary_error.write("\n!!!!" + "          " + Sample_Errors[author] + " -Empty (Exception)- " + "\n")
                        summary_empty_error.write("\n" + Sample_Errors[author] + " -Empty (Exception)- " + "\n")
                            
                    except UnicodeEncodeError:
                        
                        if Sample_Errors[author] not in Unicode_Errors:
                            Unicode_Errors.append(Sample_Errors[author])
                        
                        summary_error.write("\n!!!!" + "          " + "Author Alias # "+ str(author) + " -Empty (Exception)- " + "\n")
                        summary_unicode_error.write("\n" + "Unicode Error for author index: " + str(author) + " -Empty (Exception)- " + "\n")
                    
                    driver.quit()
                    continue                        
                    
                if (max_titles == 1):
                    nb_books = 0
                    (driver.find_element_by_css_selector("a[href*='search?searchType=7']")).click();
                    try:
                        summary_error.write("\n" + "          "+ Sample_Errors[author] + " -Good- " + titles +"\n")
                    
                    except UnicodeEncodeError:
                        
                        if Sample_Errors[author] not in Unicode_Errors:
                            Unicode_Errors.append(Sample_Errors[author])
                        
                        summary_error.write("\n" + "          " + "Author Alias # "+ str(author) + " -Good- " +  titles + "\n")
                        summary_unicode_error.write("\n" + "Unicode Error for author index " + str(author) + " -Good- " + "\n")
                    
                    time.sleep(3)
                    
                    if(driver.find_element_by_xpath("/html/body/main/article/div[2]/h1/small").text=="BOOK"):
                        nb_books = nb_books + 1
                        (driver.find_element_by_xpath("/html/body/main/article/div[2]/div/section/div/div[2]/div/a[2]")).click()
                        (driver.find_element_by_name("butExport")).click()
                        alert = driver.switch_to_alert()
                        time.sleep(5)
                        driver.quit()
                        
                    else:
                        print("No book for  author ", Sample_Errors[author])
                        Empty_Book_Set.append(Sample_Errors[author])
                        try:
                            summary_error.write("\n" + "          " + Sample_Errors[author] + " -No Book- " + titles + "\n")
                            summary_empty_error.write("\n" + Sample_Errors[author] + " -No Book- " + titles + "\n")
                            empty = 1 ## Because an IOException will be raised since there is no file to move
                        
                        except UnicodeEncodeError:
                        
                            if Sample_Errors[author] not in Unicode_Errors:
                                Unicode_Errors.append(Sample_Errors[author])
                        
                            summary_error.write("\n" + "          " + "Author Alias # "+ str(author) + " -No Book- " + titles + "\n")  
                            summary_empty_error.write("\n" + "Author Index: " + str(author) + " -No Book- " + titles + "\n")   
                            summary_unicode_error.write("\n" + "Unicode Error for author index: " + str(author) + " -No Book- " + "\n")
                            
                        driver.quit()                      
                    
                    if (nb_books == 0):
                        print ("No book to save for ", Sample_Errors[author])
                        
                    else:
                    #### MOVING FILE FOR CASE MAX_TITLES = 1
                    
                       
                        print("Moving files: ")	
                        source_dir="/Users/apple/Downloads"
                        dest_dir="/Users/apple/Desktop/RA Final"
                        files = glob.iglob(os.path.join(source_dir, "*.mrc.part"))
                        time.sleep(3)

                        for basename in os.listdir(source_dir):

                            if basename.endswith('.mrc.part'):
                                pathname = os.path.join(source_dir, basename)

                                if os.path.isfile(pathname):
                                    new_file = Sample_Errors[author] + ".mrc"
                                    print(new_file)
                                    shutil.copy2(pathname, dest_dir + "/" + new_file)
                                    os.unlink(pathname)
                                    

                        for i in range(0, Authors.nrows - 1):

                            name_sample=Author_Aliases.keys()[i]

                            for j in range (0, 17):

                                if (Sample_Errors[author]==Author_Aliases.get(name_sample)[j]):
                                    name = name_sample
                                    lc_name = j
                                    print (lc_name)

                        for i in range (0, len(Author_ID)):

                            if (name == Author_Name[i]):
                                print("Found it")
                                shutil.move(dest_dir + "/" + new_file, "/Users/apple/Desktop/RA Final/" + str(Author_ID[i])[0:(len(str(Author_ID[i])) - 2)] + "_" + Author_Name[i] + "/" + str(Author_ID[i])[0:(len(str(Author_ID[i])) - 2)] + "_" + str(lc_name + 1) + "_" + new_file )
                                summary_error.write("\n" + "          --> in Author_ID #: " + str(Author_ID[i])[0:(len(str(Author_ID[i])) - 2)] + "\n" )
                                summary_error.write("\n" + "          --> Number of pages: 1" + "\n" )                                     
                                summary_error.write("\n" + "          --> Number of books #: " + str(nb_books) + "\n" )
                                nb_books = 0


                #### IF MAX_TITLES > 2
                else:	
                    (driver.find_element_by_css_selector("a[href*='search?searchType=7']")).click();
                    try:
                        summary_error.write("\n" + "          " + Sample_Errors[author] + " -Good- " + titles + "\n")
                    
                    except UnicodeEncodeError: 
                        
                        if Sample_Errors[author] not in Unicode_Errors:
                            Unicode_Errors.append(Sample_Errors[author])
                        
                        summary_error.write("\n" + "          " + "Author Index: " + str(author) + " -Good- " +  titles + "\n")                    
                        summary_unicode_error.write("\n" + "Unicode Error for author index: " + str(author) + " -Good- " + "\n")

            #### SHOULD MAXIMIZE WINDOW AND MAXIMIZE NUMBER OF RECORDS PER PAGE FOR CONVENIENCE
                    time.sleep(3)
                    driver.maximize_window()
                    time.sleep(2)
                    record = Select(driver.find_element_by_id("record-count"))
                    record.select_by_visible_text("100")
                    time.sleep(7)
                    
            #### 2 possibilities now: 
            #1) max of titles is less than 100, so we just check the books, select the books then click save.
            #2) max of titles is more than 100, so for the first to before last page, we check all 100 options, click Next until 
            # last page, where we take the maximum of options available (<100) then click save.
            
            ############## OPTION 1 
                    if (max_titles <= 100):
                        
                        #### KEEP TRACK OF NUMBER OF BOOKS
                        nb_books = 0
                        pages = 1

                        for i in range (1, max_titles + 1):
                            
                            if (driver.find_element_by_xpath("//div[@id='search-results']/ul/li[" + str(i) + "]/div[3]/div/div[1]").text=="Book"):
                                nb_books = nb_books + 1
                                (driver.find_element_by_xpath("//div[@id='search-results']/ul/li[" + str(i) + "]").find_element_by_name("titles")).click()
                                time.sleep(0.2)
                                
                            else:
                                time.sleep(0.2)
                                
                        print("Number of books by this author: ", nb_books)
                     
                        
                        #### NO NEED TO SCROLL OF MAX_TITLES < 2        
                        if (max_titles > 2):
                            driver.execute_script('window.scrollTo(0,document.body.scrollHeight)'); 

                        if (nb_books != 0):
                            (driver.find_element_by_name("ExportDialogServlet")).click()
                            time.sleep(2)
                            (driver.find_element_by_name("butExport")).click()
                            alert = driver.switch_to_alert()
                            time.sleep(7)	
                            driver.quit()
                            
                        else:
                            print("No book for  author ", Sample_Errors[author])
                            Empty_Book_Set.append(Sample_Errors[author])
                            
                            try:
                                summary_error.write("\n" + "          " + Sample_Errors[author] + " -No Book- " + titles + "\n")
                                summary_empty_error.write("\n" + Sample_Errors[author] + " -No Book- " + titles + "\n")
                                
                            except UnicodeEncodeError:  
                                
                                if Sample_Errors[author] not in Unicode_Errors:
                                    Unicode_Errors.append(Sample_Errors[author])
                        
                                summary_error.write("\n" + "          " + "Author Alias # " + str(author) + " -No Book- " + titles + "\n") 
                                summary_empty_error.write("\n" + "Author Index: " + str(author) + " -No Book- " + titles + "\n")
                                summary_unicode_error.write("\n" + "Unicode Error for author index: " + str(author) + " -No Book- " + titles + "\n")
                            
                            driver.quit()
                            
            ############## OPTION 2                            
                    else: 
                    
                        if (max_titles > 399):
                            
                            #### KEEP TRACK OF LARGE SETS FOR FUTURE VERIFICATION
                            if (Sample_Errors[author] not in Large_Book_Set):
                                Large_Book_Set.append(Sample_Errors[author])
                                
                                try:
                                    summary_large_error.write("\n" + Sample_Errors[author] + " -Large Book Set- " + titles + "\n")
                                
                                except UnicodeEncodeError:
                                    
                                    if Sample_Errors[author] not in Unicode_Errors:
                                        Unicode_Errors.append(Sample_Errors[author])
                            
                                    summary_large_error.write("\n" + "Author Index: ", str(author) + " -Large Book Set- " + titles + "\n")
                                    summary_unicode_error.write("\n" + "Unicode Error for author index: " + str(author) + " -Large Book Set- " + titles + "\n")
                            
                            else:
                                print("Already in Large_Book_Set")
                                
                        pages = math.ceil(max_titles / 100) + 1
                        print("Nb of pages= " + str(pages))
                        k = 1
                        nb_books = 0
                        while (k < pages) :
                            
                            time.sleep(5)
                            print("Page: " , str(k))
                            
                            for i in range (1, 101):
                                
                                if (driver.find_element_by_xpath("//div[@id='search-results']/ul/li["+str(i)+"]/div[3]/div/div[1]").text=="Book"):
                                    nb_books = nb_books + 1; 
                                    (driver.find_element_by_xpath("//div[@id='search-results']/ul/li["+str(i)+"]").find_element_by_name("titles")).click()
                                    time.sleep(0.3)
                                    
                                else:
                                    time.sleep(0.3)
                                  
                            if (k == 4) :
                                time.sleep(0.5)
                                (driver.find_element_by_xpath("//*[@id='results-form']/div[1]/ul/li[" + str(k + 1) + "]/a")).click();
                                time.sleep(3)
                                
                            elif (k == 6):
                                time.sleep(0.5)
                                (driver.find_element_by_xpath("//*[@id='results-form']/div[1]/ul/li[" + str(k + 1) + "]/a")).click();
                                time.sleep(3)
                                
                            elif (k == 7):
                                time.sleep(0.5)
                                (driver.find_element_by_xpath("//*[@id='results-form']/div[1]/ul/li[" + str(k) + "]/a")).click();
                                time.sleep(3)

                            elif (k > 7):
                                time.sleep(0.5)
                                (driver.find_element_by_xpath("//*[@id='results-form']/div[1]/ul/li[7]/a")).click();
                                time.sleep(3)

                            else:
                                time.sleep(0.5)
                                (driver.find_element_by_xpath("//*[@id='results-form']/div[1]/ul/li[" + str(k + 2) + "]/a")).click();
                                time.sleep(3)
                          
                            k = k + 1
                            time.sleep(3)
                           
                        #### LAST PAGE
                        new = max_titles - (pages - 1) * 100
                        print("Last page:" , str(k), "Number left= ", str(new)) 
                        
                        for last_page in range (1, int(new + 1)):
                            
                            if (driver.find_element_by_xpath("//div[@id='search-results']/ul/li[" + str(last_page) + "]/div[3]/div/div[1]").text == "Book"):
                                nb_books = nb_books + 1; 
                                (driver.find_element_by_xpath("//div[@id='search-results']/ul/li[" + str(last_page) + "]").find_element_by_name("titles")).click()
                                time.sleep(0.7)
                            else:
                                time.sleep(0.7)
                                
                        print("Number of books by this author: ", nb_books)
                        
                        if new > 2 :
                            driver.execute_script('window.scrollTo(0,document.body.scrollHeight)'); 
                        
                        if (nb_books != 0):
                            (driver.find_element_by_name("ExportDialogServlet")).click()
                            (driver.find_element_by_name("butExport")).click()
                            alert = driver.switch_to_alert()
                        
                        #### THE MORE RECORDS, THE MORE TIME NEEDED
                            if (max_titles > 399):
                                time.sleep(118)	
                            else:
                                time.sleep(60)

                            driver.quit()
                            
                        else: 
                            print("No book for  author ", Sample_Errors[author])
                            Empty_Book_Set.append(Sample_Errors[author])
                            
                            try:
                                summary_error.write("\n" + "          " + Sample_Errors[author] + " -No Book- " + titles + ", Pages: " + str(pages) + "\n" )
                                summary_empty_error.write("\n" + Sample_Errors[author] + " -No Book- " + titles + ", Pages: " + str(pages) + "\n" )
                                
                            except UnicodeEncodeError:   
                                
                                if Sample_Errors[author] not in Unicode_Errors:
                                    Unicode_Errors.append(Sample_Errors[author])
                        
                                summary_error.write("\n" + "          " + "Author Alias # "+ str(author) + " -No Book- " + titles + ", Pages: " + str(pages) +  "\n" )                    
                                summary_empty_error.write("\n" + "          " + "Author Index: " + str(author) + " -No Book- " + titles + ", Pages: " + str(pages) +  "\n" )                    
                                summary_unicode_error.write("\n" + "Unicode Error for author index: " + str(author) + " -No Book- " + titles  + ", Pages: " + str(pages) + "\n")

                            driver.quit()
                     
                    
                    if (nb_books == 0):
                        print ("No file to move")
                        
                        
                         
                    else:
                    #### MOVING FILE FOR CASE MAX_TITLES > 1
                        print("Moving files: ")	
                        source_dir="/Users/apple/Downloads"
                        dest_dir="/Users/apple/Desktop/RA Final"
                        files = glob.iglob(os.path.join(source_dir, "*.mrc.part"))
                        time.sleep(3)

                        for basename in os.listdir(source_dir):

                            if basename.endswith('.mrc.part'):
                                pathname = os.path.join(source_dir, basename)

                                if os.path.isfile(pathname):
                                    new_file=Sample_Errors[author] + ".mrc"
                                    print(new_file)
                                    shutil.copy2(pathname, dest_dir + "/" + new_file)
                                    os.unlink(pathname)
                                    
                        for i in range(0, Authors.nrows - 1):
                    
                            name_sample = Author_Aliases.keys()[i]

                            for j in range (0, 17):

                                if (Sample_Errors[author] == Author_Aliases.get(name_sample)[j]):
                                    name = name_sample
                                    lc_name = j



                        for i in range (0, len(Author_ID)):

                            if (name ==Author_Name[i]):

                                print("Found it")
                                shutil.move(dest_dir + "/" + new_file, "/Users/apple/Desktop/RA Final/" + str(Author_ID[i])[0:(len(str(Author_ID[i])) - 2)] + "_" + Author_Name[i] + "/" + str(Author_ID[i])[0:(len(str(Author_ID[i])) - 2)] + "_" + str(lc_name + 1) + "_" + new_file)       
                                summary_error.write("\n" + "          --> in Author_ID #: " + str(Author_ID[i])[0:(len(str(Author_ID[i])) - 2)]  + "\n" )
                                summary_error.write("\n" + "          --> Number of pages: " + str(pages) + "\n" )                                     
                                summary_error.write("\n" + "          --> Number of books: " + str(nb_books) + "\n" ) 
                                
                                break

        #### EXCEPTIONS
        except StaleElementReferenceException:
            
            print("Stale Element Reference Exception at author: " + Sample_Errors[author] + ", let's count it as no such element exception")
            
            if (Sample_Errors[author] not in NoSuchElement_Sample):
                NoSuchElement_Sample.append(Sample_Errors[author])
                
                try:   
                    summary_error.write("\n!!!!" + "          " + Sample_Errors[author] + " -Stale Element Reference Exception- " + "\n")
                    summary_nosuchelement_error.write("\n" + Sample_Errors[author] + " -Stale Element Reference Exception- " + "\n")
             
                except UnicodeEncodeError: 
                    
                    if Sample_Errors[author] not in Unicode_Errors:
                        Unicode_Errors.append(Sample_Errors[author])
                        
                    summary_error.write("\n!!!!" + "          " + "Author Alias # "+ str(author) + " -Stale Element Reference Exception- " + "\n")
                    summary_nosuchelement_error.write("\n" + "Author Index: " + str(author) + " -Stale Element Reference Exception- " + "\n")
                    summary_unicode_error.write("\n" + "Unicode Error for author index: " + str(author) + " -Stale Element Reference Exception- " + "\n")
            else:
                print(Sample_Errors[author] + " is a duplicate for No Such Element Error")
                
            time.sleep(5) 
            driver.quit()
            time.sleep(60)
            continue

        except WebDriverException:
            
            print("WebDriver Exception at author: " + Sample_Errors[author] + ", let's count it as no such element exception")
            
            if (Sample_Errors[author] not in NoSuchElement_Sample):
                NoSuchElement_Sample.append(Sample_Errors[author])
                
                try:   
                    summary_error.write("\n!!!!" + "          " + Sample_Errors[author] + " -WebDriver Exception- " + "\n")
                    summary_nosuchelement_error.write("\n" + Sample_Errors[author] + " -WebDriver Exception- " + "\n")
                    
                except UnicodeEncodeError: 
                    
                    if Sample_Errors[author] not in Unicode_Errors:
                        Unicode_Errors.append(Sample_Errors[author])
                        
                    summary_error.write("\n!!!!" + "          " + "Author Alias # "+ str(author) + " -Web Driver Exception- " + "\n")
                    summary_unicode_error.write("\n" + "Unicode Error for author index: " + str(author) + " -Web Driver Exception- " + "\n")
                    summary_nosuchelement_error.write("\n" + "Author Index: " + str(author) + " -Web Driver Exception- " + "\n")
                    
            else:
                print(Sample_Errors[author] + " is a duplicate for No Such Element Error")
            
            time.sleep(5) 
            driver.quit()
            time.sleep(30)
            continue
                        
        except InvalidElementStateException:
            
            print("Invalid Element State Exception at author: " + Sample_Errors[author] + ", let's count it as no such element exception")
            
            if (Sample_Errors[author] not in NoSuchElement_Sample):
                NoSuchElement_Sample.append(Sample_Errors[author])
                
                try:   
                    summary_error.write("\n!!!!" + "          " + Sample_Errors[author] + " -Invalid Element State Exception Exception- " + "\n")
                    summary_nosuchelement_error.write("\n" + Sample_Errors[author] + " -Invalid Element State Exception Exception- " + "\n")
              
                except UnicodeEncodeError:
                    
                    if Sample_Errors[author] not in Unicode_Errors:
                        Unicode_Errors.append(Sample_Errors[author])
                        
                    summary_error.write("\n!!!!" + "          " + "Author Alias # "+ str(author) + " -Invalid Element State Exception Exception- " + "\n")
                    summary_unicode_error.write("\n" + "Unicode Error for author index: " + str(author) + " -Invalid Element State Exception Exception- " + "\n")
                    summary_nosuchelement_error.write("\n" + "Author Index: " + str(author) + " -Invalid Element State Exception Exception- " + "\n")
                    
            else:
                print(Sample_Errors[author] + " is a duplicate for No Such Element Error")
            
            time.sleep(5) 
            driver.quit()
            time.sleep(60)
            continue

        except NoSuchElementException:
            
            print("No Such Element Exception at author: " + Sample_Errors[author] + ", let's continue")
            
            if (Sample_Errors[author] not in NoSuchElement_Sample):
                NoSuchElement_Sample.append(Sample_Errors[author])
                
                try:   
                    summary_error.write("\n!!!!" + "          " + Sample_Errors[author] + " -No Such Element Exception Exception- " + "\n")
                    summary_nosuchelement_error.write("\n" + Sample_Errors[author] + " -No Such Element Exception Exception- " + "\n")
               
                except UnicodeEncodeError: 
                    
                    if Sample_Errors[author] not in Unicode_Errors:
                        Unicode_Errors.append(Sample_Errors[author])
                        
                    summary_error.write("\n!!!!" + "          " + "Author Alias # "+ str(author) + " -No Such Element Exception Exception- " + "\n")
                    summary_unicode_error.write("\n" + "Unicode Error for author index: " + str(author) + " -No Such Element Exception Exception- " + "\n")
                    summary_nosuchelement_error.write("\n" + "Author Index: " + str(author) + " -No Such Element Exception Exception- " + "\n")
                    
            else:
                print(Sample_Errors[author] + " is a duplicate for No Such Element Error")
            
            time.sleep(5) 
            driver.quit()
            continue
            
        except UnicodeEncodeError:
            
            print("Unicode Exception at author: " + Sample_Errors[author] + ", let's continue")
            
            if (Sample_Errors[author] not in Unicode_Errors):
                Unicode_Errors.append(Sample_Errors[author])
                
                try:   
                    summary_error.write("\n!!!!" + "          " + Sample_Errors[author] + " -Unicode Exception- " + "\n")
                    summary_unicode_error.write("\n" + Sample_Errors[author] + " -Unicode Exception- " + "\n")
                
                except UnicodeEncodeError: 
                    summary_error.write("\n!!!!" + "          " + "Author Alias # "+ str(author) + " -Unicode Exception- " + "\n")
                    summary_unicode_error.write("\n" + "Unicode Error for author index: " + str(author) + " -Unicode Exception- " + "\n")
                    
            else:
                print(Sample_Errors[author] + " is a duplicate for Unicode Error")
            
            time.sleep(5)   
            driver.quit()
            continue
            
        except IOError:
            
            print("IOError Exception at author: "+ Sample_Errors[author] + ", let's count it as no such element exception")
            
            if (empty == 1):
                empty = 0
                continue
                
            else:
                if (Sample_Errors[author] not in NoSuchElement_Sample):
                    NoSuchElement_Sample.append(Sample_Errors[author])
            
                try:   
                    summary_error.write("\n!!!!" + "          " + Sample_Errors[author] + " -IOError Exception- " + "\n")
                    summary_nosuchelement_error.write("\n" + Sample_Errors[author] + " -IOError Exception- " + "\n")
                
                except UnicodeEncodeError: 
                    
                    if Sample_Errors[author] not in Unicode_Errors:
                        Unicode_Errors.append(Sample_Errors[author])
                        
                    summary_error.write("\n!!!!" + "          " + "Author Alias # "+ str(author) + " -IOError Exception- " + "\n")
                    summary_unicode_error.write("\n" + "Unicode Error for author index: " + str(author) + " -IOError Exception- " + "\n")
                    summary_nosuchelement_error.write("\n" + "Author Index: " + str(author) + " -IOError Exception- " + "\n")
                    
                else:
                    print(Sample_Errors[author] + " is a duplicate for No Such Element Error")
            
            time.sleep(5)   
            driver.quit()
            
        except TimeoutException:
            
            print("Timeout Exception at author: " + Sample_Errors[author] + ", let's pause and count it as no such element exception")
            
            if (Sample_Errors[author] not in NoSuchElement_Sample):
                NoSuchElement_Sample.append(Sample_Errors[author])

                try:   
                    summary_error.write("\n!!!!" + "          " + Sample_Errors[author] + " -Timeout Exception- " + "\n")
                    summary_nosuchelement_error.write("\n" + Sample_Errors[author] + " -Timeout Exception- " + "\n")
               
                except UnicodeEncodeError: 
                    
                    if Sample_Errors[author] not in Unicode_Errors:
                        Unicode_Errors.append(Sample_Errors[author])
                        
                    summary_error.write("\n!!!!" + "          " + "Author Alias # "+ str(author) + " -Timeout Exception- " + "\n")
                    summary_unicode_error.write("\n" + "Unicode Error for author index: " + str(author) + " -Timeout Exception- " + "\n")
                    summary_nosuchelement_error.write("\n" + "Author Index: " + str(author) + " -Timeout Exception- " + "\n")
                
            else:
                print(Sample_Errors[author] + " is a duplicate for No Such Element Error")
            
            time.sleep(5)   
            driver.quit()
            time.sleep(120)

    NoSuchElement_Errors = NoSuchElement_Sample
    NoSuchElement_Sample = []

    print("NoSuchElement_Errors: ")
    print(NoSuchElement_Errors, len(NoSuchElement_Errors))





# # This only needs to be done if something happened that interrupted the script (something that couldn't be handled by exceptions handlers) [can be skipped]

# In[286]:

Sample_Errors = NoSuchElement_Errors    
print("Sample_Errors: ")
print(Sample_Errors, len(Sample_Errors))

for author in range(0 , len(Sample_Errors)):
    
    name = Sample_Errors[author]
    print("Author Alias #", author, ": ", name)
    time.sleep(2)
    #summary_nosuchelement.write("\n" + "For Author Alias #" + str(author) + " " + str(name) + ":" + "\n")
                
    try:        
        #### OPEN BROWSER
        try:
            driver = webdriver.Firefox()
            driver.get("https://catalog.loc.gov")
            driver.get("https://catalog.loc.gov/vwebv/searchBrowse")
            

        #### BROWSE PAGE
            search_code = Select(driver.find_element_by_id("search-code"))
            search_code.select_by_visible_text("AUTHORS/CREATORS beginning with (enter last name first)")
            driver.implicitly_wait(10) # seconds
            
            
        except WebDriverException:
            
            print ("WebDriver Exeption at ", Sample_Errors[author], "Let's wait 30 seconds.")
            time.sleep (30)
            driver.quit()
            driver = webdriver.Firefox()
            driver.get("https://catalog.loc.gov")
            driver.get("https://catalog.loc.gov/vwebv/searchBrowse")
            
            search_code = Select(driver.find_element_by_id("search-code"))
            search_code.select_by_visible_text("AUTHORS/CREATORS beginning with (enter last name first)")
            driver.implicitly_wait(10) # seconds

        empty = 0
        if (Sample_Errors[author] == "N/A"):
            summary_error.write("\n!!!!" + "          " + Sample_Errors[author] + " -Not in library- " + "\n")
            driver.quit()
            break
            
        else:    
            (driver.find_element_by_id("search-argument")).send_keys(Sample_Errors[author])
            (driver.find_element_by_name("page.search.search.button")).click()
            titles = (driver.find_element_by_class_name("search-results-browse-list-title-number")).text
            max_titles = int(titles[1:(len(titles)-1)])
            #print(max_titles)
            
            if (max_titles == 0):
                Empty_Book_Set.append(Sample_Errors[author])
                try:
                    summary_error.write("\n!!!!" + "          " + Sample_Errors[author] + " -Empty (Exception)- " + "\n")
                    summary_empty_error.write("\n" + Sample_Errors[author] + " -Empty (Exception)- " + "\n")
                        
                except UnicodeEncodeError:
                    
                    if Sample_Errors[author] not in Unicode_Errors:
                        Unicode_Errors.append(Sample_Errors[author])
                    
                    summary_error.write("\n!!!!" + "          " + "Author Alias # "+ str(author) + " -Empty (Exception)- " + "\n")
                    summary_unicode_error.write("\n" + "Unicode Error for author index: " + str(author) + " -Empty (Exception)- " + "\n")
                
                driver.quit()
                continue                        
                
            if (max_titles == 1):
                nb_books = 0
                (driver.find_element_by_css_selector("a[href*='search?searchType=7']")).click();
                try:
                    summary_error.write("\n" + "          "+ Sample_Errors[author] + " -Good- " + titles +"\n")
                
                except UnicodeEncodeError:
                    
                    if Sample_Errors[author] not in Unicode_Errors:
                        Unicode_Errors.append(Sample_Errors[author])
                    
                    summary_error.write("\n" + "          " + "Author Alias # "+ str(author) + " -Good- " +  titles + "\n")
                    summary_unicode_error.write("\n" + "Unicode Error for author index " + str(author) + " -Good- " + "\n")
                
                time.sleep(3)
                
                if(driver.find_element_by_xpath("/html/body/main/article/div[2]/h1/small").text=="BOOK"):
                    nb_books = nb_books + 1
                    (driver.find_element_by_xpath("/html/body/main/article/div[2]/div/section/div/div[2]/div/a[2]")).click()
                    (driver.find_element_by_name("butExport")).click()
                    alert = driver.switch_to_alert()
                    time.sleep(5)
                    driver.quit()
                    
                else:
                    print("No book for  author ", Sample_Errors[author])
                    Empty_Book_Set.append(Sample_Errors[author])
                    try:
                        summary_error.write("\n" + "          " + Sample_Errors[author] + " -No Book- " + titles + "\n")
                        summary_empty_error.write("\n" + Sample_Errors[author] + " -No Book- " + titles + "\n")
                        empty = 1 ## Because an IOException will be raised since there is no file to move
                    
                    except UnicodeEncodeError:
                    
                        if Sample_Errors[author] not in Unicode_Errors:
                            Unicode_Errors.append(Sample_Errors[author])
                    
                        summary_error.write("\n" + "          " + "Author Alias # "+ str(author) + " -No Book- " + titles + "\n")  
                        summary_empty_error.write("\n" + "Author Index: " + str(author) + " -No Book- " + titles + "\n")   
                        summary_unicode_error.write("\n" + "Unicode Error for author index: " + str(author) + " -No Book- " + "\n")
                        
                    driver.quit()                      
                
                if (nb_books == 0):
                    print ("No book to save for ", Sample_Errors[author])
                    
                else:
                #### MOVING FILE FOR CASE MAX_TITLES = 1
                
                   
                    print("Moving files: ") 
                    source_dir="/Users/apple/Downloads"
                    dest_dir="/Users/apple/Desktop/RA Final"
                    files = glob.iglob(os.path.join(source_dir, "*.mrc.part"))
                    time.sleep(3)

                    for basename in os.listdir(source_dir):

                        if basename.endswith('.mrc.part'):
                            pathname = os.path.join(source_dir, basename)

                            if os.path.isfile(pathname):
                                new_file = Sample_Errors[author] + ".mrc"
                                print(new_file)
                                shutil.copy2(pathname, dest_dir + "/" + new_file)
                                os.unlink(pathname)
                                

                    for i in range(0, Authors.nrows - 1):

                        name_sample=Author_Aliases.keys()[i]

                        for j in range (0, 17):

                            if (Sample_Errors[author]==Author_Aliases.get(name_sample)[j]):
                                name = name_sample
                                lc_name = j
                                print (lc_name)

                    for i in range (0, len(Author_ID)):

                        if (name == Author_Name[i]):
                            print("Found it")
                            shutil.move(dest_dir + "/" + new_file, "/Users/apple/Desktop/RA Final/" + str(Author_ID[i])[0:(len(str(Author_ID[i])) - 2)] + "_" + Author_Name[i] + "/" + str(Author_ID[i])[0:(len(str(Author_ID[i])) - 2)] + "_" + str(lc_name + 1) + "_" + new_file )
                            summary_error.write("\n" + "          --> in Author_ID #: " + str(Author_ID[i])[0:(len(str(Author_ID[i])) - 2)] + "\n" )
                            summary_error.write("\n" + "          --> Number of pages: 1" + "\n" )                                     
                            summary_error.write("\n" + "          --> Number of books #: " + str(nb_books) + "\n" )
                            nb_books = 0


            #### IF MAX_TITLES > 2
            else:   
                (driver.find_element_by_css_selector("a[href*='search?searchType=7']")).click();
                try:
                    summary_error.write("\n" + "          " + Sample_Errors[author] + " -Good- " + titles + "\n")
                
                except UnicodeEncodeError: 
                    
                    if Sample_Errors[author] not in Unicode_Errors:
                        Unicode_Errors.append(Sample_Errors[author])
                    
                    summary_error.write("\n" + "          " + "Author Index: " + str(author) + " -Good- " +  titles + "\n")                    
                    summary_unicode_error.write("\n" + "Unicode Error for author index: " + str(author) + " -Good- " + "\n")

        #### SHOULD MAXIMIZE WINDOW AND MAXIMIZE NUMBER OF RECORDS PER PAGE FOR CONVENIENCE
                time.sleep(3)
                driver.maximize_window()
                time.sleep(2)
                record = Select(driver.find_element_by_id("record-count"))
                record.select_by_visible_text("100")
                time.sleep(7)
                
        #### 2 possibilities now: 
        #1) max of titles is less than 100, so we just check the books, select the books then click save.
        #2) max of titles is more than 100, so for the first to before last page, we check all 100 options, click Next until 
        # last page, where we take the maximum of options available (<100) then click save.
        
        ############## OPTION 1 
                if (max_titles <= 100):
                    
                    #### KEEP TRACK OF NUMBER OF BOOKS
                    nb_books = 0
                    pages = 1

                    for i in range (1, max_titles + 1):
                        
                        if (driver.find_element_by_xpath("//div[@id='search-results']/ul/li[" + str(i) + "]/div[3]/div/div[1]").text=="Book"):
                            nb_books = nb_books + 1
                            (driver.find_element_by_xpath("//div[@id='search-results']/ul/li[" + str(i) + "]").find_element_by_name("titles")).click()
                            time.sleep(0.2)
                            
                        else:
                            time.sleep(0.2)
                            
                    print("Number of books by this author: ", nb_books)
                 
                    
                    #### NO NEED TO SCROLL OF MAX_TITLES < 2        
                    if (max_titles > 2):
                        driver.execute_script('window.scrollTo(0,document.body.scrollHeight)'); 

                    if (nb_books != 0):
                        (driver.find_element_by_name("ExportDialogServlet")).click()
                        time.sleep(2)
                        (driver.find_element_by_name("butExport")).click()
                        alert = driver.switch_to_alert()
                        time.sleep(7)   
                        driver.quit()
                        
                    else:
                        print("No book for  author ", Sample_Errors[author])
                        Empty_Book_Set.append(Sample_Errors[author])
                        
                        try:
                            summary_error.write("\n" + "          " + Sample_Errors[author] + " -No Book- " + titles + "\n")
                            summary_empty_error.write("\n" + Sample_Errors[author] + " -No Book- " + titles + "\n")
                            
                        except UnicodeEncodeError:  
                            
                            if Sample_Errors[author] not in Unicode_Errors:
                                Unicode_Errors.append(Sample_Errors[author])
                    
                            summary_error.write("\n" + "          " + "Author Alias # " + str(author) + " -No Book- " + titles + "\n") 
                            summary_empty_error.write("\n" + "Author Index: " + str(author) + " -No Book- " + titles + "\n")
                            summary_unicode_error.write("\n" + "Unicode Error for author index: " + str(author) + " -No Book- " + titles + "\n")
                        
                        driver.quit()
                        
        ############## OPTION 2                            
                else: 
                
                    if (max_titles > 399):
                        
                        #### KEEP TRACK OF LARGE SETS FOR FUTURE VERIFICATION
                        if (Sample_Errors[author] not in Large_Book_Set):
                            Large_Book_Set.append(Sample_Errors[author])
                            
                            try:
                                summary_large_error.write("\n" + Sample_Errors[author] + " -Large Book Set- " + titles + "\n")
                            
                            except UnicodeEncodeError:
                                
                                if Sample_Errors[author] not in Unicode_Errors:
                                    Unicode_Errors.append(Sample_Errors[author])
                        
                                summary_large_error.write("\n" + "Author Index: ", str(author) + " -Large Book Set- " + titles + "\n")
                                summary_unicode_error.write("\n" + "Unicode Error for author index: " + str(author) + " -Large Book Set- " + titles + "\n")
                        
                        else:
                            print("Already in Large_Book_Set")
                            
                    pages = math.ceil(max_titles / 100) + 1
                    print("Nb of pages= " + str(pages))
                    k = 1
                    nb_books = 0
                    while (k < 101) :
                        
                        time.sleep(5)
                        print("Page: " , str(k))
                        
                        for i in range (1, 101):
                            
                            if (driver.find_element_by_xpath("//div[@id='search-results']/ul/li["+str(i)+"]/div[3]/div/div[1]").text=="Book"):
                                nb_books = nb_books + 1; 
                                (driver.find_element_by_xpath("//div[@id='search-results']/ul/li["+str(i)+"]").find_element_by_name("titles")).click()
                                time.sleep(0.3)
                                
                            else:
                                time.sleep(0.3)
                              
                        if (k == 4) :
                            time.sleep(0.5)
                            (driver.find_element_by_xpath("//*[@id='results-form']/div[1]/ul/li[" + str(k + 1) + "]/a")).click();
                            time.sleep(3)

                        elif (k == 6):
                            time.sleep(0.5)
                            (driver.find_element_by_xpath("//*[@id='results-form']/div[1]/ul/li[" + str(k + 1) + "]/a")).click();
                            time.sleep(3)

                        elif (k == 7):
                            time.sleep(0.5)
                            (driver.find_element_by_xpath("//*[@id='results-form']/div[1]/ul/li[" + str(k) + "]/a")).click();
                            time.sleep(3)

                        elif (k > 7):
                            time.sleep(0.5)
                            (driver.find_element_by_xpath("//*[@id='results-form']/div[1]/ul/li[7]/a")).click();
                            time.sleep(3)

                        else:
                            time.sleep(0.5)
                            (driver.find_element_by_xpath("//*[@id='results-form']/div[1]/ul/li[" + str(k + 2) + "]/a")).click();
                            time.sleep(3)
                      
                        k = k + 1
                        time.sleep(3)
                       
                    #### LAST PAGE
                    new = max_titles - (pages - 1) * 100
                    print("Last page:" , str(k), "Number left= ", str(new)) 
                    
                    for last_page in range (1, int(new + 1)):
                        
                        if (driver.find_element_by_xpath("//div[@id='search-results']/ul/li[" + str(last_page) + "]/div[3]/div/div[1]").text == "Book"):
                            nb_books = nb_books + 1; 
                            (driver.find_element_by_xpath("//div[@id='search-results']/ul/li[" + str(last_page) + "]").find_element_by_name("titles")).click()
                            time.sleep(0.7)
                        else:
                            time.sleep(0.7)
                            
                    print("Number of books by this author: ", nb_books)
                    driver.execute_script('window.scrollTo(0,document.body.scrollHeight)'); 
                    
                    if (nb_books != 0):
                        (driver.find_element_by_name("ExportDialogServlet")).click()
                        (driver.find_element_by_name("butExport")).click()
                        alert = driver.switch_to_alert()
                    
                    #### THE MORE RECORDS, THE MORE TIME NEEDED
                        if (max_titles > 399):
                            time.sleep(118) 
                        else:
                            time.sleep(60)

                        driver.quit()
                        
                    else: 
                        print("No book for  author ", Sample_Errors[author])
                        Empty_Book_Set.append(Sample_Errors[author])
                        
                        try:
                            summary_error.write("\n" + "          " + Sample_Errors[author] + " -No Book- " + titles + ", Pages: " + str(pages) + "\n" )
                            summary_empty_error.write("\n" + Sample_Errors[author] + " -No Book- " + titles + ", Pages: " + str(pages) + "\n" )
                            
                        except UnicodeEncodeError:   
                            
                            if Sample_Errors[author] not in Unicode_Errors:
                                Unicode_Errors.append(Sample_Errors[author])
                    
                            summary_error.write("\n" + "          " + "Author Alias # "+ str(author) + " -No Book- " + titles + ", Pages: " + str(pages) +  "\n" )                    
                            summary_empty_error.write("\n" + "          " + "Author Index: " + str(author) + " -No Book- " + titles + ", Pages: " + str(pages) +  "\n" )                    
                            summary_unicode_error.write("\n" + "Unicode Error for author index: " + str(author) + " -No Book- " + titles  + ", Pages: " + str(pages) + "\n")

                        driver.quit()
                 
                
                if (nb_books == 0):
                    print ("No file to move")
                    
                    
                     
                else:
                #### MOVING FILE FOR CASE MAX_TITLES > 1
                    print("Moving files: ") 
                    source_dir="/Users/apple/Downloads"
                    dest_dir="/Users/apple/Desktop/RA Final"
                    files = glob.iglob(os.path.join(source_dir, "*.mrc.part"))
                    time.sleep(3)

                    for basename in os.listdir(source_dir):

                        if basename.endswith('.mrc.part'):
                            pathname = os.path.join(source_dir, basename)

                            if os.path.isfile(pathname):
                                new_file=Sample_Errors[author] + ".mrc"
                                print(new_file)
                                shutil.copy2(pathname, dest_dir + "/" + new_file)
                                os.unlink(pathname)
                                
                    for i in range(0, Authors.nrows - 1):
                
                        name_sample = Author_Aliases.keys()[i]

                        for j in range (0, 17):

                            if (Sample_Errors[author] == Author_Aliases.get(name_sample)[j]):
                                name = name_sample
                                lc_name = j



                    for i in range (0, len(Author_ID)):

                        if (name ==Author_Name[i]):

                            print("Found it")
                            shutil.move(dest_dir + "/" + new_file, "/Users/apple/Desktop/RA Final/" + str(Author_ID[i])[0:(len(str(Author_ID[i])) - 2)] + "_" + Author_Name[i] + "/" + str(Author_ID[i])[0:(len(str(Author_ID[i])) - 2)] + "_" + str(lc_name + 1) + "_" + new_file)       
                            summary_error.write("\n" + "          --> in Author_ID #: " + str(Author_ID[i])[0:(len(str(Author_ID[i])) - 2)]  + "\n" )
                            summary_error.write("\n" + "          --> Number of pages: " + str(pages) + "\n" )                                     
                            summary_error.write("\n" + "          --> Number of books: " + str(nb_books) + "\n" ) 
                            
                            break

    #### EXCEPTIONS
    except StaleElementReferenceException:
        
        print("Stale Element Reference Exception at author: " + Sample_Errors[author] + ", let's count it as no such element exception")
        
        if (Sample_Errors[author] not in NoSuchElement_Sample):
            NoSuchElement_Sample.append(Sample_Errors[author])
            
            try:   
                summary_error.write("\n!!!!" + "          " + Sample_Errors[author] + " -Stale Element Reference Exception- " + "\n")
                summary_nosuchelement_error.write("\n" + Sample_Errors[author] + " -Stale Element Reference Exception- " + "\n")
         
            except UnicodeEncodeError: 
                
                if Sample_Errors[author] not in Unicode_Errors:
                    Unicode_Errors.append(Sample_Errors[author])
                    
                summary_error.write("\n!!!!" + "          " + "Author Alias # "+ str(author) + " -Stale Element Reference Exception- " + "\n")
                summary_nosuchelement_error.write("\n" + "Author Index: " + str(author) + " -Stale Element Reference Exception- " + "\n")
                summary_unicode_error.write("\n" + "Unicode Error for author index: " + str(author) + " -Stale Element Reference Exception- " + "\n")
        else:
            print(Sample_Errors[author] + " is a duplicate for No Such Element Error")
            
        time.sleep(5) 
        driver.quit()
        time.sleep(60)
        continue

    except WebDriverException:
        
        print("WebDriver Exception at author: " + Sample_Errors[author] + ", let's count it as no such element exception")
        
        if (Sample_Errors[author] not in NoSuchElement_Sample):
            NoSuchElement_Sample.append(Sample_Errors[author])
            
            try:   
                summary_error.write("\n!!!!" + "          " + Sample_Errors[author] + " -WebDriver Exception- " + "\n")
                summary_nosuchelement_error.write("\n" + Sample_Errors[author] + " -WebDriver Exception- " + "\n")
                
            except UnicodeEncodeError: 
                
                if Sample_Errors[author] not in Unicode_Errors:
                    Unicode_Errors.append(Sample_Errors[author])
                    
                summary_error.write("\n!!!!" + "          " + "Author Alias # "+ str(author) + " -Web Driver Exception- " + "\n")
                summary_unicode_error.write("\n" + "Unicode Error for author index: " + str(author) + " -Web Driver Exception- " + "\n")
                summary_nosuchelement_error.write("\n" + "Author Index: " + str(author) + " -Web Driver Exception- " + "\n")
                
        else:
            print(Sample_Errors[author] + " is a duplicate for No Such Element Error")
        
        time.sleep(5) 
        driver.quit()
        time.sleep(30)
        continue
                    
    except InvalidElementStateException:
        
        print("Invalid Element State Exception at author: " + Sample_Errors[author] + ", let's count it as no such element exception")
        
        if (Sample_Errors[author] not in NoSuchElement_Sample):
            NoSuchElement_Sample.append(Sample_Errors[author])
            
            try:   
                summary_error.write("\n!!!!" + "          " + Sample_Errors[author] + " -Invalid Element State Exception Exception- " + "\n")
                summary_nosuchelement_error.write("\n" + Sample_Errors[author] + " -Invalid Element State Exception Exception- " + "\n")
          
            except UnicodeEncodeError:
                
                if Sample_Errors[author] not in Unicode_Errors:
                    Unicode_Errors.append(Sample_Errors[author])
                    
                summary_error.write("\n!!!!" + "          " + "Author Alias # "+ str(author) + " -Invalid Element State Exception Exception- " + "\n")
                summary_unicode_error.write("\n" + "Unicode Error for author index: " + str(author) + " -Invalid Element State Exception Exception- " + "\n")
                summary_nosuchelement_error.write("\n" + "Author Index: " + str(author) + " -Invalid Element State Exception Exception- " + "\n")
                
        else:
            print(Sample_Errors[author] + " is a duplicate for No Such Element Error")
        
        time.sleep(5) 
        driver.quit()
        time.sleep(60)
        continue

    except NoSuchElementException:
        
        print("No Such Element Exception at author: " + Sample_Errors[author] + ", let's continue")
        
        if (Sample_Errors[author] not in NoSuchElement_Sample):
            NoSuchElement_Sample.append(Sample_Errors[author])
            
            try:   
                summary_error.write("\n!!!!" + "          " + Sample_Errors[author] + " -No Such Element Exception Exception- " + "\n")
                summary_nosuchelement_error.write("\n" + Sample_Errors[author] + " -No Such Element Exception Exception- " + "\n")
           
            except UnicodeEncodeError: 
                
                if Sample_Errors[author] not in Unicode_Errors:
                    Unicode_Errors.append(Sample_Errors[author])
                    
                summary_error.write("\n!!!!" + "          " + "Author Alias # "+ str(author) + " -No Such Element Exception Exception- " + "\n")
                summary_unicode_error.write("\n" + "Unicode Error for author index: " + str(author) + " -No Such Element Exception Exception- " + "\n")
                summary_nosuchelement_error.write("\n" + "Author Index: " + str(author) + " -No Such Element Exception Exception- " + "\n")
                
        else:
            print(Sample_Errors[author] + " is a duplicate for No Such Element Error")
        
        time.sleep(5) 
        driver.quit()
        continue
        
    except UnicodeEncodeError:
        
        print("Unicode Exception at author: " + Sample_Errors[author] + ", let's continue")
        
        if (Sample_Errors[author] not in Unicode_Errors):
            Unicode_Errors.append(Sample_Errors[author])
            
            try:   
                summary_error.write("\n!!!!" + "          " + Sample_Errors[author] + " -Unicode Exception- " + "\n")
                summary_unicode_error.write("\n" + Sample_Errors[author] + " -Unicode Exception- " + "\n")
            
            except UnicodeEncodeError: 
                summary_error.write("\n!!!!" + "          " + "Author Alias # "+ str(author) + " -Unicode Exception- " + "\n")
                summary_unicode_error.write("\n" + "Unicode Error for author index: " + str(author) + " -Unicode Exception- " + "\n")
                
        else:
            print(Sample_Errors[author] + " is a duplicate for Unicode Error")
        
        time.sleep(5)   
        driver.quit()
        continue
        
    except IOError:
        
        print("IOError Exception at author: "+ Sample_Errors[author] + ", let's count it as no such element exception")
        
        if (empty == 1):
            empty = 0
            continue
            
        else:
            if (Sample_Errors[author] not in NoSuchElement_Sample):
                NoSuchElement_Sample.append(Sample_Errors[author])
        
            try:   
                summary_error.write("\n!!!!" + "          " + Sample_Errors[author] + " -IOError Exception- " + "\n")
                summary_nosuchelement_error.write("\n" + Sample_Errors[author] + " -IOError Exception- " + "\n")
            
            except UnicodeEncodeError: 
                
                if Sample_Errors[author] not in Unicode_Errors:
                    Unicode_Errors.append(Sample_Errors[author])
                    
                summary_error.write("\n!!!!" + "          " + "Author Alias # "+ str(author) + " -IOError Exception- " + "\n")
                summary_unicode_error.write("\n" + "Unicode Error for author index: " + str(author) + " -IOError Exception- " + "\n")
                summary_nosuchelement_error.write("\n" + "Author Index: " + str(author) + " -IOError Exception- " + "\n")
                
            else:
                print(Sample_Errors[author] + " is a duplicate for No Such Element Error")
        
        time.sleep(5)   
        driver.quit()
        
    except TimeoutException:
        
        print("Timeout Exception at author: " + Sample_Errors[author] + ", let's pause and count it as no such element exception")
        
        if (Sample_Errors[author] not in NoSuchElement_Sample):
            NoSuchElement_Sample.append(Sample_Errors[author])

            try:   
                summary_error.write("\n!!!!" + "          " + Sample_Errors[author] + " -Timeout Exception- " + "\n")
                summary_nosuchelement_error.write("\n" + Sample_Errors[author] + " -Timeout Exception- " + "\n")
           
            except UnicodeEncodeError: 
                
                if Sample_Errors[author] not in Unicode_Errors:
                    Unicode_Errors.append(Sample_Errors[author])
                    
                summary_error.write("\n!!!!" + "          " + "Author Alias # "+ str(author) + " -Timeout Exception- " + "\n")
                summary_unicode_error.write("\n" + "Unicode Error for author index: " + str(author) + " -Timeout Exception- " + "\n")
                summary_nosuchelement_error.write("\n" + "Author Index: " + str(author) + " -Timeout Exception- " + "\n")
            
        else:
            print(Sample_Errors[author] + " is a duplicate for No Such Element Error")
        
        time.sleep(5)   
        driver.quit()
        time.sleep(120)

NoSuchElement_Errors = NoSuchElement_Sample
NoSuchElement_Sample = []

print("NoSuchElement_Errors: ")
print(NoSuchElement_Errors, len(NoSuchElement_Errors))


# In[287]:

print(Large_Book_Set, len(Large_Book_Set))


# # When Finished, Check All Relevant Arrays

# In[21]:

print ("NoSuchElement_Sample: ", NoSuchElement_Sample)
print ("NoSuchElement_Errors: ", NoSuchElement_Errors)
print ("Unicode_Errors: ", Unicode_Errors)
print ("Unicode_Auth_Names_Errors: ", Unicode_Auth_Names_Errors)
print ("Empty_Book_Set: ", Empty_Book_Set)
print ("Large_Book_Set: ", Large_Book_Set)


# # Moving Done Files to Folder "Complete"

# In[ ]:

for i in range (0, len(Author_ID)):
    try:
        if  os.listdir("/Users/apple/Desktop/RA Final/" + str(Author_ID[i])[0:(len(str(Author_ID[i])) - 2)] + "_" + Author_Name[i]):
            print("Ok to move " + Author_Name[i])
            shutil.move("/Users/apple/Desktop/RA Final/" + str(Author_ID[i])[0:(len(str(Author_ID[i])) - 2)]+ "_" + Author_Name[i], "/Users/apple/Desktop/RA Final/Complete" )
    except OSError:
        continue


# # Creating New Dataset

# In[40]:



import os
from pymarc import MARCReader
import xlwt

workbook1 = xlwt.Workbook(encoding="utf-8")
#workbook2 = xlwt.Workbook(encoding="utf-8")



sheet1 = workbook1.add_sheet("Sheet 1", cell_overwrite_ok = True)
#sheet2 = workbook2.add_sheet("Sheet 2", cell_overwrite_ok = True)


sheet1.write(0, 0, "Author_ID", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 1, "Author_Name", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 2, "LC_Order", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 3, "LC_ID", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 4, "LC_Name", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 5, "LCCN", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 6, "OCLC", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 7, "ISBNall", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 8, "ISBNcor", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 9, "ISBNinc", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 10, "ISBNqual", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 11, "ISSN", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 12, "LCDew", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 13, "AuthAll", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 14, "AuthName", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 15, "AuthNum", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 16, "AuthTitle", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 17, "AuthYears", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 18, "AuthRole", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 19, "AuthPubDate", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 20, "AuthQual", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 21, "AddPerson1", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 22, "AddPerson2", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 23, "AddPerson3", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 24, "AddPerson4", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 25, "AddPerson5", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 26, "AddPerson6", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 27, "AddPerson7", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 28, "AddPerson8", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 29, "AddPerson9", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 30, "AddPerson10", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 31, "AddPerson11", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 32, "AddPerson12", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 33, "AddPerson13", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 34, "AddPerson14", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 35, "AddPerson15", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 36, "AddPerson16", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 37, "AddPerson17", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 38, "AddPerson18", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 39, "AddPerson19", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 40, "AddPerson20", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 41, "AddPerson21", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 42, "AddPerson22", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 43, "AddPerson23", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 44, "AddPerson24", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 45, "AddPerson25", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 46, "AddPerson26", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 47, "AddPerson27", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 48, "AddPerson28", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 49, "AddPerson29", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 50, "AddPerson30", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 51, "AddPerson31", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 52, "AddPerson32", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 53, "AddPerson33", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 54, "AddPerson34", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 55, "AddPerson35", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 56, "AddPerson36", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 57, "AddPerson37", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 58, "AddPerson38", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 59, "AddPerson39", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 60, "AddPerson40", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 61, "AddPerson41", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 62, "AddPerson42", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 63, "AddPerson43", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 64, "AddPerson44", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 65, "AddPerson45", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 66, "AddPerson46", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 67, "AddPerson47", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 68, "AddPerson48", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 69, "AddPerson49", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 70, "AddPerson50", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 71, "AddPerson51", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 72, "AddPerson52", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 73, "AddPerson53", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 74, "AddPerson54", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 75, "AddPerson55", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 76, "AddPerson56", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 77, "AddPerson57", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 78, "AddPerson58", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 79, "AddPerson59", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 80, "AddPerson60", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 81, "AddPerson61", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 82, "AddPerson62", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 83, "AddPerson63", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 84, "AddPerson64", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 85, "AddPerson65", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 86, "AddPerson66", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 87, "AddPerson67", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 88, "AddPerson68", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 89, "AddPerson69", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 90, "AddPerson70", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 91, "AddPerson71", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 92, "AddPerson72", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 93, "AddPerson73", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 94, "AddPerson74", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 95, "AddPerson75", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 96, "AddPerson76", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 97, "AddPerson77", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 98, "AddPerson78", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 99, "AddPerson79", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 100, "AddPerson80", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 101, "AddPerson81", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 102, "AddPerson82", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 103, "AddPerson83", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 104, "AddPerson84", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 105, "AddPerson85", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 106, "AddPerson86", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 107, "AddPerson87", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 108, "AddPerson88", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 109, "AddPerson89", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 110, "AddPerson90", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 111, "AddPerson91", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 112, "AddPerson92", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 113, "AddPerson93", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 114, "AddPerson94", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 115, "AddPerson95", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 116, "AddPerson96", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 117, "AddPerson97", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 118, "AddPerson98", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 119, "AddPerson99", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 120, "AddPerson100", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 121, "AddPerson101", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 122, "AddPerson102", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 123, "AddPerson103", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 124, "AddPerson104", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 125, "AddPerson105", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 126, "TitleAll", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 127, "TitleMain", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 128, "TitleSub", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 129, "TitleBy", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 130, "TitleMaterial", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 131, "Edition", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 132, "PubAll", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 133, "PubPlace", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 134, "PubPub", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 135, "PubYear", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 136, "ContAll", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 137, "ContPgVol", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 138, "ContIllus", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 139, "ContSize", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 140, "Copyright", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 141, "MetaTitle", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 142, "AltTitle", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 143, "PrevTitle", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 144, "Language", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 145, "RDAcont", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 146, "RDAmedia", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 147, "RDAcarrier", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 148, "Material", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 149, "Date Trans", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 150, "ContNum", xlwt.easyxf('alignment: horiz center;'))
sheet1.write(0, 151, "GeoInfo", xlwt.easyxf('alignment: horiz center;'))

total_index = 0
record_index = 0

for i in range (0, 6996):
    
    if (os.listdir("/Users/apple/Desktop/RA Final/" + str(Author_ID[i])[0 : (len(str(Author_ID[i])) - 2)] + "_" + str(Author_Name[i]))):
        
        for author in range(0, len(Author_Aliases)):             
                            
                if (Author_Name[i] == Author_Aliases.keys()[author]):
                    name = Author_Aliases.keys()[author]
                    print (name)

                    for j in range (0, 17):
                        if (os.path.exists("/Users/apple/Desktop/RA Final/" + str(Author_ID[i])[0 : (len(str(Author_ID[i])) - 2)] + "_" + Author_Name[i] + "/" + str(Author_ID[i])[0 : (len(str(Author_ID[i])) - 2)] + "_" + str(j + 1) + "_" + Author_Aliases.get(name)[j] + ".mrc")):
                            #print ("Yes, ", Author_Name[i])
                            #print (Author_Aliases.get(name)[j])
                            
                            with open("/Users/apple/Desktop/RA Final/" + str(Author_ID[i])[0 : (len(str(Author_ID[i])) - 2)] + "_" + Author_Name[i] + "/" + str(Author_ID[i])[0 : (len(str(Author_ID[i])) - 2)] + "_" + str(j + 1) + "_" + Author_Aliases.get(name)[j] + ".mrc", "rb") as fh:
                                reader = MARCReader(fh)
                                
                                record_index = 0
                                
                                List = [ "Author_ID", "Author_Name", "LC_Order", "LC_ID", "LC_Name",
                                        "LCCN", "OCLC", "ISBNall", "ISBNcor", "ISBNinc",  "ISBNqual",
                                        "ISSN",  "LCDew",  "AuthAll",  "AuthName",  "AuthNum",  "AuthTitle",
                                        "AuthYears",  "AuthRole",  "AuthPubDate",  "AuthQual",  "AddPerson1",
                                        "AddPerson2",  "AddPerson3",  "AddPerson4",  "AddPerson5",  "AddPerson6",
                                        "AddPerson7",  "AddPerson8",  "AddPerson9",  "AddPerson10",  "AddPerson11",
                                        "AddPerson12",  "AddPerson13",  "AddPerson14",  "AddPerson15",  "AddPerson16",
                                        "AddPerson17",  "AddPerson18",  "AddPerson19",  "AddPerson20",  "AddPerson21",
                                        "AddPerson22",  "AddPerson23",  "AddPerson24",  "AddPerson25",  "AddPerson26",
                                        "AddPerson27",  "AddPerson28",  "AddPerson29",  "AddPerson30",  "AddPerson31",
                                        "AddPerson32",  "AddPerson33",  "AddPerson34",  "AddPerson35",  "AddPerson36",
                                        "AddPerson37",  "AddPerson38",  "AddPerson39",  "AddPerson40",  "AddPerson41",
                                        "AddPerson42",  "AddPerson43",  "AddPerson44",  "AddPerson45",  "AddPerson46",
                                        "AddPerson47",  "AddPerson48",  "AddPerson49",  "AddPerson50",  "AddPerson51",
                                        "AddPerson52",  "AddPerson53",  "AddPerson54",  "AddPerson55",  "AddPerson56",
                                        "AddPerson57",  "AddPerson58",  "AddPerson59",  "AddPerson60",  "AddPerson61",   
                                        "AddPerson62",  "AddPerson63",  "AddPerson64",  "AddPerson65",  "AddPerson66",
                                        "AddPerson67",  "AddPerson68",  "AddPerson69",  "AddPerson70",  "AddPerson71",
                                        "AddPerson72",  "AddPerson73",  "AddPerson74",  "AddPerson75",  "AddPerson76",
                                        "AddPerson77",  "AddPerson78",  "AddPerson79",  "AddPerson80",  "AddPerson81",
                                        "AddPerson82",  "AddPerson83",  "AddPerson84",  "AddPerson85",  "AddPerson86",
                                        "AddPerson87",  "AddPerson88",  "AddPerson89",  "AddPerson90",  "AddPerson91",
                                        "AddPerson92",  "AddPerson93",  "AddPerson94",  "AddPerson95",  "AddPerson96",
                                        "AddPerson97",  "AddPerson98",  "AddPerson99",  "AddPerson100",  "AddPerson101",
                                        "AddPerson102",  "AddPerson103",  "AddPerson104",  "AddPerson105", "TitleAll",  "TitleMain",  "TitleSub",  "TitleBy",
                                        "TitleMaterial",  "Edition",  "PubAll",  "PubPlace",  "PubPub",  "PubYear",  "ContAll",
                                        "ContPgVol",  "ContIllus",  "ContSize",  "Copyright",  "MetaTitle",  "AltTitle",  "PrevTitle",
                                        "Language",  "RDAcont",  "RDAmedia",  "RDAcarrier",  "Material", "DateTrans" , "ContNum", "GeoInfo"]

                                print ("INDEX IS: ", record_index, total_index)
                                 
                                for record in reader:
                                    record_index = record_index + 1
                                    print (record_index)
                                    List[0] = str(Author_ID[i])[0 : (len(str(Author_ID[i])) - 2)]
                                    List[1] = Author_Name[i]
                                    List[2] = j + 1
                                    
                                    for k in range (0, len(Auth_LC_Name)):
                                        if Author_Aliases.get(name)[j] == Auth_LC_Name[k]:
                                            List [3] = Auth_LC[Auth_LC_Name[k]]
                                    List[4] = Author_Aliases.get(name)[j]
                                    
                                    try:
                                        List[5] = record['010']

                                    except TypeError:
                                        List[5] = "None"

                                    try:
                                        List[6] = record['035']

                                    except TypeError:
                                        List[6] = "None"


                                    try:
                                        List[7] = record['020']
                                    except TypeError:
                                        List[7] = "None"

                                    try:
                                        List[8] = record['020']['a']
                                    except TypeError:
                                        List[8] = "None"

                                    try:
                                        List[9] = record['020']['z']
                                    except TypeError:
                                        List[9] = "None"

                                    try:
                                        List[10] = record['020']['q']
                                    except TypeError:
                                        List[10] = "None"

                                    try:
                                        List[11] = record['022']
                                    except TypeError:
                                        List[11] = "None"

                                    try:
                                        List[12] = record['050']
                                    except TypeError:
                                        List[12] = "None"

                                    try:
                                        List[13] = record['100']
                                    except TypeError:
                                        List[13] = "None"

                                    try:
                                        List[14] = record['100']['a']
                                    except TypeError:
                                        List[14] = "None"

                                    try:
                                        List[15] = record['100']['b']
                                    except TypeError:
                                         List[15] = "None"

                                    try:
                                        List[16] = record['100']['c']
                                    except TypeError:
                                        List[16] = "None"

                                    try:
                                        List[17] = record['100']['d']
                                    except TypeError:
                                        List[17] = "None"

                                    try:
                                        List[18] = record['100']['e']
                                    except TypeError:
                                        List[18] = "None"

                                    try:
                                        List[19] = record['100']['f']
                                    except TypeError:
                                        List[19] = "None"

                                    try:
                                        List[20] = record['100']
                                    except TypeError:
                                        List[20] = "None"
                                            
                                    rec_700s = record.get_fields('700')
                                    
                                    if (len(rec_700s) != 0):
                                        n = 0
                                        for rec_700 in rec_700s:
                                            try:
                                                print (21 + n)
                                                List[21 + n] = rec_700
                                                
                                            except TypeError:
                                                List[21 + n]= "None"
                                            n = n + 1

                                            if (n == len(rec_700s)):
                                                for k in range (0, 105 - n ):
                                                    List[21 + n + k] = "None"
                                                break
                                        rec_700s = None
                                    else:
                                        for k in range (21, 126):
                                            List[k] = "None"
                                    try:
                                        List[126] = record['245']
                                    except TypeError:
                                        List[126] = "None"

                                    try:
                                        List[127] = record['245']['a']
                                    except TypeError:
                                        List[127] = "None"

                                    try:
                                        List[128] = record['245']['b']
                                    except TypeError:
                                        List[128] = "None"

                                    try:
                                        List[129] = record['245']['c']
                                    except TypeError:
                                        List[129] = "None"

                                    try:
                                        List[130] = record['245']['h']
                                    except TypeError:
                                        List[130] = "None"

                                    try:
                                        List[131] = record['250']
                                    except TypeError:
                                        List[131] = "None"

                                    try:
                                        List[132] = record['260']
                                    except TypeError:
                                        List[132] = "None"

                                    try:
                                        List[133] = record['260']['a']
                                    except TypeError:
                                        List[133] = "None"

                                    try:
                                        List[134] = record['260']['b']
                                    except TypeError:
                                        List[134] = "None"

                                    try:
                                        List[135] = record['260']['c']
                                    except TypeError:
                                        List[135] = "None"

                                    try:
                                        List[136] = record['300']
                                    except TypeError:
                                        List[136] = "None"

                                    try:
                                        List[137] = record['300']['a']
                                    except TypeError:
                                        List[137] = "None"

                                    try:
                                        List[138] = record['300']['b']
                                    except TypeError:
                                        List[138] = "None"

                                    try:
                                        List[139] = record['300']['c']
                                    except TypeError:
                                        List[139] = "None"

                                    try:
                                        List[140] = record['264']['c']
                                    except TypeError:
                                        List[140] = "None"

                                    try:
                                        List[141] = record['130']
                                    except TypeError:
                                        List[141] = "None"

                                    try:
                                        List[142] = record['246']['a']
                                    except TypeError:
                                        List[142] = "None"

                                    try:
                                        List[143] = record['247']
                                    except TypeError:
                                        List[143] = "None"

                                    try:
                                        List[144] = record['546']
                                    except TypeError:
                                        List[144] = "None"

                                    try:
                                        List[145] = record['336']
                                    except TypeError:
                                        List[145] = "None"

                                    try:
                                        List[146] = record['337']
                                    except TypeError:
                                        List[146] = "None"

                                    try:
                                        List[147] = record['338']
                                    except TypeError:
                                        List[147] = "None"

                                    try:
                                        List[148] = record['006']  
                                    except TypeError:
                                        List[148] = "None"
                                        
                                    try:
                                        List[149] = record['005']  
                                    except TypeError:
                                        List[149] = "None"                                        

                                    try:
                                        List[150] = record['001']  
                                    except TypeError:
                                        List[150] = "None"
                                        
                                        
                                    try:
                                        List[151] = record['651']  
                                    except TypeError:
                                        List[151] = "None"                                        
                                        
                                        
                                        
                                    for index in range (0, 152):
                                        print (total_index + record_index, index)
                                        new = total_index + record_index
                                        list_index = str(List[index]).encode('ascii', 'ignore').decode('ascii')
                                        
                                        try:
                                            sheet1.write(new, index, list_index)
                                            #sheet2.write(new, index, str(rec_700s).encode('ascii', 'ignore').decode('ascii'))
                                        #except ValueError:
                                         #   sheet2.write(new-65536, index, list_index)
                                            
                                            
                                        
                                total_index = record_index + total_index


workbook1.save("RA_Dataset.csv")
#workbook2.save("trial_RA_continue.csv")
"Done"


# # Check for Empty Folders

# In[ ]:

j = 1
for i in range (0, 6996):

    if (os.listdir("/Users/apple/Desktop/RA Final/" + str(Author_ID[i])[0:(len(str(Author_ID[i])) - 2)] + "_" + Author_Name[i]) == []):
        print (Author_Name[i], "empty")
        j = j + 1
print(j)

