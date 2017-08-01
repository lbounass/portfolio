
# coding: utf-8

# # Importing All Relevant Libraries

# In[1]:

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

# In[ ]:

book = xlrd.open_workbook("/Users/apple/Downloads/Bestsellers_CN_7-16-17.xlsx")
sheet_names = book.sheet_names()


# # Loading Books & Getting Authors Array

# In[ ]:

Books = book.sheet_by_index(0)
print ('Sheet name: %s' % Books.name)

Books_Authors = []
print(Books.nrows)


# # Create Text to Keep Track of Authors

# In[ ]:

for i in range (1, Books.nrows):
	if((Books.cell(i-1,5)).value != (Books.cell(i, 5)).value):
		Books_Authors.append((Books.cell(i,5)).value)
	else:
		i = i+1


# # Loading Authors & Getting Authors Array

# In[ ]:

Authors = book.sheet_by_index(1)	
print ('Sheet name: %s' % Authors.name, Authors.nrows)


# # Create Arrays for Authors' IDs and Names

# In[ ]:

Author_ID=[]
Author_Name=[]

for i in range (1, Authors.nrows):
	Author_ID.append((Authors.cell(i,1)).value)
    
for i in range (1, Authors.nrows):
	Author_Name.append((Authors.cell(i,0)).value)
    
print(Author_Name)


# # Create Dictionary for Authors Page (order dependent on hash value, not alphabet)

# In[ ]:

Author_Aliases = {}
for author in range (0, Authors.nrows - 1):
	Author_Aliases[Author_Name[author]] = {}
	lc = 4
	lc_index = 0
	for lc in range(4, 37, 2):
		Author_Aliases[Author_Name[author]][lc_index] = Authors.cell(author+1,lc).value; 
		lc_index = lc_index+1
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

# In[ ]:

for i in range (0, len(Author_ID)):
	newpath = "/Users/apple/Desktop/RA Trial/" + str(Author_ID[i])[0:(len(str(Author_ID[i])) - 2)] + "_" + Author_Name[i]  
	if not os.path.exists(newpath):
		os.makedirs(newpath)


# # Initiating All The Exceptions and Needed Arrays & Creating Text File for Updates

# In[ ]:

NoSuchElement_Errors = []
Unicode_Errors = []
Unicode_Auth_Names_Errors = []
Empty_Book_Set = [] #### For those with no book under their names
Large_Book_Set = [] #### For those with too many books under their name
Sample_Errors = [] #### For the future handling of NoSuchElement_Exceptions
NoSuchElement_Sample = []

summary = open ("/Users/apple/Desktop/RA Trial/SummaryOfAuthors.txt", "w")
summary.write("Starting:" + "\n")


# # Firefox Script / Web Scraping Using Selenium

# In[ ]:

for author in range(0,len(Author_Aliases)):
    
    name = Author_Aliases.keys()[author]
    print ("Author #", author, ": ", name)
    time.sleep(2)
    
    try:
        summary.write("\n" + "For #"+str(author) + " " + str(Author_Aliases.keys()[author]) + ":" + "\n")
        
    except UnicodeEncodeError:
        print("Unicode Encode Error. Continue with the alias names")
        Unicode_Auth_Names_Errors.append(Author_Aliases.keys()[author])
        
    for lc in range (0 , 17):
        
        empty = 0
        if (Author_Aliases.get(name)[lc]):
            print("Author Alias #", lc, ": ", Author_Aliases.get(name)[lc])  
            
            try:        
                #### OPEN BROWSER
                driver = webdriver.Firefox()
                driver.get("https://catalog.loc.gov")
                driver.get("https://catalog.loc.gov/vwebv/searchBrowse")
                
                #### BROWSE PAGE
                search_code = Select(driver.find_element_by_id("search-code"))
                search_code.select_by_visible_text("AUTHORS/CREATORS beginning with (enter last name first)")
                driver.implicitly_wait(10) # seconds

                if (Author_Aliases.get(name)[lc] == "N/A"):
                    summary.write("\n" + "          " + Author_Aliases.keys()[author] + "-Not in library" + "\n")
                    driver.quit()
                    break
                    
                else:    
                    (driver.find_element_by_id("search-argument")).send_keys(Author_Aliases.get(name)[lc])
                    (driver.find_element_by_name("page.search.search.button")).click()
                    titles = (driver.find_element_by_class_name("search-results-browse-list-title-number")).text
                    max_titles = int(titles[1:(len(titles)-1)])
                    print(max_titles)
                    
                    if (max_titles == 0):
                        summary.write("\n" + "          " + Author_Aliases.get(name)[lc] + "-Empty (Exception)" + "\n")
                        driver.quit()
                        continue                        
                        
                    if (max_titles == 1):
                        (driver.find_element_by_css_selector("a[href*='search?searchType=7']")).click();
                        summary.write("\n" + "          "+ Author_Aliases.get(name)[lc] + "-Good-" + titles +"\n")
                        time.sleep(5)
                        
                        if(driver.find_element_by_xpath("/html/body/main/article/div[2]/h1/small").text=="Book"):
                            (driver.find_element_by_xpath("/html/body/main/article/div[2]/div/section/div/div[2]/div/a[2]")).click()
                            (driver.find_element_by_name("butExport")).click()
                            alert = driver.switch_to_alert()
                            time.sleep(5)
                            driver.quit()
                        else:
                            print("No book for  author ", Author_Aliases.get(name)[lc])
                            Empty_Book_Set.append(Author_Aliases.get(name)[lc])
                            summary.write("\n    " + "          " + Author_Aliases.get(name)[lc] + "-No Book-" + titles + "\n")
                            empty = 1 ## Because an IOException will be raised since there is no file to move
                            driver.quit()                      
                        
                        
                        #### MOVING FILE FOR CASE MAX_TITLES = 1
                        print("Moving files: ")	
                        source_dir="/Users/apple/Downloads"
                        dest_dir="/Users/apple/Desktop/RA Trial"
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
                                shutil.move(dest_dir + "/" + new_file, "/Users/apple/Desktop/RA Trial/" + str(Author_ID[i])[0:(len(str(Author_ID[i])) - 2)] + "_" + Author_Name[i] + "/" + str(Author_ID[i])[0:(len(str(Author_ID[i])) - 2)] + "_" + str(lc + 1) + "_" + new_file)
                                summary.write("\n" + "          --> in Author_ID #: " + str(Author_ID[i])[0:(len(str(Author_ID[i])) - 2)])

                    #### IF MAX_TITLES > 2
                    else:	
                        (driver.find_element_by_css_selector("a[href*='search?searchType=7']")).click();
                        summary.write("          " + Author_Aliases.get(name)[lc] + "-Good-" + titles + "\n")

                #### SHOULD MAXIMIZE WINDOW AND MAXIMIZE NUMBER OF RECORDS PER PAGE FOR CONVENIENCE
                        time.sleep(5)
                        driver.maximize_window()
                        record = Select(driver.find_element_by_id("record-count"))
                        record.select_by_visible_text("100")
                        time.sleep(5)
                        
                #### 2 possibilities now: 
                #1) max of titles is less than 100, so we just check the books, select the books then click save.
                #2) max of titles is more than 100, so for the first to before last page, we check all 100 options, click Next until 
                # last page, where we take the maximum of options available (<100) then click save.
                
                ############## OPTION 1 
                        if (max_titles <= 100):
                            
                            #### KEEP TRACK OF NUMBER OF BOOKS
                            nb_books = 0

                            for i in range (1, max_titles + 1):
                                
                                if (driver.find_element_by_xpath("//div[@id='search-results']/ul/li[" + str(i) + "]/div[3]/div/div[1]").text=="Book"):
                                    nb_books = nb_books + 1
                                    (driver.find_element_by_xpath("//div[@id='search-results']/ul/li[" + str(i) + "]").find_element_by_name("titles")).click()
                                    time.sleep(0.3)
                                    
                                else:
                                    time.sleep(0.3)
                                    
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
                                summary.write("\n    " + "          " + Author_Aliases.get(name)[lc] + "-No Book-" + titles + "\n")
                                driver.quit()
                                empty = 1
                                
                ############## OPTION 2                            
                        else: 
                        
                            if (max_titles > 399):
                                
                                #### KEEP TRACK OF LARGE SETS FOR FUTURE VERIFICATION
                                if (Author_Aliases.get(name)[lc] not in Large_Book_Set):
                                    Large_Book_Set.append(Author_Aliases.get(name)[lc])
                                    
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
                                        time.sleep(0.5)
                                        
                                    else:
                                        time.sleep(0.5)
                                      
                                if (k == 4) :
                                    (driver.find_element_by_xpath("//*[@id='results-form']/div[1]/ul/li[" + str(k + 1) + "]/a")).click();

                                elif (k == 6):
                                    (driver.find_element_by_xpath("//*[@id='results-form']/div[1]/ul/li[" + str(k + 1) + "]/a")).click();

                                elif (k == 7):
                                    (driver.find_element_by_xpath("//*[@id='results-form']/div[1]/ul/li[" + str(k) + "]/a")).click();

                                elif (k > 7):
                                    (driver.find_element_by_xpath("//*[@id='results-form']/div[1]/ul/li[7]/a")).click();

                                else:
                                    (driver.find_element_by_xpath("//*[@id='results-form']/div[1]/ul/li[" + str(k + 2) + "]/a")).click();
                              
                                k = k + 1
                                time.sleep(5)
                               
                            #### LAST PAGE
                            new = max_titles - (pages - 1) * 100
                            print("Last page:" , str(k), "Number left= ", str(new)) 
                            
                            for last_page in range (1, int(new + 1)):
                                
                                if (driver.find_element_by_xpath("//div[@id='search-results']/ul/li[" + str(last_page) + "]/div[3]/div/div[1]").text == "Book"):
                                    nb_books = nb_books + 1; 
                                    (driver.find_element_by_xpath("//div[@id='search-results']/ul/li[" + str(last_page) + "]").find_element_by_name("titles")).click()
                                    time.sleep(1)
                                else:
                                    time.sleep(1)
                                    
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
                                summary.write("\n    " + "          " + Author_Aliases.get(name)[lc] + "-No Book-" + titles + "\n")
                                driver.quit()
                                empty = 1                        
                                
                        #### MOVING FILE FOR CASE MAX_TITLES > 1
                        print("Moving files: ")	
                        source_dir="/Users/apple/Downloads"
                        dest_dir="/Users/apple/Desktop/RA Trial"
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
                                shutil.move(dest_dir + "/" + new_file, "/Users/apple/Desktop/RA Trial/" + str(Author_ID[i])[0:(len(str(Author_ID[i])) - 2)] + "_" + Author_Name[i] + "/" + str(Author_ID[i])[0:(len(str(Author_ID[i])) - 2)] + "_" + str(lc + 1) + "_" + new_file)
                                summary.write("\n" + "          --> in Author_ID #: " + str(Author_ID[i])[0:(len(str(Author_ID[i])) - 2)] + "          --> Number of pages: " + nb_books)
                                break

            #### EXCEPTIONS
            except StaleElementReferenceException:
                
                print("Stale Element Reference Exception at author: " + Author_Aliases.get(name)[lc] + ", let's count it as no such element exception")
                
                if (Author_Aliases.get(name)[lc] not in NoSuchElement_Errors):
                    NoSuchElement_Errors.append(Author_Aliases.get(name)[lc])
                    
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
                
                else:
                    print(Author_Aliases.get(name)[lc] + " is a duplicate for No Such Element Error")
                
                time.sleep(5) 
                driver.quit()
                continue
                
            except UnicodeEncodeError:
                
                print("Unicode Exception at author: " + Author_Aliases.get(name)[lc] + ", let's continue")
                
                if (Author_Aliases.get(name)[lc] not in Unicode_Errors):
                    Unicode_Errors.append(Author_Aliases.get(name)[lc])
                
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
                
                    else:
                        print(Author_Aliases.get(name)[lc] + " is a duplicate for No Such Element Error")
                
                time.sleep(5)   
                driver.quit()
                
            except TimeoutException:
                
                print("Timeout Exception at author: " + Author_Aliases.get(name)[lc] + ", let's pause and count it as no such element exception")
                
                if (Author_Aliases.get(name)[lc] not in NoSuchElement_Errors):
                    NoSuchElement_Errors.append(Author_Aliases.get(name)[lc])
                
                else:
                    print(Author_Aliases.get(name)[lc] + " is a duplicate for TimeoutException")
                
                time.sleep(5)   
                driver.quit()
                time.sleep(120)
 
        #### IF NO OTHER ALIAS
        else:
            print("No other ALIAS names")
            break


# Because of server issues, we have to consider the instances in which the IP server gets blocked occasionally and an exception is raised. The NoSuchElement_Exceptions array takes care of that. It will take in any ALIAS name that was discarded because of the exception, and a new script adapted to that array will run. As long as NoSuchElement_Exceptions isn't empty, this script will go on and on.

# # Opening New Text File for No Such Element Exceptions

# In[ ]:

summary_nosuchelement = open ("/Users/apple/Desktop/RA Trial/SummaryOfAuthors_nosuchelement.txt", "w")
summary_nosuchelement.write("Starting:" +"\n")


# # Handling No Such Element Exceptions

# In[ ]:

#### WHILE NOSUCHELEMENT_EXCEPTIONS IS NOT EMPTY, WE KEEP ITERATING
while (NoSuchElement_Errors):
    
    Sample_Errors = NoSuchElement_Errors    
    print("Sample_Errors: ")
    print(Sample_Errors, len(Sample_Errors))
    
    for author in range(0 , len(Sample_Errors)):
        
        empty = 0
        name = Sample_Errors[author]
        print("Now: ", name, author)
        print("Author Alias #", author, ": ", name)  
        summary_nosuchelement.write("\n" + "For Author Alias #" + str(author) + " " + str(name) + ":" + "\n")

        try:
            driver = webdriver.Firefox()
            driver.get("https://catalog.loc.gov")
            driver.get("https://catalog.loc.gov/vwebv/searchBrowse")
            search_code = Select(driver.find_element_by_id("search-code"))
            search_code.select_by_visible_text("AUTHORS/CREATORS beginning with (enter last name first)")
            driver.implicitly_wait(10) # seconds
            
            if (Sample_Errors[author] == "N/A"):
                driver.quit()
                break
                
            else:    
                (driver.find_element_by_id("search-argument")).send_keys(Sample_Errors[author])
                (driver.find_element_by_name("page.search.search.button")).click()
                titles = (driver.find_element_by_class_name("search-results-browse-list-title-number")).text
                max_titles = int(titles[1:(len(titles)-1)])
                print(max_titles)
                
                if (max_titles == 0):
                    summary_nosuchelement.write("\n" + "          " + Sample_Errors[author] + "-Empty (Exception)" + "\n")
                    Empty_Book_Set.append(Sample_Errors[author])
                    driver.quit()
                    continue
                    
                if (max_titles == 1):
                    (driver.find_element_by_css_selector("a[href*='search?searchType=7']")).click();
                    time.sleep(5)
                    nb_books = 0
                    
                    if(driver.find_element_by_xpath("/html/body/main/article/div[2]/h1/small").text=="Book"):
                        (driver.find_element_by_xpath("/html/body/main/article/div[2]/div/section/div/div[2]/div/a[2]")).click()
                        (driver.find_element_by_name("butExport")).click()
                        alert = driver.switch_to_alert()
                        nb_books = 1
                        time.sleep(5)
                        driver.quit()
                        
                    else:
                        print("No book for  author ", Sample_Errors[author])
                        Empty_Book_Set.append(Sample_Errors[author])
                        summary_nosuchelement.write("\n    " + "          " + Sample_Errors[author] + "-No Book-" + titles + "\n")
                        empty = 1
                        driver.quit()                  

                    print("Moving files: ")	
                    source_dir="/Users/apple/Downloads"
                    dest_dir="/Users/apple/Desktop/RA Trial"
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
                        
                        name_sample=Author_Aliases.keys()[i]
                        
                        for j in range (0, 17):
                            
                            if (Sample_Errors[author]==Author_Aliases.get(name_sample)[j]):
                                name = name_sample
                                lc_name = j
                                
                    for i in range (0, len(Author_ID)):
                        if(name == Author_Name[i]):
                                    print("Found it")
                                    summary_nosuchelement.write("\n" + "          --> in Author_ID #: " + str(Author_ID[i])[0:(len(str(Author_ID[i])) - 2)] + "          --> Number of pages: " + nb_books)                             
                                    shutil.move(dest_dir + "/" + new_file, "/Users/apple/Desktop/RA Trial/" + str(Author_ID[i])[0:(len(str(Author_ID[i])) - 2)] + "_" + Author_Name[i] + "/" + str(Author_ID[i])[0:(len(str(Author_ID[i])) - 2)] + "_" + str(lc_name + 1) + "_" + new_file)


                else:	
                    (driver.find_element_by_css_selector("a[href*='search?searchType=7']")).click();
                    time.sleep(5)
                    driver.maximize_window()
                    record = Select(driver.find_element_by_id("record-count"))
                    record.select_by_visible_text("100")
                    time.sleep(5)

                    ############## OPTION 1 
                    if (max_titles <= 100):
                        
                        nb_books = 0
                        
                        for i in range (1, max_titles + 1):

                            if (driver.find_element_by_xpath("//div[@id='search-results']/ul/li[" + str(i) + "]/div[3]/div/div[1]").text == "Book"):
                                nb_books = nb_books + 1
                                (driver.find_element_by_xpath("//div[@id='search-results']/ul/li[" + str(i) + "]").find_element_by_name("titles")).click()
                                time.sleep(0.3)
                            
                            else:                     
                                time.sleep(0.3)
                        
                        print("Number of books by this author: ", nb_books)

                        if (max_titles > 2):
                            driver.execute_script('window.scrollTo(0,document.body.scrollHeight)'); 
                        
                        if (nb_books != 0):
                            (driver.find_element_by_name("ExportDialogServlet")).click()
                            time.sleep(2)
                            (driver.find_element_by_name("butExport")).click()
                            alert = driver.switch_to_alert()
                            nb_books=0
                            time.sleep(7)	
                            driver.quit()
                        else:
                            print("No book for  author ", Sample_Errors[author])
                            Empty_Book_Set.append(Sample_Errors[author])
                            driver.quit()
                            empty = 1
                     
                    ############## OPTION 2                    
                    else: 
                        
                        if (max_titles > 399):
                            
                            if (Sample_Errors[author] not in Large_Book_Set):
                                Large_Book_Set.append(Sample_Errors[author])

                            else:
                                print("Already in Large_Book_Set")
                              
                        pages = math.ceil(max_titles / 100) + 1
                        print("Nb of pages= " + str(pages))
                        k = 1
                        nb_books = 0
                        
                        while(k < pages) :
                            
                            time.sleep(5)
                            print(k)
                            
                            for i in range (1, 101):
                                
                                if (driver.find_element_by_xpath("//div[@id='search-results']/ul/li["+str(i)+"]/div[3]/div/div[1]").text=="Book"):
                                    nb_books=nb_books + 1
                                    (driver.find_element_by_xpath("//div[@id='search-results']/ul/li["+str(i)+"]").find_element_by_name("titles")).click()
                                    time.sleep(0.5)
                                else:
                                    time.sleep(0.5)
                                    
                            if (k == 4) :
                                (driver.find_element_by_xpath("//*[@id='results-form']/div[1]/ul/li["+str(k+1)+"]/a")).click();
   
                            elif (k == 6):
                                (driver.find_element_by_xpath("//*[@id='results-form']/div[1]/ul/li["+str(k+1)+"]/a")).click();
                
                            elif (k == 7):
                                (driver.find_element_by_xpath("//*[@id='results-form']/div[1]/ul/li["+str(k)+"]/a")).click();
                            
                            elif (k > 7):
                                (driver.find_element_by_xpath("//*[@id='results-form']/div[1]/ul/li[7]/a")).click();

                            else:
                                (driver.find_element_by_xpath("//*[@id='results-form']/div[1]/ul/li["+str(k+2)+"]/a")).click();
                                
                            k = k + 1
                            time.sleep(5)
                        
                        #LAST PAGE
                        new = max_titles - (pages - 1) * 100
                        print("Last page:" , str(k), "Number left= ", str(new)) 
                        
                        for last_page in range (1, int(new + 1)):
                            
                            if (driver.find_element_by_xpath("//div[@id='search-results']/ul/li[" + str(last_page) + "]/div[3]/div/div[1]").text == "Book"):
                                nb_books = nb_books + 1; 
                                (driver.find_element_by_xpath("//div[@id='search-results']/ul/li[" + str(last_page) + "]").find_element_by_name("titles")).click()
                                time.sleep(1)
                                
                            else:
                                time.sleep(1)
                                
                        if (new > 2):
                            driver.execute_script('window.scrollTo(0,document.body.scrollHeight)'); 

                        if (nb_books != 0):    
                            (driver.find_element_by_name("ExportDialogServlet")).click()
                            (driver.find_element_by_name("butExport")).click()
                            alert = driver.switch_to_alert()

                            if (max_titles > 399):
                                time.sleep(118)	

                            else:
                                time.sleep(60)

                            driver.quit()
                            
                        else: 
                            print("No book for  author ", Sample_Errors[author])
                            Empty_Book_Set.append(Sample_Errors[author])
                            summary_nosuchelement.write("\n    " + "          " + Sample_Errors[author] + "-No Book-" + titles + "\n")
                            driver.quit()
                            empty = 1                    

                    print("Moving files: ")	
                    source_dir="/Users/apple/Downloads"
                    dest_dir="/Users/apple/Desktop/RA Trial"
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
                        
                        name_sample = Author_Aliases.keys()[i]
                        
                        for j in range (0, 17):
                            
                            if (Sample_Errors[author] == Author_Aliases.get(name_sample)[j]):
                                name = name_sample
                                lc_name = j

                            
                    for i in range (0, len(Author_ID)):
                        
                        if (name == Author_Name[i]):
                            
                            print("Found it")
                            shutil.move(dest_dir + "/" + new_file, "/Users/apple/Desktop/RA Trial/" + str(Author_ID[i])[0:(len(str(Author_ID[i])) - 2)] + "_" + Author_Name[i] + "/" + str(Author_ID[i])[0:(len(str(Author_ID[i])) - 2)] + "_" + str(lc_name + 1) + "_" + new_file)
                            summary_nosuchelement.write("\n" + "          --> in Author_ID #: " + str(Author_ID[i])[0:(len(str(Author_ID[i])) - 2)] + "          --> Number of pages: " + nb_books)
                            break

        #### EXCEPTIONS
        except WebDriverException:
            
            print("WebDriver Exception at author: " + Sample_Errors[author] + ", let's count it as no such element exception")
            
            if (Sample_Errors[author] not in NoSuchElement_Sample):
                NoSuchElement_Sample.append(Sample_Errors[author])
                
            else:
                print(Sample_Errors[author] + " is a duplicate for No Such Element Error")
            
            time.sleep(30) 
            driver.quit()
            continue

                            
        except StaleElementReferenceException:
            
            print("Stale Element Reference Exception at author: " + Sample_Errors[author] + ", let's count it as no such element exception")
            
            if (Sample_Errors[author] not in NoSuchElement_Sample):
                NoSuchElement_Sample.append(Sample_Errors[author])
            
            else:
                print(Sample_Errors[author] + " is a duplicate for No Such Element Error")
            
            time.sleep(5) 
            driver.quit()
            time.sleep(60)
            continue                            
                             
        except InvalidElementStateException:
            
            print("Invalid Element State Exception at author: " + Sample_Errors[author] + ", let's count it as no such element exception")
            
            if (Sample_Errors[author] not in NoSuchElement_Sample):
                NoSuchElement_Sample.append(Sample_Errors[author])
            
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
            
            else:
                print("Name already in list")
            
            time.sleep(5) 
            driver.quit()
            continue
            
        except UnicodeEncodeError:
            
            print("Unicode Exception at author: " + Sample_Errors[author] + ", let's continue")
            
            if ( Sample_Errors[author] not in Unicode_Errors):
                Unicode_Errors.append(Sample_Errors[author])
            
            else:
                print ("Name already in list)
            
            time.sleep(5)   
            driver.quit()
            continue

        except IOError:
                       
            print("IOError Exception at author: " + Sample_Errors[author] + ", let's count it as no such element exception")
            
            if(empty == 1):
                empty = 0
                continue
                       
            else:
                if (Sample_Errors[author] not in NoSuchElement_Sample):
                    NoSuchElement_Sample.append(Sample_Errors[author])
                else:
                    print ("Name already in list)
            
            time.sleep(5)   
            driver.quit() 
            
        except TimeoutException:
                       
            print("Timeout Exception at author: " + Sample_Errors[author] + ", let's pause and count it as no such element exception")
                       
            if (Sample_Errors[author] not in NoSuchElement_Sample):
                NoSuchElement_Sample.append(Sample_Errors[author])
                       
            else:
                print(Sample_Errors[author] + " is a duplicate for TimeoutException")
            
            time.sleep(5)   
            driver.quit()
            time.sleep(120)


    NoSuchElement_Errors = NoSuchElement_Sample
    NoSuchElement_Sample = []
                       
    print("NoSuchElement_Errors: ")
    print(NoSuchElement_Errors, len(NoSuchElement_Errors))



# # When Finished, Check All Relevant Arrays

# In[ ]:

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
        if  os.listdir("/Users/apple/Desktop/RA Trial/" + str(Author_ID[i])[0:(len(str(Author_ID[i])) - 2)] + "_" + Author_Name[i]):
            print("Ok to move " + Author_Name[i])
            shutil.move("/Users/apple/Desktop/RA Trial/" + str(Author_ID[i])[0:(len(str(Author_ID[i])) - 2)]+ "_" + Author_Name[i], "/Users/apple/Desktop/RA Trial/Complete" )
    except OSError:
        continue

