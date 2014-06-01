#                               *       *        *****
# Author : Mohit Chawla         * *   * *       * 
# License : GNU GPL v2.0   **********************************
# ISM RESULT SCRAPER V1.0          *       *       * **
#-------------------------------*-------*--------*****-----------
# Description : This script automates the tideous task of viewing semester result.
# >>Displays the result of all students of a particular batch(year) between given two admission numbers 
# >>The results are saved to a "results.csv(MS Excel Spreadsheet)"
# Dated : 31 May 2014
import sys
import csv
try :
    import mechanize
    import cookielib
except:
    print"Unable to import mechanize ! Please install the library 'mechanize'!"
Script_version ='* * * * * * * * ISM RESULT SCRAPER V1.0 * * * * * * * * * * *'
Author_description = 'Author : Mohit Chawla Email : mohitchawla@ismu.ac.in || mohshia8@gmail.com '        
print "******** ISM RESULT SCRAPER V1.0 *********** \n\n"
print 'Author : Mohit Chawla \t Email : mohitchawla@ismu.ac.in || mohshia8@gmail.com \n'

def result_between():
    print 'Enter the two terminal Admission Numbers and results of all the students between them will be diplayed \n as well as will be written to a MS Excel File: \n'
    admNo1 = raw_input('Please Enter INITIAL Admission Number : in format 20YYJEXXXX or 20YYMTXXXX or so\n')
    admNo2 = raw_input('Please Enter FINAL Admission Number : in format 20YYJEXXXX or 20YYMTXXXX or so\n')
    if str(admNo1[4:6])!='JE' and str(admNo1[4:6])!='MT' or str(admNo2[4:6])!='JE' and str(admNo2[4:6])!='MT' or len(admNo1)!=10 or len(admNo2)!=10:
        print "Please Enter the roll number in specified format 20YYJEXXXX or 20YYMTXXXX"
        sys.exit()
    if str(admNo1[4:6])!=str(admNo2[4:6]):
        print "Both Admission numbers should be of same degree"
        sys.exit()
    prefix = admNo1[0:6]
    starting_admNo = int(admNo1[6:10])
    ending_admNo = int(admNo2[6:10])
    if starting_admNo > ending_admNo:
        print "INITIAL Admission Number should precede FINAL Admission Number \n. Please try again"
        sys.exit()
    try :
        fp = open('results.csv','a')
        do = csv.writer(fp)
        do.writerow([Script_version])
        do.writerow([Author_description])
        do.writerow(['Results from']+[admNo1]+['to']+[admNo2])
        do.writerow(['Admission No.']+['Name']+['GPA'])
        fp.close()
    except:
        print "Error in opening file!"
    while starting_admNo!=ending_admNo:
        admNo=''
        suffix=''
        name=''
        if len(str(starting_admNo))==3:
            suffix='0'+str(starting_admNo)
        else:
            suffix+=str(starting_admNo)
        admNo = str(prefix+suffix)
        print 'checking for %s'%admNo
        starting_admNo+=1
        # Browser
        br = mechanize.Browser()
        # Cookie Jar
        cj = cookielib.LWPCookieJar()
        br.set_cookiejar(cj)
        # Browser options
        br.set_handle_equiv(True)
        br.set_handle_gzip(True)
        br.set_handle_redirect(True)
        br.set_handle_referer(True)
        br.set_handle_robots(False)
        # Follows refresh 0 but not hangs on refresh > 0
        br.set_handle_refresh(mechanize._http.HTTPRefreshProcessor(), max_time=1)
        br.addheaders = [('User-agent', 'Mozilla/5.0 (X11; U; Linux i686; en-US; rv:1.9.0.1) Gecko/2008071615 Fedora/3.0.1-1.fc9 Firefox/3.0.1')]  
        # Want debugging messages?
        #br.set_debug_http(True)
        #br.set_debug_redirects(True)
        #br.set_debug_responses(True)
        # User-Agent (this is cheating, ok?)
        try:
            r = br.open('http://www.ismdhanbad.ac.in/result/admin/index.php')
        except:
            print'Error in reading page,Please Check your internet connection!!'
            sys.exit()
        html = r.read()    
        br.select_form(nr=0)
        br.form['appno'] = admNo
        br.submit()
        page_source = br.response().read()
        ##Finding user name
        nameindex =0
        nameindex=page_source.find('Name')
        nameindex=page_source.find('<td>',nameindex)
        nameindex+=4
        name = ''
        i=0
        while page_source[nameindex+i]!='<':
            name+=page_source[nameindex+i]
            i+=1
        ##
        tempindex =0
        tempindex = page_source.rfind('txtHint',0,len(page_source))
        tempindex = page_source.index('.',tempindex)
        tempindex-=1
        result_lastsem=''
        for i in range(0,4):
            result_lastsem+=page_source[tempindex+i]
        if str(result_lastsem) ==(').in') :
            name='NA'
            result_lastsem='NA' 
        print 'Result of %s(Admission Number : %s) for latest semester is %s'%(name,admNo,result_lastsem) 
        print 'Writing the data..'
        try :
            fp = open('results.csv','a')
            do = csv.writer(fp)
        
            do.writerow([admNo]+[name]+[result_lastsem])
            fp.close()
            print "Data written succesfully to results.csv "
        except:
            print "Error in opening file!"

def user_result():
    admNo = raw_input('Please Enter your Admission Number : in format 2012JEXXXX or 2012MTXXXX or so \n')
    if str(admNo[4:6])!='JE' and str(admNo[4:6])!='MT' or len(admNo)!=10:
        print "Please Enter the roll number in specified format 20YYJEXXXX or 20YYMTXXXX"
        sys.exit()
    
    # Browser
    br = mechanize.Browser()
    # Cookie Jar
    cj = cookielib.LWPCookieJar()
    br.set_cookiejar(cj)
    
    # Browser options
    br.set_handle_equiv(True)
    br.set_handle_gzip(True)
    br.set_handle_redirect(True)
    br.set_handle_referer(True)
    br.set_handle_robots(False)

    # Follows refresh 0 but not hangs on refresh > 0
    br.set_handle_refresh(mechanize._http.HTTPRefreshProcessor(), max_time=1)
    
    # Want debugging messages?
    #br.set_debug_http(True)
    #br.set_debug_redirects(True)
    #br.set_debug_responses(True)

    # User-Agent (this is cheating, ok?)
    br.addheaders = [('User-agent', 'Mozilla/5.0 (X11; U; Linux i686; en-US; rv:1.9.0.1) Gecko/2008071615 Fedora/3.0.1-1.fc9 Firefox/3.0.1')]
    try:
        print 'Connecting...'
        r = br.open('http://www.ismdhanbad.ac.in/result/admin/index.php')
    except:
        print'Error in reading page,Please Check your internet connection!!'
        sys.exit()
    html = r.read()

    br.select_form(nr=0)
    br.form['appno'] = admNo
    br.submit()
    page_source = br.response().read()
    ##Finding user name
    nameindex =0
    nameindex=page_source.find('Name')
    nameindex=page_source.find('<td>',nameindex)
    nameindex+=4
    name = ''
    i=0
    while page_source[nameindex+i]!='<':
         name+=page_source[nameindex+i]
         i+=1
    ##
    tempindex =0
    tempindex = page_source.rfind('txtHint',0,len(page_source))
    tempindex = page_source.index('.',tempindex)
    tempindex-=1
    result_lastsem=''
    for i in range(0,4):
        result_lastsem+=page_source[tempindex+i]

    if str(result_lastsem) ==(').in') :
        name='NA'
        result_lastsem='NA'
        
    print 'Result of %s(Admission number = %s) for latest semester exams is %s'%(name,admNo,result_lastsem) 
    print'\n Writing Data to %s_result.csv .. \n'%admNo
    try :
        fp = open('%s_result.csv'%admNo,'a')
        do = csv.writer(fp)
        do.writerow([Script_version])
        do . writerow([Author_description])
        do.writerow(['Admission No.']+['Name']+['GPA'])
        do.writerow([admNo]+[name]+[result_lastsem])
        fp.close()
        print "Data written succesfully to %s_result.csv "%admNo
    except:
        print "Error in opening file!"


def main():
#User Input
    print "1. Display and Write result for a single Admission Number\n "
    print "2. Display and Write results of all the students between 2 user-specified Admission Numbers" 
    user_choice = raw_input('\n Please Enter your Choice :\n ')

    if user_choice == '1':
        user_result()
    elif user_choice == '2':
        result_between()
    else:
        print 'Invalid Choice, Please Try Again'
    print "\n Author : Mohit Chawla Email : mohitchawla@ismu.ac.in || mohshia8@gmail.com \n "
main()

