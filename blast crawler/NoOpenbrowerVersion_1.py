import urllib
import urllib2
import re
import time
#excel op
import xlrd
from xlwt import easyxf
from xlutils.copy import copy
import httplib

httplib.HTTPConnection._http_vsn = 10
httplib.HTTPConnection._http_vsn_str = 'HTTP/1.0'
# Open a excel
r_xls = xlrd.open_workbook('F:/Python/realwork/excelExtract/NoOpendata/Total_findnodone1-1-1.xls')
#Open a sheet
#define Protein list 
Pro_sequence = []

#open table sheet 2
r_sheet = r_xls.sheets()[0]
#get the row numbers
nrows = r_sheet.nrows
#print nrows
wi = 1
while wi < nrows:
    
    Pro_sequence.append(r_sheet.cell(wi, 2).value)
    
    #print Pro_sequence[i-1]
    wi += 1

#for mn in Pro_sequence:
    #print mn

#==============get the data from excel


k = 0
while k < nrows-1:
    r_xls_in = xlrd.open_workbook('F:/Python/realwork/excelExtract/NoOpendata/Total_findnodone1-1-1.xls')
    w_xls = copy(r_xls_in)  
    
    #first post to get CDD_RID and RID
    header_value = {}
    
    header_value['Accept'] = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
    header_value['Accept-Encoding'] = "gzip, deflate, br"
    header_value['Accept-Language'] = "zh-CN,zh;q=0.8,en-US;q=0.5,en;q=0.3"
    header_value['Connection'] = "keep-alive"
    header_value['Host'] = "blast.ncbi.nlm.nih.gov"
    header_value['Upgrade-Insecure-Requests'] = "1"
    header_value['Referer'] = "https://blast.ncbi.nlm.nih.gov/Blast.cgi?PROGRAM=blastp&PAGE_TYPE=BlastSearch&LINK_LOC=blasthome"
    header_value['Cookie'] = "ncbi_sid=AD2A97378DE0D351_0000SID; ncbi_prevPHID=AD2AA1CD8DE0D4810000000000000001; ncbi_is_ga=true; BlastCubbyImported=passive; MyBlastUser=1jalWDI_ghY8WXMAm92AF1B38; clicknext=; unloadnext=; prevselfurl="
    header_value['User-Agent'] ="Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:52.0) Gecko/20100101 Firefox/52.0"
    
    
    values = {}
    
    values['QUERY'] = Pro_sequence[k]
    values['db'] = 'protein'
    #values['QUERYFILE'] = 'Content-Type: application/octet-stream'
    values['GENETIC_CODE'] = '1'
    values['stype'] = 'protein'
    #values['SUBJECTFILE'] = 'Content-Type: application/octet-stream'
    values['DATABASE'] = 'nr'
    values['NUM_ORG'] = '1'
    values['BLAST_PROGRAMS'] = 'blastp'
    values['MAX_NUM_SEQ'] = '100'
    values['SHORT_QUERY_ADJUST'] = 'on'
    values['EXPECT'] = '10'
    values['WORD_SIZE'] = '6'
    values['HSP_RANGE_MAX'] = '0'
    values['MATRIX_NAME'] = 'BLOSUM62'
    values['MATCH_SCORES'] = '1,-2'
    values['GAPCOSTS'] = '11 1'
    values['COMPOSITION_BASED_STATISTICS...'] = '2'
    values['REPEATS'] = '4829'
    values['TEMPLATE_LENGTH'] = '0'
    values['TEMPLATE_TYPE'] = '0'
    values['PSSM'] = 'Content-Type: application/octet-stream'
    values['SHOW_OVERVIEW'] = 'true'
    values['SHOW_LINKOUT'] = 'true'
    values['GET_SEQUENCE'] = 'true'
    values['FORMAT_OBJECT'] = 'Alignment'
    values['FORMAT_TYPE'] = 'HTML'
    values['ALIGNMENT_VIEW'] = 'Pairwise'
    values['MASK_CHAR'] = '2'
    values['MASK_COLOR'] = '1'
    values['DESCRIPTIONS'] = '100'
    values['ALIGNMENTS'] = '100'
    values['LINE_LENGTH']  = '60'
    values['NEW_VIEW'] = 'true'
    values['OLD_VIEW'] = "false"
    values['NUM_OVERVIEW'] = '100'
    values['QUERY_INDEX'] = '0'
    values['FORMAT_NUM_ORG'] = '1'
    values['CONFIG_DESCR'] = "2,3,4,5,6,7,8"
    values['CLIENT'] = 'web'
    values['SERVICE'] = 'plain'
    values['CMD'] = 'request'
    values['PAGE'] = 'Proteins'
    values['PROGRAM'] = "blastp"
    values['CDD_SEARCH'] = 'on'
    values['SELECTED_PROG_TYPE'] = 'blastp'
    values['NUM_DIFFS'] = '0'
    values['NUM_OPTS_DIFFS'] = '0'
    #values['UNIQ_DEFAULTS_NAME'] = 'A_SearchDefaults_1ctqCv_16W9_dlTNnfv6Aav_GTW6B_UFbQr'
    values['PAGE_TYPE'] = "BlastSearch"
    values['USER_DEFAULT_PROG_TYPE'] = "blastp"
    values['USER_DEFAULT_MATRIX'] = '4'
    
    
    data = urllib.urlencode(values)
    
    url = 'https://blast.ncbi.nlm.nih.gov/Blast.cgi'
    
    request = urllib2.Request(url,data,header_value)
    try:
        response = urllib2.urlopen(request)
    except:
        k = k-1
        print "1st post exception "
    else:   
    
        RID_html = str(response.read())
        #print RID_html
        
        RID_pattern = re.compile(r"<tr><td>Request ID</td><td> <b>(\w*)</b></td></tr>")
        RID_start = RID_pattern.search(RID_html).groups()
        
        print RID_start[0]
        RID = RID_start[0]
        
        #CDDRID_pattern = re.compile(r'<input name="CDD_RID" type="hidden" value="((data_cache_seq:(\d*))|(\w*))">')
        CDDRID_pattern = re.compile(r'<input name="CDD_RID" type="hidden" value="([^>]+)">')
        #print RID_html
        CDDRID_start = CDDRID_pattern.findall(RID_html)
        if CDDRID_start:
            print 1
        else:
            print 0
        
        print CDDRID_start[0]
        CDDRID = CDDRID_start[0]
        #second time post
        header_value2 = {}
    
        header_value2['Accept'] = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
        header_value2['Accept-Encoding'] = "gzip, deflate, br"
        header_value2['Accept-Language'] = "zh-CN,zh;q=0.8,en-US;q=0.5,en;q=0.5"
        header_value2['Connection'] = "keep-alive"
        header_value2['Host'] = "blast.ncbi.nlm.nih.gov"
        header_value2['Upgrade-Insecure-Requests'] = "1"
        header_value2['Referer'] = "https://blast.ncbi.nlm.nih.gov/Blast.cgi"
        header_value2['Cookie'] = "ncbi_sid=AD2A97378DE0D351_0000SID; ncbi_prevPHID=50C9A6FC8DE187110000000000000001; ncbi_is_ga=true; BlastCubbyImported=passive; MyBlastUser=1tMphgNP0IW2QNP532870E215; clicknext="
        header_value2['User-Agent'] ="Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:52.0) Gecko/20100101 Firefox/52.0"
        
        
        
        values2 = {}
        
        values2['ALIGNMENTS'] = "100"
        values2['ALIGNMENT_VIEW'] = 'Pairwise'
        values2['BLAST_PROGRAMs'] = "blastp"
        values2['CDD_RID'] = str(CDDRID)
        values2['CDD_SEARCH_STATE'] = '4'
        values2['CLIENT'] = 'web'
        values2['CMD'] = 'Get'
        values2['COMPOSITION_BASED_STATISTICS...'] = '2'
        values2['CONFIG_DESCR'] = "2,3,4,5,6,7,8"
        values2['DATABASE'] = 'nr'
        values2['DESCRIPTIONS'] = '100'
        values2['EQ_OP'] = 'AND'
        values2['EXPECT'] = '10'
        values2['FILTER'] = 'F'
        values2['FORMAT_NUM_ORG'] = '1'
        values2['FORMAT_OBJECT'] = 'Alignment'
        values2['FORMAT_TYPE'] = 'HTML'
        values2['FULL_DBNAME'] = 'nr'
        values2['GAPCOSTS'] = '11 1'
        values2['GET_SEQUENCE'] = 'on'
        values2['HSP_RANGE_MAX'] = '0'
        values2['JOB_TITLE'] = 'Protein '+'Sequece ('+str(len(values['QUERY'])) +' letters)'
        values2['LAYOUT'] = 'OneWindow'
        values2['LINE_LENGTH']  = '60'
        values2['MASK_CHAR'] = '2'
        values2['MASK_COLOR'] = '1'
        values2['MATRIX_NAME'] = 'BLOSUM62'
        values2['MAX_NUM_SEQ'] = '100'
        #values2['NCBI_GI'] = 'false'
        values2['NEW_VIEW'] = 'on'
        values2['NUM_DIFFS'] = '0'
        values2['NUM_OPTS_DIFFS'] = '0'
        values2['NUM_ORG'] = '1'
        values2['NUM_OVERVIEW'] = '100'
        values2['OLD_VIEW'] = "false"
        values2['PAGE'] = 'Proteins'
        values2['PAGE_TYPE'] = "BlastSearch"
        values2['PROGRAM'] = 'blastp'
        values2['QUERY_INFO'] = 'Protein '+'Sequece ('+str(len(values['QUERY'])) +' letters)'
        values2['QUERY_LENGTH'] = '265'
        values2['REPEATS'] = '4829'
        values2['RID'] = str(RID)
        values2['RTOE'] = '48'
        #values2['SAVED_SEARCH'] = 'true'
        values2['SEARCH_DB_STATUS'] = '31'
        values2['SELECTED_PROG_TYPE'] = 'blastp'
        values2['SERVICE'] = 'plain'
        values2['SHORT_QUERY_ADJUST'] = 'on'
        #values2['SHOW_CDS_FEATURE'] = 'false'
        values2['SHOW_LINKOUT'] = 'on'
        values2['SHOW_OVERVIEW'] = 'on'
        values2['USER_DEFAULT_MATRIX'] = '4'
        values2['USER_DEFAULT_PROG_TYPE'] = "blastp"
        values2['USER_TYPE'] = '2'
        values2['WORD_SIZE'] = '6'
        values2['_PGR'] = '6'
        values2['_PGR'] = '6'
        values2['db'] = 'protein'
        values2['stype'] = 'protein'
        
        
        data2 = urllib.urlencode(values2)
        
        url2 = 'https://blast.ncbi.nlm.nih.gov/Blast.cgi'
        #loop to find the final anwser
        i_second_req = 1
        
        while(i_second_req):
            #print "loop"
            #print k
            #k = k +1
            request2 = urllib2.Request(url2,data2,header_value2)
            try:
                response2 = urllib2.urlopen(request2)
            except:
                print "2nd post exception"
                
            else:
                #match if match final result or not
                html2 = response2.read()
            
                match = re.findall(r'<meta name="ncbi_pdid" content="blastresults" />',html2)
                if match:
                    i_second_req = 0
                    print 'y'
                    result_html = html2
                time.sleep(1)
            
        pat_match_urls = re.compile(r'<td class="c1 l lim">\n<a href=([^>]+)">')
        
        match_urls = pat_match_urls.findall(result_html)
        result_Line = len(match_urls)
        if result_Line > 5:
            loopTime = 5
        elif result_Line != 0:
            loopTime = result_Line
        else:
            w_xls.get_sheet(0).write(k + 1,25,"No match Result1")
            loopTime = 0
    #above are going to get 5 links
    #followed are going to get 5 id_dtr
        i = 0
        for i in xrange(loopTime):
            id_dtr = re.findall(r"\d+", str(match_urls[i])) 
            try:
                url_req = "https://www.ncbi.nlm.nih.gov/sviewer/viewer.fcgi?id=" + str(id_dtr[0]) + "&db=protein&report=genpept&extrafeat=984&fmt_mask=0&retmode=html&withmarkup=on&tool=portal&log$=seqview&maxplex=3&maxdownloadsize=1000000"
                time.sleep(0.5)
                page_html = urllib2.urlopen(url_req).read()
                time.sleep(0.05)                   
            except:
                k = k - 1
                break
                
            else:    
                #print url_now
                #browser.save_screenshot('C:/Users/qiuxw/Desktop/1.png')
                #get the sourse page
                #page_html = browser.page_source.encode('GB18030')
                
                result_Match = []
                
                #re find the key word
                match_Zn = re.findall(r"Zn", page_html)
                match_zinc = re.findall(r"zinc",page_html)
                match_Fe = re.findall(r"Fe\W", page_html)
                match_iron = re.findall(r"\Wiron\W",page_html)
                match_Mn = re.findall(r"Mn\W", page_html)
                match_manganese = re.findall(r"manganese\W",page_html)
                match_Mg = re.findall(r"Mg\W", page_html)
                match_magnesium = re.findall(r"magnesium\W",page_html)
                match_Ca = re.findall(r"Ca\W", page_html)
                match_calcium = re.findall(r"calcium\W",page_html)
                match_Ni = re.findall(r"Ni\W", page_html)
                match_nickel = re.findall(r"nickel\W",page_html)
                
                print "%d sequence" % i
                #define the match rule
                if match_Zn or match_zinc:
                #if match_zinc:
                    result_Match.append('Zn related')
                    #w_xls.get_sheet(0).write(k + 1,25,"Zn related, %s" % url_pro )
                if match_Fe or match_iron:
                #if match_iron:
                    result_Match.append('Fe related')
                    #w_xls.get_sheet(0).write(k + 1,25,"Fe related, %s "% url_pro )
                if match_manganese or match_Mn:
                    result_Match.append('Mn related')
                    #w_xls.get_sheet(0).write(k + 1,25,"Mn related, %s" % url_pro)
                if match_Mg or match_magnesium:
                    result_Match.append('Mg related')
                    #w_xls.get_sheet(0).write(k + 1,25,"MG related, %s" % url_pro)
                if match_Ca or match_calcium:
                    result_Match.append('Ca related')
                    #w_xls.get_sheet(0).write(k + 1,2,"Ca related, %s" % url_pro)
                if match_nickel or match_Ni:
                    result_Match.append('Ni related')
                    #w_xls.get_sheet(0).write(k + 1,30,"Ni related, %s" % url_pro)
                
                
                if len(result_Match) == 0:
                    print "Nothing Found"
                    w_xls.get_sheet(0).write(k + 1,25+i,"No Metals Found1")
                else: 
                    write_ex = ''
                
                    for istr in result_Match:
                        print istr
                        if re.findall(r"Zn", istr):
                            write_ex = write_ex + "Zn"
                        elif re.findall(r"Fe", istr):
                            write_ex = write_ex + "/Fe"
                        elif re.findall(r"Mn", istr):
                            write_ex = write_ex + "/Mn"  
                        elif re.findall(r"Mg", istr):
                            write_ex = write_ex + "/Mg" 
                        elif re.findall(r"Ca", istr):
                            write_ex = write_ex + "/Ca"  
                        elif re.findall(r"Ni", istr):
                            write_ex = write_ex + "/Ni"  
                    w_xls.get_sheet(0).write(k + 1,25+i,"%s related1, %s" % (write_ex,match_urls[i]))
    #time print out
    #print i
    print k
    w_xls.get_sheet(0).write(k+1,25+i+1,"%s" %(time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))))                
    w_xls.get_sheet(0).write(k+1,25+i+2,"%d" % 1)
    w_xls.save('F:/Python/realwork/excelExtract/NoOpendata/Total_findnodone1-1-1.xls')
    k = k +1                
     
