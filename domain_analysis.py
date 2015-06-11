import win32com.client
import time,os,string,re,sys
from urlparse import urlparse
import xml.etree.ElementTree as ET
from lxml import etree as ET
from optparse import OptionParser
__author__ = 'miaolian'

def get_xmlfile(url,filepath):
    control = win32com.client.Dispatch('HttpWatch.Controller')
    plugin = control.IE.New()

    plugin.Log.EnableFilter(False)
    plugin.Record()

    plugin.GotoURL(url)
    control.Wait(plugin, -1)

    plugin.Stop()

    if plugin.Log.Pages.Count != 0 :
        plugin.Log.ExportXML(filepath)
    plugin.CloseBrowser()

def create_filename(fold_path):
    timestamp = str(time.time())
    file_name = timestamp + ".xml"
    return os.path.join(fold_path,file_name)

def create_xmlname(url,fold_path):
    timestamp = str(int(time.time()))
    file_name = "SHYY" + timestamp +"_report.xml"
    #file_name = url + timestamp +"_report.xml"
    return os.path.join(fold_path,file_name)

def analysis_xml(file_path):
    tree = ET.parse(file_path)
    root = tree.getroot()
    dir_domain = {}
    for child in root:
        domain =  child.attrib.get("URL",None)
        dir_type = {}
        if not domain:
            continue
        for entry in child.getchildren():   
            if "{http://www.httpwatch.com/xml/log/5.1}request" == entry.tag:
                # get host
                for headers in entry.getchildren():
                    if "{http://www.httpwatch.com/xml/log/5.1}headers" != headers.tag:
                        continue
                    for n in headers.getchildren():
                        type_name = n.attrib.get("name",None)
                        if type_name == "Host":
                            type_value = n.text
                            dir_type[type_name] = type_value
            elif "{http://www.httpwatch.com/xml/log/5.1}response" == entry.tag:
                # get headers of response
                for headers in entry.getchildren():
                    if "{http://www.httpwatch.com/xml/log/5.1}headers" != headers.tag:
                        continue
                    for m in headers.getchildren():
                        type_name = m.attrib.get("name",None)
                        type_value = m.text
                        dir_type[type_name] = type_value
        dir_domain[domain] = dir_type
    return dir_domain                  

def deal_xmldata(dir_domainDatabasse):
    dir_domain = {}

    for key,value in dir_domainDatabasse.items():
        if not value:
            continue
        value_CacheControl = value.get("Cache-Control",None)
        value_Connection = value.get("Connection",None)
        value_ContentLength = value.get("Content-Length",None)
        value_Expires = value.get("Expires",None)
        value_LastModified = value.get("Last-Modified",None)
        value_ContentEncoding = value.get("Content-Encoding",None)
        value_Vary = value.get("Vary",None)
        value_ContentType = value.get("Content-Type",None)
        value_Host = value.get("Host",None)

        if not value_ContentType or not value_Host:
            continue

        if not dir_domain.has_key(value_Host):
            dir_type = {}
            dir_domain[value_Host] = dir_type

        if not dir_domain[value_Host].has_key(value_ContentType):
            dir_CacheControl = {}
            dir_Connection = {}
            dir_ContentLength = {}
            dir_Expires = {}
            dir_LastModified = {}
            dir_ContentEncoding = {}
            dir_Vary  = {}
            list_type = [dir_CacheControl,dir_Connection,dir_ContentLength,dir_Expires,dir_LastModified,dir_ContentEncoding,dir_Vary]
            dir_domain[value_Host][value_ContentType] = list_type

        if value_CacheControl:
            dir_domain[value_Host][value_ContentType][0].update({value_CacheControl:key})
        if value_Connection:
            dir_domain[value_Host][value_ContentType][1].update({"Connection":value_Connection})
        if value_ContentLength:
            dir_domain[value_Host][value_ContentType][2].update({"Content-Length":value_ContentLength})
        if value_Expires:
            dir_domain[value_Host][value_ContentType][3].update({"Expires":value_Expires})
        if value_LastModified:
            dir_domain[value_Host][value_ContentType][4].update({"Last-Modified":value_LastModified})
        if value_ContentEncoding:
            dir_domain[value_Host][value_ContentType][5].update({"Content-Encoding":value_ContentEncoding})
        if value_Vary:
            dir_domain[value_Host][value_ContentType][6].update({"Vary":value_Vary})

    return dir_domain

def put_xml(file_report,dir_domain):
    root_attrib = {"author":"miaolian"}
    root = ET.Element("WebReport",root_attrib)
    for domain,value in dir_domain.items():
        attrib = {"URL":domain}
        url = ET.SubElement(root, "domain",attrib)
        for type_name,type_value in value.items():          
            url_attrib = {"ConnectType":type_name}
            url_type = ET.SubElement(url, "type",url_attrib)           
            if type_value[0]:
                for CacheControl,value_CacheControl in type_value[0].items():
                    CacheControl_attrib = {"URL":value_CacheControl,"name":"Cache-Control"}
                    CacheControl_info = ET.SubElement(url_type, "headers",CacheControl_attrib)
                    CacheControl_info.text = CacheControl

            if type_value[1]:
                for Connection,value_Connection in type_value[1].items():
                    Connection_attrib = {"name":Connection}
                    Connection_info = ET.SubElement(url_type, "headers",Connection_attrib)
                    Connection_info.text = value_Connection

            if type_value[2]:
                for ContentLength,value_ContentLength in type_value[2].items():
                    ContentLength_attrib = {"name":ContentLength}
                    ContentLength_info = ET.SubElement(url_type, "headers",ContentLength_attrib)
                    ContentLength_info.text = value_ContentLength

            if type_value[3]:
                for Expires,value_Expires in type_value[3].items():
                    Expires_attrib = {"name":Expires}
                    Expires_info = ET.SubElement(url_type, "headers",Expires_attrib)
                    Expires_info.text = value_Expires

            if type_value[4]:
                for LastModified,value_LastModified in type_value[4].items():
                    LastModified_attrib = {"name":LastModified}
                    LastModified_info = ET.SubElement(url_type, "headers",LastModified_attrib)
                    LastModified_info.text = value_LastModified

            if type_value[5]:
                for ContentEncoding,value_ContentEncoding in type_value[5].items():
                    ContentEncoding_attrib = {"name":ContentEncoding}
                    ContentEncoding_info = ET.SubElement(url_type, "headers",ContentEncoding_attrib)
                    ContentEncoding_info.text = value_ContentEncoding

            if type_value[6]:
                for Vary,value_Vary in type_value[6].items():
                    Vary_attrib = {"name":Vary}
                    Vary_info = ET.SubElement(url_type, "headers",Vary_attrib)
                    Vary_info.text = value_Vary

    tree = ET.ElementTree(root)
    tree.write(file_report)

    parser = ET.XMLParser(
    remove_blank_text=False, resolve_entities=True, strip_cdata=True)
    xmlfile = ET.parse(file_report, parser)
    pretty_xml = ET.tostring(
    xmlfile, encoding = 'UTF-8', xml_declaration = True, pretty_print = True,
    doctype='<!DOCTYPE WebReport PUBLIC>')
    file = open(file_report, "w")
    file.writelines(pretty_xml)
    file.close()


if __name__ == '__main__':
    parser = OptionParser()
    parser.add_option("-u", "--url",dest="url",
                  help="Analysis URL (Must empty cache in IE)", metavar="www.mop.com")
    parser.add_option("-f", "--fold",dest="fold", 
                  help="report file path", metavar="E:\\")
    (options,args) = parser.parse_args()
    if options.url:
        url = options.url
    else:
        file_exitpath  = os.getcwd() + "\domain_analysis.py"
        print "None URL exit,Please use command:[ python %s -h ]for help" %(file_exitpath)
        sys.exit()
    if options.fold:
        fold_path = options.fold_path
    else:
        fold_path = os.getcwd()    
    dir_domainDatabasse = {}
    dir_domain = {}   
    file_path = create_filename(fold_path)

    get_xmlfile(url,file_path)
    dir_domainDatabasse = analysis_xml(file_path)   
    os.remove(file_path)
    dir_domain = deal_xmldata(dir_domainDatabasse)
    file_report = create_xmlname(url,fold_path)
    put_xml(file_report,dir_domain)
    print "Report:[%s] has been built" %(file_report)

    


