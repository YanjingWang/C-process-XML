##pip install pymssql
import pymssql
import time

requiredKeys = ['title',
                'link',
                'description',
                'course:classStatus',
                'course:code',
                'course:classNum',
                'course:classLanguage',
                'course:classStartDate',
                'course:classEndDate',
                'course:city',
                'course:ISOCtry',
                'course:ISOSUBL1',
                'course:ISOSUBL2',
                'course:Modality',
                'course:Duration',
                'course:DurationUnits',
                'course:IRLPSupport',
                'course:VirtualWorkstationRequired',
                'course:ClassType',
                'course:TimeZone',
                'course:StartTime',
                'course:LastDayEndTime',
                'course:gtr',
                'course:extPartner',
                'course:instructorName',
                'course:instructorEmail',
                'course:siteContactName',
                'course:siteContactPhone',
                'course:siteContactEmail',
                'course:confirmedNumberofExternalStudents',
                'course:confirmedNumberofIBMStudents',
                'course:comments',
                'pubDate',
                'guid']

def GetRequiredColList(colDesc, cols):
    colDict = {}
    # coursecode, event_id, description, classstatus, classNum, classStartDate,
    # classEndDate, city, ISOCtry, ISOSUBL1, Modality, Duration, DurationUnits,
    # ClassType, TimeZone, StartTime, LastDayEndTime, gtr, confirmedNumberofExternalStudents,
    # confirmedNumberofIBMStudents, pubDate
    for colDesc, col in zip(colDescs, cols):
        colName = colDesc[0]
        if col is None:
            colDict[colName] = ''
        else:
            colDict[colName] = col
    requiredColList = []
    requiredColList.append(colDict['coursecode'] + '-' + colDict['event_id'])
    requiredColList.append('http://db.globalknowledge.com/olm/ibmgo.aspx?coursecode=' + colDict['coursecode'])
    requiredColList.append(colDict['description'])
    requiredColList.append(colDict['classstatus'])
    requiredColList.append(colDict['coursecode'])
    requiredColList.append(colDict['classNum'])
    # TODO:
    requiredColList.append('EN')
    requiredColList.append(colDict['classStartDate'])
    requiredColList.append(colDict['classEndDate'])
    requiredColList.append(colDict['city'])
    requiredColList.append(colDict['ISOCtry'])
    requiredColList.append(colDict['ISOSUBL1'])
    # TODO:
    requiredColList.append(colDict['ISOSUBL1'])
    requiredColList.append(colDict['Modality'])
    requiredColList.append(colDict['Duration'])
    requiredColList.append(colDict['DurationUnits'])
    # TODO:
    requiredColList.append('0')
    # TODO:
    requiredColList.append('0')
    requiredColList.append(colDict['ClassType'])
    requiredColList.append(colDict['TimeZone'])
    requiredColList.append(colDict['StartTime'])
    requiredColList.append(colDict['LastDayEndTime'])
    requiredColList.append(colDict['gtr'])
    # TODO:
    requiredColList.append('')
    requiredColList.append('')
    requiredColList.append('')
    requiredColList.append('')
    requiredColList.append('')
    requiredColList.append('')
    requiredColList.append(colDict['confirmedNumberofExternalStudents'])
    requiredColList.append(colDict['confirmedNumberofIBMStudents'])
    requiredColList.append('')
    requiredColList.append(colDict['pubDate'])
    requiredColList.append(requiredColList[1])
    return requiredColList

def FormHeader(filePath):
    content = []
    content.append('<?xml version="1.0" encoding="ISO-8859-1"?>\n')
    content.append('<rss version="2.0" xmlns:atom="http://www.w3.org/2005/atom" xmlns:course="http://db.globalknowledge.com/ibm/ww/rss-ibm">\n')
    content.append('<channel>\n')
    content.append('    <title>Global Knowledge</title>\n')
    content.append('    <link>http://db.globalknowledge.com/ibm/ww/rss-ibm</link>\n')
    content.append('    <atom:link href="http://db.globalknowledge.com/ibm/ww/rss-ibm" rel="self" type="application/rss+xml" />\n')
    content.append('    <description>A comprehensive list of IBM classes offered by Global Knowledge.</description>\n')
    content.append('    <language>en-gb</language>\n')
    curTime = time.strftime("%a, %d %b %Y %H:%M:%S", time.localtime())
    content.append('    <pubDate>' + curTime + ' EDT</pubDate>\n')
    # TODO don't know how to get date yet
    content.append('    <lastBuildDate>Sat, 11 October 2021 05:01:23 EDT</lastBuildDate>\n')
    with open(filePath, 'a') as f:
        f.writelines(content)

def FormBody(filePath, colDescs, cols):
    content = []
    content.append('    <item>\n')
    requiredCols = GetRequiredColList(colDescs, cols)
    for key, col in zip(requiredKeys, requiredCols):
        line = '        <' + key + '>'
        line += col
        line += '</' + key + '>\n'
        content.append(line)
    content.append('    </item>\n')
    with open(filePath, 'a') as f:
        f.writelines(content)


def FormFooter(filePath):
    content = []
    content.append('</channel>\n')
    content.append('</rss>\n')
    with open(filePath, 'a') as f:
        f.writelines(content)

#connect to azure sql database
connect = pymssql.connect('sql-dev-dwh-amer.database.windows.net', 'FileExports_Dev', 'xm!OTsuidPaNATGi9Eccw', 'dwh_amer_dev')
if connect:
    print("Connect Successfully!")

cursor = connect.cursor()    #Create a cursor,execute sql statement through cursor in Python
sql =  "select tx.coursecode,ed.event_id,\
       ltrim(replace(replace(STUFF( tx.title,1,1,UPPER(substring(tx.title,1,1))),'&','&amp;'),CHAR(39),'&apos;')) as description ,\
       CONVERT(varchar(100), 1) as classstatus,\
       case when cd.md_num in ('32','44') then null else ed.event_id end as classNum,\
       case when cd.md_num in ('32','44') then null else CONVERT(varchar(100), ed.start_date, 23) end as classStartDate,\
       case when cd.md_num in ('32','44') then null else CONVERT(varchar(100), ed.end_date, 23) end as classEndDate,\
       case when cd.md_num in ('32','44') then null when cd.md_num='20' then 'VIRTUAL EASTERN' else upper(ed.city) end as city,\
       case when cd.md_num in ('32','44','20','42') then null else upper(substring(ed.country,1,2)) end as ISOCtry,\
       case when cd.md_num in ('32','44') then null else case when upper(substring(ed.country,1,2)) = 'US' then 'US-' else null end+upper(ed.state) end as ISOSUBL1,\
       case when cd.md_num in ('20','42') then 'ILO' \
                                                 when cd.md_num in ('32','44') and tx.modality like '%SPVC%' then 'SPVC'\
                                                 when cd.md_num in ('32','44') and tx.modality like '%WBT%' then 'WBT' \
                                                 else 'CR' end as Modality,\
       tx.duration_length as Duration,\
       tx.duration_unit as DurationUnits,\
       case when cd.ch_num='20' then 'Private' else 'Public' end as ClassType,\
       case when cd.md_num in ('32','44') then null when ROUND(et.offsetfromutc, 0) <> et.offsetfromutc then concat(ROUND(et.offsetfromutc, 0),'00')+':30' else concat(ROUND(et.offsetfromutc, 0),'00')+':00' end as TimeZone,\
       case when cd.md_num in ('32','44') then null else CONVERT(varchar(100), start_time, 108) end as StartTime,\
       case when cd.md_num in ('32','44') then null else CONVERT(varchar(100), end_time, 108) end as LastDayEndTime,\
       case when gtr.gtr_level is not null then '1' else '0' end as gtr,\
       CONVERT(varchar(100),sum(case when c1.acct_name not like 'IBM%' then 1 else 0 end)) as confirmedNumberofExternalStudents,\
       CONVERT(varchar(100),sum(case when c1.acct_name like 'IBM%' then 1 else 0 end)) as confirmedNumberofIBMStudents,\
       substring(datename(dw, getdate()+4/24),0,4)+' '+CONVERT(varchar(100),DATEPART(Day , getdate()+4/24))\
	+' ' +CONVERT(varchar(100),DATEPART(month , getdate()+4/24))+' ' +CONVERT(varchar(100),DATEPART(year , getdate()+4/24))+' '+CONVERT(varchar(100),getdate()+4/24, 108) as pubDate\
  from gkdw.event_dim ed\
       left outer join gkdw.gk_state_abbrev a on upper(ed.state) = a.state_abbrv \
       inner join gkdw.course_dim cd on ed.course_id = cd.course_id and ed.ops_country = cd.country \
       inner join gkdw.ibm_tier_xml tx on isnull(cd.mfg_course_code,cd.short_name) = tx.coursecode \
       inner join base.evxevent ev on ed.event_id = ev.evxeventid \
       inner join gkdw.gk_all_event_instr_mv ie on ed.event_id = ie.event_id \
       left outer join base.evxtimezone et on ev.evxtimezoneid = et.evxtimezoneid \
       left outer join gkdw.order_fact f on ed.event_id = f.event_id and f.enroll_status = 'Confirmed' \
       left outer join gkdw.cust_dim c1 on f.cust_id = c1.cust_id \
       left outer join gkdw.gk_gtr_events gtr on ed.event_id = gtr.event_id \
 where cd.course_pl = 'IBM' and substring(cd.course_code,-1,1) <> 'W' \
   and cd.country in ('USA','CANADA') \
   and ed.end_date >= CONVERT(varchar(100), GETDATE(), 23) \
   and ed.status = 'Open' \
   and cd.ch_num = '10' \
   and isnull(cd.mfg_course_code,cd.short_name) is not null \
   and not exists (select 1 from gkdw.ibm_rss_feed_tbl r where ed.event_id = r.event_id) \
  group by ed.event_id,cd.course_code,cd.course_name,tx.coursecode, \
         ltrim(replace(replace(STUFF( tx.title,1,1,UPPER(substring(tx.title,1,1))),'&','&amp;'),CHAR(39),'&apos;')),ed.start_date,ed.end_date, \
         case when cd.md_num in ('32','44') then null when cd.md_num='20' then 'VIRTUAL EASTERN' else upper(ed.city) end,upper(substring(ed.country,1,2)),upper(ed.state), \
          case when cd.md_num in ('20','42') then 'ILO' \
               when cd.md_num in ('32','44') and tx.modality like '%SPVC%' then 'SPVC' \
               when cd.md_num in ('32','44') and tx.modality like '%WBT%' then 'WBT' \
               else 'CR' end, \
          tx.duration_length,tx.duration_unit, \
          case when cd.ch_num='20' then 'Private' else 'Public' end,cd.md_num,cd.country, \
          et.offsetfromutc,start_time,end_time,ie.firstname1,ie.lastname1,ie.email1,gtr.gtr_level"
#print(sql)
cursor.execute(sql)   #execute sql statement
cols = cursor.fetchone()
colDescs = cursor.description
filePath = 'PythonDemo1.xml'
FormHeader(filePath)
while cols:
    FormBody(filePath, colDescs, cols)
FormFooter(filePath)
cursor.close()
connect.close()
print("XML file has been generated successfully")
