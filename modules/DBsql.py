#!/usr/bin/env python
# coding: utf-8

# In[2]:


import pyodbc


# In[ ]:


class N41DB:
    def __init__(self, username ='', password ='', server = ''):
        self.username = username
        self.password = password
        self.server = server
        self.conn = None
        self.cursor = None
    #db.key
    def loadKey(self, path):
        with open(path, 'r') as f:
            self.username = f.readline().strip()
            self.password = f.readline().strip()
            self.server = f.readline().strip()
        return
    
    def connectDB(self)
        self.conn = pyodbc.connect('Trusted_Connection=yes;'+'DRIVER={SQL Server Native Client 11.0};SERVER='+server+';DATABASE=N41'+';UID='+username+';PWD='+ password+';Encrypt=no;autocommit=True')
        self.cursor = conn.cursor()
    def dropDB(self)
        self.cursor.close()
        self.conn.close()
        
    def CSVtoDB_styleAttribute(self, siteid):
        self.connectDB()
        
        query = '''
        declare @siteid as int = '''+ str(siteid) +'''

        create table #temp(
        ind int,
        style varchar(20),
        attributeValue varchar(4000)
        )

        BULK INSERT #temp
        FROM 'C:\Ninexis.csv' -- file path
        WITH
        (
            FIRSTROW = 2, -- as 1st one is header
            FIELDTERMINATOR = ',',
            ROWTERMINATOR = '\n',
            TABLOCK
        )

        alter table #temp ADD siteid int
        UPDATE #temp SET siteid = @siteid;

        WITH CTE AS(
           SELECT style,
               RN = ROW_NUMBER()OVER(PARTITION BY style ORDER BY style)
           FROM #temp
        )
        DELETE FROM CTE WHERE RN > 1

        UPDATE #temp SET attributeValue = REPLACE(attributeValue, '"', '')

        delete from #temp where style in 
        (select style from nvlt_stylecolor where division = '' or bundle = '')

        insert into nvlt_eav_styleattribute (siteid, style, attribute,attributevalue)
        select distinct tmp.siteid, tmp.style, 'Category', tmp.attributevalue
        from #temp tmp
        where tmp.style not in (select style from nvlt_eav_styleattribute where siteid = @siteid and attribute = 'Category')

        insert into nvlt_eav_styleattribute (siteid, style, attribute)
        select tmp.siteid, tmp.style, 'Description'
        from #temp tmp
        where tmp.style not in (select style from nvlt_eav_styleattribute where siteid = @siteid and attribute = 'Description')

        insert into nvlt_eav_styleattribute (siteid, style, attribute)
        select tmp.siteid, tmp.style, 'Detail'
        from #temp tmp
        where tmp.style not in (select style from nvlt_eav_styleattribute where siteid = @siteid and attribute = 'Detail')

        insert into nvlt_eav_styleattribute (siteid, style, attribute)
        select tmp.siteid, tmp.style, 'MetaDescription'
        from #temp tmp
        where tmp.style not in (select style from nvlt_eav_styleattribute where siteid = @siteid and attribute = 'MetaDescription')

        insert into nvlt_eav_styleattribute (siteid, style, attribute)
        select tmp.siteid, tmp.style, 'MetaTitle'
        from #temp tmp
        where tmp.style not in (select style from nvlt_eav_styleattribute where siteid = @siteid and attribute = 'MetaTitle')

        insert into nvlt_eav_styleattribute (siteid, style, attribute)
        select tmp.siteid, tmp.style, 'MetaKeyword'
        from #temp tmp
        where tmp.style not in (select style from nvlt_eav_styleattribute where siteid = @siteid and attribute = 'MetaKeyword')

        drop table #temp
        '''
        self.cursor.execute(query)
        self.dropDB()


# In[ ]:


def CSVtoDB_sitePublish(self, siteid):
    self.connectDB()
    
    query = '''
    declare @siteid as int = '''+ str(siteid) +'''

    select sc.style, sc.color
    into #temp2
    from nvlt_stylecolor sc
    where division != '' and
         bundle !='' and
         style in(select distinct style
                from nvlt_eav_styleattribute
                where siteid = 1)

    alter table #temp2
    ADD siteid int, toPublish int, active int, published int;

    UPDATE #temp2
    SET siteid = @siteid, toPublish = 1, active = 1, published = 0

    insert into NVLT_EAV_SITEPUBLISH (siteid, style, color, toPublish, active,published)
    select tmp.siteid, tmp.style, tmp.color, tmp.toPublish, tmp.active, tmp.published
    from #temp2 tmp
    WHERE style not in (Select style from nvlt_eav_sitePublish
                        where siteid = '''+ str(siteid) '''
                        )

    DROP TABLE #temp2;
    '''
    self.cursor.execute(query)
    self.dropDB()

