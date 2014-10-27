
	-- create tables
    CREATE TABLE [dbo].[qlikViewUser](
    	[entityid] [uniqueidentifier] NOT NULL,
    	[name] [nvarchar](100) NULL,
    	[descr] [nvarchar](300) NULL,
    	[email] [nvarchar](300) NULL,
     CONSTRAINT [PK_qlikViewUser] PRIMARY KEY CLUSTERED
    (
    	[entityid] ASC
    )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
    ) ON [PRIMARY]

	CREATE TABLE [dbo].[qlikViewGroup](
		[groupid] [uniqueidentifier] NOT NULL,
		[memberid] [uniqueidentifier] NOT NULL
	) ON [PRIMARY]

	--select * from qlikViewUser
	--select * from qlikViewGroup

	-- insert GAIM users
	Insert Into qlikviewUser
	select contactid [entityid], FirstName + '.' + LastName [name], AccountIdName + '/' + FullName as [descr], emailaddress1 [email]
	from crm_contact where new_gaimaccess = 1

	-- insert myself
	Insert Into qlikviewUser
	select contactid [entityid], FirstName + '.' + LastName [name], AccountIdName + '/' + FullName as [descr], emailaddress1 [email]
	from crm_contact where emailaddress1 = 'arthur.wang@greenwich.com'

	-- create groups
	insert into qlikviewUser values (newid(), 'Banking', 'Banking Security Group', 'contactus@greenwichassociates.com')
	insert into qlikviewUser values (newid(), 'IM', 'IM Security Group', 'contactus@greenwichassociates.com')
	insert into qlikviewUser values (newid(), 'S&T', 'S&T Security Group', 'contactus@greenwichassociates.com')
	insert into qlikviewUser values (newid(), 'Overall', 'Overall Security Group', 'contactus@greenwichassociates.com')
	insert into qlikviewUser values (newid(), 'ALL', 'ALL Security Group', 'contactus@greenwichassociates.com')

	-- assign Banking to "53"
	insert into qlikviewGroup
		select u1.entityid , u2.entityid
		from (select entityid from qlikviewUser where name='Banking') u1
		cross join (select entityid from qlikviewUser where email like '%@53%') u2
	-- assign S&Tto "BOA"
	insert into qlikviewGroup
		select u1.entityid , u2.entityid
		from (select entityid from qlikviewUser where name='S&T') u1
		cross join (select entityid from qlikviewUser where email like '%@bankofamerica%') u2
	-- assign IM to "chase"
	insert into qlikviewGroup
		select u1.entityid , u2.entityid
		from (select entityid from qlikviewUser where name='IM') u1
		cross join (select entityid from qlikviewUser where email like '%@chase%') u2
	-- assign Overall to chase Tara
	insert into qlikviewGroup
		select u1.entityid , u2.entityid
		from (select entityid from qlikviewUser where name='Overall') u1
		cross join (select entityid from qlikviewUser where name in ('Tara Riley')) u2
	-- assign ALL to internal members
	insert into qlikviewGroup
		select u1.entityid , u2.entityid
		from (select entityid from qlikviewUser where name='ALL') u1
		cross join (select entityid from qlikviewUser where email like '%@greenwich.com') u2
	-- assign all groups (except ALL itself) to ALL group -> assume 3 levels here
	insert into qlikviewGroup
		select u1.entityid , u2.groupid
		from (select entityid from qlikviewUser where name='ALL') u1
		cross join (select distinct groupid from qlikviewGroup where groupid not in (select entityid from qlikviewUser where name='ALL')) u2
