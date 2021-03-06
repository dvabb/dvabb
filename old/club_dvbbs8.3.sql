if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_Appraise]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_Appraise]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_Badlanguage]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_Badlanguage]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_BbsLink]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_BbsLink]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_BbsNews]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_BbsNews]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_BestTopic]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_BestTopic]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_Bill_Online_Payment]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_Bill_Online_Payment]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_Bill_Orders]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_Bill_Orders]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_Board]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_Board]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_BoardPermission]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_BoardPermission]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_BookMark]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_BookMark]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_ChallengeInfo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_ChallengeInfo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_ChanOrders]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_ChanOrders]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_Friend]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_Friend]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_GroupName]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_GroupName]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_GroupUser]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_GroupUser]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_Group_Board]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_Group_Board]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_Group_Class]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_Group_Class]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_Group_Topic]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_Group_Topic]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_Group_bbs]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_Group_bbs]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_Log]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_Log]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_Medal]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_Medal]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_MedalLog]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_MedalLog]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_Message]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_Message]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_MoneyLog]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_MoneyLog]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_Online]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_Online]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_Plus]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_Plus]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_Plus_Tools_Buss]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_Plus_Tools_Buss]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_Plus_Tools_Info]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_Plus_Tools_Info]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_Plus_Tools_MagicFace]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_Plus_Tools_MagicFace]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_Qcomic]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_Qcomic]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_Setup]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_Setup]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_SmallPaper]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_SmallPaper]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_Space_skin]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_Space_skin]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_Space_user]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_Space_user]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_Stylehelp]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_Stylehelp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_TableList]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_TableList]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_Templates]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_Templates]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_Topic]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_Topic]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_Upfile]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_Upfile]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_User]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_User]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_UserAccess]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_UserAccess]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_UserGroups]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_UserGroups]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_Vote]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_Vote]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_VoteUser]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_VoteUser]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_admin]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_admin]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_banzhu_config]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_banzhu_config]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_banzhu_log]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_banzhu_log]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_banzhu_user]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_banzhu_user]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dv_help]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dv_help]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[dv_album]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[dv_album]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[dv_albumbbs]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[dv_albumbbs]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[dv_albumconfig]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[dv_albumconfig]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[dv_albumfav]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[dv_albumfav]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[dv_albumimg]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[dv_albumimg]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[dv_albumtype]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[dv_albumtype]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[dv_bbs1]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[dv_bbs1]
GO

CREATE TABLE [dbo].[Dv_Appraise] (
	[AppraiseID] [int] IDENTITY (1, 1) NOT NULL ,
	[Boardid] [int] NOT NULL ,
	[TopicID] [int] NOT NULL ,
	[PostID] [int] NOT NULL ,
	[AType] [int] NOT NULL ,
	[ATitle] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[AContent] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[UserID] [int] NOT NULL ,
	[UserName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[DateTime] [smalldatetime] NULL ,
	[IP] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_Badlanguage] (
	[id] [int] IDENTITY (1, 1) NOT NULL ,
	[content] [ntext] COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_BbsLink] (
	[id] [int] NOT NULL ,
	[boardname] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[readme] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[url] [nvarchar] (150) COLLATE Chinese_PRC_CI_AS NULL ,
	[logo] [varchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[islogo] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_BbsNews] (
	[id] [int] IDENTITY (1, 1) NOT NULL ,
	[boardid] [int] NULL ,
	[title] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[content] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[username] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[addtime] [smalldatetime] NULL ,
	[bgs] [varchar] (100) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_BestTopic] (
	[id] [int] IDENTITY (1, 1) NOT NULL ,
	[Announceid] [int] NULL ,
	[RootID] [int] NULL ,
	[BoardID] [int] NULL ,
	[Title] [varchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[PostUserName] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[PostUserID] [int] NULL ,
	[DateAndTime] [smalldatetime] NULL ,
	[Expression] [varchar] (100) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_Bill_Online_Payment] (
	[id] [int] IDENTITY (1, 1) NOT NULL ,
	[title] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[content] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[api_uname] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[api_pass] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[api_ms] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[active] [tinyint] NULL ,
	[orders] [int] NULL ,
	[indeximg] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[orther] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[usermoney] [int] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_Bill_Orders] (
	[id] [int] IDENTITY (1, 1) NOT NULL ,
	[uname] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[uid] [int] NULL ,
	[acount] [float] NULL ,
	[ip] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[addtime] [datetime] NULL ,
	[orderid] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[active] [tinyint] NULL ,
	[proname] [nvarchar] (250) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_Board] (
	[boardid] [int] NOT NULL ,
	[BoardType] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[ParentID] [int] NOT NULL ,
	[ParentStr] [varchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[Depth] [smallint] NULL ,
	[RootID] [int] NULL ,
	[Child] [smallint] NULL ,
	[readme] [nvarchar] (250) COLLATE Chinese_PRC_CI_AS NULL ,
	[BoardMaster] [nvarchar] (250) COLLATE Chinese_PRC_CI_AS NULL ,
	[PostNum] [int] NULL ,
	[TopicNum] [int] NULL ,
	[indexIMG] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[todayNum] [int] NULL ,
	[boarduser] [text] COLLATE Chinese_PRC_CI_AS NULL ,
	[LastPost] [varchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[Orders] [int] NULL ,
	[sid] [int] NULL ,
	[Board_Setting] [text] COLLATE Chinese_PRC_CI_AS NULL ,
	[Board_Ads] [text] COLLATE Chinese_PRC_CI_AS NULL ,
	[Board_user] [varchar] (500) COLLATE Chinese_PRC_CI_AS NULL ,
	[IsGroupSetting] [varchar] (200) COLLATE Chinese_PRC_CI_AS NULL ,
	[BoardTopStr] [varchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[CID] [tinyint] NULL ,
	[Rules] [text] COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_BoardPermission] (
	[Pid] [int] IDENTITY (1, 1) NOT NULL ,
	[BoardID] [int] NULL ,
	[GroupID] [int] NULL ,
	[PSetting] [ntext] COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_BookMark] (
	[id] [int] IDENTITY (1, 1) NOT NULL ,
	[username] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[url] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[topic] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[addtime] [smalldatetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_ChallengeInfo] (
	[D_ForumID] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[D_UserName] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[D_Password] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[D_RealName] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[D_identityNo] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[D_sex] [varchar] (5) COLLATE Chinese_PRC_CI_AS NULL ,
	[D_postcode] [varchar] (10) COLLATE Chinese_PRC_CI_AS NULL ,
	[D_address] [varchar] (500) COLLATE Chinese_PRC_CI_AS NULL ,
	[D_receiver] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[D_email] [varchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[D_forumname] [varchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[D_forumurl] [varchar] (200) COLLATE Chinese_PRC_CI_AS NULL ,
	[D_telephone] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[D_mobile] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[D_forumProvider] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[D_version] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[D_challengePassWord] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_ChanOrders] (
	[O_ID] [int] IDENTITY (1, 1) NOT NULL ,
	[O_type] [smallint] NULL ,
	[O_mobile] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[O_Username] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[O_isApply] [tinyint] NULL ,
	[O_issuc] [tinyint] NULL ,
	[O_PayMoney] [float] NULL ,
	[O_Paycode] [varchar] (30) COLLATE Chinese_PRC_CI_AS NULL ,
	[O_BoardID] [int] NULL ,
	[O_TopicID] [int] NULL ,
	[O_AddTime] [smalldatetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_Friend] (
	[F_id] [int] IDENTITY (1, 1) NOT NULL ,
	[F_username] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[F_friend] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[F_addtime] [smalldatetime] NULL ,
	[F_userid] [int] NULL ,
	[F_Mod] [tinyint] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_GroupName] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[ClassID] [int] NULL ,
	[GroupName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[GroupInfo] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[AppUserID] [int] NOT NULL ,
	[AppUserName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[UserNum] [int] NOT NULL ,
	[Stats] [int] NOT NULL ,
	[PostNum] [int] NOT NULL ,
	[TopicNum] [int] NOT NULL ,
	[TodayNum] [int] NOT NULL ,
	[YesterdayNum] [int] NOT NULL ,
	[LimitUser] [int] NOT NULL ,
	[AppDate] [smalldatetime] NULL ,
	[PassDate] [smalldatetime] NULL ,
	[visitDate] [smalldatetime] NULL ,
	[Locked] [tinyint] NOT NULL ,
	[viewflag] [tinyint] NOT NULL ,
	[GroupLogo] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_GroupUser] (
	[id] [int] IDENTITY (1, 1) NOT NULL ,
	[GroupID] [int] NOT NULL ,
	[UserID] [int] NOT NULL ,
	[UserName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[IsLock] [int] NOT NULL ,
	[Intro] [nvarchar] (250) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_Group_Board] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[BoardName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[BoardInfo] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[IndexIMG] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[TopicNum] [int] NULL ,
	[TodayNum] [int] NULL ,
	[PostNum] [int] NULL ,
	[RootID] [int] NULL ,
	[LastPost] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[FoundDate] [smalldatetime] NULL ,
	[Rules] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[BoardStats] [int] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_Group_Class] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[ClassName] [nvarchar] (250) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[GroupCount] [int] NULL ,
	[Orders] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_Group_Topic] (
	[TopicID] [int] IDENTITY (1, 1) NOT NULL ,
	[Title] [nvarchar] (250) COLLATE Chinese_PRC_CI_AS NULL ,
	[GroupID] [int] NOT NULL ,
	[BoardID] [int] NOT NULL ,
	[PollID] [int] NOT NULL ,
	[LockTopic] [int] NOT NULL ,
	[Child] [int] NOT NULL ,
	[PostUserName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[PostUserID] [int] NOT NULL ,
	[DateAndTime] [smalldatetime] NULL ,
	[hits] [int] NULL ,
	[Expression] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[LastPost] [nvarchar] (250) COLLATE Chinese_PRC_CI_AS NULL ,
	[istop] [tinyint] NULL ,
	[LastPostTime] [datetime] NULL ,
	[isbest] [tinyint] NULL ,
	[Mode] [tinyint] NULL ,
	[TopicMode] [tinyint] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_Group_bbs] (
	[AnnounceID] [int] IDENTITY (1, 1) NOT NULL ,
	[ParentID] [int] NOT NULL ,
	[GroupID] [int] NOT NULL ,
	[BoardID] [int] NOT NULL ,
	[UserName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[postuserid] [int] NOT NULL ,
	[Topic] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[Body] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[DateAndTime] [smalldatetime] NULL ,
	[length] [int] NULL ,
	[RootID] [int] NOT NULL ,
	[layer] [int] NULL ,
	[orders] [int] NULL ,
	[isbest] [tinyint] NOT NULL ,
	[ip] [nvarchar] (40) COLLATE Chinese_PRC_CI_AS NULL ,
	[Expression] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[locktopic] [int] NOT NULL ,
	[signflag] [tinyint] NOT NULL ,
	[isagree] [nvarchar] (250) COLLATE Chinese_PRC_CI_AS NULL ,
	[isupload] [tinyint] NULL ,
	[UbbList] [nvarchar] (250) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_Log] (
	[l_id] [int] IDENTITY (1, 1) NOT NULL ,
	[l_announceid] [int] NULL ,
	[l_boardid] [int] NULL ,
	[l_touser] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[l_username] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[l_content] [varchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[l_addtime] [smalldatetime] NULL ,
	[l_ip] [varchar] (40) COLLATE Chinese_PRC_CI_AS NULL ,
	[l_type] [tinyint] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_Medal] (
	[id] [int] IDENTITY (1, 1) NOT NULL ,
	[MedalName] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[MedalDesc] [varchar] (255) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[MedalPic] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_MedalLog] (
	[id] [int] IDENTITY (1, 1) NOT NULL ,
	[UserId] [int] NOT NULL ,
	[UserName] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[MedalId] [int] NOT NULL ,
	[AwardUser] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[AwardDesc] [varchar] (255) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[AddTime] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_Message] (
	[id] [int] IDENTITY (1, 1) NOT NULL ,
	[sender] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[incept] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[title] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[content] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[flag] [int] NULL ,
	[sendtime] [smalldatetime] NOT NULL ,
	[delR] [int] NOT NULL ,
	[delS] [int] NOT NULL ,
	[isSend] [int] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_MoneyLog] (
	[Log_ID] [int] IDENTITY (1, 1) NOT FOR REPLICATION  NOT NULL ,
	[ToolsID] [int] NOT NULL ,
	[CountNum] [int] NOT NULL ,
	[Log_Money] [int] NULL ,
	[Log_Ticket] [int] NULL ,
	[AddUserName] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[AddUserID] [int] NULL ,
	[Log_IP] [varchar] (40) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Log_Time] [smalldatetime] NOT NULL ,
	[Log_Type] [tinyint] NULL ,
	[BoardID] [int] NULL ,
	[Conect] [text] COLLATE Chinese_PRC_CI_AS NULL ,
	[HMoney] [varchar] (250) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_Online] (
	[id] [float] NOT NULL ,
	[username] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[userclass] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[stats] [varchar] (250) COLLATE Chinese_PRC_CI_AS NULL ,
	[ip] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[actforip] [varchar] (30) COLLATE Chinese_PRC_CI_AS NULL ,
	[startime] [smalldatetime] NULL ,
	[lastimebk] [smalldatetime] NULL ,
	[boardid] [int] NULL ,
	[browser] [varchar] (200) COLLATE Chinese_PRC_CI_AS NULL ,
	[UserGroupID] [int] NULL ,
	[actCome] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[userhidden] [int] NULL ,
	[userid] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_Plus] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[Plus_Type] [varchar] (100) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Plus_Name] [varchar] (100) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Isuse] [tinyint] NOT NULL ,
	[Plus_Setting] [text] COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Mainpage] [varchar] (100) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[IsShowMenu] [tinyint] NOT NULL ,
	[plus_adminpage] [varchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[plus_id] [varchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[plus_Copyright] [varchar] (200) COLLATE Chinese_PRC_CI_AS NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_Plus_Tools_Buss] (
	[ID] [int] IDENTITY (1, 1) NOT FOR REPLICATION  NOT NULL ,
	[UserID] [int] NOT NULL ,
	[UserName] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[ToolsID] [int] NOT NULL ,
	[ToolsName] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[ToolsCount] [int] NOT NULL ,
	[SaleCount] [int] NOT NULL ,
	[UpdateTime] [smalldatetime] NOT NULL ,
	[SaleMoney] [int] NOT NULL ,
	[SaleTicket] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_Plus_Tools_Info] (
	[ID] [int] IDENTITY (1, 1) NOT FOR REPLICATION  NOT NULL ,
	[ToolsName] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[ToolsInfo] [varchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[ToolsImg] [varchar] (150) COLLATE Chinese_PRC_CI_AS NULL ,
	[IsStar] [smallint] NOT NULL ,
	[SysStock] [int] NOT NULL ,
	[UserStock] [int] NOT NULL ,
	[UserTicket] [int] NOT NULL ,
	[UserMoney] [int] NOT NULL ,
	[UserPost] [int] NOT NULL ,
	[UserWealth] [int] NOT NULL ,
	[UserEp] [int] NOT NULL ,
	[UserCp] [int] NOT NULL ,
	[UserGroupID] [varchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[BoardID] [text] COLLATE Chinese_PRC_CI_AS NULL ,
	[BuyType] [tinyint] NOT NULL ,
	[ToolsSetting] [text] COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_Plus_Tools_MagicFace] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[Title] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[MagicFace_s] [int] NULL ,
	[MagicFace_l] [int] NULL ,
	[MagicType] [smallint] NULL ,
	[iMoney] [int] NULL ,
	[iTicket] [int] NULL ,
	[tMoney] [int] NULL ,
	[tTicket] [int] NULL ,
	[MagicSetting] [varchar] (250) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_Qcomic] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[senable] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[semail] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[sid] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[spassword] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[skey] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[owidth] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[oheight] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[iwidth] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[iheight] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[ext1] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[ext2] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[ext3] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[ext4] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[ext5] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_Setup] (
	[id] [int] IDENTITY (1, 1) NOT NULL ,
	[Forum_Setting] [text] COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Forum_ads] [text] COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Forum_Badwords] [text] COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Forum_rBadword] [text] COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Forum_Maxonline] [int] NOT NULL ,
	[Forum_MaxonlineDate] [smalldatetime] NOT NULL ,
	[Forum_TopicNum] [int] NOT NULL ,
	[Forum_PostNum] [int] NOT NULL ,
	[Forum_TodayNum] [int] NOT NULL ,
	[Forum_UserNum] [int] NOT NULL ,
	[Forum_YesTerdayNum] [int] NOT NULL ,
	[Forum_MaxPostNum] [int] NOT NULL ,
	[Forum_MaxPostDate] [smalldatetime] NOT NULL ,
	[Forum_lastUser] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Forum_LastPost] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Forum_BirthUser] [text] COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Forum_Sid] [int] NOT NULL ,
	[Forum_Version] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Forum_NowUseBBS] [varchar] (8) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Forum_IsInstall] [tinyint] NOT NULL ,
	[Forum_challengePassWord] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Forum_Ad] [text] COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Forum_ChanName] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Forum_ChanSetting] [varchar] (250) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Forum_LockIP] [text] COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Forum_Cookiespath] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Forum_Boards] [text] COLLATE Chinese_PRC_CI_AS NULL ,
	[Forum_alltopnum] [varchar] (250) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Forum_pack] [varchar] (250) COLLATE Chinese_PRC_CI_AS NULL ,
	[Forum_Cid] [tinyint] NOT NULL ,
	[Forum_AvaSiteID] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[Forum_AvaSign] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[Forum_AdminFolder] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Forum_BoardXML] [text] COLLATE Chinese_PRC_CI_AS NULL ,
	[Forum_Css] [text] COLLATE Chinese_PRC_CI_AS NULL ,
	[Forum_apis] [text] COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_SmallPaper] (
	[S_ID] [int] IDENTITY (1, 1) NOT NULL ,
	[S_BoardID] [int] NULL ,
	[S_UserName] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[S_Title] [varchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[S_Content] [varchar] (1000) COLLATE Chinese_PRC_CI_AS NULL ,
	[S_Hits] [int] NULL ,
	[S_Addtime] [smalldatetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_Space_skin] (
	[id] [int] IDENTITY (1, 1) NOT NULL ,
	[s_name] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[s_username] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[s_userid] [int] NULL ,
	[s_css] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[s_style] [int] NULL ,
	[s_path] [nvarchar] (150) COLLATE Chinese_PRC_CI_AS NULL ,
	[s_lock] [int] NOT NULL ,
	[s_addtime] [smalldatetime] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_Space_user] (
	[id] [int] IDENTITY (1, 1) NOT NULL ,
	[userid] [int] NOT NULL ,
	[username] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[title] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[intro] [nvarchar] (250) COLLATE Chinese_PRC_CI_AS NULL ,
	[s_left] [nvarchar] (250) COLLATE Chinese_PRC_CI_AS NULL ,
	[s_right] [nvarchar] (250) COLLATE Chinese_PRC_CI_AS NULL ,
	[s_center] [nvarchar] (250) COLLATE Chinese_PRC_CI_AS NULL ,
	[s_css] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[s_style] [int] NOT NULL ,
	[s_path] [nvarchar] (150) COLLATE Chinese_PRC_CI_AS NULL ,
	[updatetime] [smalldatetime] NULL ,
	[lock] [int] NOT NULL ,
	[set] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[cachedb] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[ownercachedb] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[plusdb] [ntext] COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_Stylehelp] (
	[id] [int] IDENTITY (1, 1) NOT NULL ,
	[StyleName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Main_Style] [ntext] COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Style_Pic] [ntext] COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[page_index] [ntext] COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[page_dispbbs] [ntext] COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[page_showerr] [ntext] COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[page_login] [ntext] COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[page_online] [ntext] COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[page_usermanager] [ntext] COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[page_fmanage] [ntext] COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[page_boardstat] [ntext] COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[page_paper_even_toplist] [ntext] COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[page_query] [ntext] COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[page_show] [ntext] COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[page_dispuser] [ntext] COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[page_help_permission] [ntext] COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[page_postjob] [ntext] COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[page_post] [ntext] COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[page_boardhelp] [ntext] COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[page_indivgroup] [ntext] COLLATE Chinese_PRC_CI_AS NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_TableList] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[TableName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[TableType] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_Templates] (
	[Id] [int] IDENTITY (1, 1) NOT NULL ,
	[Type] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[Folder] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[Layout] [ntext] COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_Topic] (
	[TopicID] [int] IDENTITY (1, 1) NOT NULL ,
	[Title] [varchar] (255) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Boardid] [int] NOT NULL ,
	[PollID] [int] NULL ,
	[LockTopic] [int] NOT NULL ,
	[Child] [int] NULL ,
	[PostUsername] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[PostUserid] [int] NOT NULL ,
	[DateAndTime] [smalldatetime] NOT NULL ,
	[hits] [int] NOT NULL ,
	[Expression] [varchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[VoteTotal] [int] NULL ,
	[LastPost] [varchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[LastPostTime] [datetime] NOT NULL ,
	[istop] [tinyint] NOT NULL ,
	[isvote] [tinyint] NULL ,
	[isbest] [tinyint] NULL ,
	[PostTable] [varchar] (8) COLLATE Chinese_PRC_CI_AS NULL ,
	[SmsUserList] [text] COLLATE Chinese_PRC_CI_AS NULL ,
	[IsSmsTopic] [tinyint] NULL ,
	[LastSmsTime] [smalldatetime] NULL ,
	[TopicMode] [tinyint] NULL ,
	[Mode] [smallint] NULL ,
	[GetMoney] [int] NOT NULL ,
	[UseTools] [varchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[GetMoneyType] [tinyint] NOT NULL ,
	[HideName] [tinyint] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_Upfile] (
	[F_ID] [int] IDENTITY (1, 1) NOT NULL ,
	[F_AnnounceID] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[F_BoardID] [int] NULL ,
	[F_UserID] [int] NULL ,
	[F_Username] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[F_Filename] [varchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[F_Viewname] [varchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[F_FileType] [varchar] (10) COLLATE Chinese_PRC_CI_AS NULL ,
	[F_Type] [smallint] NULL ,
	[F_FileSize] [int] NULL ,
	[F_Readme] [varchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[F_DownNum] [int] NULL ,
	[F_ViewNum] [int] NULL ,
	[F_DownUser] [text] COLLATE Chinese_PRC_CI_AS NULL ,
	[F_Flag] [tinyint] NULL ,
	[F_AddTime] [smalldatetime] NULL ,
	[F_OldName] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[f_topicid] [int] NULL ,
	[f_bbsid] [int] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_User] (
	[UserID] [int] IDENTITY (1, 1) NOT NULL ,
	[UserName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[UserPassword] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[UserEmail] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[UserPost] [int] NULL ,
	[UserTopic] [int] NULL ,
	[UserSign] [varchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[UserSex] [tinyint] NULL ,
	[UserFace] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[UserWidth] [int] NULL ,
	[UserHeight] [int] NULL ,
	[UserIM] [text] COLLATE Chinese_PRC_CI_AS NULL ,
	[JoinDate] [smalldatetime] NULL ,
	[LastLogin] [smalldatetime] NULL ,
	[UserLogins] [int] NULL ,
	[UserViews] [int] NULL ,
	[LockUser] [tinyint] NULL ,
	[UserClass] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[UserGroup] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[UserMoney] [int] NULL ,
	[UserTicket] [int] NULL ,
	[userWealth] [int] NULL ,
	[userEP] [int] NULL ,
	[userCP] [int] NULL ,
	[UserPower] [int] NULL ,
	[UserDel] [int] NULL ,
	[UserIsBest] [int] NULL ,
	[UserTitle] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[UserBirthday] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[UserQuesion] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[UserAnswer] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[UserLastIP] [varchar] (30) COLLATE Chinese_PRC_CI_AS NULL ,
	[UserPhoto] [varchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[UserFav] [varchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[UserInfo] [text] COLLATE Chinese_PRC_CI_AS NULL ,
	[UserSetting] [varchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[UserGroupID] [int] NOT NULL ,
	[TitlePic] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[UserHidden] [tinyint] NULL ,
	[UserMsg] [varchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[IsChallenge] [tinyint] NULL ,
	[UserMobile] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[TruePassWord] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[UserToday] [varchar] (250) COLLATE Chinese_PRC_CI_AS NULL ,
	[UserIsAva] [tinyint] NULL ,
	[UserAvaSetting] [text] COLLATE Chinese_PRC_CI_AS NULL ,
	[FollowMsgID] [text] COLLATE Chinese_PRC_CI_AS NULL ,
	[Vip_StarTime] [smalldatetime] NULL ,
	[Vip_EndTime] [smalldatetime] NULL ,
	[Passport] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[Fav_boards] [varchar] (250) COLLATE Chinese_PRC_CI_AS NULL ,
	[UserMedal] [varchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[RLActTimeT] [int] NULL ,
	[UserIsAudit_Custom] [int] NULL ,
	[UserIsAudit] [int] NULL ,
	[regip] [nvarchar] (40) COLLATE Chinese_PRC_CI_AS NULL ,
	[LastMsg] [smalldatetime] NULL ,
	[lastipinfo] [nvarchar] (250) COLLATE Chinese_PRC_CI_AS NULL ,
	[TyMedaled] [int] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_UserAccess] (
	[uc_UserID] [int] NULL ,
	[uc_BoardID] [int] NULL ,
	[uc_Setting] [ntext] COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_UserGroups] (
	[UserGroupID] [int] IDENTITY (1, 1) NOT NULL ,
	[title] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[usertitle] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[GroupSetting] [text] COLLATE Chinese_PRC_CI_AS NULL ,
	[Orders] [smallint] NULL ,
	[MinArticle] [int] NULL ,
	[TitlePic] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[GroupPic] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[ParentGID] [int] NULL ,
	[IsSetting] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[UserGroupIsAudit] [int] NULL ,
	[TyClsGroup] [int] NULL ,
	[TyClsGroupM] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_Vote] (
	[voteid] [int] IDENTITY (1, 1) NOT NULL ,
	[vote] [text] COLLATE Chinese_PRC_CI_AS NULL ,
	[votenum] [text] COLLATE Chinese_PRC_CI_AS NULL ,
	[votetype] [int] NULL ,
	[LockVote] [int] NULL ,
	[voters] [int] NULL ,
	[TimeOut] [smalldatetime] NULL ,
	[UArticle] [int] NULL ,
	[UWealth] [int] NULL ,
	[UEP] [int] NULL ,
	[UCP] [int] NULL ,
	[UPower] [int] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_VoteUser] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[VoteID] [int] NULL ,
	[UserID] [int] NULL ,
	[VoteDate] [smalldatetime] NULL ,
	[VoteOption] [text] COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_admin] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[Username] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Password] [varchar] (40) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Flag] [text] COLLATE Chinese_PRC_CI_AS NULL ,
	[LastLogin] [smalldatetime] NULL ,
	[LastLoginIP] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[Adduser] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[AcceptIP] [varchar] (255) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_banzhu_config] (
	[Gzdays] [smallint] NULL ,
	[gzifmess] [nvarchar] (5) COLLATE Chinese_PRC_CI_AS NULL ,
	[pagenum] [smallint] NULL ,
	[Gzjbgz] [int] NULL ,
	[gzsuperfj] [int] NULL ,
	[gzifarticle] [nvarchar] (5) COLLATE Chinese_PRC_CI_AS NULL ,
	[gzparticle] [real] NULL ,
	[gziflogins] [nvarchar] (5) COLLATE Chinese_PRC_CI_AS NULL ,
	[gzplogins] [real] NULL ,
	[gziflogs] [nvarchar] (5) COLLATE Chinese_PRC_CI_AS NULL ,
	[gzplogs] [real] NULL ,
	[ifautocancel] [nvarchar] (5) COLLATE Chinese_PRC_CI_AS NULL ,
	[autocday] [smallint] NULL ,
	[Gzlbt] [int] NULL ,
	[gzhmon] [int] NULL ,
	[gzwealth] [int] NULL ,
	[gzuserep] [int] NULL ,
	[gzusercp] [int] NULL ,
	[gzpower] [int] NULL ,
	[LeaveTimeout] [int] NULL ,
	[rUrl] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_banzhu_log] (
	[username] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[LeaveTime] [datetime] NULL ,
	[LeaveDays] [int] NULL ,
	[Leave] [nvarchar] (200) COLLATE Chinese_PRC_CI_AS NULL ,
	[CancelLeave] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_banzhu_user] (
	[username] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[FuLiDate] [smalldatetime] NULL ,
	[oldlogs] [int] NULL ,
	[oldlogins] [int] NULL ,
	[lastinfo] [nvarchar] (120) COLLATE Chinese_PRC_CI_AS NULL ,
	[oldarticle] [int] NULL ,
	[IsLeave] [int] NULL ,
	[Leave] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[LeaveTime] [smalldatetime] NULL ,
	[Leavedays] [int] NOT NULL ,
	[IsSendMail] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dv_help] (
	[H_ID] [int] IDENTITY (1, 1) NOT NULL ,
	[H_ParentID] [int] NULL ,
	[H_title] [varchar] (250) COLLATE Chinese_PRC_CI_AS NULL ,
	[H_content] [text] COLLATE Chinese_PRC_CI_AS NULL ,
	[H_type] [tinyint] NULL ,
	[H_stype] [int] NULL ,
	[H_bgimg] [varchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[H_Addtime] [smalldatetime] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[dv_album] (
	[id] [int] IDENTITY (1, 1) NOT NULL ,
	[userid] [int] NOT NULL ,
	[username] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[albumtitle] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[albumtype] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[albumsubtype] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[albumtext] [nvarchar] (200) COLLATE Chinese_PRC_CI_AS NULL ,
	[topimgid] [int] NOT NULL ,
	[ishowdo] [int] NOT NULL ,
	[accesspwd] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[albumclass] [int] NULL ,
	[favcount] [int] NOT NULL ,
	[photonum] [int] NOT NULL ,
	[albumsize] [int] NULL ,
	[albummaxsize] [int] NULL ,
	[createdate] [datetime] NOT NULL ,
	[albumcover] [nvarchar] (200) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[dv_albumbbs] (
	[id] [int] IDENTITY (1, 1) NOT NULL ,
	[imgid] [int] NOT NULL ,
	[userid] [int] NOT NULL ,
	[username] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[bbstext] [text] COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[newphoto] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[newphototitle] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[createdate] [datetime] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[dv_albumconfig] (
	[albumglobal] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[album] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[albumimg] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[albumbbs] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[albumfav] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[albumtype] [ntext] COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[dv_albumfav] (
	[userid] [int] NOT NULL ,
	[username] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[favtype] [int] NOT NULL ,
	[favid] [int] NOT NULL ,
	[recdate] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[dv_albumimg] (
	[id] [int] IDENTITY (1, 1) NOT NULL ,
	[albumid] [int] NOT NULL ,
	[userid] [int] NOT NULL ,
	[username] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[imgtitle] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[imgtext] [nvarchar] (200) COLLATE Chinese_PRC_CI_AS NULL ,
	[clickcount] [int] NOT NULL ,
	[bbscount] [int] NOT NULL ,
	[favcount] [int] NOT NULL ,
	[imgfile] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[zoomfile] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[filesize] [int] NOT NULL ,
	[uploaddate] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[dv_albumtype] (
	[id] [int] IDENTITY (1, 1) NOT NULL ,
	[albumtype] [nvarchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[dv_bbs1] (
	[AnnounceID] [int] IDENTITY (1, 1) NOT NULL ,
	[ParentID] [int] NOT NULL ,
	[BoardID] [int] NOT NULL ,
	[UserName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Topic] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[Body] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[DateAndTime] [smalldatetime] NOT NULL ,
	[length] [int] NULL ,
	[RootID] [int] NOT NULL ,
	[layer] [smallint] NULL ,
	[orders] [int] NOT NULL ,
	[isbest] [tinyint] NULL ,
	[ip] [varchar] (40) COLLATE Chinese_PRC_CI_AS NULL ,
	[Expression] [varchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[locktopic] [int] NOT NULL ,
	[signflag] [tinyint] NULL ,
	[emailflag] [tinyint] NULL ,
	[isagree] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[PostUserID] [int] NULL ,
	[IsAudit] [tinyint] NULL ,
	[IsUpload] [tinyint] NULL ,
	[PostBuyUser] [text] COLLATE Chinese_PRC_CI_AS NULL ,
	[Ubblist] [varchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[GetMoney] [int] NOT NULL ,
	[UseTools] [varchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[GetMoneyType] [tinyint] NOT NULL ,
	[FlashId] [varchar] (255) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

ALTER TABLE [dbo].[Dv_Appraise] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dv_Appraise] PRIMARY KEY  CLUSTERED 
	(
		[AppraiseID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_Badlanguage] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dv_Badlanguage] PRIMARY KEY  CLUSTERED 
	(
		[id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_BbsNews] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dv_BbsNews] PRIMARY KEY  CLUSTERED 
	(
		[id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_BestTopic] WITH NOCHECK ADD 
	CONSTRAINT [PK_BestTopic] PRIMARY KEY  CLUSTERED 
	(
		[id]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_Bill_Online_Payment] WITH NOCHECK ADD 
	CONSTRAINT [PK_Ht_Online_Payment] PRIMARY KEY  CLUSTERED 
	(
		[id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_Bill_Orders] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dv_Bill_Orders] PRIMARY KEY  CLUSTERED 
	(
		[id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_BoardPermission] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dv_BoardPermission] PRIMARY KEY  CLUSTERED 
	(
		[Pid]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_BookMark] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dv_BookMark] PRIMARY KEY  CLUSTERED 
	(
		[id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_ChanOrders] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dv_ChanOrders] PRIMARY KEY  CLUSTERED 
	(
		[O_ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_Friend] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dv_Friend] PRIMARY KEY  CLUSTERED 
	(
		[F_id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_GroupName] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dv_GroupName] PRIMARY KEY  CLUSTERED 
	(
		[ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_GroupUser] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dv_GroupUser] PRIMARY KEY  CLUSTERED 
	(
		[id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_Group_Board] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dv_Group_Board] PRIMARY KEY  CLUSTERED 
	(
		[ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_Group_Topic] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dv_Group_Topic] PRIMARY KEY  CLUSTERED 
	(
		[TopicID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_Group_bbs] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dv_Group_bbs] PRIMARY KEY  CLUSTERED 
	(
		[AnnounceID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_Log] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dv_Log] PRIMARY KEY  CLUSTERED 
	(
		[l_id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_Medal] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_MedalLog] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_MoneyLog] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dv_Plus_Tools_Log] PRIMARY KEY  CLUSTERED 
	(
		[Log_ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_Plus] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dv_Plus] PRIMARY KEY  CLUSTERED 
	(
		[ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_Plus_Tools_Buss] WITH NOCHECK ADD 
	CONSTRAINT [PK_Tools_Buss] PRIMARY KEY  CLUSTERED 
	(
		[ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_Plus_Tools_Info] WITH NOCHECK ADD 
	CONSTRAINT [PK_Plus_Tools_Info] PRIMARY KEY  CLUSTERED 
	(
		[ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_Plus_Tools_MagicFace] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dv_Plus_Tools_MagicFace] PRIMARY KEY  CLUSTERED 
	(
		[ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_Qcomic] WITH NOCHECK ADD 
	CONSTRAINT [PrimaryKey] PRIMARY KEY  CLUSTERED 
	(
		[ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_Setup] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dv_Setup] PRIMARY KEY  CLUSTERED 
	(
		[id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_SmallPaper] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dv_SmallPaper] PRIMARY KEY  CLUSTERED 
	(
		[S_ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_Space_skin] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dv_Space_skin] PRIMARY KEY  CLUSTERED 
	(
		[id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_Space_user] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dv_Space_user] PRIMARY KEY  CLUSTERED 
	(
		[id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_Stylehelp] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dv_Stylehelp] PRIMARY KEY  CLUSTERED 
	(
		[id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_TableList] WITH NOCHECK ADD 
	CONSTRAINT [PK_TableList] PRIMARY KEY  CLUSTERED 
	(
		[ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_Templates] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dv_Templates] PRIMARY KEY  CLUSTERED 
	(
		[Id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_Topic] WITH NOCHECK ADD 
	CONSTRAINT [PK_Topic] PRIMARY KEY  CLUSTERED 
	(
		[TopicID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_Upfile] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dv_Upfile] PRIMARY KEY  CLUSTERED 
	(
		[F_ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_User] WITH NOCHECK ADD 
	CONSTRAINT [PK_user] PRIMARY KEY  CLUSTERED 
	(
		[UserID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_UserGroups] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dv_UserGroups] PRIMARY KEY  CLUSTERED 
	(
		[UserGroupID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_Vote] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dv_Vote] PRIMARY KEY  CLUSTERED 
	(
		[voteid]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_VoteUser] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dv_VoteUser] PRIMARY KEY  CLUSTERED 
	(
		[ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_admin] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dv_admin] PRIMARY KEY  CLUSTERED 
	(
		[ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_help] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dv_help] PRIMARY KEY  CLUSTERED 
	(
		[H_ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[dv_album] WITH NOCHECK ADD 
	CONSTRAINT [PK_dv_album] PRIMARY KEY  CLUSTERED 
	(
		[id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[dv_albumbbs] WITH NOCHECK ADD 
	CONSTRAINT [PK_dv_albumbbs] PRIMARY KEY  CLUSTERED 
	(
		[id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[dv_albumimg] WITH NOCHECK ADD 
	CONSTRAINT [PK_dv_albumimg] PRIMARY KEY  CLUSTERED 
	(
		[id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[dv_albumtype] WITH NOCHECK ADD 
	CONSTRAINT [PK_dv_albumtype] PRIMARY KEY  CLUSTERED 
	(
		[id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[dv_bbs1] WITH NOCHECK ADD 
	CONSTRAINT [PK_bbs1] PRIMARY KEY  CLUSTERED 
	(
		[AnnounceID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_Appraise] ADD 
	CONSTRAINT [DF_Dv_Appraise_Boardid] DEFAULT (0) FOR [Boardid],
	CONSTRAINT [DF_Dv_Appraise_TopicID] DEFAULT (0) FOR [TopicID],
	CONSTRAINT [DF_Dv_Appraise_PostID] DEFAULT (0) FOR [PostID],
	CONSTRAINT [DF_Dv_Appraise_AType] DEFAULT (0) FOR [AType],
	CONSTRAINT [DF_Dv_Appraise_UserID] DEFAULT (0) FOR [UserID],
	CONSTRAINT [DF_Dv_Appraise_DateTime] DEFAULT (getdate()) FOR [DateTime]
GO

 CREATE  INDEX [PKlist] ON [dbo].[Dv_Appraise]([PostID], [AType]) ON [PRIMARY]
GO

 CREATE  INDEX [PKboard] ON [dbo].[Dv_Appraise]([Boardid]) ON [PRIMARY]
GO

 CREATE  INDEX [PKtopic] ON [dbo].[Dv_Appraise]([TopicID]) ON [PRIMARY]
GO

ALTER TABLE [dbo].[Dv_BbsLink] ADD 
	CONSTRAINT [DF_bbslink_islogo] DEFAULT (0) FOR [islogo]
GO

ALTER TABLE [dbo].[Dv_BbsNews] ADD 
	CONSTRAINT [DF_Dv_BbsNews_boardid] DEFAULT (0) FOR [boardid]
GO

 CREATE  INDEX [boardid] ON [dbo].[Dv_BbsNews]([boardid]) ON [PRIMARY]
GO

 CREATE  INDEX [PostUserID] ON [dbo].[Dv_BestTopic]([PostUserID]) ON [PRIMARY]
GO

 CREATE  INDEX [BoardID] ON [dbo].[Dv_BestTopic]([BoardID]) ON [PRIMARY]
GO

ALTER TABLE [dbo].[Dv_Bill_Online_Payment] ADD 
	CONSTRAINT [DF_Ht_Online_Payment_active] DEFAULT (0) FOR [active],
	CONSTRAINT [DF_Ht_Online_Payment_orders] DEFAULT (0) FOR [orders],
	CONSTRAINT [DF_Dv_Bill_Online_Payment_usermoney] DEFAULT (0) FOR [usermoney]
GO

ALTER TABLE [dbo].[Dv_Bill_Orders] ADD 
	CONSTRAINT [DF_Dv_Bill_Orders_uid] DEFAULT (0) FOR [uid],
	CONSTRAINT [DF_Dv_Bill_Orders_acount] DEFAULT (0) FOR [acount],
	CONSTRAINT [DF_Dv_Bill_Orders_active] DEFAULT (0) FOR [active]
GO

ALTER TABLE [dbo].[Dv_Board] ADD 
	CONSTRAINT [DF_board_ParentID] DEFAULT (0) FOR [ParentID],
	CONSTRAINT [DF_board_Depth] DEFAULT (0) FOR [Depth],
	CONSTRAINT [DF_board_RootID] DEFAULT (0) FOR [RootID],
	CONSTRAINT [DF_board_Child] DEFAULT (0) FOR [Child],
	CONSTRAINT [DF_board_CID] DEFAULT (0) FOR [CID],
	CONSTRAINT [PK_board] PRIMARY KEY  NONCLUSTERED 
	(
		[boardid]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dv_ChanOrders] ADD 
	CONSTRAINT [DF_DV_ChanOrders_O_isApply] DEFAULT (0) FOR [O_isApply],
	CONSTRAINT [DF_DV_ChanOrders_O_issuc] DEFAULT (0) FOR [O_issuc],
	CONSTRAINT [DF_DV_ChanOrders_O_BoardID] DEFAULT (0) FOR [O_BoardID],
	CONSTRAINT [DF_DV_ChanOrders_O_TopicID] DEFAULT (0) FOR [O_TopicID],
	CONSTRAINT [DF_DV_ChanOrders_O_AddTime] DEFAULT (getdate()) FOR [O_AddTime]
GO

ALTER TABLE [dbo].[Dv_Friend] ADD 
	CONSTRAINT [DF_Friend_F_addtime] DEFAULT (getdate()) FOR [F_addtime],
	CONSTRAINT [DF__Dv_FriEnd__F_Mod__1387E197] DEFAULT (0) FOR [F_Mod]
GO

ALTER TABLE [dbo].[Dv_GroupName] ADD 
	CONSTRAINT [DF__Dv_GroupN__Class__03317E3D] DEFAULT (0) FOR [ClassID],
	CONSTRAINT [DF_Dv_GroupName_AppUserID] DEFAULT (0) FOR [AppUserID],
	CONSTRAINT [DF_Dv_GroupName_UserNum] DEFAULT (0) FOR [UserNum],
	CONSTRAINT [DF_Dv_GroupName_Stats] DEFAULT (0) FOR [Stats],
	CONSTRAINT [DF_Dv_GroupName_PostNum] DEFAULT (0) FOR [PostNum],
	CONSTRAINT [DF_Dv_GroupName_TopicNum] DEFAULT (0) FOR [TopicNum],
	CONSTRAINT [DF_Dv_GroupName_TodayNum] DEFAULT (0) FOR [TodayNum],
	CONSTRAINT [DF_Dv_GroupName_YesterdayNum] DEFAULT (0) FOR [YesterdayNum],
	CONSTRAINT [DF_Dv_GroupName_LimitUser] DEFAULT (0) FOR [LimitUser],
	CONSTRAINT [DF_Dv_GroupName_AppDate] DEFAULT (getdate()) FOR [AppDate],
	CONSTRAINT [DF_Dv_GroupName_PassDate] DEFAULT (getdate()) FOR [PassDate],
	CONSTRAINT [DF_Dv_GroupName_visitDate] DEFAULT (getdate()) FOR [visitDate]
GO

ALTER TABLE [dbo].[Dv_GroupUser] ADD 
	CONSTRAINT [DF_Dv_GroupUser_GroupID] DEFAULT (0) FOR [GroupID],
	CONSTRAINT [DF_Dv_GroupUser_UserID] DEFAULT (0) FOR [UserID],
	CONSTRAINT [DF_Dv_GroupUser_IsLock] DEFAULT (0) FOR [IsLock]
GO

ALTER TABLE [dbo].[Dv_Group_Board] ADD 
	CONSTRAINT [DF_Dv_Group_Board_TopicNum] DEFAULT (0) FOR [TopicNum],
	CONSTRAINT [DF_Dv_Group_Board_TodayNum] DEFAULT (0) FOR [TodayNum],
	CONSTRAINT [DF_Dv_Group_Board_PostNum] DEFAULT (0) FOR [PostNum],
	CONSTRAINT [DF_Dv_Group_Board_RootID] DEFAULT (0) FOR [RootID],
	CONSTRAINT [DF_Dv_Group_Board_FoundDate] DEFAULT (getdate()) FOR [FoundDate],
	CONSTRAINT [DF_Dv_Group_Board_BoardStats] DEFAULT (0) FOR [BoardStats]
GO

ALTER TABLE [dbo].[Dv_Group_Class] ADD 
	CONSTRAINT [DF_Dv_Group_Class_GroupCount] DEFAULT (0) FOR [GroupCount],
	CONSTRAINT [DF_Dv_Group_Class_Orders] DEFAULT (0) FOR [Orders]
GO

ALTER TABLE [dbo].[Dv_Group_Topic] ADD 
	CONSTRAINT [DF_Dv_Group_Topic_GroupID] DEFAULT (0) FOR [GroupID],
	CONSTRAINT [DF_Dv_Group_Topic_BoardID] DEFAULT (0) FOR [BoardID],
	CONSTRAINT [DF_Dv_Group_Topic_PollID] DEFAULT (0) FOR [PollID],
	CONSTRAINT [DF_Dv_Group_Topic_LockTopic] DEFAULT (0) FOR [LockTopic],
	CONSTRAINT [DF_Dv_Group_Topic_Child] DEFAULT (0) FOR [Child],
	CONSTRAINT [DF_Dv_Group_Topic_PostUserID] DEFAULT (0) FOR [PostUserID],
	CONSTRAINT [DF_Dv_Group_Topic_DateAndTime] DEFAULT (getdate()) FOR [DateAndTime],
	CONSTRAINT [DF_Dv_Group_Topic_hits] DEFAULT (0) FOR [hits],
	CONSTRAINT [DF_Dv_Group_Topic_istop] DEFAULT (0) FOR [istop],
	CONSTRAINT [DF_Dv_Group_Topic_LastPostTime] DEFAULT (getdate()) FOR [LastPostTime],
	CONSTRAINT [DF_Dv_Group_Topic_isbest] DEFAULT (0) FOR [isbest],
	CONSTRAINT [DF_Dv_Group_Topic_Mode] DEFAULT (0) FOR [Mode],
	CONSTRAINT [DF_Dv_Group_Topic_TopicMode_1] DEFAULT (0) FOR [TopicMode]
GO

ALTER TABLE [dbo].[Dv_Group_bbs] ADD 
	CONSTRAINT [DF_Dv_Group_bbs_ParentID] DEFAULT (0) FOR [ParentID],
	CONSTRAINT [DF_Dv_Group_bbs_GroupID] DEFAULT (0) FOR [GroupID],
	CONSTRAINT [DF_Dv_Group_bbs_BoardID] DEFAULT (0) FOR [BoardID],
	CONSTRAINT [DF_Dv_Group_bbs_postuserid] DEFAULT (0) FOR [postuserid],
	CONSTRAINT [DF_Dv_Group_bbs_DateAndTime] DEFAULT (getdate()) FOR [DateAndTime],
	CONSTRAINT [DF_Dv_Group_bbs_length] DEFAULT (0) FOR [length],
	CONSTRAINT [DF_Dv_Group_bbs_RootID] DEFAULT (0) FOR [RootID],
	CONSTRAINT [DF_Dv_Group_bbs_layer] DEFAULT (0) FOR [layer],
	CONSTRAINT [DF_Dv_Group_bbs_orders] DEFAULT (0) FOR [orders],
	CONSTRAINT [DF_Dv_Group_bbs_isbest] DEFAULT (0) FOR [isbest],
	CONSTRAINT [DF_Dv_Group_bbs_locktopic] DEFAULT (0) FOR [locktopic],
	CONSTRAINT [DF_Dv_Group_bbs_signflag] DEFAULT (0) FOR [signflag],
	CONSTRAINT [DF_Dv_Group_bbs_isupload] DEFAULT (0) FOR [isupload]
GO

ALTER TABLE [dbo].[Dv_Log] ADD 
	CONSTRAINT [DF_log_l_addtime] DEFAULT (getdate()) FOR [l_addtime],
	CONSTRAINT [DF_log_l_type] DEFAULT (0) FOR [l_type]
GO

 CREATE  INDEX [l_boardid] ON [dbo].[Dv_Log]([l_boardid]) ON [PRIMARY]
GO

ALTER TABLE [dbo].[Dv_MedalLog] ADD 
	CONSTRAINT [DF__Dv_MedalL__AddTi__3EDC53F0] DEFAULT (getdate()) FOR [AddTime]
GO

ALTER TABLE [dbo].[Dv_Message] ADD 
	CONSTRAINT [DF_Dv_Message_flag] DEFAULT (0) FOR [flag],
	CONSTRAINT [DF_Dv_Message_sendtime] DEFAULT (getdate()) FOR [sendtime],
	CONSTRAINT [DF_message_delR] DEFAULT (0) FOR [delR],
	CONSTRAINT [DF_message_delS] DEFAULT (0) FOR [delS],
	CONSTRAINT [DF_message_isSend] DEFAULT (0) FOR [isSend],
	CONSTRAINT [PK_message] PRIMARY KEY  NONCLUSTERED 
	(
		[id]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

 CREATE  INDEX [IX_message] ON [dbo].[Dv_Message]([sender], [isSend], [delS]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_message_1] ON [dbo].[Dv_Message]([incept], [isSend], [delR], [flag]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [read] ON [dbo].[Dv_Message]([id], [sender], [incept]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

ALTER TABLE [dbo].[Dv_MoneyLog] ADD 
	CONSTRAINT [DF_Dv_Plus_Tools_Log_ToolsID] DEFAULT (0) FOR [ToolsID],
	CONSTRAINT [DF_Dv_Plus_Tools_Log_CountNum] DEFAULT (0) FOR [CountNum],
	CONSTRAINT [DF_Dv_Plus_Tools_Log_Log_Money] DEFAULT (0) FOR [Log_Money],
	CONSTRAINT [DF_Dv_Plus_Tools_Log_Log_Ticket] DEFAULT (0) FOR [Log_Ticket],
	CONSTRAINT [DF_Dv_Plus_Tools_Log_AddUserID] DEFAULT (0) FOR [AddUserID],
	CONSTRAINT [DF_Dv_Plus_Tools_Log_AddTime] DEFAULT (getdate()) FOR [Log_Time],
	CONSTRAINT [DF_Dv_Plus_Tools_Log_Log_Type] DEFAULT (0) FOR [Log_Type],
	CONSTRAINT [DF_Dv_Plus_Tools_Log_BoardID] DEFAULT (0) FOR [BoardID],
	CONSTRAINT [DF_Dv_Plus_Tools_Log_HMoney] DEFAULT (0) FOR [HMoney]
GO

 CREATE  INDEX [UserID] ON [dbo].[Dv_MoneyLog]([AddUserID]) ON [PRIMARY]
GO

 CREATE  INDEX [BoardID] ON [dbo].[Dv_MoneyLog]([BoardID]) ON [PRIMARY]
GO

 CREATE  INDEX [Log_Type] ON [dbo].[Dv_MoneyLog]([Log_Type]) ON [PRIMARY]
GO

 CREATE  INDEX [ToolsID] ON [dbo].[Dv_MoneyLog]([ToolsID]) ON [PRIMARY]
GO

 CREATE  INDEX [UserTools] ON [dbo].[Dv_MoneyLog]([ToolsID], [AddUserID]) ON [PRIMARY]
GO

ALTER TABLE [dbo].[Dv_Online] ADD 
	CONSTRAINT [DF_online_UserGroupID] DEFAULT (0) FOR [UserGroupID],
	CONSTRAINT [DF_online_userhidden] DEFAULT (2) FOR [userhidden],
	CONSTRAINT [DF_online_userid] DEFAULT (0) FOR [userid]
GO

 CREATE  INDEX [o_1] ON [dbo].[Dv_Online]([id]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [o_2] ON [dbo].[Dv_Online]([userid]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [o_3] ON [dbo].[Dv_Online]([userid], [userhidden]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [o_4] ON [dbo].[Dv_Online]([boardid], [userid]) ON [PRIMARY]
GO

ALTER TABLE [dbo].[Dv_Plus_Tools_Buss] ADD 
	CONSTRAINT [DF_Dv_Plus_Tools_Buss_ToolsCount] DEFAULT (0) FOR [ToolsCount],
	CONSTRAINT [DF_Dv_Plus_Tools_Buss_SaleCount] DEFAULT (0) FOR [SaleCount],
	CONSTRAINT [DF_Dv_Plus_Tools_Buss_UpdateTime] DEFAULT (getdate()) FOR [UpdateTime],
	CONSTRAINT [DF_Dv_Plus_Tools_Buss_SaleMoney] DEFAULT (0) FOR [SaleMoney],
	CONSTRAINT [DF_Dv_Plus_Tools_Buss_SaleTicks] DEFAULT (0) FOR [SaleTicket]
GO

 CREATE  INDEX [Tools_UserID] ON [dbo].[Dv_Plus_Tools_Buss]([UserID]) ON [PRIMARY]
GO

 CREATE  INDEX [Tools_ToolsID] ON [dbo].[Dv_Plus_Tools_Buss]([ToolsID]) ON [PRIMARY]
GO

 CREATE  INDEX [Buy_Tools] ON [dbo].[Dv_Plus_Tools_Buss]([UserID], [ToolsID]) ON [PRIMARY]
GO

ALTER TABLE [dbo].[Dv_Plus_Tools_Info] ADD 
	CONSTRAINT [DF_Plus_Tools_Info_IsStar] DEFAULT (1) FOR [IsStar],
	CONSTRAINT [DF_Plus_Tools_Info_Stock] DEFAULT (0) FOR [SysStock],
	CONSTRAINT [DF_Dv_Plus_Tools_Info_UserStock] DEFAULT (0) FOR [UserStock],
	CONSTRAINT [DF_Dv_Plus_Tools_Info_UserTicket] DEFAULT (0) FOR [UserTicket],
	CONSTRAINT [DF_Plus_Tools_Info_Price] DEFAULT (0) FOR [UserMoney],
	CONSTRAINT [DF_Dv_Plus_Tools_Info_UserPost] DEFAULT (0) FOR [UserPost],
	CONSTRAINT [DF_Dv_Plus_Tools_Info_UserWealth] DEFAULT (0) FOR [UserWealth],
	CONSTRAINT [DF_Dv_Plus_Tools_Info_UserEp] DEFAULT (0) FOR [UserEp],
	CONSTRAINT [DF_Dv_Plus_Tools_Info_UserCp] DEFAULT (0) FOR [UserCp],
	CONSTRAINT [DF_Dv_Plus_Tools_Info_BuyType] DEFAULT (0) FOR [BuyType]
GO

ALTER TABLE [dbo].[Dv_Plus_Tools_MagicFace] ADD 
	CONSTRAINT [DF_Dv_Plus_Tools_MagicFace_iMoney] DEFAULT (0) FOR [iMoney],
	CONSTRAINT [DF_Dv_Plus_Tools_MagicFace_iTicket] DEFAULT (0) FOR [iTicket],
	CONSTRAINT [DF_Dv_Plus_Tools_MagicFace_tMoney] DEFAULT (0) FOR [tMoney],
	CONSTRAINT [DF_Dv_Plus_Tools_MagicFace_tTicket] DEFAULT (0) FOR [tTicket]
GO

ALTER TABLE [dbo].[Dv_Setup] ADD 
	CONSTRAINT [DF_Dv_Setup_Forum_Setting_1] DEFAULT ('') FOR [Forum_Setting],
	CONSTRAINT [DF_Dv_Setup_Forum_ads_1] DEFAULT ('') FOR [Forum_ads],
	CONSTRAINT [DF_Dv_Setup_Forum_Badwords_1] DEFAULT ('') FOR [Forum_Badwords],
	CONSTRAINT [DF_Dv_Setup_Forum_rBadword_1] DEFAULT ('') FOR [Forum_rBadword],
	CONSTRAINT [DF_Dv_Setup_Forum_Maxonline_1] DEFAULT (0) FOR [Forum_Maxonline],
	CONSTRAINT [DF_Dv_Setup_Forum_MaxonlineDate_1] DEFAULT (getdate()) FOR [Forum_MaxonlineDate],
	CONSTRAINT [DF_Dv_Setup_Forum_TopicNum_1] DEFAULT (0) FOR [Forum_TopicNum],
	CONSTRAINT [DF_Dv_Setup_Forum_PostNum_1] DEFAULT (0) FOR [Forum_PostNum],
	CONSTRAINT [DF_Dv_Setup_Forum_TodayNum_1] DEFAULT (0) FOR [Forum_TodayNum],
	CONSTRAINT [DF_Dv_Setup_Forum_UserNum] DEFAULT (0) FOR [Forum_UserNum],
	CONSTRAINT [DF_Dv_Setup_Forum_YesTerdayNum_1] DEFAULT (0) FOR [Forum_YesTerdayNum],
	CONSTRAINT [DF_Dv_Setup_Forum_MaxPostNum_1] DEFAULT (0) FOR [Forum_MaxPostNum],
	CONSTRAINT [DF_Dv_Setup_Forum_MaxPostDate_1] DEFAULT (getdate()) FOR [Forum_MaxPostDate],
	CONSTRAINT [DF_Dv_Setup_Forum_lastUser_1] DEFAULT ('') FOR [Forum_lastUser],
	CONSTRAINT [DF_Dv_Setup_Forum_LastPost_1] DEFAULT ('') FOR [Forum_LastPost],
	CONSTRAINT [DF_Dv_Setup_Forum_BirthUser] DEFAULT ('2003-10-13 15:03:00|') FOR [Forum_BirthUser],
	CONSTRAINT [DF_Dv_Setup_Forum_Sid_1] DEFAULT (1) FOR [Forum_Sid],
	CONSTRAINT [DF_Dv_Setup_Forum_Version_1] DEFAULT ('') FOR [Forum_Version],
	CONSTRAINT [DF_Dv_Setup_Forum_NowUseBBS_1] DEFAULT ('dv_bbs1') FOR [Forum_NowUseBBS],
	CONSTRAINT [DF_Dv_Setup_Forum_IsInstall_1] DEFAULT (0) FOR [Forum_IsInstall],
	CONSTRAINT [DF_Dv_Setup_Forum_challengePassWord] DEFAULT ('') FOR [Forum_challengePassWord],
	CONSTRAINT [DF_Dv_Setup_Forum_Ad] DEFAULT ('') FOR [Forum_Ad],
	CONSTRAINT [DF_Dv_Setup_Forum_ChanName_1] DEFAULT ('') FOR [Forum_ChanName],
	CONSTRAINT [DF_Dv_Setup_Forum_ChanSetting_1] DEFAULT ('1,1,1,1,1,1,1,1,1,1,1,1,1') FOR [Forum_ChanSetting],
	CONSTRAINT [DF_Dv_Setup_Forum_LockIP] DEFAULT ('') FOR [Forum_LockIP],
	CONSTRAINT [DF_Dv_Setup_Forum_Cookiespath] DEFAULT ('/') FOR [Forum_Cookiespath],
	CONSTRAINT [DF_Dv_Setup_Forum_BoardS] DEFAULT ('') FOR [Forum_Boards],
	CONSTRAINT [DF_Dv_Setup_Forum_alltopnum] DEFAULT ('') FOR [Forum_alltopnum],
	CONSTRAINT [DF_Dv_Setup_Forum_Cid_1] DEFAULT (0) FOR [Forum_Cid],
	CONSTRAINT [DF_Dv_Setup_Forum_AdminFolder] DEFAULT ('Admin') FOR [Forum_AdminFolder],
	CONSTRAINT [DF_Dv_Setup_Forum_BoardXML] DEFAULT ('') FOR [Forum_BoardXML]
GO

ALTER TABLE [dbo].[Dv_SmallPaper] ADD 
	CONSTRAINT [DF_SmallPaper_S_Hits] DEFAULT (0) FOR [S_Hits],
	CONSTRAINT [DF_SmallPaper_S_Addtime] DEFAULT (getdate()) FOR [S_Addtime]
GO

 CREATE  INDEX [S_BoardID] ON [dbo].[Dv_SmallPaper]([S_BoardID]) ON [PRIMARY]
GO

ALTER TABLE [dbo].[Dv_Space_skin] ADD 
	CONSTRAINT [DF_Dv_Space_skin_s_userid] DEFAULT (0) FOR [s_userid],
	CONSTRAINT [DF_Dv_Space_skin_s_style] DEFAULT (0) FOR [s_style],
	CONSTRAINT [DF_Dv_Space_skin_s_lock] DEFAULT (0) FOR [s_lock],
	CONSTRAINT [DF_Dv_Space_skin_s_addtime] DEFAULT (getdate()) FOR [s_addtime]
GO

ALTER TABLE [dbo].[Dv_Space_user] ADD 
	CONSTRAINT [DF_Dv_Space_user_userid] DEFAULT (0) FOR [userid],
	CONSTRAINT [DF_Dv_Space_user_s_style] DEFAULT (0) FOR [s_style],
	CONSTRAINT [DF_Dv_Space_user_updatetime] DEFAULT (getdate()) FOR [updatetime],
	CONSTRAINT [DF_Dv_Space_user_lock] DEFAULT (0) FOR [lock]
GO

 CREATE  INDEX [Dv_Space_userid] ON [dbo].[Dv_Space_user]([userid]) ON [PRIMARY]
GO

ALTER TABLE [dbo].[Dv_TableList] ADD 
	CONSTRAINT [DF_TableList_TableName] DEFAULT ('') FOR [TableName],
	CONSTRAINT [DF_TableList_TableType] DEFAULT ('') FOR [TableType]
GO

ALTER TABLE [dbo].[Dv_Templates] ADD 
	CONSTRAINT [DF_Templates_Folder] DEFAULT (',') FOR [Folder]
GO

 CREATE  INDEX [ID] ON [dbo].[Dv_Templates]([Id]) ON [PRIMARY]
GO

ALTER TABLE [dbo].[Dv_Topic] ADD 
	CONSTRAINT [DF_Topic_PollID] DEFAULT (0) FOR [PollID],
	CONSTRAINT [DF_Topic_LockTopic] DEFAULT (0) FOR [LockTopic],
	CONSTRAINT [DF_Topic_Child] DEFAULT (0) FOR [Child],
	CONSTRAINT [DF_Topic_DateAndTime] DEFAULT (getdate()) FOR [DateAndTime],
	CONSTRAINT [DF_Topic_hits] DEFAULT (0) FOR [hits],
	CONSTRAINT [DF_Topic_VoteTotal] DEFAULT (0) FOR [VoteTotal],
	CONSTRAINT [DF_Topic_LastPostTime] DEFAULT (getdate()) FOR [LastPostTime],
	CONSTRAINT [DF_Topic_istop] DEFAULT (0) FOR [istop],
	CONSTRAINT [DF_Topic_isvote] DEFAULT (0) FOR [isvote],
	CONSTRAINT [DF_Topic_isbest] DEFAULT (0) FOR [isbest],
	CONSTRAINT [DF_Topic_IsSmsTopic] DEFAULT (0) FOR [IsSmsTopic],
	CONSTRAINT [DF_Topic_LastSmsTime] DEFAULT (getdate()) FOR [LastSmsTime],
	CONSTRAINT [DF_Topic_TopicMode] DEFAULT (0) FOR [TopicMode],
	CONSTRAINT [DF_Dv_Topic_Mode] DEFAULT (0) FOR [Mode],
	CONSTRAINT [DF_Dv_Topic_GetMoney] DEFAULT (0) FOR [GetMoney],
	CONSTRAINT [DF_Dv_Topic_GetMoneyType] DEFAULT (0) FOR [GetMoneyType],
	CONSTRAINT [DF_Dv_Topic_HideName] DEFAULT (0) FOR [HideName]
GO

 CREATE  INDEX [list_1] ON [dbo].[Dv_Topic]([Boardid], [istop], [LastPostTime]) ON [PRIMARY]
GO

 CREATE  INDEX [list_2] ON [dbo].[Dv_Topic]([Boardid], [istop], [Mode], [LastPostTime]) ON [PRIMARY]
GO

 CREATE  INDEX [topicwithme] ON [dbo].[Dv_Topic]([PostUserid], [Child]) ON [PRIMARY]
GO

 CREATE  INDEX [SearchUser] ON [dbo].[Dv_Topic]([PostUserid]) ON [PRIMARY]
GO

 CREATE  INDEX [SearchTitle] ON [dbo].[Dv_Topic]([Title]) ON [PRIMARY]
GO

 CREATE  INDEX [GetMoneyType_List] ON [dbo].[Dv_Topic]([Boardid], [GetMoneyType]) ON [PRIMARY]
GO

 CREATE  INDEX [Dv_Topic_hits] ON [dbo].[Dv_Topic]([hits]) ON [PRIMARY]
GO

ALTER TABLE [dbo].[Dv_Upfile] ADD 
	CONSTRAINT [DF_Dv_Upfile_F_BoardID] DEFAULT (0) FOR [F_BoardID],
	CONSTRAINT [DF_DV_Upfile_F_Type] DEFAULT (0) FOR [F_Type],
	CONSTRAINT [DF_Dv_Upfile_F_FileSize] DEFAULT (0) FOR [F_FileSize],
	CONSTRAINT [DF_DV_Upfile_F_DownNum] DEFAULT (0) FOR [F_DownNum],
	CONSTRAINT [DF_DV_Upfile_F_ViewNum] DEFAULT (0) FOR [F_ViewNum],
	CONSTRAINT [DF_DV_Upfile_F_Flag] DEFAULT (0) FOR [F_Flag],
	CONSTRAINT [DF_DV_Upfile_F_AddTime] DEFAULT (getdate()) FOR [F_AddTime],
	CONSTRAINT [DF_Dv_Upfile_f_topicid] DEFAULT (0) FOR [f_topicid],
	CONSTRAINT [DF_Dv_Upfile_f_bbsid] DEFAULT (0) FOR [f_bbsid]
GO

 CREATE  INDEX [F_BoardID] ON [dbo].[Dv_Upfile]([F_BoardID]) ON [PRIMARY]
GO

 CREATE  INDEX [F_UserID] ON [dbo].[Dv_Upfile]([F_UserID]) ON [PRIMARY]
GO

 CREATE  INDEX [F_Username] ON [dbo].[Dv_Upfile]([F_Username]) ON [PRIMARY]
GO

 CREATE  INDEX [pk_Dv_Upfile_annid] ON [dbo].[Dv_Upfile]([F_AnnounceID]) ON [PRIMARY]
GO

ALTER TABLE [dbo].[Dv_User] ADD 
	CONSTRAINT [DF_Dv_User_UserPost] DEFAULT (0) FOR [UserPost],
	CONSTRAINT [DF_Dv_User_UserTopic] DEFAULT (0) FOR [UserTopic],
	CONSTRAINT [DF_Dv_User_UserWidth] DEFAULT (0) FOR [UserWidth],
	CONSTRAINT [DF_Dv_User_UserHeight] DEFAULT (0) FOR [UserHeight],
	CONSTRAINT [DF_Dv_User_UserLogins] DEFAULT (0) FOR [UserLogins],
	CONSTRAINT [DF_Dv_User_UserViews] DEFAULT (0) FOR [UserViews],
	CONSTRAINT [DF_Dv_User_UserMoney] DEFAULT (0) FOR [UserMoney],
	CONSTRAINT [DF_Dv_User_UserTicket] DEFAULT (0) FOR [UserTicket],
	CONSTRAINT [DF_Dv_User_userWealth] DEFAULT (0) FOR [userWealth],
	CONSTRAINT [DF_Dv_User_userEP] DEFAULT (0) FOR [userEP],
	CONSTRAINT [DF_Dv_User_userCP] DEFAULT (0) FOR [userCP],
	CONSTRAINT [DF_user_UserPower] DEFAULT (0) FOR [UserPower],
	CONSTRAINT [DF_user_UserDel] DEFAULT (0) FOR [UserDel],
	CONSTRAINT [DF_user_UserIsBest] DEFAULT (0) FOR [UserIsBest],
	CONSTRAINT [DF_user_UserGroupID] DEFAULT (0) FOR [UserGroupID],
	CONSTRAINT [DF_user_UserHidden] DEFAULT (0) FOR [UserHidden],
	CONSTRAINT [DF_user_IsChallenge] DEFAULT (0) FOR [IsChallenge],
	CONSTRAINT [DF_user_UserIsAva] DEFAULT (0) FOR [UserIsAva],
	CONSTRAINT [DF__Dv_User__RLActTi__6225902D] DEFAULT (0) FOR [RLActTimeT],
	CONSTRAINT [DF_Dv_User_UserIsAudit_Custom] DEFAULT (2) FOR [UserIsAudit_Custom],
	CONSTRAINT [DF_Dv_User_UserIsAudit] DEFAULT (0) FOR [UserIsAudit],
	CONSTRAINT [DF_Dv_User_TyMedaled] DEFAULT (0) FOR [TyMedaled]
GO

 CREATE  INDEX [toplist_1] ON [dbo].[Dv_User]([UserPost] DESC ) ON [PRIMARY]
GO

 CREATE  INDEX [IX_user] ON [dbo].[Dv_User]([UserName]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [userinfo] ON [dbo].[Dv_User]([UserGroupID], [UserID]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [userann] ON [dbo].[Dv_User]([UserID], [UserClass]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [toplist_2] ON [dbo].[Dv_User]([JoinDate] DESC ) ON [PRIMARY]
GO

 CREATE  INDEX [toplist_3] ON [dbo].[Dv_User]([userWealth] DESC ) ON [PRIMARY]
GO

 CREATE  INDEX [chan_1] ON [dbo].[Dv_User]([UserMobile], [IsChallenge]) ON [PRIMARY]
GO

 CREATE  INDEX [UserEmail] ON [dbo].[Dv_User]([UserEmail]) ON [PRIMARY]
GO

 CREATE  INDEX [Dv_UserMoney] ON [dbo].[Dv_User]([UserMoney]) ON [PRIMARY]
GO

 CREATE  INDEX [Dv_UserTicket] ON [dbo].[Dv_User]([UserTicket]) ON [PRIMARY]
GO

 CREATE  INDEX [Dv_UserGroupID] ON [dbo].[Dv_User]([UserGroupID]) ON [PRIMARY]
GO

 CREATE  INDEX [Passport] ON [dbo].[Dv_User]([Passport]) ON [PRIMARY]
GO

 CREATE  INDEX [UserPer] ON [dbo].[Dv_UserAccess]([uc_BoardID], [uc_UserID]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

ALTER TABLE [dbo].[Dv_UserGroups] ADD 
	CONSTRAINT [DF_Dv_UserGroups1_MinArticle] DEFAULT (0) FOR [MinArticle],
	CONSTRAINT [DF_Dv_UserGroups1_ParentGID] DEFAULT (0) FOR [ParentGID],
	CONSTRAINT [DF_Dv_UserGroups_UserGroupIsAudit] DEFAULT (0) FOR [UserGroupIsAudit]
GO

ALTER TABLE [dbo].[Dv_Vote] ADD 
	CONSTRAINT [DF_vote_votetype] DEFAULT (0) FOR [votetype],
	CONSTRAINT [DF_vote_LockVote] DEFAULT (0) FOR [LockVote],
	CONSTRAINT [DF_vote_voters] DEFAULT (0) FOR [voters],
	CONSTRAINT [DF_vote_TimeOut] DEFAULT (getdate()) FOR [TimeOut],
	CONSTRAINT [DF_Dv_Vote_UArticle] DEFAULT (0) FOR [UArticle],
	CONSTRAINT [DF_Dv_Vote_UWealth] DEFAULT (0) FOR [UWealth],
	CONSTRAINT [DF_Dv_Vote_UEP] DEFAULT (0) FOR [UEP],
	CONSTRAINT [DF_Dv_Vote_UCP] DEFAULT (0) FOR [UCP],
	CONSTRAINT [DF_Dv_Vote_UPower] DEFAULT (0) FOR [UPower]
GO

ALTER TABLE [dbo].[Dv_VoteUser] ADD 
	CONSTRAINT [DF_voteuser_VoteDate] DEFAULT (getdate()) FOR [VoteDate]
GO

ALTER TABLE [dbo].[Dv_admin] ADD 
	CONSTRAINT [DF_admin_LastLogin] DEFAULT (getdate()) FOR [LastLogin]
GO

 CREATE  INDEX [CHKLOGIN] ON [dbo].[Dv_admin]([Username], [Adduser]) ON [PRIMARY]
GO

ALTER TABLE [dbo].[Dv_banzhu_config] ADD 
	CONSTRAINT [DF_banzhu_config_Gzdays] DEFAULT (7) FOR [Gzdays],
	CONSTRAINT [DF_banzhu_config_gzifmess] DEFAULT (1) FOR [gzifmess]
GO

ALTER TABLE [dbo].[Dv_banzhu_log] ADD 
	CONSTRAINT [DF_banzhu_log_LeaveDays] DEFAULT (0) FOR [LeaveDays],
	CONSTRAINT [DF_banzhu_log_CancelLeave] DEFAULT (2100 - 1 - 1) FOR [CancelLeave]
GO

ALTER TABLE [dbo].[Dv_banzhu_user] ADD 
	CONSTRAINT [DF_banzhu_user_oldlogs] DEFAULT (0) FOR [oldlogs],
	CONSTRAINT [DF_banzhu_user_oldlogins] DEFAULT (0) FOR [oldlogins],
	CONSTRAINT [DF_banzhu_user_lastinfo] DEFAULT (0) FOR [lastinfo],
	CONSTRAINT [DF_banzhu_user_oldarticle] DEFAULT (0) FOR [oldarticle],
	CONSTRAINT [DF_banzhu_user_IsLeave] DEFAULT (0) FOR [IsLeave],
	CONSTRAINT [DF_banzhu_user_Leavedays] DEFAULT (0) FOR [Leavedays],
	CONSTRAINT [DF_banzhu_user_IsSendMail] DEFAULT (0) FOR [IsSendMail]
GO

ALTER TABLE [dbo].[Dv_help] ADD 
	CONSTRAINT [DF_dv_help_H_type] DEFAULT (0) FOR [H_type],
	CONSTRAINT [DF_dv_help_H_orders] DEFAULT (0) FOR [H_stype],
	CONSTRAINT [DF_dv_help_H_Addtime] DEFAULT (getdate()) FOR [H_Addtime]
GO

 CREATE  INDEX [IX_Dv_help] ON [dbo].[Dv_help]([H_ParentID]) ON [PRIMARY]
GO

ALTER TABLE [dbo].[dv_bbs1] ADD 
	CONSTRAINT [DF_bbs1_ParentID] DEFAULT (0) FOR [ParentID],
	CONSTRAINT [DF_bbs1_locktopic] DEFAULT (0) FOR [locktopic],
	CONSTRAINT [DF_bbs1_IsAudit] DEFAULT (0) FOR [IsAudit],
	CONSTRAINT [DF_bbs1_IsUpload] DEFAULT (0) FOR [IsUpload],
	CONSTRAINT [DF_dv_bbs1_GetMoney] DEFAULT (0) FOR [GetMoney],
	CONSTRAINT [DF_dv_bbs1_GetMoneyType] DEFAULT (0) FOR [GetMoneyType],
	CONSTRAINT [DF_dv_bbs1_FlashId] DEFAULT ('0') FOR [FlashId]
GO

 CREATE  INDEX [dispbbs] ON [dbo].[dv_bbs1]([BoardID], [RootID]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [save_1] ON [dbo].[dv_bbs1]([RootID], [orders]) ON [PRIMARY]
GO

 CREATE  INDEX [disp] ON [dbo].[dv_bbs1]([BoardID]) ON [PRIMARY]
GO

 CREATE  INDEX [PostUserID] ON [dbo].[dv_bbs1]([PostUserID]) ON [PRIMARY]
GO



if exists (select * from sysobjects where id = object_id(N'[dv_Dispbbs]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dv_Dispbbs]
GO

if exists (select * from sysobjects where id = object_id(N'[dv_list]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dv_list]
GO

if exists (select * from sysobjects where id = object_id(N'[Dv_loadSetup]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [Dv_loadSetup]
GO

if exists (select * from sysobjects where id = object_id(N'[dv_toplist]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dv_toplist]
GO

if exists (select * from sysobjects where id = object_id(N'[Dv_TSQL]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [Dv_TSQL]
GO

SET QUOTED_IDENTIFIER  ON    SET ANSI_NULLS  OFF 
GO

exec(decrypt(0x05DFE6B96BC9EFAE8D019E2D95B4C88B17C6F821D5ABF69B9432816D5122D96284A4804A15BFD4B5558E0DCAC71A7A2A4D8EA99E1C85BEC12D43F367FDEF43271282487E2918173CCDDFB07A55015629FE59C51C1A964D4C3C408AD7A1A0AAF5DBEDB33162D06B05B632D22EF82071F6FE8402717152B409B9E284A4FA06E8A815F21700A8766701C6CEA661BD7B4C5B60A6B8094639D3A029BA7DD4085372BA5D70CB07B9066D9AC0C9453016FB2F4BECAFFB9ADA702860B74C27B421BD02DFF0A85A90C04A2F7A990614010C5714EAF2889C894791D2DCFEA16EE8A43C7027A9538500B55270B6E803A0DD393AE322960ABFBACEBD20CFE8B96132C988AF1D4D76C8A7EB975580FA3CFE21C1251D1ACAC287754E7DA10F7DDF698B3808827832CD7927C4F471D8AF033A3006FFB9AF7E654B91863136BB916C24DDC311C96A67DDBE6346A8417BC2DDF203EE573FDE28ECD04FBD2EFA33A9FC65F40AF6270EBB9C3E6AC02501A8E192E442DC7898093DFC304CB20840E5C48E18249E8C5B968D62A57EB61D86D5A0790813888D7E86DF00B79B8BBC671E6301415C3B06EDF7F48963A79125099E7AA214A1E80E34CC2A46E895B114B15C2E35F03855F7BA8E803461EBB14990E67CD7CAEA2337D69A0C3F15D6CDED76478D0A0A643D39A27E469F08FA58BF6103F3DB71B57DDC14DF38CC4F3B425C9AF46764BCEC6819D46453685D93F47AC31CA6C70969C1FA411148F9B96F09392A0B13C52D0D60A4BBDAE77EA35C2B66711ECC8BBAC76D0A41B535FB8E5C7EC4613AD9CF08DDE752C840F07C9CA9385EAF2D0E36B326C533D1CE9DD61AE286444C4D83176475DF00613B0750BF2887765707D5F11612843D399BDA90EDD6A2D52AEBBD67B980
))
GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  ON    SET ANSI_NULLS  OFF 
GO

exec(decrypt(0x05DFE6B96BC9EFAE8D019E2D95B4C88B17C6F821D5ABF69B9432816D5122D96284A4804A15BFD4B5558E0DCAEFDA547DA9784BF3C66735E1C44946C3C24E4B2EFDCFDC5065D4DA687C3A1A653AA6E205706316990FEF8729B748B24295665F76BAF48FC8D80D0D8F3EC53B23842B54F6566E8E23F0C6BE27CABC3A8F0C30271E583AFE023D7D94264F708D34791EC920FB995B1BACCF6DB81AC6FAD12D92CEAA8C853B0E25631FBFED4281F4E3D36D62BF86213E07C1D0BF98A3F93A8676F6D52C82EA27082104C376064EB4C89F4ED8FA3D4ABB6C6148614BF938630F0A7CA3437877FA02806921F28E4AA20D0CDEC383387581CBBC72C4A84ACF42E880FA1A7923CE998568960779464BA2C82E025EE4F8BE76B08ED5D256DCC0924CD4E8A7C33FC3528D297C14779EDDF81CF435BB2D64A6A7B7365B1B47646F96010E95906CAC09C43E49ADD70F9C7B21EEED611064D44A2BDE3C1885776643CBA917338D5BA1A1A44AD27FEE3AC289EED056B088366ACD563DF93FEC3FCE8C0B3729A24F0C2B3C703594D6362913AC9C46C2B9137BF102D6040D92520C9159E6EFEEC5A6C3255048FD93AAC2D2982D5F61E3DFFACE0F288E9AECE1EB3D8CEB864C1C8003FC937773E2935C7CF771C949BA4AAD78887AF574928D6725BB939CFF813A081ADEF690F0098048A9EA45A5302235D87B773A118F92DCD2CD9C1CB4840509CB55F44EF4026E92F70C1FA5A81415FBCBD12B5AC248B948F8497065CADCC34B18EDC0E3BAF7574E1C68328DFF2FC5BAB9F25305023A0831BFAE1899216DEE20BED7EE09B0844407C6A422070FD4C8857F43C9059F268B9377E0C5D1221E0381C5CFFC04D47DDAC9F4ED88718BC17C977F4071F58C5C6B69F6BD51D29D780D789D226C8580CF29157E811291CBD3998BD560184AD628B02BB106F7810D10CA6D5E67E308F54F0D9CAA7BF68BA9B09186FAE278B392A7CBF7CC0550C7CE53E5D522F5D333E0E7657F28AFA3C2C71A11A15BC699ABBE9D489934C8B689BF12A7948F41F9178612E4C59FBB002D5F8D7ABCED8F11A97EF15C0CB95274F7B79CCF8E6B9B2841D14C1A8F49E4CCD3B6FB63CFEEF082D3A1A1FD1AC59DBD77968FEE734585DC219FE6C6BE321AD8CBAE126D95B32E7F6B4D38BE4C4F21BB1C4134FD0CA58894706B51DA30A0EACFBDDA4919F604DDDCA170D1372A8518D13BF9F8136548455525FBF1462D0972FD7ED49000D8654952281459B87DC066D9E2B1184CC9101B472B54AFB199E695E3595ED9CB5E1D3B25EF9530BF65B602AB4D554777D9F55AAC64DF8CB5CB64ACE73A28B762A1177F4DEFE04AEA01981F9AD7B80CC766A4A1F3E6A4EEE62B965FCB49B37601D700DE552569A6D8252F924235F79F85F47C527318B4CB8E2043CC294B3894BAFD804FD04D55D83CE60252C241BC3581A56FC07ACE9746F4EB731064EAD35179BCA3808220FB4C87FCD0B4B74E035E9F75119A2E2643A55FD0F636558672C8218183A9DFC3E67E5F4B3F18B2D5E3CE66556A339702D3F2CD93403AAE1C01DFE4ADCA9B2F0499FB19EC04643BAD71F235B7633A7EC7646220FFA289F968CE923B6688D2583D8E9B6CDE59D48F45D5363E58EEC1B488C962EA0BD7B7016BF4E27E8CA48F6F3941A6AFEA395A9A4030277669520074FDD58C55F0A58E51C3E69FBAF8EA29C0F75CD1DE30E8158F6CCF16C557F08B32DB2172CAB631A5D38B768A466D7E6C8FF8B78A74F979C67746CC5D9E8D90C72CF0C3A95C785339E27362AAEB3E08C8EE6301E4A068F2BEE851180F94E389896E28108433954B9E99DE3E80210A57F02B7C930B716760257044258BC3BF1BF8AEB9ACB0BBF088C4F0F6E497F459196F3CE0302FCE00E6A14768D37F96A8073337BB0C5A6E6EEF5A3E190C1CC2DA52FB735080C74276451F0F8D2158A78B3BB5966850AF23477FB438FB8F0D5DBAC5A9F4305AE29969D979745CD85AFC3ED258D709135735DE68221F8E8257500F3675CA2EAE7587030512C6FA339FF395707DAB2039F6F57EB36A1CDC473984E0838C5BFD31DB7C1134E4B3C20D2DDEDB4BB9C5552E1578182EEFAB0C064DD0D0EE3C17052F869196F0818CF47C25DD66EAF6ED3D657877D69ECD6DE8901BD81ABD862F864ECE7A67EC6A8C7CD815D6EE0A2D39B755BB611306B690A99CB5702FE5DC83B29EA3CDFB8305DBB39911E19F271E74579BB9BD910F8EC11AE46CA54E2930140F838F1953C610663F104DF5D18B4398CC286065227107C311E0D16A795B512698E52E3DD02601E5646C44AF4C12AA90CB04B04E8AA387E20B4587F2F13D264B456CD84893DF1D55AB1E912DF6080C52A6C9169FA098B88A0C382B2EB9CFA7C3FB4B7DBD6215B50246BC0827FC3C8D93DFEB936D3212D00B2D112A791CFC3C172807DF5961CDBA7943444BE65EA13CC60B5C7711B54D13A598528465DEE6645BC55A4DF2B592021BD31766B0CE6D8B560C868F50BA972E6C221567BD80CF2B2AF47B4FD7DD9A43258845B78D337407C8BA9AD7E38174952EEA570F2EFDDB475EC3168C6047514BFEDF78530EF1C99E12E02E7DA052AEBAB593A506663BDFCC14B148A78C72D414DBA7B48ADA6BF19015E7A63C88338E19E692D220265F83124BF7F1193EDAC54EEF3067066E644BF549C5C1F07570A2AB4709A834E08024E13B8BCBB283CE643312A989A5501825767783BEF8ED421EC843CC7BD201906E2D35EB4CCE9253886ECCA84963E603FFA33DA0BEFBF1925F7A13F23E20359B01E0E4FB1A5EC68F109DA745E53F5FB89FC01971F142F6C71028F33DB13FE4E1E0DA3E0C4BF435EF32BAA966B10A46B8E22520169C8DF58E1C0398A062C44CEAB20897E85255B0872D32B51B4DFE1C5A8C084BC20DA1B7E3FA0351E318B7BFCBBA1EA0EC27D17149126E75F273B1BFDF96C2A8BD5B3F6DCB00DAF27CB37F6F9AB0AC52696F13F8A3CEE0715D29F36BA110E2EEADFE6F7290D67B299AC1926AA8941CB4A50480629BF40389BB0B4B7CAC2F7D0E34A67F55EEDBFE4C9A0333E91918DB9DAA452F6C0CCD082963D6FF69E3F54A5B79930D82D5EB4D22F5DA2C3B35A2EE18088325682517D9148FF20A9B50E6ED4399FD56303F1EAFAC87156A6598A0A94158A69EC4BD96AF3C63DA8459B01D509F45BC8BCA5A0563A3EBA68E56F1CE97603606A72C18CD72C1D3932504370124C089E583276C2E1553152C939D9A077A7857C8F55CC6CD729622EF2555AF43FE9D75EA4F756BCC2C8343F3DEF58BD7CCC9B2EE357B856A004CDCE3EB983DE4724769D8EAECAE0DD74F5250DB314AFADC08A022CE624482C8FABF0D3CCBDD47BFAA9F185F7C0736A561638905CFCC36840F6B19137AC603AD787D9B1586E4D9B453F28A3FCBD12F14E7EA7DF70F2032A21AC6A1D4BB950E31525374C1754B969210331264F5DA0CCF43208001D619B62336B63AB1CE53BF91E67AA5155D599D4E8572087CA23B47488D48D90063C8C27812E4E97A8E4C6786177E5852CCC96344977AC8FCF9E34EE1A2FBBEF03DDA9A96DBA4B763405399A831EED5C36E30F1EB225885C9CDCF581471E403887BFC0DE4A648AA6A18CC38185C1FB932007778679247E3B8BDEB4C0BC2D2AEE99395CDCE96F5350AF385BCCF8139752C2ACFA910866F2B1C3A97C82D2F02FADD069C0E15C80CBD1503A37B4F924E7A2F7C50F3E05A2D76E4515240DC8533BDFF6538FA5CCADA3F5485303C7B7D13283683E182EF5CE494A71FC5A697768E1B77758600991D4EF771FD07C0AD5549B72CF1F82A3D44C348589FD6189DC8C61C81648855217836AC287DF4D07143E24F0A6A1BA1EC307FC38162503D4F773CB8742295263575F2534C9397A6C89C54E468D179B5AB28823E01F928836A0827E916BE2238883EE7ED9DFCC517D7A145D2397202D57E9652AB1907BF13EBA71563D3BA1B84D39888F0216C38DC5360F808E4816013F618E03096F1317232D8335570278458F077C0864890C02F2281812F1754500422D3F45AC746653FD2D024AC7A191BB0ED86FBDC9CB3BAF029D6154CD3A2724F273B15A8349C7C84DD42F1BEE14F473295D953D948A7C29FFCADFB78D1785AE61BCE76A3C171D01115D36E055E353968F85446CC1412A0371C689979EB9A80A6CD880AB52C6A2E348CA6CC0565DC085AE142A7D7919389949074284F0522201FF4C9B271BAE384E64E3CEA0E3B938026ECF2DCA9C4AC33F17EB135F58A5790F6A27133988DC46108DB3A883BDDFA184C20CF8C9C3A4DC9A7DF64F8EED3AAB7E51348A6A34339749AA99E10C82435756C946CB03AF2E41B625ADD113B579C35EB01E5962F1327C4DF94E5B35BF58A33EB0A2C206D3F415A70DE1B9AB608D32D414221BCA8A963FA26A7ECF96A37C2029AD829A59B092ECD17165399FAA60C0B8E0CDD0403C5D0B7F1D126713D2B3DA5D31A9EDBF065F224CB3098F8598F9D8096E079A0D2347CD40A22AB9BA2EF364BC9BE30B1AA7F55793689D2D1A2AC282B088B83FE92A2F606689816090263DF3AF35B1F5C6A406580A40CD2223468DE2B0A22E243AB2395D1F40DF27950FD8B7527FF548DFE5D5FFAC4FA65B31DA73793663E30E9B1AE2FAB5A4CCD1DEFA49CA8A5D431BE4401E8F4542408B3821A9A8BBDD92B61694D0941CCE96A8DA83DD6B017EB9411AA686AC7CAA49CB6FDCBEBD31A3210E1076E9F97CD1F28BF0DC2C956E7DEDFEBC09DCA0447B766169A8EE6D5B28F0611D90709B51ED61DCBB3F443A5FC7CEB2F41D791A1D28F426C42BD52DAEF0AB062AA0D8A0F934A39EA50E88AC071E7A2E5DCA9F823559E02E06545BD958516017F407C478AA3D93A225CCD6175AE41CD7CC8BCD6E0AEB2CCAEE25217CD7EC93C58970024863D8C4941A8ACCE01903E78E79FA8F812EFB68D02C1D94364A1B8404415C912491492C42E4C5BA76E526F8B45F6855490A0548D0EE78C3F8D814F0B112F50C459BD5F7D6CEA86C8A0052DB30AE8CDCD202C43EF4A2275A0D9A893058C945742607165A1BB070BA8E0A3A5A1EFC8E21EFC1FDBC2EEE449B7FD9275F37C1420ACCF86FA414C47CE3C89DEEB5A3EA8123776D949DFE714EEB19C7AF45E2697989D0156ABEAFE09515D061A3788237270446EC0EF1F41EC70DB0F35FC030C13DD954557EEF2865641C7A4F4CBCFAAA5FCAA17106536F49E31FEF78741B9873A36EB2BCEF34136924611F91EF3F40956B5958A3705A774A07BCB9A2462F4A9468024E96F7C4F756572E73EFF52BCE5B4C072B392A8A1E25D659C408E4CE74E51A8399587E4BDB5AFC3A452ABEC94137DB0732C4FCB31974BAA48FCFC9B198356300F29370C3CDBF775FD842940D26A81B6F451C9E9781C867540A118B74562628E1238FE91FB36C8648BFCCBC8077191E19AD77CA8EA1F774F138A6856BE4B6521A3EFA93737060E5C184C6B2A999BBFD36088D802FFC28CC149DA448AC7E44A5A4EB461609884F19F6C0A4F1B560923930921D032D91727BD4A1CFC67B0C79B84931DC1079789E158BEDC80068EC2CB6459DF882572A8678205D594018A94364D05D035A57E4A3FECC34EC25749983EB9B3F90A2DA8AB1A3259BD7A211B34E3A09D37FFEBC3E8A35316F6C28863B09F4DE73CDD864F5D09F1DCA49B4FF2E173CCB4F06A8A85CBDEBDD70CE4D3845D4A4CFE3F631187E2F8ADB238657293FD1C828EC42C180EF7522BEFA6ED4EDCE44B0CA0F9E187D2E5DE3282305C450AFF4D5CAD46FB75B3E41F9FB675E879EFD4923EEFB706911E8ADC054C98F9D9472B1DD56B58406B6EE41056C104AA79E4965D87E2CE9E9680C97C9A876B7F73A7451C09446FA61E5808E7C40A6C9D93D4E8BC65270C99DCE95D1E313EEC8C81BD171A58CD61990BB61CFA81C8F3B9A4F040B94DFE4BF7A8AF80CBA95421F2B9428B42D735BECB7DBCBBE19D63C82258EDA0BF7C4F04D7A2DAACC2C30435459012ECBF5F64A362ED25F33FB683DA38A2F290E89106663F5B31FFFDFFDB9BA98EC30CD93E676A42127F6024A2961AE1F56844FBE822E0306750A4526F08E3C91B75B563AA9ED89B55CD9CBFB7F12C0FF387C6FD3950E3AE1D3F3883FFC367DDBB4D974550E17C1F75872713469C5C1D442C4273D6A2AB086550F410F706DC9E376F2DDA9BCA704A6D445586D3411458542287A49C1FA3FA4D926F1610A9909FC690CF8BA10012ADB3C7C02D7698BC47FBFD4AF5ED4B5EFB3AA43CFD3260031B76D79477F81590B47DBC6062EC1A813EB9607B433C598686CE95625EA7741FCFC061A0113E627DCCC0318556C3A3775C9FF141518CBD410042279AF1A843B8A109632DD76828CEC4FB4E6DD78CB3A698E6822E7614A088AA3BDC9EE94328AE3AEE34F2F9EF268ED57B005B4FEE5DF7C8C297242EF5
))
GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  OFF 
GO

exec(decrypt(0x05DFAFE2EA01570A7B970F9F700ACD20B377A0B529B3AD3938A518FEF021372F885733C6BB8B52CF48677E8F3E8C8F2A7BE448C3819757063DCCE839A432E62517004DB3C3DE973856F6E7F45A5E0ABDCA627ACB94CDAC556EC5B3C05BD6106D6E6B13C7F587D0A1B8C87602C9F3B1CDD32EA161FE10D1304AAE65A1DD055D8C220262CE5F32241243C362D78C5C5123D8AFFD77CA21632FFBAA14F46497E756B0B7481A699136228633C531AC00367E6144
))
GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  ON    SET ANSI_NULLS  OFF 
GO

exec(decrypt(0x05DFAFE2EA011E81392C623662D4FCCD49E839B87B43E9CC5E66DD4EBFF5EE7C4D47669F6B3749D59BB74B2AF0CBD2D92B34954A757A235B4CEBEF09C595BAC20B31DAE27DAF103BBF72C082799316E7CC93FFDBDA6EF0DC8D1457DCA6FD1DA152B82B6BEE9351D7152180FB5649BAF6607ADEF8C641E305BEF6F7A4B28B0A8FBC8664B2D5719C40FC9689C2A06106B2EE6A6FECA9B86BAEDEFF4AF4A6860BB8B8EADD60BD2A2CF4A5C34E344C673F545E5630B574F96F825BA8FE4E273FBD1A43206C56A22C770B31E230434274648A77C9DA2832578F89804CBAFA8EAE5E5FB8DEE554AD12F129DE244F75A3275419E4D1B85EDD28E908BBC7B3C6359DE74D8F5134EC5D405AD479B8383FF7F5C4F503B409167017D50AF479A5A8253EEA575D4BD03177E23593FA97042218138CD9A8F6738340328D7D9D4465A2BEDBF4B0313F52F5CC7E7BB31C78A240266F1AD1AD9A24046EF849218E1CD55B21A6B27DA76EA09C7280AFEDE8B42832E21761163FCE8069CF5A40631F4A3EB869DF86552ABDAAE58F825A716A62A4F7340425A5E58061D966FFD94E7413B96D9150163270C8F8CD45B70ABA07884D5D29DD7454D8BB4DDA19B387F9F75CD7B17A75C5A9910D5399A2AD1BA1B148A598F11F36679789557DC5B71135C1439D964E7B83166F5C19D789B46755BBBF80E760DC02E7A43DA3B0A1478384827CE429B1FD02DD4AF2C3DC6BD77BF130193C577BE1CD0D0A910F837BC65894C3930A53B424BD809CC97316B4D658DC6F9FCFDDCF17A6956A6E97CAEEA09A85D6EAAFDCC55BABFFE7F924CE38DB3116206FB887D70F167286A38176D44FFBF6B9141B4BD80608A0BCCCE549E9587BD234CC292FE696DEF79C80B3EDF5F40067B62449777EC842BBA10E0BF1C32602AA06A2BEC8E99D387C91D3AEEE0B56A612332F15997AA6A12FACD66674594DB81886355052371AB305FAE961AC33BC6F122F4209BBC8AE34115406F57F9D6A7F934744962C3D6C8E8E99D913753007C90828E38D58854E83F21B509663D64B6DC5DFC076E73828A52F9B8DEEA07379524F643B801E6559457E9F25192E9A726D24CA166F70D7BE88E41CF35AF4C76B86ED431AEFE80B2AE8717D153BB2D6DCC12BF70B62B03F272401E34C6F3233A2CAD739B5EAFC1EBAE467E61AEE0DC2A4B77EFCC964A56EAA98CDE499E98F4DFD7DE55135F3047AFF7605C3E4790A7EBB88FE73EE8FC2BE39E1AB38679B619AF92332FFCC728CCA185D63B8745B742F0C99CE0EF199A78C88BAA8FDE2C6DCFB9D576A34ED734BF55BA2A6395CEDECD927D2682263B4AFDFF6A529F37AFDE46CEFF38953977CEEC0DBDBD20934ED44D40C720BD2EAD2356BB5DF2D8B3CE102AB9AC37C413533A9F2BD050E6C8703FAE8A9B91548A8405B313025A0E320DEB3548F7DF344E7DA85068950E05DCA4B7D57E23D72C1AF75880A5E955A2603A01A4792DE7782DC7411AEDB83ABE51F4576841805E64004B283543D25C6870AB79987A53967EAD0FF8C35A238B421143839951AC6F493429AE0A1199BFCA730EFAB6037FDB79C34714BA151A070B43E878E1B9FD3E250F3977BECA62647E0A8CBABC6A331348C9DBC0E96B8306528879F8915F520250769BD4346847A279DDD34AF69D3A489DC1C9E98B3DC5ED5FE0E4AA02DD3E0E79286D195C2D9E0963FCBEE3316DA07A807BD8712F44B9C72D69CBB4F20578F02CD88298AF02EEDA258336CAB9CC65F11844E101E6C7A969182B5517EF11A7C7823A3929DCD4ECAAFA9C9C6EB6E48F54466C9092FB4C8079877F3DF871884E40DD7C1D93493BA5D74AB8B1E047FBC90B23C4588E1C7F965B5C6101D5365AA1B4E51C17C53084E0179AB2DE275AB40EC260297B80DA52A2EEF4728012CD956D6611C44B461BE19D3051F88211379168E4908A75A39A2BA56602ADDE9F2C57ACFC8E9841D8EC3EFBB3537F27D8E90B9202CD50448626ECB4AAAB0A6A7DC466DBD642ED43CCE47D06608C140F327BCA018B813F6F055849C4C3FA551C13E212FFF42624B7DFF445C4F54CE7877998F3D4C639C6FB563A6A859D4004F4681ADF668475B35CF90385C6EBF1C787D0AA0DB16905213ADC6547F523CB55F9C17AB840F5C6CCA2BE31B02E48C86BA4CA2AE0F6DDCEB36DCED8AD36721897105C2416AA3F57C9E56333D838FB38C245D59C343D23FDA0BC2528F5D59D5A88A0C07016B3AE435D17A3BC981597DA6F5F48AF225F76D63100CEE0D5832A4FD6422285F59013525E0E80D57BC8EC19F5B1C8273452EADE76867DB0CA087706E94A36C67E7FE78532BEC2F1F59DAA76D43FB670714715110B26AB6474DE9A8FDE5C59D59982D6A96547C7F564A6F2F7D117A84D036EBB5F9E5E2EC7816164BC98121FAB717535368DCB9AB4298CBE1F4C0EADCCA2E9A857B45A1951F266539E6760C15ECDFF1517ECDF3A35651D4DD4EB17FC30B8364DAA5410ED0A15087811CF4D22C2E50E34D631E285B89DA89649E6F4F1D5DA243ADA9E0677B7EDD3AA55D5852E0B46D254030402E567FB17CB2EF58C47C7C7D704C998771545E9590C55B16C2DDE9D98EDD9DB334FA77C919F76E113F5D13F9033B270936913A1DF8AB8FD75966B09B7AA59E323FD16EB6502E6B177E843FD12772C86655674CAE9E3A87E2DAF015BBF2284C17499D18438F82A699EC60758110D5295E0E3AFC324D39BBE72B7CF16F860B4AA9B57955CABB7EBAA956AF7B84B37DCF96988F1EFB5F49150555DCA661C9AB55F5678447FCE5B6F0AC1C9F7B4953F6CEC6B7433A9AA8CB7AF1F3C2D5EB3FAACB3A5531B53336319094DF61C87A4DD3CDD7EC9351C0A505D3A44A3993E1A1FD78B257F9FC8B7E4A2A2DBCEE571B2F4AA0121733DAE449C52B5F1EDB45DD8B9FCAE3364AA7576278CF83C4F8E67119BA3D249002A07BE13323519D5DB4E40257189198FDB158F4A150C782348D9731BA12C466175DE67BE3D643E5617E4535F1EBA3F5E777436F52B765F8DE0238B3158D6C199CC8AB63C42EC7ECE5E9FAABDA766A874039923FC4AE78EF0E59AF8D8E4BA6F6AE87A47552AEA831849C40485F805BAD9B0D50C3B3270829AAAAE8A78ED8A725375BDAB2D87F724CB014030795D6D06778E304FD24EC9E26B171D662B4
))
GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  OFF 
GO

exec(decrypt(0x05DFAFE2A33A32060660F7F6695C73CD5DE651D1CA7E4C6BDAB0904C423C5CC1AD8D73E01CA3A221B15829862D60BC80FEA3D6B2D2453E94699D2103D97C56563B6B77D91AE514694C96BF86579E6B149241013F6943983862733CA1CD809689995EF0B8C4219B2D0585C91D927A18C5B9D4EDAE9A8E37BA1CC680523361C87E21F942846C3F0F6FC354222BB1160FF45D55BBF7A6473B5811AF1BA79068CE142101BAF84352D27C3A5EFCA0338E
))
GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO
/****** 对象:  StoredProcedure     脚本日期: 02/26/2008 15:56:47 ******/
IF  EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dv_disp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
DROP PROCEDURE [dv_disp]

IF  EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dv_getTopTopic]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
DROP PROCEDURE [dv_getTopTopic]

IF  EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dv_TopicList]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
DROP PROCEDURE [dv_TopicList]

IF  EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[Dv_loadSetup]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
DROP PROCEDURE [Dv_loadSetup]

IF  EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[Dv_dispbbs]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
DROP PROCEDURE [Dv_dispbbs]

IF  EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dv_list]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
DROP PROCEDURE [dv_list]


/****** 对象:  StoredProcedure [dv_getTopTopic]    脚本日期: 02/26/2008 15:59:46 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dv_getTopTopic]
@topidlist nvarchar(500)

AS
declare @str_sql nvarchar(500)

set @str_sql='Select topicid,boardid,title,postusername,postuserid,dateandtime,child,hits,votetotal,lastpost,lastposttime,istop,isvote,isbest,locktopic,expression,topicmode,mode,getmoney,getmoneytype,usetools,issmstopic,hidename from dv_topic Where istop > 0 and topicid in ('+@topidlist +') Order By istop desc, Lastposttime Desc'
	exec sp_executesql @str_sql
	

set nocount off

/****** 对象:  StoredProcedure [dv_TopicList]    脚本日期: 02/26/2008 16:00:08 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dv_TopicList]
@boardid int=1,
@pagenow int=1,	--当前页数             
@pagesize int=1,	--定义每页面帖子数目
@topicmode int=0,	--专题
@inConditions varchar(250)=' ',
@inOrder int=0,
@inSort int=0,
@totalrec int output --SET @TotalRec=@@ROWCOUNT
 AS
SET nocount on
DECLARE @int_topnum int
DECLARE @int_timenum int
DECLARE @var_times varchar(5000)
DECLARE @OrderField varchar(20)
DECLARE @SortStr varchar(5)
DECLARE @strSQL nvarchar(4000)
DECLARE @Compare varchar(1)
DECLARE @nRet int
Declare @Compare1 nvarchar(20)
IF	@inOrder=0
	SET @OrderField='LastPostTime'
ELSE IF @inOrder=1
	SET @OrderField='TopicId'
ELSE IF @inOrder=2
	SET @OrderField='hits'
ELSE IF @inOrder=3
	SET @OrderField='child'
ELSE
	SET @OrderField='LastPostTime'

IF	@inSort=0
	BEGIN
	SET @SortStr='DESC'
	SET @Compare = '<'
	Set @Compare1='Min'
	END
ELSE
	BEGIN
	SET @SortStr='ASC'
	SET @Compare = '>'
	Set @Compare1='Max'
	END
	

IF @pagenow>1
	IF @topicmode>0
		BEGIN
			SELECT @int_timenum=(@pagenow-1) * @pagesize
			--SET ROWCOUNT @int_timenum
			SET @strSQL='SELECT @var_times ='+@Compare1+'(' + @OrderField  + ')  FROM (Select Top '+str(@int_timenum)+' ' + @OrderField + ' From Dv_Topic WHERE mode=@3 And boardID=@2 AND istop = 0 ' + @inConditions + 'ORDER BY ' + @OrderField + ' ' + @SortStr +') as t'  
			EXEC sp_executesql @strSQL,N'@var_times varchar(5000) output,@2 int,@3 int',@var_times output,@2=@boardID,@3=@topicmode

			SET ROWCOUNT @pagesize
			SET @strSQL='SELECT TopicID,boardid,title,postusername,postuserid,dateandtime,child,hits,votetotal,lastpost,lastposttime,istop,isvote,isbest,locktopic,Expression,TopicMode,Mode,GetMoney,GetMoneyType,UseTools,IsSmsTopic,HideName FROM dv_topic WHERE mode=@3 AND boardID=@2 AND istop = 0 AND ' + @OrderField + @Compare + ' @1 ' + @inConditions + ' ORDER BY ' + @OrderField + ' ' + @SortStr
			EXEC sp_executesql @strSQL,N'@1 varchar(5000),@2 int,@3 int',@1=@var_times,@2=@boardID,@3=@topicmode

			SET @strSQL='SELECT @nRet=COUNT(1) FROM Dv_Topic WHERE mode=@3 AND boardID=@2 AND istop=0 ' + @inConditions
			EXEC sp_executesql @strSQL,N'@nRet int output , @2 int,@3 int',@nRet output ,@2=@boardID,@3=@topicmode
			SELECT @totalrec=@nRet

			SET nocount OFF
			RETURN
		END
	ELSE	--@topicmode
		BEGIN
			SELECT @int_timenum=(@pagenow-1) * @pagesize
			--SET ROWCOUNT @int_timenum
			--SET @strSQL='SELECT @var_times=' + @OrderField + '  FROM Dv_Topic WHERE boardID=@2 AND istop=0 ' + @inConditions +' ORDER BY ' + @OrderField + ' ' + @SortStr
			SET @strSQL='SELECT @var_times ='+@Compare1+'(' + @OrderField  + ')  FROM (Select Top '+str(@int_timenum)+' ' + @OrderField + ' From Dv_Topic WHERE boardID=@2 AND istop=0 ' + @inConditions +' ORDER BY ' + @OrderField + ' ' + @SortStr +') as t'  
			EXEC sp_executesql @strSQL,N'@var_times varchar(5000) output,@2 int',@var_times output,@2=@boardID

			SET ROWCOUNT @pagesize
			SET @strSQL='SELECT TopicID,boardid,title,postusername,postuserid,dateandtime,child,hits,votetotal,lastpost,lastposttime,istop,isvote,isbest,locktopic,Expression,TopicMode,Mode,GetMoney,GetMoneyType,UseTools,IsSmsTopic,HideName FROM dv_topic WHERE boardID=@2 AND istop = 0 AND ' + @OrderField + @Compare + ' @1 ' + @inConditions + ' ORDER BY ' + @OrderField + ' ' + @SortStr
			EXEC sp_executesql @strSQL,N'@1 varchar(5000),@2 int',@1=@var_times,@2=@boardID

			SET @strSQL='SELECT @nRet=COUNT(1) FROM Dv_Topic WHERE boardID=@2 AND istop=0 ' + @inConditions
			EXEC sp_executesql @strSQL,N'@nRet int output ,@2 int',@nRet output , @2=@boardID
			SELECT @totalrec=@nRet

			SET nocount OFF
			RETURN
		END
ELSE	--pagenow
	IF @topicmode>0
		BEGIN
			SET ROWCOUNT @pagesize
			SET @strSQL='SELECT TopicID,boardid,title,postusername,postuserid,dateandtime,child,hits,votetotal,lastpost,lastposttime,istop,isvote,isbest,locktopic,Expression,TopicMode,Mode,GetMoney,GetMoneyType,UseTools,IsSmsTopic,HideName FROM Dv_topic WHERE mode=@3 AND boardID=@2 AND istop = 0 ' + @inConditions + ' ORDER BY ' + @OrderField + ' ' + @SortStr
			EXEC sp_executesql @strSQL,N'@2 int,@3 int',@2=@boardID,@3=@topicmode

			SET @strSQL='SELECT @nRet=COUNT(1) FROM Dv_Topic WHERE mode=@3 And boardID=@2 AND istop=0 ' + @inConditions
			EXEC sp_executesql @strSQL,N'@nRet int output,@2 int,@3 int',@nRet output,@2=@boardID,@3=@topicmode
			SELECT @totalrec=@nRet
		END
	ELSE	--topicmode
		BEGIN
			SET ROWCOUNT @pagesize
			SET @strSQL='SELECT TopicID,boardid,title,postusername,postuserid,dateandtime,child,hits,votetotal,lastpost,lastposttime,istop,isvote,isbest,locktopic,Expression,TopicMode,Mode,GetMoney,GetMoneyType,UseTools,IsSmsTopic,HideName FROM Dv_topic WHERE boardID=@2 AND istop=0 ' + @inConditions + ' ORDER BY ' + @OrderField + ' ' + @SortStr
			EXEC sp_executesql @strSQL,N'@2 int',@2=@boardID

			SET @strSQL='SELECT @nRet=COUNT(TopicID) FROM Dv_Topic WHERE boardID=@2 AND istop=0 ' + @inConditions
			EXEC sp_executesql @strSQL,N'@nRet int output,@2 int',@nRet output,@2=@boardID
			SELECT @totalrec=@nRet
		END


/****** 对象:  StoredProcedure [dv_disp]    脚本日期: 02/26/2008 15:59:30 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dv_disp]
@boardid int=0,
@rootid int=0,
@pagenow int=1,     
@pagesize int=1,
@totalusetable nvarchar(20)='dv_bbs1'
AS
set nocount on
declare @int_top int
declare @int_begin int
declare @str_sql nvarchar(500)

if @pagenow>1
	begin
		select @int_top=(@pagenow-1)*@pagesize
		SET @str_sql ='SELECT @int_begin = Max(announceid ) FROM (Select Top '+str(@int_top)+'  announceid From '+@totalusetable+'  where  RootID='+str(@rootid)+' and Boardid='+str(@boardid)+' Order By Announceid ) as t'
		
		exec sp_executesql @str_sql,N'@int_begin int output ',@int_begin output

		set @str_sql='select AnnounceID,UserName,Topic,dateandtime,body,Expression,ip,RootID,signflag,isbest,PostUserid,layer,isagree,GetMoneyType,IsUpload,Ubblist,LockTopic,GetMoney,UseTools,PostBuyUser,ParentID,FlashId From '+@totalusetable+' where RootID='+str(@rootid)+' and announceid>'+str(@int_begin)+'  and Boardid='+str(@boardid)+' Order By Announceid'
		set rowcount @pagesize
		exec sp_executesql @str_sql
		set nocount off
		return
	end
else
	begin
	set rowcount @pagesize
	set @str_sql='Select AnnounceID,UserName,Topic,dateandtime,body,Expression,ip,RootID,signflag,isbest,PostUserid,layer,isagree,GetMoneyType,IsUpload,Ubblist,LockTopic,GetMoney,UseTools,PostBuyUser,ParentID,FlashId From '+@totalusetable+' where  RootID='+str(@rootid)+' and Boardid='+str(@boardid)+' Order By Announceid'
	exec sp_executesql @str_sql
	return
	end


GO


--dvbbs8.3新版增加的最快的分页算法 by niutou--

IF  EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[Dv_GetRecordCount]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
DROP PROCEDURE [Dv_GetRecordCount]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE Dv_GetRecordCount
    @tblName      varchar(255),       -- 表名
    @strWhere     varchar(2000) = ''  -- 查询条件 (注意: 不要加 where)
AS
declare @sql varchar(255)
if @strWhere = ''
begin
	set @sql='Select Count(*) From '+@tblName
end
else
begin
	set @sql='Select Count(*) From '+@tblName+' Where ' +@strWhere
end
exec(@sql)
GO



IF  EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[Dv_GetRecordFromPage]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
DROP PROCEDURE [Dv_GetRecordFromPage]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO
/*
  Dv_GetRecordFromPage
  参数说明: @tblName      包含数据的表名
            @fldName      关键字段名
            @PageSize     每页记录数
            @PageIndex    要获取的页码
            @OrderType    排序类型, 0 - 升序, 1 - 降序
            @strWhere     查询条件 (注意: 不要加 where)
  author: 牛头
*/
CREATE PROCEDURE Dv_GetRecordFromPage
    @tblName      varchar(255),       -- 表名
    @fldName      varchar(255),       -- 字段名
    @PageSize     int = 10,           -- 页尺寸
    @PageIndex    int = 1,            -- 页码
    @OrderType    bit = 0,            -- 设置排序类型, 非 0 值则降序
    @strWhere     varchar(2000) = '',  -- 查询条件 (注意: 不要加 where)
    @strtablezd     varchar(2000) = ''
AS

declare @strSQL   varchar(6000)       -- 主语句
declare @strTmp   varchar(1000)       -- 临时变量
declare @strOrder varchar(500)        -- 排序类型

if @OrderType != 0
begin
    set @strTmp = '<(select min'
    set @strOrder = ' order by [' + @fldName + '] desc'
end
else
begin
    set @strTmp = '>(select max'
    set @strOrder = ' order by [' + @fldName +'] asc'
end

set @strSQL = 'select top ' + str(@PageSize) + @strtablezd + 'from ['
    + @tblName + '] where [' + @fldName + ']' + @strTmp + '(['
    + @fldName + ']) from (select top ' + str((@PageIndex-1)*@PageSize) + ' ['
    + @fldName + '] from [' + @tblName + ']' + @strOrder + ') as tblTmp)'
    + @strOrder

if @strWhere != ''
    set @strSQL = 'select top ' + str(@PageSize) + @strtablezd + ' from ['
        + @tblName + '] where [' + @fldName + ']' + @strTmp + '(['
        + @fldName + ']) from (select top ' + str((@PageIndex-1)*@PageSize) + ' ['
        + @fldName + '] from [' + @tblName + '] where ' + @strWhere + ' '
        + @strOrder + ') as tblTmp) and ' + @strWhere + ' ' + @strOrder

if @PageIndex = 1
begin
    set @strTmp = ''
    if @strWhere != ''
        set @strTmp = ' where (' + @strWhere + ')'

    set @strSQL = 'select top ' + str(@PageSize) + @strtablezd + ' from ['
        + @tblName + ']' + @strTmp + ' ' + @strOrder
end

exec (@strSQL)