if exists (select * from dbo.sysobjects where id = object_id(N'[Dv_Boke_KeyWord]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [Dv_Boke_KeyWord]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[Dv_Boke_Post]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [Dv_Boke_Post]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[Dv_Boke_Skins]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [Dv_Boke_Skins]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[Dv_Boke_SysCat]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [Dv_Boke_SysCat]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[Dv_Boke_System]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [Dv_Boke_System]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[Dv_Boke_Topic]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [Dv_Boke_Topic]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[Dv_Boke_Upfile]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [Dv_Boke_Upfile]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[Dv_Boke_User]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [Dv_Boke_User]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[Dv_Boke_UserCat]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [Dv_Boke_UserCat]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[Dv_Boke_UserSave]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [Dv_Boke_UserSave]
GO

CREATE TABLE [Dv_Boke_KeyWord] (
	[KeyID] [int] IDENTITY (1, 1) NOT NULL ,
	[UserID] [int] NULL ,
	[KeyWord] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[nKeyWord] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[LinkUrl] [nvarchar] (250) COLLATE Chinese_PRC_CI_AS NULL ,
	[LinkTitle] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[NewWindows] [tinyint] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [Dv_Boke_Post] (
	[PostID] [int] IDENTITY (1, 1) NOT NULL ,
	[BokeUserID] [int] NOT NULL ,
	[CatID] [int] NOT NULL ,
	[sCatID] [int] NOT NULL ,
	[ParentID] [int] NOT NULL ,
	[RootID] [int] NOT NULL ,
	[UserID] [int] NOT NULL ,
	[UserName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Title] [nvarchar] (250) COLLATE Chinese_PRC_CI_AS NULL ,
	[Content] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[JoinTime] [smalldatetime] NOT NULL ,
	[IP] [nvarchar] (40) COLLATE Chinese_PRC_CI_AS NULL ,
	[sType] [tinyint] NOT NULL ,
	[IsUpfile] [tinyint] NOT NULL ,
	[IsLock] [int] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [Dv_Boke_Skins] (
	[S_ID] [int] IDENTITY (1, 1) NOT NULL ,
	[S_SkinName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[S_Path] [nvarchar] (150) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[S_ViewPic] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[S_Info] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[S_Builder] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [Dv_Boke_SysCat] (
	[sCatID] [int] IDENTITY (1, 1) NOT NULL ,
	[sCatTitle] [nvarchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[sCatNote] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[uCatNum] [int] NOT NULL ,
	[TopicNum] [int] NOT NULL ,
	[PostNum] [int] NOT NULL ,
	[TodayNum] [int] NOT NULL ,
	[sType] [tinyint] NOT NULL ,
	[LastUpTime] [smalldatetime] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [Dv_Boke_System] (
	[S_id] [int] IDENTITY (1, 1) NOT NULL ,
	[S_Name] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[S_Note] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[S_LastPostTime] [smalldatetime] NOT NULL ,
	[S_TopicNum] [int] NOT NULL ,
	[S_PhotoNum] [int] NOT NULL ,
	[S_FavNum] [int] NOT NULL ,
	[S_UserNum] [int] NOT NULL ,
	[S_TodayNum] [int] NOT NULL ,
	[S_PostNum] [int] NOT NULL ,
	[S_Setting] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[S_Url] [nvarchar] (250) COLLATE Chinese_PRC_CI_AS NULL ,
	[S_sDomain] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[SkinID] [int] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [Dv_Boke_Topic] (
	[TopicID] [int] IDENTITY (1, 1) NOT NULL ,
	[CatID] [int] NOT NULL ,
	[sCatID] [int] NOT NULL ,
	[UserID] [int] NOT NULL ,
	[UserName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Title] [nvarchar] (250) COLLATE Chinese_PRC_CI_AS NULL ,
	[TitleNote] [nvarchar] (250) COLLATE Chinese_PRC_CI_AS NULL ,
	[PostTime] [smalldatetime] NOT NULL ,
	[Child] [int] NULL ,
	[Hits] [int] NULL ,
	[IsView] [tinyint] NULL ,
	[IsLock] [tinyint] NOT NULL ,
	[sType] [tinyint] NOT NULL ,
	[LastPostTime] [smalldatetime] NOT NULL ,
	[IsBest] [int] NOT NULL ,
	[S_Key] [nvarchar] (250) COLLATE Chinese_PRC_CI_AS NULL ,
	[Weather] [smallint] NULL ,
	[VisitUser] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[PayMoney] [money] NULL ,
	[PayNumber] [smallint] NULL ,
	[PayTime] [smallint] NULL ,
	[TrackBacks] [int] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [Dv_Boke_Upfile] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[BokeUserID] [int] NOT NULL ,
	[UserID] [int] NOT NULL ,
	[UserName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[CatID] [int] NOT NULL ,
	[sType] [tinyint] NOT NULL ,
	[TopicID] [int] NOT NULL ,
	[PostID] [int] NOT NULL ,
	[IsTopic] [int] NOT NULL ,
	[Title] [nvarchar] (250) COLLATE Chinese_PRC_CI_AS NULL ,
	[FileName] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[sFileName] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[FileType] [int] NOT NULL ,
	[FileSize] [int] NOT NULL ,
	[FileNote] [nvarchar] (250) COLLATE Chinese_PRC_CI_AS NULL ,
	[DownNum] [int] NOT NULL ,
	[ViewNum] [int] NOT NULL ,
	[DateAndTime] [smalldatetime] NULL ,
	[PreviewImage] [nvarchar] (250) COLLATE Chinese_PRC_CI_AS NULL ,
	[IsLock] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [Dv_Boke_User] (
	[UserID] [int] NOT NULL ,
	[UserName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[NickName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[BokeName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[PassWord] [nvarchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[BokeTitle] [nvarchar] (150) COLLATE Chinese_PRC_CI_AS NULL ,
	[BokeChildTitle] [nvarchar] (250) COLLATE Chinese_PRC_CI_AS NULL ,
	[BokeNote] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[JoinBokeTime] [smalldatetime] NOT NULL ,
	[PageView] [int] NOT NULL ,
	[TopicNum] [int] NOT NULL ,
	[FavNum] [int] NOT NULL ,
	[PhotoNum] [int] NOT NULL ,
	[PostNum] [int] NOT NULL ,
	[TodayNum] [int] NOT NULL ,
	[Trackbacks] [int] NOT NULL ,
	[SpaceSize] [float] NOT NULL ,
	[XmlData] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[SysCatID] [int] NOT NULL ,
	[BokeSetting] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[LastUpTime] [smalldatetime] NOT NULL ,
	[SkinID] [int] NOT NULL ,
	[Stats] [int] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [Dv_Boke_UserCat] (
	[uCatID] [int] IDENTITY (1, 1) NOT NULL ,
	[UserID] [int] NOT NULL ,
	[uCatTitle] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[uCatNote] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[OpenTime] [smalldatetime] NOT NULL ,
	[uType] [tinyint] NOT NULL ,
	[TopicNum] [int] NOT NULL ,
	[PostNum] [int] NOT NULL ,
	[TodayNum] [int] NOT NULL ,
	[IsView] [tinyint] NOT NULL ,
	[LastUpTime] [smalldatetime] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [Dv_Boke_UserSave] (
	[UserID] [int] NOT NULL ,
	[SaveDate] [int] NOT NULL ,
	[SaveNum] [int] NOT NULL 
) ON [PRIMARY]
GO

ALTER TABLE [Dv_Boke_KeyWord] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dv_Boke_KeyWord] PRIMARY KEY  CLUSTERED 
	(
		[KeyID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [Dv_Boke_Post] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dv_Boke_Post] PRIMARY KEY  CLUSTERED 
	(
		[PostID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [Dv_Boke_Skins] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dv_Boke_Skins] PRIMARY KEY  CLUSTERED 
	(
		[S_ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [Dv_Boke_SysCat] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dv_Boke_SysCat] PRIMARY KEY  CLUSTERED 
	(
		[sCatID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [Dv_Boke_System] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dv_Boke_System] PRIMARY KEY  CLUSTERED 
	(
		[S_id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [Dv_Boke_Topic] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dv_Boke_Topic] PRIMARY KEY  CLUSTERED 
	(
		[TopicID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [Dv_Boke_Upfile] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dv_Boke_Upfile] PRIMARY KEY  CLUSTERED 
	(
		[ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [Dv_Boke_User] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dv_Boke_User] PRIMARY KEY  CLUSTERED 
	(
		[UserID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [Dv_Boke_UserCat] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dv_Boke_UserCat] PRIMARY KEY  CLUSTERED 
	(
		[uCatID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [Dv_Boke_KeyWord] ADD 
	CONSTRAINT [DF_Dv_Boke_KeyWord_UserID] DEFAULT (0) FOR [UserID],
	CONSTRAINT [DF_Dv_Boke_KeyWord_NewWindows] DEFAULT (0) FOR [NewWindows]
GO

ALTER TABLE [Dv_Boke_Post] ADD 
	CONSTRAINT [DF_Dv_Boke_Post_BokeUserID] DEFAULT (0) FOR [BokeUserID],
	CONSTRAINT [DF_Dv_Boke_Post_CatID] DEFAULT (0) FOR [CatID],
	CONSTRAINT [DF_Dv_Boke_Post_sCatID] DEFAULT (0) FOR [sCatID],
	CONSTRAINT [DF_Dv_Boke_Post_ParentID] DEFAULT (0) FOR [ParentID],
	CONSTRAINT [DF_Dv_Boke_Post_RootID] DEFAULT (0) FOR [RootID],
	CONSTRAINT [DF_Dv_Boke_Post_UserID] DEFAULT (0) FOR [UserID],
	CONSTRAINT [DF_Dv_Boke_Post_JoinTime] DEFAULT (getdate()) FOR [JoinTime],
	CONSTRAINT [DF_Dv_Boke_Post_sType] DEFAULT (0) FOR [sType],
	CONSTRAINT [DF_Dv_Boke_Post_IsUpfile] DEFAULT (0) FOR [IsUpfile],
	CONSTRAINT [DF_Dv_Boke_Post_IsLock0] DEFAULT (0) FOR [IsLock]
GO

 CREATE  INDEX [IX_Dv_Boke_BokeUserID] ON [Dv_Boke_Post]([BokeUserID]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_Dv_Boke_CatID] ON [Dv_Boke_Post]([CatID]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_Dv_Boke_sCatID] ON [Dv_Boke_Post]([sCatID]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_Dv_Boke_ParentID] ON [Dv_Boke_Post]([ParentID]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_Dv_Boke_RootID] ON [Dv_Boke_Post]([RootID]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_Dv_Boke_sType] ON [Dv_Boke_Post]([sType]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_Dv_Boke_IsLock] ON [Dv_Boke_Post]([IsLock]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_Dv_Boke_UserName] ON [Dv_Boke_Post]([UserName]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_Dv_Boke_Post] ON [Dv_Boke_Post]([PostID]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_Dv_Boke_Post_Title] ON [Dv_Boke_Post]([Title]) ON [PRIMARY]
GO

ALTER TABLE [Dv_Boke_SysCat] ADD 
	CONSTRAINT [DF_Dv_Boke_SysCat_uCatNum] DEFAULT (0) FOR [uCatNum],
	CONSTRAINT [DF_Dv_Boke_SysCat_TopicNum] DEFAULT (0) FOR [TopicNum],
	CONSTRAINT [DF_Dv_Boke_SysCat_PostNum] DEFAULT (0) FOR [PostNum],
	CONSTRAINT [DF_Dv_Boke_SysCat_TodayNum] DEFAULT (0) FOR [TodayNum],
	CONSTRAINT [DF_Dv_Boke_SysCat_sType] DEFAULT (0) FOR [sType],
	CONSTRAINT [DF_Dv_Boke_SysCat_LastUpTime] DEFAULT (getdate()) FOR [LastUpTime]
GO

 CREATE  INDEX [IX_Dv_Boke_sType] ON [Dv_Boke_SysCat]([sType]) ON [PRIMARY]
GO

ALTER TABLE [Dv_Boke_System] ADD 
	CONSTRAINT [DF_Dv_Boke_System_S_LastPostTime] DEFAULT (getdate()) FOR [S_LastPostTime],
	CONSTRAINT [DF_Dv_Boke_System_S_TopicNum] DEFAULT (0) FOR [S_TopicNum],
	CONSTRAINT [DF_Dv_Boke_System_S_PhotoNum] DEFAULT (0) FOR [S_PhotoNum],
	CONSTRAINT [DF_Dv_Boke_System_S_FavNum] DEFAULT (0) FOR [S_FavNum],
	CONSTRAINT [DF_Dv_Boke_System_S_UserNum] DEFAULT (0) FOR [S_UserNum],
	CONSTRAINT [DF_Dv_Boke_System_S_TodayNum] DEFAULT (0) FOR [S_TodayNum],
	CONSTRAINT [DF_Dv_Boke_System_S_PostNum] DEFAULT (0) FOR [S_PostNum],
	CONSTRAINT [DF_Dv_Boke_System_SkinID] DEFAULT (0) FOR [SkinID]
GO

ALTER TABLE [Dv_Boke_Topic] ADD 
	CONSTRAINT [DF_Dv_Boke_Topic_CatID] DEFAULT (0) FOR [CatID],
	CONSTRAINT [DF_Dv_Boke_Topic_sCatID] DEFAULT (0) FOR [sCatID],
	CONSTRAINT [DF_Dv_Boke_Topic_UserID] DEFAULT (0) FOR [UserID],
	CONSTRAINT [DF_Dv_Boke_Topic_PostTime] DEFAULT (getdate()) FOR [PostTime],
	CONSTRAINT [DF_Dv_Boke_Topic_Child] DEFAULT (0) FOR [Child],
	CONSTRAINT [DF_Dv_Boke_Topic_Hits] DEFAULT (0) FOR [Hits],
	CONSTRAINT [DF_Dv_Boke_Topic_IsView] DEFAULT (0) FOR [IsView],
	CONSTRAINT [DF_Dv_Boke_Topic_IsLock] DEFAULT (0) FOR [IsLock],
	CONSTRAINT [DF_Dv_Boke_Topic_sType] DEFAULT (0) FOR [sType],
	CONSTRAINT [DF_Dv_Boke_Topic_LastPostTime] DEFAULT (getdate()) FOR [LastPostTime],
	CONSTRAINT [DF_Dv_Boke_Topic_IsBest] DEFAULT (0) FOR [IsBest],
	CONSTRAINT [DF_Dv_Boke_Topic_Weather] DEFAULT (0) FOR [Weather],
	CONSTRAINT [DF_Dv_Boke_Topic_PayMoney] DEFAULT (0) FOR [PayMoney],
	CONSTRAINT [DF_Dv_Boke_Topic_PayNumber] DEFAULT (0) FOR [PayNumber],
	CONSTRAINT [DF_Dv_Boke_Topic_PayTime] DEFAULT (0) FOR [PayTime],
	CONSTRAINT [DF_Dv_Boke_Topic_TrackBacks] DEFAULT (0) FOR [TrackBacks]
GO

 CREATE  INDEX [IX_Dv_Boke_UserID] ON [Dv_Boke_Topic]([UserID]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_Dv_Boke_CatID] ON [Dv_Boke_Topic]([CatID]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_Dv_Boke_sCatID] ON [Dv_Boke_Topic]([sCatID]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_Dv_Boke_IsLock] ON [Dv_Boke_Topic]([IsLock]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_Dv_Boke_sType] ON [Dv_Boke_Topic]([sType]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_Dv_Boke_IsBest] ON [Dv_Boke_Topic]([IsBest]) ON [PRIMARY]
GO

ALTER TABLE [Dv_Boke_Upfile] ADD 
	CONSTRAINT [DF_Dv_Boke_Upfile_BokeUserID] DEFAULT (0) FOR [BokeUserID],
	CONSTRAINT [DF_Dv_Boke_Upfile_CatID] DEFAULT (0) FOR [CatID],
	CONSTRAINT [DF_Dv_Boke_Upfile_sType] DEFAULT (0) FOR [sType],
	CONSTRAINT [DF_Dv_Boke_Upfile_TopicID] DEFAULT (0) FOR [TopicID],
	CONSTRAINT [DF_Dv_Boke_Upfile_PostID] DEFAULT (0) FOR [PostID],
	CONSTRAINT [DF_Dv_Boke_Upfile_IsTopic] DEFAULT (0) FOR [IsTopic],
	CONSTRAINT [DF_Dv_Boke_Upfile_FileType] DEFAULT (0) FOR [FileType],
	CONSTRAINT [DF_Dv_Boke_Upfile_FileSize] DEFAULT (0) FOR [FileSize],
	CONSTRAINT [DF_Dv_Boke_Upfile_DownNum] DEFAULT (0) FOR [DownNum],
	CONSTRAINT [DF_Dv_Boke_Upfile_ViewNum] DEFAULT (0) FOR [ViewNum],
	CONSTRAINT [DF_Dv_Boke_Upfile_DateAndTime] DEFAULT (getdate()) FOR [DateAndTime],
	CONSTRAINT [DF_Dv_Boke_Upfile_IsLock] DEFAULT (0) FOR [IsLock]
GO

 CREATE  INDEX [IX_Dv_Boke_BokeUserID] ON [Dv_Boke_Upfile]([BokeUserID]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_Dv_Boke_PostID] ON [Dv_Boke_Upfile]([PostID]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_Dv_Boke_sType] ON [Dv_Boke_Upfile]([sType]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_Dv_Boke_IsTopic] ON [Dv_Boke_Upfile]([IsTopic]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_Dv_Boke_CatID] ON [Dv_Boke_Upfile]([CatID]) ON [PRIMARY]
GO

ALTER TABLE [Dv_Boke_User] ADD 
	CONSTRAINT [DF_Dv_Boke_User_JoinBokeTime] DEFAULT (getdate()) FOR [JoinBokeTime],
	CONSTRAINT [DF_Dv_Boke_User_PageView] DEFAULT (0) FOR [PageView],
	CONSTRAINT [DF_Dv_Boke_User_TopicNum] DEFAULT (0) FOR [TopicNum],
	CONSTRAINT [DF_Dv_Boke_User_FavNum] DEFAULT (0) FOR [FavNum],
	CONSTRAINT [DF_Dv_Boke_User_PhotoNum] DEFAULT (0) FOR [PhotoNum],
	CONSTRAINT [DF_Dv_Boke_User_PostNum] DEFAULT (0) FOR [PostNum],
	CONSTRAINT [DF_Dv_Boke_User_TodayNum] DEFAULT (0) FOR [TodayNum],
	CONSTRAINT [DF_Dv_Boke_User_Trackbacks] DEFAULT (0) FOR [Trackbacks],
	CONSTRAINT [DF_Dv_Boke_User_SpaceSize] DEFAULT (0) FOR [SpaceSize],
	CONSTRAINT [DF_Dv_Boke_User_SysCatID] DEFAULT (0) FOR [SysCatID],
	CONSTRAINT [DF_Dv_Boke_User_LastUpTime] DEFAULT (getdate()) FOR [LastUpTime],
	CONSTRAINT [DF_Dv_Boke_User_SkinID] DEFAULT (1) FOR [SkinID],
	CONSTRAINT [DF_Dv_Boke_User_Stats] DEFAULT (0) FOR [Stats]
GO

ALTER TABLE [Dv_Boke_UserCat] ADD 
	CONSTRAINT [DF_Dv_Boke_UserCat_UserID] DEFAULT (0) FOR [UserID],
	CONSTRAINT [DF_Dv_Boke_UserCat_OpenTime] DEFAULT (getdate()) FOR [OpenTime],
	CONSTRAINT [DF_Dv_Boke_UserCat_uType] DEFAULT (0) FOR [uType],
	CONSTRAINT [DF_Dv_Boke_UserCat_TopicNum] DEFAULT (0) FOR [TopicNum],
	CONSTRAINT [DF_Dv_Boke_UserCat_PostNum] DEFAULT (0) FOR [PostNum],
	CONSTRAINT [DF_Dv_Boke_UserCat_TodayNum] DEFAULT (0) FOR [TodayNum],
	CONSTRAINT [DF_Dv_Boke_UserCat_IsView] DEFAULT (0) FOR [IsView],
	CONSTRAINT [DF_Dv_Boke_UserCat_LastUpTime] DEFAULT (getdate()) FOR [LastUpTime]
GO

 CREATE  INDEX [IX_Dv_Boke_UserID] ON [Dv_Boke_UserCat]([UserID]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_Dv_Boke_UType] ON [Dv_Boke_UserCat]([uType]) ON [PRIMARY]
GO

ALTER TABLE [Dv_Boke_UserSave] ADD 
	CONSTRAINT [DF_Dv_Boke_UserSave_UserID] DEFAULT (0) FOR [UserID],
	CONSTRAINT [DF_Dv_Boke_UserSave_SaveDate] DEFAULT (0) FOR [SaveDate],
	CONSTRAINT [DF_Dv_Boke_UserSave_SaveNum] DEFAULT (0) FOR [SaveNum]
GO

 CREATE  INDEX [Dv_Boke_UserSave] ON [Dv_Boke_UserSave]([UserID]) ON [PRIMARY]
GO


InSert Into Dv_Boke_Skins (S_SkinName,S_Path,S_ViewPic,S_Builder) Values ('iBoke 默认风格','Boke/Skins/Default/','Boke/Skins/Default/viewlogo.png','AspSky.Net')
go
InSert Into Dv_Boke_Skins (S_SkinName,S_Path,S_ViewPic,S_Builder) Values ('iBoke 默认风格2','Boke/Skins/dvskin/','Boke/Skins/dvskin/viewlogo.png','AspSky.Net')
go

InSert Into Dv_Boke_SysCat (sCatTitle,sType) Values ('个人空间',0)
go
InSert Into Dv_Boke_SysCat (sCatTitle,sType) Values ('电脑网络',0)
go
InSert Into Dv_Boke_SysCat (sCatTitle,sType) Values ('经济金融',0)
go
InSert Into Dv_Boke_SysCat (sCatTitle,sType) Values ('教育学习',0)
go
InSert Into Dv_Boke_SysCat (sCatTitle,sType) Values ('情感绿洲',0)
go
InSert Into Dv_Boke_SysCat (sCatTitle,sType) Values ('娱乐休闲',0)
go
InSert Into Dv_Boke_SysCat (sCatTitle,sType) Values ('少儿乐园',0)
go
InSert Into Dv_Boke_SysCat (sCatTitle,sType) Values ('原创文学',0)
go
InSert Into Dv_Boke_SysCat (sCatTitle,sType) Values ('生活百科',0)
go
InSert Into Dv_Boke_SysCat (sCatTitle,sType) Values ('大话天下',0)
go
InSert Into Dv_Boke_SysCat (sCatTitle,sType) Values ('体育竞技',0)
go
InSert Into Dv_Boke_SysCat (sCatTitle,sType) Values ('旅游自然',0)
go
InSert Into Dv_Boke_SysCat (sCatTitle,sType) Values ('网络游戏',0)
go
InSert Into Dv_Boke_SysCat (sCatTitle,sType) Values ('班级集体',0)
go
InSert Into Dv_Boke_SysCat (sCatTitle,sType) Values ('政法军事',0)
go
InSert Into Dv_Boke_SysCat (sCatTitle,sType) Values ('电子商务',0)
go
InSert Into Dv_Boke_SysCat (sCatTitle,sType) Values ('日记',1)
go
InSert Into Dv_Boke_SysCat (sCatTitle,sType) Values ('知识',1)
go
InSert Into Dv_Boke_SysCat (sCatTitle,sType) Values ('读书',1)
go
InSert Into Dv_Boke_SysCat (sCatTitle,sType) Values ('心情',1)
go
InSert Into Dv_Boke_SysCat (sCatTitle,sType) Values ('生活',1)
go
InSert Into Dv_Boke_SysCat (sCatTitle,sType) Values ('体育',1)
go
InSert Into Dv_Boke_SysCat (sCatTitle,sType) Values ('下载',1)
go
InSert Into Dv_Boke_SysCat (sCatTitle,sType) Values ('数码',1)
go
InSert Into Dv_Boke_SysCat (sCatTitle,sType) Values ('帖图',1)
go
InSert Into Dv_Boke_SysCat (sCatTitle,sType) Values ('电影',1)
go
InSert Into Dv_Boke_SysCat (sCatTitle,sType) Values ('音乐',1)
go
InSert Into Dv_Boke_SysCat (sCatTitle,sType) Values ('网络',1)
go
InSert Into Dv_Boke_SysCat (sCatTitle,sType) Values ('软件',1)
go
InSert Into Dv_Boke_SysCat (sCatTitle,sType) Values ('开发',1)
go
InSert Into Dv_Boke_SysCat (sCatTitle,sType) Values ('传媒',1)
go
InSert Into Dv_Boke_SysCat (sCatTitle,sType) Values ('交易',1)
go

InSert Into Dv_Boke_System (S_Name,S_Setting,S_Url,SkinID) Values ('动网博客','1,1,1,1,1,1,20,20,15,3,1,1,1|0|0|999|bbs.dvbbs.net|12|1|Arial|0|images/WaterMap.gif|0.7|110|35|4|120|100|1|1|1|Boke/UploadFile/|0,晴天|和煦|阴天|清爽|多云|有雾|小雨|中雨|雷雨|彩虹|酷热|寒冷|小雪|大雪|月圆|月缺,sun.gif|sun2.gif|yin.gif|yin.gif|yun.gif|wu.gif|xiaoyu.gif|yinyu.gif|leiyu.gif|caihong.gif|sun.gif|feng.gif|xue.gif|daxue.gif|moon.gif|moon2.gif,50,gif|jpg|jpeg|bmp|png|rar|txt|zip|mid,欢迎大家来到动网博客建立自己的网上家园,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1','http://www.iboker.com',1)
go