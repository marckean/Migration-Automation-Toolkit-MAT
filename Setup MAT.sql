/******    Migration Automation Script    ******/
/******           Database setup          ******/
/******           version 1.5.2           ******/
/******                                   ******/
/******     Copyright 2014 Microsoft      ******/
/******                                   ******/
/******    http:\\aka.ms\buildingclouds   ******/


USE [master]

/****** Create Database [MAT]   ******/
DECLARE @data_path nvarchar(256);
SET @data_path = (SELECT SUBSTRING(physical_name, 1, CHARINDEX(N'master.mdf', LOWER(physical_name)) - 1)
                  FROM master.sys.master_files
                  WHERE database_id = 1 AND file_id = 1);

IF EXISTS(SELECT * FROM sys.sysdatabases where name='MAT')
        GOTO MATDB_Exists
ELSE 
        Print 'Creating the MAT Database v1.5.2'

Execute('CREATE DATABASE [MAT]
ON  PRIMARY 
	(
	NAME = MAT
	,FILENAME = ''' + @data_path + 'MAT.mdf''
	,SIZE = 4096KB 
	,MAXSIZE = UNLIMITED
	,FILEGROWTH = 1024KB )
	LOG ON 
	(
	NAME = MAT_log
	,FILENAME = ''' + @data_path + 'MAT_log.ldf''
	,SIZE = 1024KB 
	,MAXSIZE = 2048MB 
	,FILEGROWTH = 10%
	)'
)
ALTER DATABASE [MAT] SET COMPATIBILITY_LEVEL = 100


IF (1 = FULLTEXTSERVICEPROPERTY('IsFullTextInstalled'))
begin
EXEC [MAT].[dbo].[sp_fulltext_database] @action = 'enable'
end


ALTER DATABASE [MAT] SET ANSI_NULL_DEFAULT OFF 
ALTER DATABASE [MAT] SET ANSI_NULLS OFF 
ALTER DATABASE [MAT] SET ANSI_PADDING OFF 
ALTER DATABASE [MAT] SET ANSI_WARNINGS OFF 
ALTER DATABASE [MAT] SET ARITHABORT OFF 
ALTER DATABASE [MAT] SET AUTO_CLOSE OFF 
ALTER DATABASE [MAT] SET AUTO_CREATE_STATISTICS ON 
ALTER DATABASE [MAT] SET AUTO_SHRINK OFF 
ALTER DATABASE [MAT] SET AUTO_UPDATE_STATISTICS ON 
ALTER DATABASE [MAT] SET CURSOR_CLOSE_ON_COMMIT OFF 
ALTER DATABASE [MAT] SET CURSOR_DEFAULT  GLOBAL 
ALTER DATABASE [MAT] SET CONCAT_NULL_YIELDS_NULL OFF 
ALTER DATABASE [MAT] SET NUMERIC_ROUNDABORT OFF 
ALTER DATABASE [MAT] SET QUOTED_IDENTIFIER OFF 
ALTER DATABASE [MAT] SET RECURSIVE_TRIGGERS OFF 
ALTER DATABASE [MAT] SET  DISABLE_BROKER 
ALTER DATABASE [MAT] SET AUTO_UPDATE_STATISTICS_ASYNC OFF 
ALTER DATABASE [MAT] SET DATE_CORRELATION_OPTIMIZATION OFF 
ALTER DATABASE [MAT] SET TRUSTWORTHY OFF 
ALTER DATABASE [MAT] SET ALLOW_SNAPSHOT_ISOLATION OFF 
ALTER DATABASE [MAT] SET PARAMETERIZATION SIMPLE 
ALTER DATABASE [MAT] SET READ_COMMITTED_SNAPSHOT OFF 
ALTER DATABASE [MAT] SET HONOR_BROKER_PRIORITY OFF 
ALTER DATABASE [MAT] SET  READ_WRITE 
ALTER DATABASE [MAT] SET RECOVERY FULL 
ALTER DATABASE [MAT] SET  MULTI_USER 
ALTER DATABASE [MAT] SET PAGE_VERIFY CHECKSUM  
ALTER DATABASE [MAT] SET DB_CHAINING OFF
Print 'MAT Database created.' 
GOTO MATDB_Done

MATDB_Exists:
Print 'The MAT Database already exists.'
GOTO MATDB_Done

MATDB_Done: 
GO

/****** Finshied Creating Database [MAT]   ******/
USE [MAT]

/****** Creat Version Table  ******/
IF EXISTS (SELECT 1 FROM sysobjects WHERE xtype='u' AND name='VERSION') 
        GOTO VERSION_Exists
ELSE 
	SET ANSI_NULLS ON
	SET QUOTED_IDENTIFIER ON
	
CREATE TABLE [dbo].[VERSION](
	[InternalVersion] [int] NOT NULL,
	[DBVersion] [varchar](10),
 CONSTRAINT [PK_VERSION] PRIMARY KEY CLUSTERED 
(
	[InternalVersion] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
INSERT INTO [dbo].[VERSION] VALUES ('152', '1.5.2')
Print 'Version Table created.'

GOTO VERSION_Done

VERSION_Exists:
Print 'VERSION already exists.'
/****** Verify newer version is NOT Installed   ******/
	IF (Select [InternalVersion]from [VERSION]) > 152
		BEGIN
		Print 'The Exitsing database version is newer than this version. Terminating.'
		Print 'No changes were made.' 
		set noexec on 
		END
	IF (Select [InternalVersion]from [VERSION]) < 152
		BEGIN
		Print 'Updating the database version.'
		Update [Version] SET [InternalVersion] = '152'
		Update [Version] SET [DBVersion] = '1.5.2'
		
		END
GOTO VERSION_Done

VERSION_Done: 
GO


/******  Create Table [dbo].[VMQueue] ******/
IF EXISTS (SELECT 1 FROM sysobjects WHERE xtype='u' AND name='VMQueue') 
        BEGIN
        Drop Table VMQueue 
        Print 'Dropped Table VMQueue'
        END
	ELSE         
	SET ANSI_NULLS ON
	SET QUOTED_IDENTIFIER ON

CREATE TABLE [dbo].[VMQueue](
	[JobID] [int] IDENTITY(1,1) NOT NULL,
	[VMName] [nvarchar](255) NOT NULL,
	[objectID] [nvarchar](255) NULL,
	[Supported] [tinyint] NOT NULL,
	[ReadytoConvert] [bit] NULL,
	[ConvertServer] [nvarchar](255) NULL,
	[StartTime] [datetime] NULL,
	[EndTime] [datetime] NULL,
	[Status] [tinyint] NULL,
	[Warning] [bit] NULL,
	[Summary] [nvarchar](255) NULL,
	[Notes] [nvarchar](255) NULL,
	[Completed] [bit] NULL,
	[PID] [int] NULL,
	[Position] [int] NULL,
	[InUse] [bit] NULL,
 CONSTRAINT [PK_VMQueue] PRIMARY KEY CLUSTERED 
(
	[JobID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

Print 'VMQueue Table created.'
GOTO VMQueue_Done

VMQueue_Done: 
GO

/****** Creat  Table [dbo].[VMTools]  ******/
IF EXISTS (SELECT 1 FROM sysobjects WHERE xtype='u' AND name='VMTools') 
        BEGIN
        Drop Table VMTools 
        Print 'Dropped Table VMTools'
        END 

ELSE 
	SET ANSI_NULLS ON
	SET QUOTED_IDENTIFIER ON
	SET ANSI_PADDING ON

CREATE TABLE [dbo].[VMTools](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[VMwareToolsVerion] [int] NOT NULL,
	[GUID] [nvarchar] (max) NOT NULL,
	[GUESTID] [char](50) NOT NULL,
 CONSTRAINT [PK_VMTools] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]


Print 'VMTools Table created.'

/****** Populate VMTools   ******/
Print 'Populating VMTools Table...'
INSERT INTO VMTools  VALUES (8196,'{FE2F6A2C-196E-4210-9C04-2B1BC21F07EF}','winNetEnterpriseGuest')
INSERT INTO VMTools  VALUES (8198,'{FE2F6A2C-196E-4210-9C04-2B1BC21F07EF}','winNetEnterpriseGuest')
INSERT INTO VMTools  VALUES (8295,'{FE2F6A2C-196E-4210-9C04-2B1BC21F07EF}','winNetEnterpriseGuest')
INSERT INTO VMTools  VALUES (8384,'{A5CD39D8-F8A7-494F-9357-878A4AB6537F}','win7Server64Guest')
INSERT INTO VMTools  VALUES (8384,'{62198C42-974B-4F90-9AD2-12763AB58C97}','winNetEnterpriseGuest')
INSERT INTO VMTools  VALUES (8384,'{62198C42-974B-4F90-9AD2-12763AB58C97}','winLonghornGuest')
INSERT INTO VMTools  VALUES (8384,'{A5CD39D8-F8A7-494F-9357-878A4AB6537F}','winLonghorn64Guest')
INSERT INTO VMTools  VALUES (8384,'{A5CD39D8-F8A7-494F-9357-878A4AB6537F}','winNetEnterprise64Guest')
INSERT INTO VMTools  VALUES (8196,'{FE2F6A2C-196E-4210-9C04-2B1BC21F07EF}','win7Server64Guest')
INSERT INTO VMTools  VALUES (8196,'{FE2F6A2C-196E-4210-9C04-2B1BC21F07EF}','winLonghornGuest')
INSERT INTO VMTools  VALUES (8196,'{FE2F6A2C-196E-4210-9C04-2B1BC21F07EF}','winLonghorn64Guest')
INSERT INTO VMTools  VALUES (8196,'{FE2F6A2C-196E-4210-9C04-2B1BC21F07EF}','winNetEnterprise64Guest')
INSERT INTO VMTools  VALUES (8198,'{FE2F6A2C-196E-4210-9C04-2B1BC21F07EF}','win7Server64Guest')
INSERT INTO VMTools  VALUES (8198,'{FE2F6A2C-196E-4210-9C04-2B1BC21F07EF}','winLonghornGuest')
INSERT INTO VMTools  VALUES (8198,'{FE2F6A2C-196E-4210-9C04-2B1BC21F07EF}','winLonghorn64Guest')
INSERT INTO VMTools  VALUES (8198,'{FE2F6A2C-196E-4210-9C04-2B1BC21F07EF}','winNetEnterprise64Guest')
INSERT INTO VMTools  VALUES (8295,'{FE2F6A2C-196E-4210-9C04-2B1BC21F07EF}','winNetEnterprise64Guest')
INSERT INTO VMTools  VALUES (8295,'{FE2F6A2C-196E-4210-9C04-2B1BC21F07EF}','winLonghorn64Guest')
INSERT INTO VMTools  VALUES (8295,'{FE2F6A2C-196E-4210-9C04-2B1BC21F07EF}','winLonghornGuest')
INSERT INTO VMTools  VALUES (8295,'{FE2F6A2C-196E-4210-9C04-2B1BC21F07EF}','win7Server64Guest')
INSERT INTO VMTools  VALUES (8389,'{DD3770AA-8012-453F-AD8D-0B6D91ED40D5}','winNetEnterpriseGuest')  
INSERT INTO VMTools  VALUES (8389,'{A2CC6F0B-E888-4485-82F5-587699B3CDB7}','winLonghorn64Guest')
INSERT INTO VMTools  VALUES (8384,'{A5CD39D8-F8A7-494F-9357-878A4AB6537F}','windows7Server64Guest')
INSERT INTO VMTools  VALUES (8384,'{62198C42-974B-4F90-9AD2-12763AB58C97}','windowsLonghornGuest')
INSERT INTO VMTools  VALUES (8384,'{A5CD39D8-F8A7-494F-9357-878A4AB6537F}','windowsLonghorn64Guest')
INSERT INTO VMTools  VALUES (8196,'{FE2F6A2C-196E-4210-9C04-2B1BC21F07EF}','windows7Server64Guest')
INSERT INTO VMTools  VALUES (8196,'{FE2F6A2C-196E-4210-9C04-2B1BC21F07EF}','windowsLonghornGuest')
INSERT INTO VMTools  VALUES (8196,'{FE2F6A2C-196E-4210-9C04-2B1BC21F07EF}','windowsLonghorn64Guest')
INSERT INTO VMTools  VALUES (8198,'{FE2F6A2C-196E-4210-9C04-2B1BC21F07EF}','windows7Server64Guest')
INSERT INTO VMTools  VALUES (8198,'{FE2F6A2C-196E-4210-9C04-2B1BC21F07EF}','windowsLonghornGuest')
INSERT INTO VMTools  VALUES (8198,'{FE2F6A2C-196E-4210-9C04-2B1BC21F07EF}','windowsLonghorn64Guest')
INSERT INTO VMTools  VALUES (8295,'{FE2F6A2C-196E-4210-9C04-2B1BC21F07EF}','windowsLonghorn64Guest')
INSERT INTO VMTools  VALUES (8295,'{FE2F6A2C-196E-4210-9C04-2B1BC21F07EF}','windowsLonghornGuest')
INSERT INTO VMTools  VALUES (8295,'{FE2F6A2C-196E-4210-9C04-2B1BC21F07EF}','windows7Server64Guest')
INSERT INTO VMTools  VALUES (8389,'{A2CC6F0B-E888-4485-82F5-587699B3CDB7}','windowsLonghorn64Guest')
INSERT INTO VMTools  VALUES (9216,'{4D80C805-67C3-4525-A7BA-DC43215E9167}','windows7Server64Guest')
INSERT INTO VMTools  VALUES (8384,'/usr/bin/vmware-uninstall-tools.pl | shutdown -h +3','centos64Guest')
INSERT INTO VMTools  VALUES (8384,'/usr/bin/vmware-uninstall-tools.pl | shutdown -h +3','redhatGuest')
INSERT INTO VMTools  VALUES (8384,'/usr/bin/vmware-uninstall-tools.pl | shutdown -h +3','sles11_64Guest')
INSERT INTO VMTools  VALUES (8384,'/usr/bin/vmware-uninstall-tools.pl | shutdown -h +3','centos64Guest')
INSERT INTO VMTools  VALUES (8384,'/usr/bin/vmware-uninstall-tools.pl | shutdown -h +3','redhatGuest')
INSERT INTO VMTools  VALUES (8384,'/usr/bin/vmware-uninstall-tools.pl | shutdown -h +3','sles11_64Guest') 
INSERT INTO VMTools  VALUES (8295,'/usr/bin/vmware-uninstall-tools.pl | shutdown -h +3','centos64Guest')
INSERT INTO VMTools  VALUES (8295,'/usr/bin/vmware-uninstall-tools.pl | shutdown -h +3','rhel5_64Guest') 
INSERT INTO VMTools  VALUES (8295,'/usr/bin/vmware-uninstall-tools.pl | shutdown -h +3','sles10_64Guest') 
INSERT INTO VMTools  VALUES (8295,'/usr/bin/vmware-uninstall-tools.pl | shutdown -h +3','rhel6_64Guest') 
INSERT INTO VMTools  VALUES (8295,'/usr/bin/vmware-uninstall-tools.pl | shutdown -h +3','sles11_64Guest')

GOTO VMTools_Done

VMTools_Exists:
Print 'VMTools already exists.'
GOTO VMTools_Done

VMTools_Done: 
GO



/******  Create Table [dbo].[VMConversionData_PS] ******/
IF EXISTS (SELECT 1 FROM sysobjects WHERE xtype='u' AND name='VMConversionData_PS') 
        BEGIN
        Drop Table VMConversionData_PS
        Print 'Dropped Table VMConversionData_PS'
        END
ELSE         
	SET ANSI_NULLS ON
	SET QUOTED_IDENTIFIER ON

CREATE TABLE [dbo].[VMConversionData_PS](
	[MoRef] [nvarchar](10) NULL,
	[DisplayName] [nvarchar](255) NULL,
	[VMHostName] [nvarchar](255) NULL,
	[VMHostID] [nvarchar](max) NULL,
	[GuestVMFQDN] [nvarchar](max) NULL,
	[GuestVMID] [nvarchar](255) NULL,
	[GUESTID] [char](50) NULL,
	[GuestOS] [nvarchar](255) NULL,
	[GuestVMMB] [int] NULL,
	[GuestVMCPUCount] [int] NULL,
	[GuestVMHDProvisionedGB] [float] NULL,
	[GuestVMHDUsedGB] [float] NULL,
	[VMwareToolsVerion] [int] NULL,
	[VMwareHardwareVersion] [nvarchar](10) NULL,
	[FtInfo] [nvarchar](max) NULL,
	[DiskLocation0] [nvarchar] (max) NULL,
	[DiskLocation1] [nvarchar] (max) NULL,
	[DiskLocation2] [nvarchar] (max) NULL,
	[Network0] [nvarchar](255) NULL,
	[IPv6Address0] [nvarchar](50) NULL,
	[IPv4Address0] [nvarchar](15) NULL,
	[Network1] [nvarchar](255) NULL,
	[IPv6Address1] [nvarchar](50) NULL,
	[IPv4Address1] [nvarchar](15) NULL,
	[Network2] [nvarchar](255) NULL,
	[IPv6Address2] [nvarchar](50) NULL,
	[IPv4Address2] [nvarchar](15) NULL,
	[MAC0] [nvarchar](255) NULL,
	[MAC1] [nvarchar](255) NULL,
	[MAC2] [nvarchar](255) NULL,
	[VLAN0] [int] NULL,
	[VLAN1] [int] NULL,
	[VLAN2] [int] NULL,
	[DataInsertTimeStamp] [datetime] NULL
) ON [PRIMARY]


Print 'VMConversionData_PS'
GOTO VMConversionData_PS_Done

VMConversionData_PS_Done: 
GO

/****** Create Details View   ******/
IF EXISTS(SELECT * from INFORMATION_SCHEMA.VIEWS WHERE table_name = 'VMDetails_VIEW')
Begin
		DROP VIEW [VMDetails_VIEW]
		Print 'VMDetails_VIEW already exists, droppping view.'
END

SET ANSI_NULLS ON
SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[VMDetails_VIEW]
AS
SELECT        dbo.VMConversionData_PS.GuestVMFQDN, dbo.VMConversionData_PS.GuestVMID, dbo.VMConversionData_PS.VMHostID, 
                         dbo.VMConversionData_PS.VMHostName, dbo.VMConversionData_PS.GUESTID, dbo.VMConversionData_PS.GuestOS, dbo.VMQueue.StartTime, 
                         dbo.VMQueue.EndTime, dbo.VMQueue.Notes, dbo.VMQueue.PID, dbo.VMConversionData_PS.GuestVMMB, dbo.VMConversionData_PS.GuestVMCPUCount, 
                         dbo.VMConversionData_PS.GuestVMHDProvisionedGB, dbo.VMConversionData_PS.GuestVMHDUsedGB, dbo.VMConversionData_PS.FtInfo, 
                         dbo.VMConversionData_PS.DiskLocation0, dbo.VMConversionData_PS.DiskLocation1, dbo.VMConversionData_PS.DiskLocation2, 
                         dbo.VMConversionData_PS.Network0, dbo.VMConversionData_PS.IPv6Address0, dbo.VMConversionData_PS.IPv4Address0, 
                         dbo.VMConversionData_PS.Network1, dbo.VMConversionData_PS.IPv6Address1, dbo.VMConversionData_PS.IPv4Address1, 
                         dbo.VMConversionData_PS.Network2, dbo.VMConversionData_PS.IPv6Address2, dbo.VMConversionData_PS.IPv4Address2, 
                         dbo.VMConversionData_PS.DataInsertTimeStamp, dbo.VMQueue.Supported, dbo.VMQueue.ReadytoConvert, dbo.VMQueue.InUse, dbo.VMQueue.Completed, 
                         dbo.VMQueue.Status, dbo.VMQueue.Position, dbo.VMQueue.ConvertServer, dbo.VMConversionData_PS.MoRef, dbo.VMQueue.VMName, dbo.VMQueue.JobID, 
                         dbo.VMConversionData_PS.MAC0, dbo.VMConversionData_PS.MAC1, dbo.VMConversionData_PS.MAC2, dbo.VMConversionData_PS.VLAN0, 
                         dbo.VMConversionData_PS.VLAN1, dbo.VMConversionData_PS.VLAN2, dbo.VMConversionData_PS.DisplayName, dbo.VMQueue.Summary, 
                         dbo.VMQueue.Warning, dbo.VMConversionData_PS.VMwareToolsVerion
FROM            dbo.VMConversionData_PS INNER JOIN
                         dbo.VMQueue ON dbo.VMConversionData_PS.GuestVMFQDN = dbo.VMQueue.VMName
GO


Print 'VMDetails_VIEW created.'

/****** Create sp_TransferVMConversionData_PS  ******/
IF object_id('sp_TransferVMConversionData_PS') IS NOT NULL
Begin
		DROP PROCEDURE [sp_TransferVMConversionData_PS]
		Print 'sp_TransferVMConversionData_PS already exists, droppping stored proccedure.'
END

SET ANSI_NULLS ON
SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[sp_TransferVMConversionData_PS] 

AS
BEGIN

--SET NOCOUNT ON
--INSERT UNIQUE RECORDS (only specific fields) FROM VMConversionData_PS INTO VMQueue
INSERT INTO [VMQueue] ([objectID],[VMName],[Supported],[ReadytoConvert],[Completed],[Status],[Notes])
SELECT vmcd.[MoRef] AS [objectID]
	  ,vmcd.[GuestVMFQDN] AS [VMName]
      ,CASE
			WHEN vmcd.[FtInfo] = '' AND vmcd.[GuestOS] = 'Microsoft Windows Server 2008 R2 (64-bit)' THEN 1
			WHEN vmcd.[FtInfo] = '' AND vmcd.[GuestOS] = 'Microsoft Windows Server 2008 (32-bit)' THEN 1
			WHEN vmcd.[FtInfo] = '' AND vmcd.[GuestOS] = 'Microsoft Windows Server 2008 (64-bit)' THEN 1
			WHEN vmcd.[FtInfo] = '' AND vmcd.[GuestOS] = 'Microsoft Windows XP Professional (32-bit)' THEN 3
			WHEN vmcd.[FtInfo] = '' AND vmcd.[GuestOS] = 'Microsoft Windows 7 (32-bit)' THEN 1
			WHEN vmcd.[FtInfo] = '' AND vmcd.[GuestOS] = 'Microsoft Windows 7 (64-bit)' THEN 1
			WHEN vmcd.[FtInfo] = '' AND vmcd.[GuestOS] = 'Microsoft Windows 8 (32-bit)' THEN 3
			WHEN vmcd.[FtInfo] = '' AND vmcd.[GuestOS] = 'Microsoft Windows 8 (64-bit)' THEN 3
			WHEN vmcd.[FtInfo] = '' AND vmcd.[GuestOS] = 'Microsoft Windows Vista (32-bit)' THEN 1
			WHEN vmcd.[FtInfo] = '' AND vmcd.[GuestOS] = 'Microsoft Windows Vista (64-bit)' THEN 1
			WHEN vmcd.[FtInfo] = '' AND vmcd.[GuestOS] = 'Microsoft Windows Server 2003 Standard (32-bit)' THEN 1
			WHEN vmcd.[FtInfo] = '' AND vmcd.[GuestOS] = 'Microsoft Windows Server 2003 Standard (64-bit)' THEN 1
			WHEN vmcd.[FtInfo] = '' AND vmcd.[GuestOS] = 'Microsoft Windows Server 2003, Enterprise Edition (32-bit)' THEN 1
			WHEN vmcd.[FtInfo] = '' AND vmcd.[GuestOS] = 'Microsoft Windows Server 2003, Enterprise Edition (64-bit)' THEN 1
			WHEN vmcd.[FtInfo] = '' AND vmcd.[GuestOS] = 'Microsoft Windows Server 2003, Datacenter Edition (32-bit)' THEN 1
			WHEN vmcd.[FtInfo] = '' AND vmcd.[GuestOS] = 'Microsoft Windows Server 2003, Datacenter Edition (64-bit)' THEN 1
			WHEN vmcd.[FtInfo] = '' AND vmcd.[GuestOS] = 'Microsoft Windows Server 2003, Standard Edition (32-bit)' THEN 1
			WHEN vmcd.[FtInfo] = '' AND vmcd.[GuestOS] = 'Microsoft Windows Server 2003, Standard Edition (64-bit)' THEN 1
			WHEN vmcd.[FtInfo] = '' AND vmcd.[GuestOS] = 'Microsoft Windows Server 2003 (32-bit)' THEN 1
			WHEN vmcd.[FtInfo] = '' AND vmcd.[GuestOS] = 'Microsoft Windows Server 2003, Web Edition' THEN 1
			WHEN vmcd.[FtInfo] = '' AND vmcd.[GuestOS] = 'Microsoft Windows Small Business Server 2003' THEN 1
			ELSE 3
	   END AS [Supported]
	   ,0 AS [ReadytoConvert]
	   ,0 AS [Completed]
	   ,0 AS [Status]
	   ,'VM automatically collected by script' AS [Notes]
  FROM [VMConversionData_PS] vmcd
			LEFT JOIN
	   [VMQueue] vmq
			ON vmcd.[MoRef] = vmq.[objectID]
WHERE vmq.[objectID] IS NULL

SET NOCOUNT OFF

END

GO
Print 'Stored Proceedure sp_TransferVMConversionData_PS created.'


/****** Done with MAT   ******/
/****** MAT Database Changes Complete  ******/
Print 'MAT Database Setup Complete.'
