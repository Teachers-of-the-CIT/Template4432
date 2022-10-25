
-- --------------------------------------------------
-- Entity Designer DDL Script for SQL Server 2005, 2008, 2012 and Azure
-- --------------------------------------------------
-- Date Created: 10/24/2022 13:11:47
-- Generated from EDMX file: D:\Документы\Шарага\ИСРПО\Лабораторная работа 2\Template4432\Template4432\Model1.edmx
-- --------------------------------------------------

SET QUOTED_IDENTIFIER OFF;
GO
USE [D:\ДОКУМЕНТЫ\ШАРАГА\ИСРПО\ЛАБОРАТОРНАЯ РАБОТА 2\TEMPLATE4432\TEMPLATE4432\BIN\DEBUG\DB.MDF];
GO
IF SCHEMA_ID(N'dbo') IS NULL EXECUTE(N'CREATE SCHEMA [dbo]');
GO

-- --------------------------------------------------
-- Dropping existing FOREIGN KEY constraints
-- --------------------------------------------------


-- --------------------------------------------------
-- Dropping existing tables
-- --------------------------------------------------

IF OBJECT_ID(N'[dbo].[OrderSet]', 'U') IS NOT NULL
    DROP TABLE [dbo].[OrderSet];
GO

-- --------------------------------------------------
-- Creating all tables
-- --------------------------------------------------

-- Creating table 'OrderSet'
CREATE TABLE [dbo].[OrderSet] (
    [ID] int IDENTITY(1,1) NOT NULL,
    [Code] nvarchar(max)  NOT NULL,
    [Date] datetime  NOT NULL,
    [Time] datetime  NOT NULL,
    [ClientCode] int  NOT NULL,
    [Services] nvarchar(max)  NOT NULL,
    [Status] nvarchar(max)  NOT NULL,
    [ClosingDate] datetime  NULL,
    [RentTime] nvarchar(max)  NOT NULL
);
GO

-- --------------------------------------------------
-- Creating all PRIMARY KEY constraints
-- --------------------------------------------------

-- Creating primary key on [ID] in table 'OrderSet'
ALTER TABLE [dbo].[OrderSet]
ADD CONSTRAINT [PK_OrderSet]
    PRIMARY KEY CLUSTERED ([ID] ASC);
GO

-- --------------------------------------------------
-- Creating all FOREIGN KEY constraints
-- --------------------------------------------------

-- --------------------------------------------------
-- Script has ended
-- --------------------------------------------------