-- IQC Barcode Scanner - SQL Server schema
-- Run once in SSMS / Azure Data Studio / sqlcmd.
-- Column ORDER matters: the Records viewer reads SELECT * by position.

IF DB_ID('IQCScanner') IS NULL CREATE DATABASE IQCScanner;
GO
USE IQCScanner;
GO

CREATE TABLE dbo.scan_records (
    id             INT IDENTITY(1,1) PRIMARY KEY,
    project_id     NVARCHAR(50)   NOT NULL,
    project_name   NVARCHAR(200)  NOT NULL,
    gr_id          NVARCHAR(100)  NOT NULL,
    barcode        NVARCHAR(200)  NOT NULL,
    panel_sn       NVARCHAR(200)  NULL,      -- NULL in single-barcode mode
    array_index    INT            NULL,      -- 1-based slot in the panel
    [timestamp]    NVARCHAR(25)   NOT NULL,  -- ISO-8601 text, e.g. 2026-10-10T14:03:22
    operator       NVARCHAR(100)  NOT NULL,  -- empty string if operator not required
    station        NVARCHAR(50)   NOT NULL,
    result         NVARCHAR(10)   NOT NULL CHECK (result IN ('PASS','FAIL')),
    failure_reason NVARCHAR(500)  NULL
);
GO

-- Speeds up duplicate detection (barcode + project + PASS)
CREATE INDEX IX_scan_records_dup ON dbo.scan_records (barcode, project_id, result);
GO
