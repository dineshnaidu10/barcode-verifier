-- OPTIONAL: dummy rows so the Records viewer has something to show.
-- Run after schema.sql. Safe to delete afterwards:  DELETE FROM dbo.scan_records WHERE gr_id = 'GR-DEMO';
USE IQCScanner;
GO
INSERT INTO dbo.scan_records
(project_id, project_name, gr_id, barcode, panel_sn, array_index, [timestamp], operator, station, result, failure_reason)
VALUES
('DEMO-001','Demo Main Board (structured)','GR-DEMO','MA2641000001',NULL,NULL,'2026-10-05T09:12:01','alice','IQC-01','PASS',NULL),
('DEMO-001','Demo Main Board (structured)','GR-DEMO','MA2641000002',NULL,NULL,'2026-10-05T09:12:09','alice','IQC-01','PASS',NULL),
('DEMO-001','Demo Main Board (structured)','GR-DEMO','MA2541000003',NULL,NULL,'2026-10-05T09:12:20','alice','IQC-01','FAIL','Year mismatch — barcode ''25'', expected ''26'''),
('DEMO-002','Demo Board (regex)','GR-DEMO','ABC-1234-X9Z2',NULL,NULL,'2026-10-06T14:30:45','bob','IQC-01','PASS',NULL),
('DEMO-002','Demo Board (regex)','GR-DEMO','ABC-123-X9Z2',NULL,NULL,'2026-10-06T14:31:02','bob','IQC-01','FAIL','Format mismatch — expected: 3 letters - 4 digits - 4 alphanumerics'),
('DEMO-004','Demo Panel 2x2','GR-DEMO','ARR2641000001','PNL0000000000001',1,'2026-10-07T10:00:10','alice','IQC-01','PASS',NULL),
('DEMO-004','Demo Panel 2x2','GR-DEMO','ARR2641000002','PNL0000000000001',2,'2026-10-07T10:00:18','alice','IQC-01','PASS',NULL),
('DEMO-004','Demo Panel 2x2','GR-DEMO','ARR2641000003','PNL0000000000001',3,'2026-10-07T10:00:25','alice','IQC-01','PASS',NULL),
('DEMO-004','Demo Panel 2x2','GR-DEMO','ARR2641000004','PNL0000000000001',4,'2026-10-07T10:00:31','alice','IQC-01','PASS',NULL);
GO
