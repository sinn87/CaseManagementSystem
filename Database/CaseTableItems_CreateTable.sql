-- 创建CaseTableItems表，用于存储DataGridView的表格数据
CREATE TABLE CaseTableItems (
    ItemID AUTOINCREMENT PRIMARY KEY,
    CaseID INTEGER NOT NULL,
    TabIndex INTEGER NOT NULL,
    DataGridViewName VARCHAR(255),
    DataGridViewIndex VARCHAR(50),
    ItemName VARCHAR(255) NOT NULL,
    ItemValue TEXT,
    LastUpdate DATETIME NOT NULL,
    ReviewTime DATETIME NOT NULL,
    Status VARCHAR(50) NOT NULL,
    Reviewer VARCHAR(100) NOT NULL,
    FOREIGN KEY (CaseID) REFERENCES Cases(CaseID) ON DELETE CASCADE
);

-- 创建索引以提高查询性能
CREATE INDEX idx_CaseTableItems_CaseID ON CaseTableItems(CaseID);
CREATE INDEX idx_CaseTableItems_TabIndex ON CaseTableItems(TabIndex);
CREATE INDEX idx_CaseTableItems_Status ON CaseTableItems(Status); 