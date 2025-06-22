# CaseManagementSystem

## 项目信息

- **项目名称**: CaseManagementSystem
- **开发语言**: VB.NET
- **框架**: .NET Framework 4.8 (WinForms)
- **数据库**: Microsoft Access
- **架构**: 三层架构 (UI/BLL/DAL)

## 项目结构

```
CaseManagementSystem/
├── BusinessLogic/          # 业务逻辑层
├── DataAccess/            # 数据访问层
├── Forms/                 # 界面层
├── Models/                # 数据模型
├── Controls/              # 自定义控件
├── Interfaces/            # 接口定义
├── Enums/                 # 枚举定义
├── Constants/             # 常量定义
├── Utils/                 # 工具类
├── Doc/                   # 项目文档
├── UIDesign/              # UI设计文件
└── README.md              # 项目说明
```

## 主要功能

1. **用户管理**: 登录、密码变更
2. **案件管理**: 新建、编辑、查询、审查
3. **数据录入**: 多标签页详细录入
4. **权限控制**: 多角色权限管理
5. **数据导出**: 案件数据导出功能

## 技术特点

- 三层架构设计
- 工厂模式实现Tab模板
- 参数化SQL查询
- 完善的异常处理
- Bootstrap风格UI

## 版本信息

- **当前版本**: 1.0.0
- **最后更新**: 2024年
- **开发状态**: 开发中 