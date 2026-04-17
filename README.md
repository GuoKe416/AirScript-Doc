# AirScript 文档

<p align="center">
  <img src="https://trae-api-cn.mchost.guru/api/ide/v1/text_to_image?prompt=AirScript%20logo%20with%20cloud%20and%20code%20symbols%2C%20modern%20clean%20design%2C%20blue%20and%20white%20color%20scheme&image_size=square_hd" alt="AirScript Logo" width="200">
</p>

<p align="center">
  <a href="#">
    <img src="https://img.shields.io/badge/version-2.0-blue.svg" alt="Version">
  </a>
  <a href="#">
    <img src="https://img.shields.io/badge/license-MIT-green.svg" alt="License">
  </a>
  <a href="#">
    <img src="https://img.shields.io/badge/language-JavaScript-yellow.svg" alt="Language">
  </a>
</p>

## 项目简介

**AirScript** 是金山文档推出的一个简单快速的轻量级脚本应用开发平台，基于云技术构建，可让您快速轻松地创建与金山文档 Office 文件交互的业务应用。
- **官方网站**：[金山文档](https://www.kdocs.cn)
- **文档地址**：[AirScript 文档](https://www.kdocs.cn/airscript/docs)
- **项目用途**：本项目仅做为将官方文档转换为离线Markdown使用，非实时更新，建议前往官方文档参考最新教程

## 核心功能

- **云开发环境**：无需搭建本地环境，直接在文档内进行脚本云开发
- **智能编辑器**：内置定制化的全局 Application 对象，编辑器智能提示，开发、调试、运行一条龙服务
- **同步执行**：同步获取属性，同步执行方法，减少传统的异步调用带来的心智负担
- **强大 API**：内置强大的组件（表格）API，支持对表格数据的增删查改、单元格式修改、属性设置等能力
- **多场景支持**：无论是创建定时任务，还是批量处理数据，亦或是自动化生成文档，开发者可以在这里尽情发挥自己的想象力

## 技术特点

- **语言支持**：采用标准 JavaScript 语言进行编写，支持大部分 ES6 语法
- **内置对象**：提供 Application（文档 OpenApi 对象）和 Context（脚本运行上下文对象）
- **安全运行**：代码将会安全地运行在服务端上
- **多文档类型**：支持在线表格、多维表等多种金山文档类型

## 目录结构

```
AirScript脚本文档/
├── 产品介绍/
│   ├── 概述.md
│   └── 脚本语言.md
├── 快速上手/
│   ├── 开始.md
│   ├── 最佳实践.md
│   └── 配置视图.md
├── 脚本令牌/
│   ├── 应用场景.md
│   ├── 接口说明.md
│   └── 简介.md
├── 示范案例/
│   ├── 多维表.md
│   └── 表格.md
├── API文档(1.0)/
│   ├── 多维表格/
│   ├── 智能表格/
│   ├── 高级服务/
│   └── 内置函数.md
├── API文档(2.0)/
│   ├── 智能表格/
│   └── 高级服务/
├── AirScript完整文档.md
├── combine_md.py
└── README.md
```

## 快速开始

### 1. 打开 AirScript 编辑器

1. 打开在线表
2. 切换至「效率」Tab
3. 在下方二级工具栏找到「AirScript 编辑工具」
4. 点击即可调起 AirScript 编辑器

> 提示：有文件编辑权限的协作者才能打开开发工具

### 2. 编写第一个脚本

```javascript
// 示例：在表格中插入当前时间
function main() {
  // 获取当前工作表
  let sheet = Application.ActiveSheet;
  
  // 在 A1 单元格插入当前时间
  sheet.Range("A1").Value = new Date().toLocaleString();
  
  // 打印结果
  console.log("脚本执行完成，已在 A1 单元格插入当前时间");
}

// 执行主函数
main();
```

### 3. 运行脚本

在编辑器中点击「运行」按钮，或使用快捷键 `F5` 运行脚本。执行结果会在日志面板中显示。

## 语法支持

### 支持的语法

- 大部分 ES6 语法
- 箭头函数
- 模板字符串
- 解构赋值
- 扩展运算符
- 模块化语法（部分）

### 不支持的语法

- Class 语法
- Object 里直接定义方法
- import、export
- 可选链操作符 (?.)
- await
- yield

## API 文档

AirScript 提供了丰富的 API 文档，包括：

- **API文档(1.0)**：包含多维表格、智能表格、高级服务和内置函数的 API 文档
- **API文档(2.0)**：包含智能表格和高级服务的 API 文档，提供了更丰富的对象和方法

详细 API 文档请参考 `API文档(1.0)` 和 `API文档(2.0)` 目录。

## 脚本令牌

AirScript 提供了脚本令牌功能，允许您通过 HTTP 接口执行脚本。支持同步执行和异步执行两种方式：

- **同步执行**：接口调用后会直接返回执行结果，适用于执行耗时一般的场景
- **异步执行**：接口调用后返回一个 task_id，需要根据此 task_id 轮询脚本执行的日志，适用于执行耗时比较大的场景

详细使用方法请参考 `脚本令牌` 目录。

## 示范案例

本项目提供了一些示范案例，展示了 AirScript 的实际应用场景：

- **多维表**：展示如何使用 AirScript 操作多维表
- **表格**：展示如何使用 AirScript 操作表格

详细案例请参考 `示范案例` 目录。

## 最佳实践

- 使用模块化思想组织代码
- 合理使用 try-catch 处理异常
- 避免在循环中执行大量操作
- 使用 console.log 进行调试
- 定期保存代码，避免意外丢失

详细最佳实践请参考 `快速上手/最佳实践.md`。

## 常见问题

### Q: 脚本执行失败怎么办？

A: 查看日志面板中的错误信息，定位错误位置并修复。

### Q: 如何访问表格数据？

A: 使用 `Application.ActiveSheet` 获取当前工作表，然后使用 `Range` 方法获取单元格或区域。

### Q: 脚本执行超时怎么办？

A: 对于耗时较长的操作，建议使用异步执行方式，或优化代码结构，减少执行时间。

### Q: 如何调用外部 API？

A: 可以使用内置的网络 API 进行 HTTP 请求，详细使用方法请参考 `API文档(1.0)/高级服务/网络 API.md`。

## 贡献指南

欢迎对本项目进行贡献！如果您有任何建议或问题，请：

1. Fork 本仓库
2. 创建您的特性分支 (`git checkout -b feature/amazing-feature`)
3. 提交您的更改 (`git commit -m 'Add some amazing feature'`)
4. 推送到分支 (`git push origin feature/amazing-feature`)
5. 打开一个 Pull Request

---

