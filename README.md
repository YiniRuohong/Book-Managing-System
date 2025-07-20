# 图书管理系统 (Book Management System)

一个基于 Python 和 Tkinter 开发的智能图书管理系统，集成了二维码识别、IC卡识别等现代化管理功能。

## 功能特性

### 基础功能
- 📚 **图书管理**: 图书信息的增删改查
- 👥 **读者管理**: 读者信息管理和维护
- 📖 **借阅功能**: 图书借阅与归还流程
- 📊 **统计报表**: 借阅统计和数据分析
- 🔍 **智能搜索**: 多条件搜索功能

### 高级功能
- 🤖 **AI 助手**: 智能客服对话功能
- 📷 **图像识别**: 图书封面识别
- 💳 **IC卡读取**: 通过 PN532 IC卡读卡器
- 📊 **数据统计**: 数据可视化分析
- 🔒 **权限管理**: 多级用户权限控制

## 技术栈

- **界面框架**: CustomTkinter (现代化UI)
- **数据处理**: openpyxl, pandas
- **图像处理**: OpenCV, PIL
- **二维码**: qrcode
- **硬件通信**: pyserial, adafruit-pn532
- **AI对话**: 百度千帆 API (可选)
- **数据可视化**: QuickChart API

## 系统要求

### 环境要求
- Python 3.12+
- 操作系统: Windows/macOS/Linux

### 安装步骤

1. **克隆项目**
```bash
git clone https://github.com/YiniRuohong/Book-Managing-System.git
cd Book-Managing-System
```

2. **安装依赖**
```bash
# 使用 uv (推荐)
uv install

# 或使用 pip
pip install -r requirements.txt
```

3. **配置数据库** (可选)
```sql
-- 创建 MySQL 数据库
CREATE DATABASE library;
-- 配置相应的用户权限
```

4. **配置 API 密钥**
在 `library_management_system/library_system.py` 中配置百度 API 密钥:
```python
api_key = "YOUR_API_KEY_HERE"  # 替换为实际的API密钥
```

## 使用指南

### 启动程序
```bash
python main.py
```

### 默认账户
- 管理员账号: `admin`
- 管理员密码: `admin123`

### 操作说明

#### 图书搜索
1. 支持多种搜索方式
2. 正则表达式搜索
   - 书名模糊: 输入关键词
   - 正则模式: `/pattern/`
   - 组合搜索: `关键词1 AND 关键词2`
   - 排除搜索: `+关键词 NOT 排除词`

#### 图书借阅
1. 选择读者和图书
2. 确认借阅信息
3. 扫描条形码或IC卡确认
4. 系统自动记录借阅信息

#### 图书归还
1. 查找借阅记录
2. 确认归还图书
3. 系统更新图书状态
4. 生成归还凭证(可选)

## 系统配置

### 数据库结构
```
library_management_system/
├── book.xlsx          # 图书信息
├── name.xlsx          # 读者信息  
├── borrow.xlsx        # 当前借阅记录
├── borrow_history.csv # 历史借阅记录
├── log.txt           # 系统日志
└── usrs_info.pickle  # 用户登录信息
```

### Excel 文件格式

**book.xlsx** - 图书信息表:
| 书名 | 作者 | 出版社 | ISBN号 | 类别 | 借阅状态 | 位置信息 |
|------|------|--------|--------|------|----------|----------|

**name.xlsx** - 读者信息表:
| 读者姓名 | 学号工号 | 联系方式 |
|----------|----------|----------|

### 硬件配置 (可选)

#### IC卡读卡器
- 支持 PN532 芯片的读卡器
- 通过 USB 串口连接
- 自动识别连接

#### 条码扫描器
- 支持标准条码扫描
- 通过 USB HID 连接

## 系统截图

系统提供友好的图形化界面:
- 主控制面板
- 图书管理界面
- AI 智能助手
- 数据统计图表

## 项目结构

### 目录结构
```
book_managing/
├── main.py                          # 程序入口
├── library_management_system/       # 核心模块
│   ├── library_system.py          # 业务逻辑
│   ├── windowui.py                 # UI界面组件
│   └── [其他模块]
├── pyproject.toml                   # 项目配置
└── README.md                        # 项目说明
```

### 代码规范
- 遵循 PEP 8 Python 编码规范
- 使用类型注解
- 完善的异常处理
- 模块化设计理念

## 贡献指南

欢迎提交 Issue 和 Pull Request！

### 贡献流程
1. Fork 本项目
2. 创建特性分支: `git checkout -b feature-name`
3. 提交修改: `git commit -m 'Add feature'`
4. 推送分支: `git push origin feature-name`
5. 提交 Pull Request

## 开源协议

本项目采用 MIT 开源协议 - 查看 [LICENSE](LICENSE) 了解详情

## 致谢

- [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) - 现代 Tkinter 框架
- [OpenCV](https://opencv.org/) - 计算机视觉库
- [Adafruit](https://www.adafruit.com/) - 硬件支持库

## 技术支持

如有问题请通过以下方式联系:
- 提交 [GitHub Issue](https://github.com/YiniRuohong/Book-Managing-System/issues)
- 发送邮件至项目维护者

---

**如果这个项目对你有帮助，请给个 Star ⭐**