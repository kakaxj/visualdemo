# Macro Dashboard — reorganized

## 当前目录

```text
macro_dashboard_reorganized/
├── index.html
├── tab1.html
├── tab2.html
├── css/
│   ├── style.css
│   ├── tab1.css
│   └── tab2.css
└── js/
    ├── config.js
    ├── index.js
    ├── tab1.js
    └── tab2.js
```

## 结构原则

- HTML：只负责页面结构。
- CSS：全部独立存放，不再把大段 `<style>` 写进 tab1 / tab2。
- JS：全部独立存放，不再把业务脚本写进 HTML。
- `js/config.js`：统一维护 Excel 数据源地址。
- `index.js`：首页导航、数据 Tab、首页摘要卡读取。
- `tab1.js`：ICHI 页面数据读取、计算、图表。
- `tab2.js`：财政观测页面数据读取、计算、图表。

## GitHub Excel 数据源

默认配置仍使用：

```js
ICHI_EXCEL_URL: "ichi data.xlsx"
FISCAL_EXCEL_URL: "fiscal data.xlsx"
```

如果 Excel 已放到其他 GitHub 仓库，把 `js/config.js` 中两个地址替换成 Raw URL 即可。
其他 HTML / JS 文件不需要再修改。

## 已处理

1. 原页面中的机构名称已统一显示为“XXXX”。
2. 修正 `tab2.html` 顶部导航 active 状态：财政观测为当前页。
3. 修正首页免责声明文件名为 `disclaimer.html`。
4. 首页数据卡改为动态读取 ICHI / fiscal Excel 最新值。
5. 首页、tab1、tab2 均使用独立 JS。
6. tab1、tab2 的内嵌 CSS 已移出 HTML。
7. 保留 tab3–5 的现有导航入口，但本次未创建或修改对应页面。
