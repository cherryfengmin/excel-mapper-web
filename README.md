# Excel 翻译映射工具（纯前端）

一个纯前端网页工具：上传 `.xlsx`，选择两列生成翻译映射表，并支持把粘贴的 JSON（默认处理 `settings`）按映射关系自动替换为目标语言文案。

## 功能

### 1) Excel → 映射表
- 上传 `.xlsx`（不上传服务器）
- 选择工作表、起始行、源列/目标列
- 生成映射记录并在表格中筛选/搜索（`OK / NEED_REVIEW / UNMATCHED / ERROR`）
- 一键复制当前筛选结果为：`"源文案"：“目标文案”`

#### 分段/切分规则（摘要）
- **Block 分段**：按空行分块（支持 `\\n\\n` 以及 `\\n + 空白 + \\n`）
- **Item 切分**（优先级）：\n  1) 行首条目标记（例如 `* / • / - / [Icon] / [Symbol]`）\n  2) inline `*`（文本中出现多个 `*` 的情况）\n  3) 句子兜底拆分（会标 `NEED_REVIEW`）

> 说明：浏览器端解析通常无法稳定读取单元格富文本（颜色/删除线等），当前以文本与真实换行为准。

### 2) JSON 智能替换（默认只处理 `settings`）
- 粘贴 JSON，点击 **应用映射** 生成替换后的 JSON
- 支持选择替换方向：\n  - 自动判断方向\n  - 源列 → 目标列\n  - 目标列 → 源列
- 替换统计：`replaced / unchanged / skipped / unmatched / review`，可展开查看未匹配示例

#### 智能匹配策略（摘要）
- 精确匹配 / 规范化匹配（统一换行、trim、Unicode NFKC、去零宽字符、统一连字符等）
- HTML 实体解码匹配
- 去 HTML 标签匹配（用于命中带标签的文本）
- 大小写不敏感匹配
- 多行优化：多行文本按行建立额外映射 key
- HTML 标签保护：`<br> / <span ...> / <nobr>` 等标签在替换时保留结构，仅替换文本 token
- 规格参数拆分：`Label: Value` / `Label：Value` 仅替换 label，value 默认不翻译
- 技术字段跳过：URL、资源路径、颜色、数字/布尔、对齐关键字、短 CSS 标识符等

> 说明：若 JSON 根对象没有 `settings` 字段，会自动把整个 JSON 当作 settings 递归替换（便于直接粘贴片段测试）。

## 本地运行

```bash
cd excel-mapper-web
npm install
npm run dev
```

## 构建

```bash
npm run build
npm run preview
```

## 常见问题

### 为什么点“应用映射”后 JSON 没变化？
- 可能是替换方向不对（建议用“自动判断方向”）
- JSON 中待替换的字符串在映射表中不存在或存在不可见字符差异（可查看“未匹配示例”）
- 字段不在 `settings` 内（或根对象没有 `settings` 时未开启根替换；当前默认已开启）

### 为什么下拉选项/系统控件字体颜色不一致？
- 原生 `<select>` 受浏览器/系统主题影响，样式可控性有限；如需完全统一可改为自定义下拉组件。

# React + TypeScript + Vite

This template provides a minimal setup to get React working in Vite with HMR and some ESLint rules.

Currently, two official plugins are available:

- [@vitejs/plugin-react](https://github.com/vitejs/vite-plugin-react/blob/main/packages/plugin-react) uses [Oxc](https://oxc.rs)
- [@vitejs/plugin-react-swc](https://github.com/vitejs/vite-plugin-react/blob/main/packages/plugin-react-swc) uses [SWC](https://swc.rs/)

## React Compiler

The React Compiler is not enabled on this template because of its impact on dev & build performances. To add it, see [this documentation](https://react.dev/learn/react-compiler/installation).

## Expanding the ESLint configuration

If you are developing a production application, we recommend updating the configuration to enable type-aware lint rules:

```js
export default defineConfig([
  globalIgnores(['dist']),
  {
    files: ['**/*.{ts,tsx}'],
    extends: [
      // Other configs...

      // Remove tseslint.configs.recommended and replace with this
      tseslint.configs.recommendedTypeChecked,
      // Alternatively, use this for stricter rules
      tseslint.configs.strictTypeChecked,
      // Optionally, add this for stylistic rules
      tseslint.configs.stylisticTypeChecked,

      // Other configs...
    ],
    languageOptions: {
      parserOptions: {
        project: ['./tsconfig.node.json', './tsconfig.app.json'],
        tsconfigRootDir: import.meta.dirname,
      },
      // other options...
    },
  },
])
```

You can also install [eslint-plugin-react-x](https://github.com/Rel1cx/eslint-react/tree/main/packages/plugins/eslint-plugin-react-x) and [eslint-plugin-react-dom](https://github.com/Rel1cx/eslint-react/tree/main/packages/plugins/eslint-plugin-react-dom) for React-specific lint rules:

```js
// eslint.config.js
import reactX from 'eslint-plugin-react-x'
import reactDom from 'eslint-plugin-react-dom'

export default defineConfig([
  globalIgnores(['dist']),
  {
    files: ['**/*.{ts,tsx}'],
    extends: [
      // Other configs...
      // Enable lint rules for React
      reactX.configs['recommended-typescript'],
      // Enable lint rules for React DOM
      reactDom.configs.recommended,
    ],
    languageOptions: {
      parserOptions: {
        project: ['./tsconfig.node.json', './tsconfig.app.json'],
        tsconfigRootDir: import.meta.dirname,
      },
      // other options...
    },
  },
])
```
