# AI 生成 PPT 助手（主题 + 资料直出 PPTX）

输入主题和资料（`.md` / `.docx` / 文本），自动生成大纲、扩写页内容并导出 `PPTX`。

## 当前能力

- 主题 + 资料生成大纲
- 基于大纲生成逐页内容并导出 PPTX
- 风格切换：管理版 / 技术版
- 模板选择 + 历史记录回看

## 技术栈

- 后端：Python + FastAPI + python-docx + python-pptx
- 前端：React + Vite + TypeScript
- 导出：Bridge Export（`export_engines/pptx_export_bridge`）

## 项目结构

- `backend/`：API、任务流、模型调用、模板与导出
- `frontend/`：页面交互与任务发起
- `export_engines/pptx_export_bridge/`：PPTX 导出桥接引擎（PHP）
- `start_backend.bat`：后端启动脚本
- `start_frontend.bat`：前端启动脚本

## 快速启动（Windows，推荐）

### 1) 启动后端

```powershell
.\start_backend.bat 8002
```

- 端口参数可省略，默认 `8001`
- 脚本会优先使用：
  - `backend\.venv311\Scripts\python.exe`
  - `backend\.venv\Scripts\python.exe`

### 2) 启动前端

```powershell
.\start_frontend.bat 5173 8002
```

- 第一个参数：前端端口（默认 `5173`）
- 第二个参数：后端端口（默认读取 `.backend_port`，没有则 `8001`）
- 脚本会检查 `node_modules\.bin\vite.cmd`，缺失时自动执行 `npm install`

## 手动启动（可选）

### 后端

```powershell
cd backend
pip install -r requirements.txt
set PPT_EXPORT_ENGINE=bridge
uvicorn app.main:app --reload --host 127.0.0.1 --port 8001
```

### 前端

```powershell
cd frontend
npm install
npm run dev -- --host 127.0.0.1 --port 5173
```

## 模型配置

编辑 `backend/model_provider.json`：

```json
{
  "provider": "doubao",
  "use_mock": false,
  "base_url": "https://ark.cn-beijing.volces.com/api",
  "api_key": "YOUR_API_KEY",
  "endpoint_id": "YOUR_ENDPOINT_ID",
  "model": "YOUR_MODEL_NAME",
  "chat_path": "/v3/chat/completions",
  "timeout": 60
}
```

可用环境变量覆盖：

- `MODEL_PROVIDER`
- `MODEL_BASE_URL`
- `MODEL_API_KEY`
- `MODEL_ENDPOINT_ID`
- `MODEL_NAME`
- `MODEL_CHAT_PATH`
- `MODEL_TIMEOUT`
- `MODEL_STREAM_TIMEOUT`
- `USE_MOCK_LLM`

## 导出引擎配置（Bridge Export）

- 默认：`PPT_EXPORT_ENGINE=bridge`
- 可选：`bridge | auto | legacy`
- 需要 `PHP >= 7.4` 且启用 `zip` 扩展
- 可用 `AIPPT_PHP_BIN` 指定 PHP 路径：

```powershell
set AIPPT_PHP_BIN=C:\path\to\php.exe
```

## 访问地址

- 前端：`http://127.0.0.1:5173`
- 后端健康检查：`http://127.0.0.1:8001/health`
- 后端 API 前缀：`/api`
- 导出文件静态路径：`/exports/...`
- 模板预览静态路径：`/assets/...`

说明：后端根路径 `/` 返回 `404` 是正常现象，接口都在 `/api` 下。

## 常见问题

- 报错 `'vite' 不是内部或外部命令`：
  - 重新执行 `./start_frontend.bat`，脚本会自动补装依赖
- npm 提示 `Unknown user config`：
  - 来自用户级 `.npmrc` 的拼写或配置问题，通常不影响运行

## 安全建议

- 不要把真实 `API Key` 提交到仓库
- 建议通过环境变量注入敏感配置
