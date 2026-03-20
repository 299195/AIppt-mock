const API_BASE = import.meta.env.VITE_API_BASE ?? "http://127.0.0.1:8000/api";
const FILE_BASE = import.meta.env.VITE_FILE_BASE ?? "http://127.0.0.1:8000";
export const fileUrl = (url) => (url ? `${FILE_BASE}${url}` : "");
async function readError(res, fallback) {
    try {
        const data = (await res.json());
        throw new Error(data.detail || fallback);
    }
    catch {
        throw new Error(fallback);
    }
}
export async function getModelConfig() {
    const res = await fetch(`${API_BASE}/model/config`);
    if (!res.ok)
        return readError(res, `模型配置获取失败: HTTP ${res.status}`);
    return (await res.json());
}
export async function getTemplates() {
    const res = await fetch(`${API_BASE}/templates`);
    if (!res.ok)
        return readError(res, `模板列表获取失败: HTTP ${res.status}`);
    return (await res.json());
}
export async function previewOutline(payload) {
    const res = await fetch(`${API_BASE}/outline/preview`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
    });
    if (!res.ok)
        return readError(res, `大纲预览失败: HTTP ${res.status}`);
    return (await res.json());
}
export async function createJob(payload) {
    const res = await fetch(`${API_BASE}/jobs`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
    });
    if (!res.ok)
        return readError(res, `创建任务失败: HTTP ${res.status}`);
    return (await res.json());
}
export async function getJob(jobId) {
    const res = await fetch(`${API_BASE}/jobs/${jobId}`);
    if (!res.ok)
        return readError(res, `查询任务失败: HTTP ${res.status}`);
    return (await res.json());
}
export async function rewriteJob(jobId, action) {
    const res = await fetch(`${API_BASE}/jobs/${jobId}/rewrite`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action }),
    });
    if (!res.ok)
        return readError(res, `重写失败: HTTP ${res.status}`);
    return (await res.json());
}
export async function getHistory() {
    const res = await fetch(`${API_BASE}/jobs`);
    if (!res.ok)
        return readError(res, `获取历史失败: HTTP ${res.status}`);
    return (await res.json());
}
export async function parseUpload(file) {
    const fd = new FormData();
    fd.append("file", file);
    const res = await fetch(`${API_BASE}/parse-upload`, {
        method: "POST",
        body: fd,
    });
    if (!res.ok)
        return readError(res, `文件解析失败: HTTP ${res.status}`);
    const payload = (await res.json());
    return payload.extracted_text;
}
