import type { HistoryItem, JobDetail, ModelConfig, RewriteAction, StyleType, TemplateId, TemplateItem } from "./types";

const API_BASE = import.meta.env.VITE_API_BASE ?? "http://127.0.0.1:8000/api";
const FILE_BASE = import.meta.env.VITE_FILE_BASE ?? "http://127.0.0.1:8000";

export const fileUrl = (url: string | null): string => (url ? `${FILE_BASE}${url}` : "");

async function readError(res: Response, fallback: string): Promise<never> {
  try {
    const data = (await res.json()) as { detail?: string };
    throw new Error(data.detail || fallback);
  } catch {
    throw new Error(fallback);
  }
}

export async function getModelConfig(): Promise<ModelConfig> {
  const res = await fetch(`${API_BASE}/model/config`);
  if (!res.ok) return readError(res, `模型配置获取失败: HTTP ${res.status}`);
  return (await res.json()) as ModelConfig;
}

export async function getTemplates(): Promise<TemplateItem[]> {
  const res = await fetch(`${API_BASE}/templates`);
  if (!res.ok) return readError(res, `模板列表获取失败: HTTP ${res.status}`);
  return (await res.json()) as TemplateItem[];
}

export async function previewOutline(payload: {
  title: string;
  material_text: string;
  outline_text: string;
  style: StyleType;
  target_pages: number;
}) {
  const res = await fetch(`${API_BASE}/outline/preview`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  if (!res.ok) return readError(res, `大纲预览失败: HTTP ${res.status}`);
  return (await res.json()) as { outline: string[] };
}

export async function createJob(payload: {
  title: string;
  material_text: string;
  outline_text: string;
  outline: string[];
  style: StyleType;
  template_id: string;
  target_pages: number;
}) {
  const res = await fetch(`${API_BASE}/jobs`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  if (!res.ok) return readError(res, `创建任务失败: HTTP ${res.status}`);
  return (await res.json()) as { job_id: string };
}

export async function getJob(jobId: string): Promise<JobDetail> {
  const res = await fetch(`${API_BASE}/jobs/${jobId}`);
  if (!res.ok) return readError(res, `查询任务失败: HTTP ${res.status}`);
  return (await res.json()) as JobDetail;
}

export async function rewriteJob(jobId: string, action: RewriteAction) {
  const res = await fetch(`${API_BASE}/jobs/${jobId}/rewrite`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ action }),
  });
  if (!res.ok) return readError(res, `重写失败: HTTP ${res.status}`);
  return (await res.json()) as { job_id: string };
}

export async function getHistory(): Promise<HistoryItem[]> {
  const res = await fetch(`${API_BASE}/jobs`);
  if (!res.ok) return readError(res, `获取历史失败: HTTP ${res.status}`);
  return (await res.json()) as HistoryItem[];
}

export async function parseUpload(file: File): Promise<string> {
  const fd = new FormData();
  fd.append("file", file);
  const res = await fetch(`${API_BASE}/parse-upload`, {
    method: "POST",
    body: fd,
  });
  if (!res.ok) return readError(res, `文件解析失败: HTTP ${res.status}`);
  const payload = (await res.json()) as { extracted_text: string };
  return payload.extracted_text;
}

