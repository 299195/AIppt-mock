interface ImportMetaEnv {
  readonly VITE_API_BASE?: string;
  readonly VITE_FILE_BASE?: string;
}

interface ImportMeta {
  readonly env: ImportMetaEnv;
}
