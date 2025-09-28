export type Role = "system" | "user" | "assistant";

export interface ChatMessage {
  id: string;
  role: Role;
  content: string;
  createdAt: number;
}

export type MCPActionType = "insertText" | "replaceSelection" | "insertHeading";

export interface MCPAction {
  type: MCPActionType;
  payload: any;
  explain?: string;
}

export interface Settings {
  apiBaseUrl: string;
  apiKey: string;
  modelsUrl?: string;
  model?: string;
  systemPrompt?: string;
  allowEdit?: boolean;
}

export interface ChatResult {
  assistantText: string;
  actions?: MCPAction[];
}