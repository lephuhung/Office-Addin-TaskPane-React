import * as React from "react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import TextInsertion from "./TextInsertion";
import { Chat } from "./Chat";
import { makeStyles, shorthands, Tab, TabList, tokens, Button, Field, Input, Textarea, Switch, Text, Dropdown, Option } from "@fluentui/react-components";
import { Ribbon24Regular, LockOpen24Regular, DesignIdeas24Regular } from "@fluentui/react-icons";
import { ChatMessage, Settings, MCPAction } from "../types";
import { insertText, replaceSelection, insertHeading } from "../word";

export interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    ...shorthands.padding("20px"),
  },
  container: {
    width: "100%",
  },
  sidebar: {
    height: "100vh",
    overflow: "auto",
    ...shorthands.padding("12px"),
  },
  sidebarNoScroll: {
    height: "100vh",
    overflow: "hidden",
    padding: 0,
  },
  chatBox: {
    display: "grid",
    gridTemplateRows: "1fr auto",
    height: "100vh",
    gap: "8px",
  },
  messages: {
    overflow: "auto",
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    background: tokens.colorNeutralBackground2,
    ...shorthands.padding("8px"),
    borderRadius: tokens.borderRadiusMedium,
  },
  message: {
    whiteSpace: "pre-wrap",
    background: tokens.colorNeutralBackground3,
    ...shorthands.padding("8px"),
    borderRadius: tokens.borderRadiusSmall,
  },
  inputRow: {
    display: "grid",
    gridTemplateColumns: "1fr auto",
    gap: "8px",
  },
});

const listItems: HeroListItem[] = [
  {
    icon: <Ribbon24Regular />,
    primaryText: "This add-in illustrates a basic UI built using Fluent UI React.",
  },
  {
    icon: <LockOpen24Regular />,
    primaryText: "It also includes a simple boilerplate for calling the OpenAI-style backend and applying MCP actions.",
  },
  {
    icon: <DesignIdeas24Regular />,
    primaryText: "Customize the assistant to author content and apply document edits via MCP.",
  },
];

function useSettings() {
  const [settings, setSettings] = React.useState<Settings>({
    apiBaseUrl: "http://10.8.0.8:8000",
    apiKey: "",
    modelsUrl: "http://10.8.0.8:8000/v1/models",
    model: "gpt-4o-mini",
    systemPrompt: "You are a helpful assistant for Word authoring.",
    allowEdit: true,
  });
  React.useEffect(() => {
    (async () => {
      try {
        const stored = await (OfficeRuntime as any).storage?.getItem("settings");
        if (stored) setSettings(JSON.parse(stored));
      } catch (e) {
        const raw = localStorage.getItem("settings");
        if (raw) setSettings(JSON.parse(raw));
      }
    })();
  }, []);

  const save = async (s: Settings) => {
    setSettings(s);
    try {
      await (OfficeRuntime as any).storage?.setItem("settings", JSON.stringify(s));
    } catch (e) {
      localStorage.setItem("settings", JSON.stringify(s));
    }
  };

  return { settings, save };
}

async function applyActions(actions?: MCPAction[], log?: (entry: string) => void) {
  if (!actions) return;
  for (const a of actions) {
    try {
      if (a.type === "insertText") {
        const text = String(a.payload?.text ?? "");
        await insertText(text);
        log?.(`insertText: ${text.slice(0, 64)}`);
      } else if (a.type === "replaceSelection") {
        const text = String(a.payload?.text ?? "");
        await replaceSelection(text);
        log?.(`replaceSelection: ${text.slice(0, 64)}`);
      } else if (a.type === "insertHeading") {
        const text = String(a.payload?.text ?? "");
        const level = Number(a.payload?.level ?? 1);
        await insertHeading(text, level);
        log?.(`insertHeading(level ${level}): ${text.slice(0, 64)}`);
      } else {
        log?.(`Unknown action: ${a.type}`);
      }
    } catch (e: any) {
      log?.(`Action failed (${a.type}): ${e?.message || String(e)}`);
    }
  }
}

export default function App({ title }: AppProps) {
  const styles = useStyles();
  const { settings, save } = useSettings();
  const [tab, setTab] = React.useState<string>("chat");
  const [messages, setMessages] = React.useState<ChatMessage[]>([{
    id: "m0", role: "system", content: settings.systemPrompt || "", createdAt: Date.now()
  }]);
  const [input, setInput] = React.useState("");
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState<string | null>(null);

  const [actionLogs, setActionLogs] = React.useState<string[]>([]);
  const [conn, setConn] = React.useState<"unknown" | "ok" | "error">("unknown");
  const [connMsg, setConnMsg] = React.useState<string | null>(null);
  const [availableModels, setAvailableModels] = React.useState<string[]>(["gpt-4o-mini", "gpt-4o", "gpt-3.5-turbo"]);
  const [isLoggedIn, setIsLoggedIn] = React.useState(false);
  const [loginForm, setLoginForm] = React.useState({ username: "", password: "" });
  const [userInfo, setUserInfo] = React.useState<{ username: string } | null>(null);
  const messagesEndRef = React.useRef<HTMLDivElement | null>(null);

  React.useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages]);

  React.useEffect(() => {
    // keep system prompt sync when settings change
    setMessages((prev) => {
      const others = prev.filter((m) => m.role !== "system");
      return [{ id: "m0", role: "system", content: settings.systemPrompt || "", createdAt: Date.now() }, ...others];
    });
  }, [settings.systemPrompt]);

  const checkConnection = React.useCallback(async () => {
    const modelsUrl = settings.modelsUrl || `${settings.apiBaseUrl?.replace(/\/$/, "")}/v1/models`;
    if (!modelsUrl) {
      setConn("error");
      setConnMsg("Thiếu Models URL hoặc API Base URL");
      return;
    }
    try {
      let res = await fetch(modelsUrl, {
        headers: settings.apiKey ? { Authorization: `Bearer ${settings.apiKey}` } : undefined,
      });
      if (res.ok) {
        const data = await res.json();
        if (data?.data && Array.isArray(data.data)) {
          const models = data.data.map((m: any) => m.id || m.name).filter(Boolean);
          if (models.length > 0) {
            setAvailableModels(models);
          }
        }
        setConn("ok");
        setConnMsg(null);
      } else {
        // fallback thử /health nếu models URL không khả dụng
        const baseUrl = settings.apiBaseUrl?.replace(/\/$/, "");
        if (baseUrl) {
          try {
            const healthRes = await fetch(`${baseUrl}/health`, {
              headers: settings.apiKey ? { Authorization: `Bearer ${settings.apiKey}` } : undefined,
            });
            if (healthRes.ok) {
              setConn("ok");
              setConnMsg(null);
            } else {
              setConn("error");
              setConnMsg(`HTTP ${res.status}`);
            }
          } catch {
            setConn("error");
            setConnMsg(`HTTP ${res.status}`);
          }
        } else {
          setConn("error");
          setConnMsg(`HTTP ${res.status}`);
        }
      }
    } catch (e: any) {
      setConn("error");
      setConnMsg(e?.message || String(e));
    }
  }, [settings.apiBaseUrl, settings.modelsUrl, settings.apiKey]);

  React.useEffect(() => {
    // Attempt auto-check when URL or key changes
    const modelsUrl = settings.modelsUrl || settings.apiBaseUrl;
    if (modelsUrl) {
      checkConnection();
    } else {
      setConn("unknown");
      setConnMsg(null);
    }
  }, [settings.apiBaseUrl, settings.modelsUrl, settings.apiKey, checkConnection]);

  const send = async () => {
    if (!input.trim()) return;
    if (!settings.apiBaseUrl || !settings.apiKey) {
      setError("Vui lòng cấu hình API Base URL và API Key trong Settings.");
      setTab("settings");
      return;
    }
    setError(null);
    const userMsg: ChatMessage = { id: `u-${Date.now()}`, role: "user", content: input, createdAt: Date.now() };
    setMessages((prev) => [...prev, userMsg]);
    setInput("");
    setLoading(true);

    try {
      const res = await fetch(`${settings.apiBaseUrl.replace(/\/$/, "")}/v1/chat/completions`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Authorization": `Bearer ${settings.apiKey}`,
        },
        body: JSON.stringify({
          model: settings.model || "gpt-4o-mini",
          messages: messages.concat(userMsg).map((m) => ({ role: m.role, content: m.content })),
          stream: false,
          // You can instruct server to return MCP actions in a structured field
          extra: { mcp: true, host: "word" },
        }),
      });

      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const data = await res.json();
      const assistantText: string = data?.choices?.[0]?.message?.content ?? "";
      const actions: MCPAction[] | undefined = data?.choices?.[0]?.message?.mcp_actions || data?.mcp_actions;

      const assistantMsg: ChatMessage = { id: `a-${Date.now()}`, role: "assistant", content: assistantText, createdAt: Date.now() };
      setMessages((prev) => [...prev, assistantMsg]);

      if (settings.allowEdit) {
        await applyActions(actions, (entry) => setActionLogs((prev) => [...prev, `${new Date().toLocaleTimeString()}: ${entry}`]));
      }
      setConn("ok");
      setConnMsg(null);
    } catch (e: any) {
      setError(e?.message || String(e));
      setConn("error");
      setConnMsg(e?.message || String(e));
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className={styles.root}>
      <div className={styles.container}>
        <div className={tab === "chat" ? styles.sidebarNoScroll : styles.sidebar}>
          <div style={{ padding: tab === "chat" ? "12px 12px 0 12px" : "12px" }}>
            <TabList selectedValue={tab} onTabSelect={(_, d) => setTab(String(d.value))}>
              <Tab value="chat">Chat</Tab>
              <Tab value="settings">Settings</Tab>
              <Tab value="quickstart">Quick Start</Tab>
            </TabList>
          </div>

          {tab === "chat" && <Chat />}

          {tab === "settings" && (
            <SettingsPanel
              value={settings}
              onChange={save}
              logs={actionLogs}
              onClearLogs={() => setActionLogs([])}
              connectionStatus={conn}
              connectionMessage={connMsg || undefined}
              onCheckConnection={checkConnection}
              availableModels={availableModels}
            />
          )}

          {tab === "quickstart" && (
            <div style={{ marginTop: 8 }}>
              {/* Logo Section */}
              <div style={{ textAlign: "center", marginBottom: 16 }}>
                <img src="assets/logo-filled.png" alt={title} style={{ width: 64, height: 64, marginBottom: 8 }} />
                <Text size={500} weight="semibold" style={{ display: "block" }}>{title}</Text>
              </div>

              {/* Login Section */}
              {!isLoggedIn ? (
                <div style={{ marginBottom: 16, padding: 16, background: tokens.colorNeutralBackground2, borderRadius: tokens.borderRadiusMedium }}>
                  <Text size={400} weight="semibold" style={{ display: "block", marginBottom: 12 }}>Đăng nhập tài khoản</Text>
                  <Field label="Tên đăng nhập" style={{ marginBottom: 8 }}>
                    <Input 
                      value={loginForm.username} 
                      onChange={(_, d) => setLoginForm(prev => ({ ...prev, username: d.value }))}
                      placeholder="Nhập tên đăng nhập"
                    />
                  </Field>
                  <Field label="Mật khẩu" style={{ marginBottom: 12 }}>
                    <Input 
                      type="password"
                      value={loginForm.password} 
                      onChange={(_, d) => setLoginForm(prev => ({ ...prev, password: d.value }))}
                      placeholder="Nhập mật khẩu"
                    />
                  </Field>
                  <Button 
                    appearance="primary" 
                    onClick={() => {
                      if (loginForm.username && loginForm.password) {
                        setIsLoggedIn(true);
                        setUserInfo({ username: loginForm.username });
                        setLoginForm({ username: "", password: "" });
                      }
                    }}
                    disabled={!loginForm.username || !loginForm.password}
                  >
                    Đăng nhập
                  </Button>
                </div>
              ) : (
                <div style={{ marginBottom: 16, padding: 12, background: tokens.colorBrandBackground2, borderRadius: tokens.borderRadiusMedium, display: "flex", alignItems: "center", gap: 8 }}>
                  <img src="assets/logo-filled.png" alt="User" style={{ width: 24, height: 24, borderRadius: "50%" }} />
                  <Text size={300}>Xin chào, {userInfo?.username}</Text>
                  <Button size="small" onClick={() => { setIsLoggedIn(false); setUserInfo(null); }}>Đăng xuất</Button>
                </div>
              )}

              {/* Features Introduction */}
              <div style={{ textAlign: "center" }}>
                <Text size={400} weight="semibold" style={{ display: "block", marginBottom: 16 }}>Tính năng chính</Text>
                <div style={{ maxWidth: "400px", margin: "0 auto" }}>
                  <HeroList message="" items={listItems} />
                </div>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function SettingsPanel({ value, onChange, logs, onClearLogs, connectionStatus, connectionMessage, onCheckConnection, availableModels }: { value: Settings; onChange: (s: Settings) => void; logs: string[]; onClearLogs: () => void; connectionStatus: "unknown" | "ok" | "error"; connectionMessage?: string; onCheckConnection: () => void; availableModels: string[]; }) {
  const [form, setForm] = React.useState<Settings>(value);
  React.useEffect(() => setForm(value), [value]);

  const update = (patch: Partial<Settings>) => setForm((prev) => ({ ...prev, ...patch }));
  const save = () => onChange(form);

  const statusColor = connectionStatus === "ok" ? tokens.colorPaletteGreenForeground1 : connectionStatus === "error" ? tokens.colorPaletteRedForeground1 : tokens.colorNeutralForeground3;
  const statusText = connectionStatus === "ok" ? "Connected" : connectionStatus === "error" ? "Disconnected" : "Unknown";

  return (
    <div style={{ display: "grid", gap: 8, marginTop: 8 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
        <div style={{ width: 8, height: 8, borderRadius: 999, background: statusColor }} />
        <Text>{statusText}{connectionMessage ? ` - ${connectionMessage}` : ""}</Text>
        <Button size="small" onClick={onCheckConnection}>Kiểm tra kết nối</Button>
      </div>
      <Field label="API Base URL">
        <Input value={form.apiBaseUrl} onChange={(_, d) => update({ apiBaseUrl: d.value })} placeholder="https://your-server" />
      </Field>
      <Field label="Models URL">
        <Input value={form.modelsUrl || ""} onChange={(_, d) => update({ modelsUrl: d.value })} placeholder="https://your-server/v1/models" />
      </Field>
      <Field label="API Key">
        <Input type="password" value={form.apiKey} onChange={(_, d) => update({ apiKey: d.value })} placeholder="sk-..." />
      </Field>
      <Field label="Model">
        <Dropdown
          value={form.model}
          selectedOptions={[form.model]}
          onOptionSelect={(_, data) => update({ model: data.optionValue || "" })}
          placeholder="Chọn model"
        >
          {availableModels.map((model) => (
            <Option key={model} value={model}>
              {model}
            </Option>
          ))}
        </Dropdown>
      </Field>
      <Field label="System Prompt">
        <Textarea value={form.systemPrompt} onChange={(_, d) => update({ systemPrompt: d.value })} />
      </Field>
      <Field label="Cho phép add-in chỉnh sửa tài liệu">
        <Switch checked={!!form.allowEdit} onChange={(_, d) => update({ allowEdit: d.checked })} />
      </Field>
      <div style={{ display: "flex", gap: 8 }}>
        <Button appearance="primary" onClick={save}>Lưu cấu hình</Button>
        <Button onClick={onClearLogs} appearance="secondary">Xóa nhật ký</Button>
      </div>
      {logs.length > 0 && (
        <Field label="Nhật ký thao tác đã áp dụng">
          <div style={{ maxHeight: 160, overflow: "auto", background: tokens.colorNeutralBackground2, padding: 8, borderRadius: 6 }}>
            {logs.map((l, i) => (
              <div key={i} style={{ fontSize: 12, color: tokens.colorNeutralForeground3 }}>{l}</div>
            ))}
          </div>
        </Field>
      )}
    </div>
  );
}
