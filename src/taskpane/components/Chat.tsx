import * as React from "react";
import {
  Button,
  Input,
  Text,
  Card,
  CardHeader,
  CardPreview,
  makeStyles,
  tokens,
  Textarea,
  Spinner,
} from "@fluentui/react-components";
import { Send24Regular, Bot24Regular, Person24Regular } from "@fluentui/react-icons";
import { Settings } from "../types";
import ReactMarkdown from "react-markdown";

interface Message {
  id: string;
  text: string;
  sender: "user" | "bot";
  timestamp: Date;
}

const useStyles = makeStyles({
  chatContainer: {
    display: "flex",
    flexDirection: "column",
    height: "calc(100vh - 60px)", // Subtract space for TabList
    maxHeight: "calc(100vh - 60px)",
    overflow: "hidden",
  },
  chatHeader: {
    padding: "12px 16px",
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
    flexShrink: 0,
  },
  messagesArea: {
    flex: 1,
    overflowY: "auto",
    padding: "16px",
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    minHeight: 0, // Important for flex child to shrink
  },
  messageCard: {
    maxWidth: "75%",
    wordWrap: "break-word",
    marginBottom: "8px",
    // Override Fluent UI Card defaults
    "& .fui-Card": {
      backgroundColor: "inherit !important",
      border: "inherit !important",
    }
  },
  userMessage: {
    alignSelf: "flex-start", // Align to left side
    backgroundColor: "#0078d4 !important", // Blue background for user
    color: "#ffffff !important", // White text for contrast
    borderRadius: "18px 18px 18px 4px !important", // Rounded corners for left side
    marginRight: "20%", // Space on right side
    border: "none !important",
    // Override nested components
    "& .fui-CardHeader": {
      backgroundColor: "transparent !important",
      color: "#ffffff !important",
    },
    "& .fui-CardPreview": {
      backgroundColor: "transparent !important",
      color: "#ffffff !important",
    },
    "& .fui-Text": {
      color: "#ffffff !important",
    }
  },
  botMessage: {
    alignSelf: "flex-end", // Align to right side
    backgroundColor: "#ffffff !important", // White background for better contrast
    color: "#000000 !important", // Black text for maximum contrast
    borderRadius: "18px 18px 4px 18px !important", // Rounded corners for right side
    marginLeft: "20%", // Space on left side
    border: `1px solid #d1d1d1 !important`, // Darker border for definition
    // Override nested components
    "& .fui-CardHeader": {
      backgroundColor: "transparent !important",
      color: "#000000 !important",
    },
    "& .fui-CardPreview": {
      backgroundColor: "transparent !important",
      color: "#000000 !important",
    },
    "& .fui-Text": {
      color: "#000000 !important",
    }
  },
  inputArea: {
    padding: "16px",
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
    flexShrink: 0,
  },
  inputContainer: {
    display: "flex",
    gap: "8px",
    alignItems: "flex-end",
  },
  messageInput: {
    flex: 1,
    minHeight: "36px",
    maxHeight: "120px",
    resize: "none",
  },
  sendButton: {
    minWidth: "40px",
    height: "36px",
  },
  messageHeader: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    marginBottom: "4px",
  },
  messageContent: {
    padding: "12px 16px",
  },
  userMessageContent: {
    padding: "12px 16px",
    color: "#ffffff !important", // Ensure white text for user messages
    "& *": {
      color: "#ffffff !important",
    }
  },
  botMessageContent: {
    padding: "12px 16px",
    color: "#000000 !important", // Black text for maximum contrast
    "& *": {
      color: "#000000 !important",
    },
    "& p": {
      color: "#000000 !important",
      margin: "0 0 8px 0",
    },
    "& ul, & ol": {
      color: "#000000 !important",
      paddingLeft: "20px",
    },
    "& li": {
      color: "#000000 !important",
    },
    "& code": {
      backgroundColor: "#f5f5f5 !important",
      color: "#000000 !important",
      padding: "2px 4px",
      borderRadius: "3px",
    },
    "& pre": {
      backgroundColor: "#f5f5f5 !important",
      color: "#000000 !important",
      padding: "8px",
      borderRadius: "4px",
      overflow: "auto",
    }
  },
  emptyState: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    height: "100%",
    textAlign: "center",
    color: tokens.colorNeutralForeground3,
  },
});

export const Chat: React.FC = () => {
  const styles = useStyles();
  const [messages, setMessages] = React.useState<Message[]>([]);
  const [inputValue, setInputValue] = React.useState("");
  const [loading, setLoading] = React.useState(false);
  const [settings, setSettings] = React.useState<Settings>({
    apiBaseUrl: "http://10.8.0.8:8000",
    apiKey: "",
    model: "gpt-4o-mini",
  });
  const messagesEndRef = React.useRef<HTMLDivElement>(null);

  // Debug log for initial state
  React.useEffect(() => {
    console.log("Chat: Component mounted, initial messages state:", messages);
  }, []);

  // Load settings from storage
  React.useEffect(() => {
    (async () => {
      try {
        const stored = await (OfficeRuntime as any).storage?.getItem("settings");
        if (stored) {
          const parsedSettings = JSON.parse(stored);
          console.log("Chat: Loaded settings from OfficeRuntime storage:", parsedSettings);
          setSettings(parsedSettings);
        } else {
          console.log("Chat: No settings found in OfficeRuntime storage");
        }
      } catch (e) {
        console.log("Chat: OfficeRuntime storage failed, trying localStorage:", e);
        const raw = localStorage.getItem("settings");
        if (raw) {
          const parsedSettings = JSON.parse(raw);
          console.log("Chat: Loaded settings from localStorage:", parsedSettings);
          setSettings(parsedSettings);
        } else {
          console.log("Chat: No settings found in localStorage either");
        }
      }
    })();
  }, []);

  const scrollToBottom = () => {
    messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
  };

  React.useEffect(() => {
    console.log("Chat: Messages state changed:", messages);
    scrollToBottom();
  }, [messages]);

  const handleSendMessage = async () => {
    if (!inputValue.trim()) return;
    
    console.log("Chat: Current settings when sending message:", settings);
    
    if (!settings.apiBaseUrl || !settings.apiKey) {
      const errorMessage: Message = {
        id: Date.now().toString(),
        text: "Vui lòng cấu hình API Base URL và API Key trong tab Settings trước khi chat.",
        sender: "bot",
        timestamp: new Date(),
      };
      setMessages(prev => {
        console.log("Chat: Adding error message, prev messages:", prev);
        return Array.isArray(prev) ? [...prev, errorMessage] : [errorMessage];
      });
      return;
    }

    const newMessage: Message = {
      id: Date.now().toString(),
      text: inputValue.trim(),
      sender: "user",
      timestamp: new Date(),
    };

    setMessages(prev => {
      console.log("Chat: Adding user message, prev messages:", prev);
      return Array.isArray(prev) ? [...prev, newMessage] : [newMessage];
    });
    setInputValue("");
    setLoading(true);

    try {
      const response = await fetch(`${settings.apiBaseUrl.replace(/\/$/, "")}/v1/chat/completions`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Authorization": `Bearer ${settings.apiKey}`,
        },
        body: JSON.stringify({
          model: settings.model || "gpt-4o-mini",
          messages: [
            { role: "system", content: "You are a helpful AI assistant for Word documents. Respond in Vietnamese." },
            { role: "user", content: inputValue.trim() }
          ],
          stream: false,
        }),
      });

      if (!response.ok) {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      }

      const data = await response.json();
      const assistantText = data?.choices?.[0]?.message?.content ?? "Xin lỗi, tôi không thể tạo phản hồi.";

      const botResponse: Message = {
        id: (Date.now() + 1).toString(),
        text: assistantText,
        sender: "bot",
        timestamp: new Date(),
      };
      
      setMessages(prev => {
        console.log("Chat: Adding bot response, prev messages:", prev);
        return Array.isArray(prev) ? [...prev, botResponse] : [botResponse];
      });
    } catch (error: any) {
      const errorMessage: Message = {
        id: (Date.now() + 1).toString(),
        text: `Lỗi kết nối API: ${error.message}. Vui lòng kiểm tra cấu hình API trong tab Settings.`,
        sender: "bot",
        timestamp: new Date(),
      };
      setMessages(prev => {
        console.log("Chat: Adding error message, prev messages:", prev);
        return Array.isArray(prev) ? [...prev, errorMessage] : [errorMessage];
      });
    } finally {
      setLoading(false);
    }
  };

  const handleKeyPress = (event: React.KeyboardEvent) => {
    if (event.key === "Enter" && !event.shiftKey) {
      event.preventDefault();
      handleSendMessage();
    }
  };

  const formatTime = (date: Date) => {
    return date.toLocaleTimeString("vi-VN", {
      hour: "2-digit",
      minute: "2-digit",
    });
  };

  return (
    <div className={styles.chatContainer}>
      {/* Chat Header */}
      <div className={styles.chatHeader}>
        <Text size={500} weight="semibold">
          AI Chat Assistant
        </Text>
        <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
          Trò chuyện với AI để được hỗ trợ
        </Text>
      </div>

      {/* Messages Area */}
      <div className={styles.messagesArea}>
        {messages.length === 0 ? (
          <div className={styles.emptyState}>
            <Bot24Regular style={{ fontSize: "48px", marginBottom: "16px" }} />
            <Text size={400} weight="semibold">
              Chào mừng đến với AI Chat!
            </Text>
            <Text size={300}>
              Hãy bắt đầu cuộc trò chuyện bằng cách gửi tin nhắn đầu tiên
            </Text>
          </div>
        ) : (
          messages.map((message) => (
            <Card
              key={message.id}
              className={`${styles.messageCard} ${
                message.sender === "user" ? styles.userMessage : styles.botMessage
              }`}
            >
              <CardHeader
                header={
                  <div className={styles.messageHeader}>
                    {message.sender === "user" ? (
                      <Person24Regular />
                    ) : (
                      <Bot24Regular />
                    )}
                    <Text size={300} weight="semibold">
                      {message.sender === "user" ? "Bạn" : "AI Assistant"}
                    </Text>
                    <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
                      {formatTime(message.timestamp)}
                    </Text>
                  </div>
                }
              />
              <CardPreview>
                 <div className={message.sender === "user" ? styles.userMessageContent : styles.botMessageContent}>
                   {message.sender === "user" ? (
                     <Text size={300}>
                       {message.text}
                     </Text>
                   ) : (
                     <div>
                       <ReactMarkdown>{message.text}</ReactMarkdown>
                     </div>
                   )}
                 </div>
               </CardPreview>
            </Card>
          ))
        )}
        <div ref={messagesEndRef} />
      </div>

      {/* Input Area */}
      <div className={styles.inputArea}>
        <div className={styles.inputContainer}>
          <Textarea
            className={styles.messageInput}
            placeholder="Nhập tin nhắn của bạn..."
            value={inputValue}
            onChange={(_, data) => setInputValue(data.value)}
            onKeyDown={handleKeyPress}
            resize="vertical"
          />
          <Button
            className={styles.sendButton}
            appearance="primary"
            icon={<Send24Regular />}
            onClick={handleSendMessage}
            disabled={!inputValue.trim()}
          />
        </div>
      </div>
    </div>
  );
};