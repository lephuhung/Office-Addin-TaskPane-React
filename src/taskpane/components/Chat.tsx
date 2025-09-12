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
} from "@fluentui/react-components";
import { Send24Regular, Bot24Regular, Person24Regular } from "@fluentui/react-icons";

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
  },
  userMessage: {
    alignSelf: "flex-end",
    backgroundColor: "#0078d4", // Blue background for user
    color: "#ffffff", // White text for contrast
    borderRadius: "18px 18px 4px 18px",
    marginLeft: "20%",
    border: "none",
  },
  botMessage: {
    alignSelf: "flex-start",
    backgroundColor: "#f3f2f1", // Light gray for bot
    color: "#323130", // Dark text for contrast
    borderRadius: "18px 18px 18px 4px",
    marginRight: "20%",
    border: `1px solid #e1dfdd`,
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
    color: "#ffffff", // Ensure white text for user messages
  },
  botMessageContent: {
    padding: "12px 16px",
    color: "#323130", // Ensure dark text for bot messages
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
  const messagesEndRef = React.useRef<HTMLDivElement>(null);

  const scrollToBottom = () => {
    messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
  };

  React.useEffect(() => {
    scrollToBottom();
  }, [messages]);

  const handleSendMessage = () => {
    if (inputValue.trim()) {
      const newMessage: Message = {
        id: Date.now().toString(),
        text: inputValue.trim(),
        sender: "user",
        timestamp: new Date(),
      };

      setMessages(prev => [...prev, newMessage]);
      setInputValue("");

      // Simulate bot response
      setTimeout(() => {
        const botResponse: Message = {
          id: (Date.now() + 1).toString(),
          text: "Tôi đã nhận được tin nhắn của bạn. Đây là phản hồi tự động từ AI Assistant.",
          sender: "bot",
          timestamp: new Date(),
        };
        setMessages(prev => [...prev, botResponse]);
      }, 1000);
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
                  <Text size={300} style={{ color: message.sender === "user" ? tokens.colorNeutralForegroundOnBrand : tokens.colorNeutralForeground1 }}>
                    {message.text}
                  </Text>
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