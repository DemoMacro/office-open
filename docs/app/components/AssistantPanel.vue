<script setup lang="ts">
import { useChat } from "@ai-sdk/vue";
import { Ai, Drawer } from "@bysages/vue";
import { DefaultChatTransport, type ToolUIPart } from "ai";

const {
  Conversation,
  Message,
  MessageContent,
  Response,
  Reasoning,
  Tool,
  Suggestion,
  PromptInput,
  Loader,
  Action,
} = Ai;
const { Root, Backdrop, Positioner, Content, Title, CloseTrigger } = Drawer;

const { isOpen, open, close, draft } = useAssistant();

const { t } = useDocsI18n();

const { messages, sendMessage, status, error, stop, setMessages } = useChat({
  transport: new DefaultChatTransport({ api: "/api/assistant" }),
});

const text = ref("");

// The floating input handed its text over — send it and clear the hand-off.
watch(draft, (value) => {
  if (!value) return;
  draft.value = "";
  void sendMessage({ text: value });
});

const busy = computed(() => status.value !== "ready" && status.value !== "error");

/** Zag-free mapping from a tool part's lifecycle to our status dot. */
const TOOL_STATUS: Record<string, "running" | "completed" | "error"> = {
  "input-streaming": "running",
  "input-available": "running",
  "output-available": "completed",
  "output-error": "error",
};

/** The tool call in words — Ai.Tool prints its name verbatim, so the
 * friendly label is computed here per tool and lifecycle. */
function toolLabel(part: ToolUIPart): string {
  const done = part.state === "output-available";
  const input = (part.input ?? {}) as Record<string, unknown>;
  const path = input.path ? ` ${input.path}` : "";
  const pairs: Record<string, [string, string]> = {
    "list-pages": ["Searching pages", "Searched pages"],
    "get-page": [`Reading${path}`, `Read${path}`],
    "generate-docx": ["Creating Word document", "Created Word document"],
    "generate-pptx": ["Creating presentation", "Created presentation"],
    "generate-xlsx": ["Creating spreadsheet", "Created spreadsheet"],
    "office-open-schema-lookup": ["Fetching option schema", "Fetched option schema"],
  };
  const pair = pairs[part.type.slice(5)];
  if (pair) return done ? pair[1] : pair[0];
  return done ? `Ran ${part.type.slice(5)}` : `Running ${part.type.slice(5)}`;
}

/** Tool payloads run long; the collapsible only needs a peek. */
function format(value: unknown): string {
  const text = typeof value === "string" ? value : JSON.stringify(value, null, 2);
  return text.length > 500 ? `${text.slice(0, 500)}…` : text;
}

interface DocumentData {
  filename: string;
  base64: string;
  mimeType: string;
  size: number;
}

// data-document parts are emitted by our assistant API (the office-open
// generate tools), not part of the AI SDK's typed UIMessage parts.
function documentData(part: { data?: unknown }): DocumentData {
  return part.data as DocumentData;
}

function clearMessages() {
  if (busy.value) stop();
  setMessages([]);
  text.value = "";
}

function toggleShortcut(event: KeyboardEvent) {
  if (!(event.metaKey || event.ctrlKey) || event.key.toLowerCase() !== "i") return;
  event.preventDefault();
  if (isOpen.value) close();
  else open();
}

onMounted(() => window.addEventListener("keydown", toggleShortcut));
onBeforeUnmount(() => window.removeEventListener("keydown", toggleShortcut));
</script>

<template>
  <ClientOnly>
    <Root :open="isOpen" swipe-direction="end" @update:open="(value: boolean) => value || close()">
      <Backdrop />
      <Positioner>
        <Content aria-label="AI assistant" class="bs-docs-assistant">
          <div class="bs-docs-assistant-head">
            <Title>{{ t("docs.assistantTitle") }}</Title>
            <div class="bs-docs-assistant-actions">
              <Action :label="t('docs.assistantClear')" @click="clearMessages">
                <Icon name="i-lucide-eraser" />
              </Action>
              <CloseTrigger as-child>
                <Action label="Close assistant">✕</Action>
              </CloseTrigger>
            </div>
          </div>

          <Conversation class="bs-docs-assistant-log">
            <template v-if="messages.length === 0">
              <div class="bs-docs-assistant-empty">
                <p class="bs-docs-assistant-greeting">{{ t("docs.assistantGreeting") }}</p>
                <Suggestion
                  v-for="starter in [
                    t('docs.assistantStarter1'),
                    t('docs.assistantStarter2'),
                    t('docs.assistantStarter3'),
                  ]"
                  :key="starter"
                  :prompt="starter"
                  @select="sendMessage({ text: $event })"
                />
              </div>
            </template>

            <template v-for="message in messages" :key="message.id">
              <Message :role="message.role">
                <MessageContent v-if="message.role === 'user'">
                  {{
                    message.parts
                      .flatMap((part) => (part.type === "text" ? [part.text] : []))
                      .join("")
                  }}
                </MessageContent>
                <template v-else>
                  <template v-for="(part, index) in message.parts" :key="index">
                    <DocumentCard
                      v-if="part.type === 'data-document' && part.data"
                      v-bind="documentData(part)"
                    />
                    <Response v-else-if="part.type === 'text'" :content="part.text" />
                    <Reasoning
                      v-else-if="part.type === 'reasoning'"
                      :label="t('docs.assistantThinking')"
                    >
                      {{ part.text }}
                    </Reasoning>
                    <!-- A `tool-` prefix marks a tool invocation, but TS
                         cannot narrow a union by prefix — the assertion
                         carries what the check just proved. -->
                    <Tool
                      v-else-if="part.type.startsWith('tool-')"
                      :name="toolLabel(part as ToolUIPart)"
                      :status="TOOL_STATUS[(part as ToolUIPart).state] ?? 'running'"
                    >
                      <template v-if="(part as ToolUIPart).input !== undefined" #input>{{
                        format((part as ToolUIPart).input)
                      }}</template>
                      <template v-if="(part as ToolUIPart).output !== undefined" #output>{{
                        format((part as ToolUIPart).output)
                      }}</template>
                    </Tool>
                  </template>
                </template>
              </Message>
            </template>

            <Loader v-if="status === 'submitted'" />
            <p v-if="error" class="bs-docs-assistant-error" role="alert">
              {{ t("docs.assistantError") }}
            </p>
          </Conversation>

          <PromptInput v-model="text" :disabled="busy" @submit="sendMessage({ text: $event })" />
        </Content>
      </Positioner>
    </Root>
  </ClientOnly>
</template>
