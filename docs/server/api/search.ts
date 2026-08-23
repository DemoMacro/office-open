import { createMCPClient } from "@ai-sdk/mcp";
import { createOpenAICompatible } from "@ai-sdk/openai-compatible";
import {
  streamText,
  convertToModelMessages,
  createUIMessageStream,
  createUIMessageStreamResponse,
} from "ai";
import type { UIMessageStreamWriter, ToolSet } from "ai";
import type { H3Event } from "h3";
import { officeOpenTools } from "office-open/ai";

const GENERATE_TOOL_EXTENSIONS = {
  "generate-docx": "docx",
  "generate-pptx": "pptx",
  "generate-xlsx": "xlsx",
} as const;

// Enough headroom for schema lookups plus a validation retry before the final answer.
const MAX_STEPS = 20;

function createLocalFetch(event: H3Event): typeof fetch {
  const origin = getRequestURL(event).origin;

  return (input, init) => {
    const requestUrl =
      input instanceof URL
        ? input
        : typeof input === "string"
          ? new URL(input, origin)
          : new URL(input.url);
    const localPath =
      requestUrl.origin === origin
        ? `${requestUrl.pathname}${requestUrl.search}`
        : requestUrl.toString();

    return event.fetch(localPath, init);
  };
}

function stopWhenResponseComplete({ steps }: { steps: any[] }): boolean {
  const lastStep = steps.at(-1);
  if (!lastStep) return false;

  const hasText = Boolean(lastStep.text && lastStep.text.trim().length > 0);
  const hasNoToolCalls = !lastStep.toolCalls || lastStep.toolCalls.length === 0;

  if (hasText && hasNoToolCalls) return true;

  return steps.length >= MAX_STEPS;
}

function getSystemPrompt(siteName: string) {
  return `You are the documentation assistant for ${siteName}. Help users navigate and understand the project documentation.

**Your identity:**
- You are an assistant helping users with ${siteName} documentation
- NEVER use first person ("I", "me", "my") - always refer to the project by name: "${siteName} provides...", "${siteName} supports...", "The project offers..."
- Be confident and knowledgeable about the project
- Speak as a helpful guide, not as the documentation itself

**Tool usage (CRITICAL):**
- You have tools: list-pages (discover pages), get-page (read a page), generate-docx / generate-pptx / generate-xlsx (create Office files), and office-open-schema-lookup (fetch option schemas on demand)
- If a page title clearly matches the question, read it directly without listing first
- ALWAYS respond with text after using tools - never end with just tool calls

**Guidelines:**
- If you can't find something, say "There is no documentation on that yet" or "${siteName} doesn't cover that topic yet"
- Be concise, helpful, and direct
- Guide users like a friendly expert would

**Links and exploration:**
- Tool results include a \`url\` for each page — prefer markdown links \`[label](url)\` so users can open the doc in one click
- When it helps, add extra links (related pages, "read more", side topics) — make the answer easy to dig into, not a wall of text
- Stick to URLs from tool results (\`url\` / \`path\`) so links stay valid

**FORMATTING RULES (CRITICAL):**
- NEVER use markdown headings (#, ##, ###, etc.)
- Use **bold text** for emphasis and section labels
- Start responses with content directly, never with a heading
- Use bullet points for lists
- Keep code examples focused and minimal

**Response style:**
- Conversational but professional
- "Here's how you can do that:" instead of "The documentation shows:"
- "${siteName} supports TypeScript out of the box" instead of "I support TypeScript"
- Provide actionable guidance, not just information dumps

**Document Generation:**
- When a user asks to create/generate/build an Office document, ALWAYS read the relevant documentation pages FIRST to understand the correct JSON structure
- Unsure about an option type's fields? Call office-open-schema-lookup with { type, definitions: [...] } (e.g. ["ParagraphOptions"]) — it returns the full schema slice; names come from the generate tool's skeleton stubs
- The generate tools validate input with JSON Schema before generating; if validation reports errors, fix the reported instance paths and call the same tool again
- CRITICAL STRUCTURE RULES:
  - Section/slide children MUST use wrapper keys: { paragraph: {...} }, { table: {...} }, NOT bare objects
  - Text runs MUST have a "text" key: { text: "Hello", bold?: true }, NOT { bold: true } alone
  - Colors are hex WITHOUT "#": "FF0000", not "#FF0000"
- For docx: options must include { sections: [{ children: [{ paragraph: { children: [{ text: "..." }] } }] }] }
- For pptx: options must include { title: "...", slides: [{ children: [{ shape: { x: 100, y: 100, width: 600, height: 60, textBody: { text: "..." } } }] }] }
- For xlsx: options must include { worksheets: [{ rows: [{ cells: [{ value: "Name" }] }] }] }
- Set the "title" field in options to customize the download filename without extension (e.g. "My Report")
- Call the generate tool once per document — retry only to fix validation errors
- ALWAYS describe what you generated after the tool completes
- Keep generated documents focused and reasonable in size`;
}

export default defineEventHandler(async (event) => {
  const { messages } = await readBody(event);
  const config = useRuntimeConfig();
  const siteConfig = getSiteConfig(event);

  const apiBaseURL = process.env.OPENAI_COMPATIBLE_BASE_URL;
  const apiKey = process.env.OPENAI_COMPATIBLE_API_KEY;
  const modelName = process.env.OPENAI_COMPATIBLE_MODEL;

  if (!apiBaseURL || !apiKey || !modelName) {
    throw createError({ statusCode: 503, statusMessage: "AI assistant is not configured" });
  }

  const siteName = siteConfig.name || "Documentation";

  const provider = createOpenAICompatible({
    name: "openai-compatible",
    baseURL: apiBaseURL,
    apiKey,
  });

  const model = provider(modelName);

  const mcpServer = config.assistant.mcpServer;
  const isExternalUrl = mcpServer.startsWith("http://") || mcpServer.startsWith("https://");
  const appBaseURL = config.app?.baseURL?.replace(/\/$/, "") || "";

  // Local MCP server: route through event.fetch (nitro's local dispatcher) —
  // no network round-trip, and no port assumption (the dev server does not
  // have to sit on :3000; external URLs fetch directly).
  let transport: Parameters<typeof createMCPClient>[0]["transport"];
  if (isExternalUrl) {
    transport = { type: "http", url: mcpServer };
  } else {
    transport = {
      type: "http",
      url: `${getRequestURL(event).origin}${appBaseURL}${mcpServer}`,
      fetch: createLocalFetch(event),
    };
  }

  const httpClient = await createMCPClient({ transport });
  const mcpTools = await httpClient.tools();

  const stream = createUIMessageStream({
    execute: async ({ writer }: { writer: UIMessageStreamWriter }) => {
      const modelMessages = await convertToModelMessages(messages);
      const result = streamText({
        model,
        maxOutputTokens: 4000,
        maxRetries: 2,
        stopWhen: stopWhenResponseComplete,
        system: getSystemPrompt(siteName),
        messages: modelMessages,
        tools: { ...mcpTools, ...officeOpenTools } as ToolSet,
        onStepFinish: ({ toolCalls, toolResults }) => {
          if (toolCalls.length > 0) {
            writer.write({
              id: toolCalls[0]?.toolCallId,
              type: "data-tool-calls",
              data: {
                tools: toolCalls.map((tc) => {
                  const args = "args" in tc ? tc.args : "input" in tc ? tc.input : {};
                  return {
                    toolName: tc.toolName,
                    toolCallId: tc.toolCallId,
                    args,
                  };
                }),
              },
            });
          }

          const argsByCallId = new Map(
            toolCalls.map((tc) => {
              const args = "args" in tc ? tc.args : "input" in tc ? tc.input : {};
              return [tc.toolCallId, args] as const;
            }),
          );

          for (const tr of toolResults) {
            const extension =
              GENERATE_TOOL_EXTENSIONS[tr.toolName as keyof typeof GENERATE_TOOL_EXTENSIONS];
            if (extension && tr.output?.base64) {
              const args = (argsByCallId.get(tr.toolCallId) ?? {}) as { title?: string };
              const base64 = tr.output.base64 as string;
              writer.write({
                id: tr.toolCallId,
                type: "data-document",
                data: {
                  filename: `${args.title || "generated"}.${extension}`,
                  base64,
                  mimeType: tr.output.mimeType,
                  size: Math.ceil((base64.length * 3) / 4),
                },
              });
            }
          }
        },
      });

      writer.merge(result.toUIMessageStream());
    },
    onFinish: async () => {
      await httpClient.close();
    },
  });

  return createUIMessageStreamResponse({ stream });
});
