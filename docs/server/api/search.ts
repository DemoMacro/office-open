import { createMCPClient } from "@ai-sdk/mcp";
import { createOpenAICompatible } from "@ai-sdk/openai-compatible";
import {
  streamText,
  convertToModelMessages,
  createUIMessageStream,
  createUIMessageStreamResponse,
  tool,
} from "ai";
import type { UIMessageStreamWriter, ToolSet } from "ai";
import type { H3Event } from "h3";
import { formatToolError } from "office-open/ai";
import { generate } from "office-open/generate";
import type { GenerateType } from "office-open/generate";
import { z } from "zod";

const MIME_TYPES = {
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
} as const;

async function generateDocument(type: GenerateType, options: Record<string, unknown>) {
  try {
    const base64 = (await generate({ type, options, outputType: "base64" })) as string;
    return {
      filename: `${(options.title as string) || "generated"}.${type}`,
      base64,
      mimeType: MIME_TYPES[type],
      size: Math.ceil((base64.length * 3) / 4),
    };
  } catch (error) {
    throw new Error(formatToolError(type, error));
  }
}

const generateDocumentTool = tool({
  description:
    "Generate a downloadable Office document (.docx, .pptx, or .xlsx). " +
    "Pass type and options as separate parameters. " +
    "options is a JSON object (NOT a string) containing the document structure directly. " +
    "IMPORTANT: " +
    "Section/slide children must use wrapper keys: { paragraph: {...} }, { table: {...} }, { shape: {...} }, etc. " +
    "Text runs MUST have a 'text' key: { text: '...', bold?: true }. " +
    "Colors are hex WITHOUT '#': 'FF0000', not '#FF0000'. " +
    "For docx, options must include 'sections'. For pptx, options must include 'slides'. For xlsx, options must include 'worksheets'.",
  inputSchema: z.object({
    type: z.enum(["docx", "pptx", "xlsx"]).describe("Document type to generate"),
    options: z
      .object({})
      .passthrough()
      .describe(
        "Document options object. " +
          "docx: { sections: [{ children: [...] }] }, " +
          "pptx: { title: '...', slides: [{ children: [...] }] }, " +
          "xlsx: { worksheets: [{ rows: [{ cells: [...] }] }] }",
      ),
  }),
  execute: async ({ type, options }: { type: GenerateType; options: Record<string, unknown> }) => {
    return generateDocument(type, options);
  },
});

const MAX_STEPS = 8;

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
- You have tools: list-pages (discover pages), get-page (read a page), and generate-document (create Office files)
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
- Use get-page to read the docx/pptx/xlsx documentation before calling generate-document
- Call generate-document with two parameters: type ("docx"/"pptx"/"xlsx") and options (a JSON object, NOT a string)
- The entire options object is the document definition — pass it directly as an object with proper nesting
- CRITICAL STRUCTURE RULES:
  - Section/slide children MUST use wrapper keys: { paragraph: {...} }, { table: {...} }, NOT bare objects
  - Text runs MUST have a "text" key: { text: "Hello", bold?: true }, NOT { bold: true } alone
  - Colors are hex WITHOUT "#": "FF0000", not "#FF0000"
- For docx: { type: "docx", options: { sections: [{ properties: {}, children: [{ paragraph: { children: [{ text: "..." }] } }] }] }] } }
- For pptx: { type: "pptx", options: { title: "...", slides: [{ children: [{ shape: { x: 100, y: 100, width: 600, height: 60, textBody: { text: "..." } } }] }] } }
- For xlsx: { type: "xlsx", options: { worksheets: [{ rows: [{ cells: [{ value: "Name" }] }] }] } }
- Set the "title" field in options to customize the download filename without extension (e.g. "My Report")
- Call generate-document exactly ONCE — never retry or call it multiple times for the same request
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

  let transport: Parameters<typeof createMCPClient>[0]["transport"];
  if (isExternalUrl) {
    transport = { type: "http", url: mcpServer };
  } else if (import.meta.dev) {
    transport = { type: "http", url: `http://localhost:3000${appBaseURL}${mcpServer}` };
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
        tools: { ...mcpTools, "generate-document": generateDocumentTool } as ToolSet,
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

          for (const tr of toolResults) {
            if (tr.toolName === "generate-document" && tr.output?.base64) {
              const { filename, base64, mimeType, size } = tr.output;
              writer.write({
                id: tr.toolCallId,
                type: "data-document",
                data: { filename, base64, mimeType, size },
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
