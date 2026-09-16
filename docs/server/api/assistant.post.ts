import { createMCPClient } from "@ai-sdk/mcp";
import { createOpenAICompatible } from "@ai-sdk/openai-compatible";
import {
  convertToModelMessages,
  createUIMessageStream,
  createUIMessageStreamResponse,
  isStepCount,
  smoothStream,
  streamText,
  type ToolSet,
  type UIMessage,
  type UIMessageStreamWriter,
} from "ai";
import { createError, defineEventHandler, getRequestURL, readBody, type H3Event } from "h3";
import { useRuntimeConfig } from "nitropack/runtime";
import { officeOpenTools } from "office-open/ai";

/**
 * The docs assistant, speaking any OpenAI-compatible endpoint via ai-sdk
 * and carrying the office-open generation tools. Configure with environment
 * variables:
 *
 * - `OPENAI_COMPATIBLE_BASE_URL` — e.g. https://api.example.com/v1
 * - `OPENAI_COMPATIBLE_API_KEY`  — bearer credential for that endpoint
 * - `OPENAI_COMPATIBLE_MODEL`    — model id
 */

const GENERATE_TOOL_EXTENSIONS = {
  "generate-docx": "docx",
  "generate-pptx": "pptx",
  "generate-xlsx": "xlsx",
} as const;

// Enough headroom for schema lookups plus a validation retry before the
// final answer.
const MAX_STEPS = 20;

/** The internal MCP route is not always reachable over the network in
 * production (TLS terminated upstream, hairpin restrictions), so
 * same-origin requests ride the event instead of the wire. */
function createLocalFetch(event: H3Event): typeof fetch {
  const origin = getRequestURL(event).origin;
  return (input, init) => {
    const url =
      input instanceof URL
        ? input
        : typeof input === "string"
          ? new URL(input, origin)
          : new URL(input.url);
    return url.origin === origin
      ? event.fetch(`${url.pathname}${url.search}`, init)
      : fetch(url, init);
  };
}

/** Docus's documentation-tuned prompt, extended with the document
 * generation contract the office-open tools rely on. */
function systemPrompt(siteName: string): string {
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

function stepArgs(step: { args?: unknown; input?: unknown }): Record<string, unknown> {
  return (step.args ?? step.input ?? {}) as Record<string, unknown>;
}

export default defineEventHandler(async (event) => {
  const baseURL = process.env.OPENAI_COMPATIBLE_BASE_URL;
  const apiKey = process.env.OPENAI_COMPATIBLE_API_KEY;
  const model = process.env.OPENAI_COMPATIBLE_MODEL;
  if (!baseURL || !apiKey || !model) {
    throw createError({
      statusCode: 503,
      statusMessage:
        "Assistant not configured: set OPENAI_COMPATIBLE_BASE_URL, OPENAI_COMPATIBLE_API_KEY and OPENAI_COMPATIBLE_MODEL.",
    });
  }

  const { messages } = await readBody<{ messages: UIMessage[] }>(event);
  const provider = createOpenAICompatible({ name: "openai-compatible", baseURL, apiKey });

  // The assistant browses this site the same way an agent does: through
  // our own /mcp server (list-pages, get-page).
  const mcpClient = await createMCPClient({
    transport: {
      type: "http",
      url: `${getRequestURL(event).origin}/mcp`,
      fetch: createLocalFetch(event),
    },
  });
  const mcpTools = (await mcpClient.tools()) as ToolSet;

  // Aborting when the reader walks away — generation stops server-side.
  const abortController = new AbortController();
  event.node.req.on("close", () => abortController.abort());

  const siteName = useRuntimeConfig(event).site?.name ?? "Documentation";

  const stream = createUIMessageStream({
    execute: async ({ writer }: { writer: UIMessageStreamWriter }) => {
      const result = streamText({
        model: provider(model),
        system: systemPrompt(siteName),
        tools: { ...mcpTools, ...officeOpenTools } as ToolSet,
        // Convergence: a step ceiling with tool calls denied on the final
        // step, and an output budget — reasoning models cannot think forever.
        stopWhen: isStepCount(MAX_STEPS),
        prepareStep: ({ stepNumber }) =>
          stepNumber >= MAX_STEPS - 1 ? { toolChoice: "none" } : {},
        maxOutputTokens: 4000,
        maxRetries: 2,
        abortSignal: abortController.signal,
        // Chars arrive in readable pulses, not in awkward chunk boundaries.
        experimental_transform: smoothStream(),
        onStepFinish: ({ toolCalls, toolResults }) => {
          if (toolCalls.length > 0) {
            writer.write({
              id: toolCalls[0]?.toolCallId,
              type: "data-tool-calls",
              data: {
                tools: toolCalls.map((tc) => ({
                  toolName: tc.toolName,
                  toolCallId: tc.toolCallId,
                  args: stepArgs(tc),
                })),
              },
            });
          }

          const argsByCallId = new Map(toolCalls.map((tc) => [tc.toolCallId, stepArgs(tc)]));
          for (const result of toolResults) {
            const extension =
              GENERATE_TOOL_EXTENSIONS[result.toolName as keyof typeof GENERATE_TOOL_EXTENSIONS];
            const output = result.output as { base64?: string; mimeType?: string } | undefined;
            if (extension && output?.base64) {
              const args = argsByCallId.get(result.toolCallId) ?? {};
              const base64 = output.base64;
              writer.write({
                id: result.toolCallId,
                type: "data-document",
                data: {
                  filename: `${(args.title as string) || "generated"}.${extension}`,
                  base64,
                  mimeType: output.mimeType,
                  size: Math.ceil((base64.length * 3) / 4),
                },
              });
            }
          }
        },
        messages: await convertToModelMessages(messages ?? []),
      });

      writer.merge(result.toUIMessageStream());
    },
    onEnd: () => event.waitUntil(mcpClient.close()),
    onError: (error) => {
      console.error("[docs] assistant error:", error);
      event.waitUntil(mcpClient.close());
      return "The assistant hit an error while responding.";
    },
  });

  return createUIMessageStreamResponse({ stream });
});
