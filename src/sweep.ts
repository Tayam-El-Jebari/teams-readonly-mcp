import * as z from 'zod';
import { redactSensitiveData } from './auth.js';
import { Conversations, ReadConversationInput, ReadConversationOutput } from './conversations.js';

const MAX_OUTPUT_BYTES = 12_000;
export const SweepInput = z.object({
  since: z.iso.datetime({ offset: true }),
  targets: z.array(ReadConversationInput.shape.chat).min(1).max(20),
  response_format: ReadConversationInput.shape.response_format,
});
export const SweepOutput = z.object({
  results: z.array(z.object({ target: z.string(), ...ReadConversationOutput.shape })),
  unreachable: z.array(z.object({ target: z.string(), reason: z.string() })),
  skipped: z.array(z.string()),
  truncated: z.boolean(),
  detail: z.string(),
});

export async function sweep(conversations: Conversations, input: z.input<typeof SweepInput>) {
  const options = SweepInput.parse(input);
  const output: z.infer<typeof SweepOutput> = {
    results: [], unreachable: [], skipped: [...options.targets], truncated: true,
    detail: 'Partial sweep. Check unreachable, skipped, truncated and bodyTruncated.',
  };
  const fits = () => Buffer.byteLength(JSON.stringify(output), 'utf8') <= MAX_OUTPUT_BYTES;
  if (!fits()) throw new Error('Target names exceed the output budget. Supply fewer or shorter targets.');
  for (const target of options.targets) {
    // Remove the current target from pending work; restore it if its outcome cannot fit.
    output.skipped.shift();
    try {
      const read = await conversations.read({ chat: target, since: options.since, response_format: options.response_format });
      const result = { target, ...read, messages: [...read.messages] };
      output.results.push(result);
      while (!fits() && result.messages.length > 0) {
        result.messages.pop();
        result.truncated = true;
        result.detail = 'Sweep output limit reached. Read this chat separately or narrow since.';
      }
      if (!fits()) {
        output.results.pop();
        output.skipped.unshift(target);
        break;
      }
    } catch (error) {
      const reason = redactSensitiveData(error instanceof Error ? error.message : 'Read failed. Try this chat separately.').slice(0, 300);
      output.unreachable.push({ target, reason });
      if (!fits()) {
        output.unreachable.pop();
        output.skipped.unshift(target);
        break;
      }
    }
  }
  output.truncated = output.unreachable.length > 0 || output.skipped.length > 0 ||
    output.results.some((result) => result.truncated || result.messages.some((message) => message.bodyTruncated));
  if (!output.truncated) output.detail = 'All requested chats read without omissions.';
  return output;
}